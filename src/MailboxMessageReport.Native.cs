// Mailbox Message Report - compiled helpers, loaded once by MailboxMessageReport.psm1 (Add-Type).
//
// A mailbox can hold hundreds of thousands of messages. Each page of Microsoft Graph (up to 1,000 messages) is turned
// into report rows here, without PowerShell objects, and appended at once to the part file of its folder (one JSON
// array per line). The report is then merged from the part files, in the order of the report, into the CSV files,
// the HTML report and the preview of the window. Compiled so that the tool reads as fast as Graph answers, with a
// memory that does not grow with the number of messages.
//
//   Body        the body of a Graph answer, read as bytes: a page never goes through a PowerShell string (a .NET
//               method called from PowerShell with a string of 1 MB costs about 100 ms: the argument is scanned)
//   Columns     the columns of the messages, in the order of the CSV files
//   PartWriter  the part file of one folder: the rows of each page appended to it
//   CsvTarget   a CSV file written row by row (UTF-8 with BOM, formula injection neutralised, cells within the
//               32,767 characters of an Excel cell)
//   RowBuffer   the first rows of a merge, kept as JSON for the preview of the window
//   HtmlReport  the HTML report: its template, and every message in compressed blocks of columns (as the HTML
//               report of Purview DLP Report: JSON -> gzip -> base64, repeated texts once in dictionaries)
//   Merge       a part file appended to CSV files, row buffers and HTML reports
//   Fast        dates, CSV cells
//   PreviewRow  a message in the window; BulkCollection, the list bound to the window
//
// Author : Nicolas Fabert
// Version: 2.0.0

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace MailboxMessageReportNative
{
    /// <summary>The body of a Graph answer, as received (UTF-8); its text only when asked (errors, small lists).</summary>
    public sealed class Body
    {
        string _text;
        public Body(byte[] bytes) { Bytes = bytes ?? new byte[0]; }
        public byte[] Bytes { get; private set; }
        public int Length { get { return Bytes.Length; } }
        public string Text { get { return _text ?? (_text = Encoding.UTF8.GetString(Bytes)); } }
        public override string ToString() { return Text; }

        /// <summary>The content of an HTTP answer, read at once.</summary>
        public static Body Read(System.Net.Http.HttpContent content)
        {
            return new Body(content == null ? null : content.ReadAsByteArrayAsync().GetAwaiter().GetResult());
        }
    }

    public static class Columns
    {
        /// <summary>One row per message: the columns of the CSV files, the HTML report and the part files, in this order.</summary>
        public static readonly string[] Messages = {
            "Mailbox", "MailboxName", "Location", "RecoverableItems", "FolderPath", "Folder", "Received", "Sent", "Subject",
            "From", "FromName", "Sender", "To", "Cc", "Bcc", "RecipientCount", "InternetMessageId", "HasAttachments", "Importance",
            "IsRead", "Type", "ReceivedUtc", "ItemId" };

        // The position of each column in a row (part files, CSV files): the code reads the cells by these.
        public static readonly int IMailbox = Index("Mailbox"), IMailboxName = Index("MailboxName"), ILocation = Index("Location"),
            IRecoverable = Index("RecoverableItems"), IFolderPath = Index("FolderPath"), IFolder = Index("Folder"), IReceived = Index("Received"),
            ISent = Index("Sent"), ISubject = Index("Subject"), IFrom = Index("From"), IFromName = Index("FromName"), ISender = Index("Sender"),
            ITo = Index("To"), ICc = Index("Cc"), IBcc = Index("Bcc"), IRecipientCount = Index("RecipientCount"), IMessageId = Index("InternetMessageId"),
            IAttachments = Index("HasAttachments"), IImportance = Index("Importance"), IIsRead = Index("IsRead"), IType = Index("Type"),
            IReceivedUtc = Index("ReceivedUtc"), IItemId = Index("ItemId");

        static int Index(string name)
        {
            var i = Array.IndexOf(Messages, name);
            if (i < 0) { throw new InvalidOperationException("Unknown column " + name); }
            return i;
        }

        /// <summary>The Microsoft Graph properties read for each message ($select).</summary>
        public const string Select = "id,subject,receivedDateTime,sentDateTime,from,sender,toRecipients,ccRecipients,bccRecipients,internetMessageId,hasAttachments,importance,isRead";
    }

    public static class Fast
    {
        public const string Version = "2.0.0";
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        /// <summary>A UTC date shown in a time zone: yyyy-MM-dd HH:mm (or HH:mm:ss, or yyyy-MM-dd). -PeriodEnd: 00:00 shows the day before.</summary>
        public static string FormatDate(object utc, TimeZoneInfo zone, bool dateOnly, bool periodEnd, bool seconds)
        {
            if (utc == null) { return ""; }
            var p = utc as System.Management.Automation.PSObject;
            if (p != null) { utc = p.BaseObject; }
            DateTime d;
            if (utc is DateTime) { d = (DateTime)utc; }
            else
            {
                var text = Convert.ToString(utc, Inv);
                if (string.IsNullOrEmpty(text)) { return ""; }
                d = DateTime.Parse(text, Inv, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
            }
            d = DateTime.SpecifyKind(d.Kind == DateTimeKind.Local ? d.ToUniversalTime() : d, DateTimeKind.Utc);
            var local = TimeZoneInfo.ConvertTimeFromUtc(d, zone ?? TimeZoneInfo.Local);
            if (periodEnd && local.TimeOfDay == TimeSpan.Zero) { local = local.AddDays(-1); dateOnly = true; }
            return local.ToString(dateOnly ? "yyyy-MM-dd" : (seconds ? "yyyy-MM-dd HH:mm:ss" : "yyyy-MM-dd HH:mm"), Inv);
        }

        /// <summary>The most characters of a cell of Excel; a longer cell breaks the rows of a CSV file opened in Excel.</summary>
        public const int ExcelCellMax = 32767;

        /// <summary>
        /// A text within an Excel cell (32,000 characters: room for the apostrophe and the quotes). A list (To, Cc, Bcc:
        /// "a; b; c") is cut after its last whole address, with the number of the others: "a; b; ... (+9,800 more)".
        /// </summary>
        public static string ExcelText(string text)
        {
            const int max = 32000;
            if (text == null || text.Length <= max) { return text; }
            var cut = text.LastIndexOf("; ", max - 40, StringComparison.Ordinal);
            if (cut > 0)
            {
                long more = 0;
                for (var i = cut; (i = text.IndexOf("; ", i, StringComparison.Ordinal)) >= 0; i += 2) { more++; }
                return text.Substring(0, cut) + "; \u2026 (+" + more.ToString("N0", CultureInfo.InvariantCulture) + " more)";
            }
            return text.Substring(0, max - 20) + " \u2026 (cut)";
        }

        /// <summary>A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe; quoted when needed; within an Excel cell.</summary>
        public static string CsvCell(string text, string delimiter)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            text = ExcelText(text);
            if ("=+-@\t\r".IndexOf(text[0]) >= 0) { text = "'" + text; }
            if (text.Contains(delimiter) || text.Contains("\"") || text.IndexOf('\r') >= 0 || text.IndexOf('\n') >= 0) { text = "\"" + text.Replace("\"", "\"\"") + "\""; }
            return text;
        }

        /// <summary>JSON options of the part files and the HTML report: every language as is, the characters of HTML escaped (no &lt; &gt; &amp; ' ").</summary>
        public static readonly JavaScriptEncoder Encoder = JavaScriptEncoder.Create(UnicodeRanges.All);

        /// <summary>The cells of a row of a part file (a JSON array of strings).</summary>
        public static string[] ParseRow(string line)
        {
            using (var doc = JsonDocument.Parse(line))
            {
                var root = doc.RootElement;
                var cells = new string[root.GetArrayLength()];
                int i = 0;
                foreach (var e in root.EnumerateArray()) { cells[i++] = e.ValueKind == JsonValueKind.String ? e.GetString() : (e.ValueKind == JsonValueKind.Null ? "" : e.GetRawText()); }
                return cells;
            }
        }
    }

    /// <summary>The rows of one page of messages and the link to the next page.</summary>
    public sealed class PageResult
    {
        public int Rows { get; set; }
        public string NextLink { get; set; }
    }

    /// <summary>
    /// The part file of one folder: each page of messages is appended to it, one JSON array per message and per line, in
    /// the order of Graph (received date, newest first). The fixed cells (mailbox, location, folder) are the same on
    /// every row of the file.
    /// </summary>
    public sealed class PartWriter : IDisposable
    {
        readonly FileStream _stream;
        // The rows of a page are written to memory, then to the file at once.
        readonly MemoryStream _buffer = new MemoryStream(1 << 20);
        readonly Utf8JsonWriter _json;
        readonly string[] _fixed;
        readonly TimeZoneInfo _zone;
        // The offset of the time zone for a quarter of an hour (the messages of a page are in date order).
        DateTime _offsetFrom = DateTime.MaxValue;
        TimeSpan _offset;
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public PartWriter(string path, string mailbox, string mailboxName, string location, bool recoverable, string folderPath, string folder, TimeZoneInfo zone)
        {
            Path = path;
            _stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Read, 65536);
            _json = new Utf8JsonWriter(_buffer, new JsonWriterOptions { Encoder = Fast.Encoder, SkipValidation = true });
            _fixed = new[] { mailbox ?? "", mailboxName ?? "", location ?? "", recoverable ? "Yes" : "No", folderPath ?? "", folder ?? "" };
            _zone = zone ?? TimeZoneInfo.Local;
        }

        public string Path { get; private set; }
        public long Count { get; private set; }

        /// <summary>
        /// A page answered by Graph (a Body, or JSON text): its messages appended as rows; returns how many and the
        /// nextLink ("" on the last page).
        /// </summary>
        public PageResult AddPage(object content)
        {
            var result = new PageResult { Rows = 0, NextLink = "" };
            var ps = content as System.Management.Automation.PSObject;
            if (ps != null) { content = ps.BaseObject; }
            var body = content as Body;
            string text = body == null ? Convert.ToString(content, Inv) : null;
            if (body == null && string.IsNullOrEmpty(text)) { return result; }
            if (body != null && body.Length == 0) { return result; }
            using (var doc = body != null ? JsonDocument.Parse(new ReadOnlyMemory<byte>(body.Bytes)) : JsonDocument.Parse(text))
            {
                var root = doc.RootElement;
                JsonElement value;
                if (root.TryGetProperty("value", out value) && value.ValueKind == JsonValueKind.Array)
                {
                    foreach (var m in value.EnumerateArray())
                    {
                        WriteRow(m);
                        result.Rows++;
                    }
                }
                JsonElement next;
                if (root.TryGetProperty("@odata.nextLink", out next) && next.ValueKind == JsonValueKind.String) { result.NextLink = next.GetString(); }
            }
            _buffer.WriteTo(_stream);
            _buffer.SetLength(0);
            _stream.Flush();
            Count += result.Rows;
            return result;
        }

        void WriteRow(JsonElement m)
        {
            _json.Reset();
            _json.WriteStartArray();
            foreach (var f in _fixed) { _json.WriteStringValue(f); }
            var received = Str(m, "receivedDateTime");
            var receivedUtc = ParseUtc(received);
            _json.WriteStringValue(receivedUtc.HasValue ? Local(receivedUtc.Value) : "");
            var sent = ParseUtc(Str(m, "sentDateTime"));
            _json.WriteStringValue(sent.HasValue ? Local(sent.Value) : "");
            _json.WriteStringValue(Str(m, "subject"));
            _json.WriteStringValue(Address(m, "from", "address"));
            _json.WriteStringValue(Address(m, "from", "name"));
            _json.WriteStringValue(Address(m, "sender", "address"));
            int to, cc, bcc;
            _json.WriteStringValue(Recipients(m, "toRecipients", out to));
            _json.WriteStringValue(Recipients(m, "ccRecipients", out cc));
            _json.WriteStringValue(Recipients(m, "bccRecipients", out bcc));
            // Empty when the recipients were not asked (-SkipRecipients), not 0.
            JsonElement any;
            var asked = m.TryGetProperty("toRecipients", out any) || m.TryGetProperty("ccRecipients", out any) || m.TryGetProperty("bccRecipients", out any);
            _json.WriteStringValue(asked ? (to + cc + bcc).ToString(Inv) : "");
            _json.WriteStringValue(Str(m, "internetMessageId"));
            _json.WriteStringValue(Bool(m, "hasAttachments"));
            _json.WriteStringValue(Title(Str(m, "importance")));
            _json.WriteStringValue(Bool(m, "isRead"));
            _json.WriteStringValue(Kind(Str(m, "@odata.type")));
            _json.WriteStringValue(receivedUtc.HasValue ? receivedUtc.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", Inv) : "");
            _json.WriteStringValue(Str(m, "id"));
            _json.WriteEndArray();
            _json.Flush();
            _buffer.WriteByte((byte)'\n');
        }

        string Local(DateTime utc)
        {
            if (utc < _offsetFrom || utc >= _offsetFrom.AddMinutes(15))
            {
                _offsetFrom = new DateTime(utc.Ticks - utc.Ticks % (15 * TimeSpan.TicksPerMinute), DateTimeKind.Utc);
                _offset = _zone.GetUtcOffset(_offsetFrom);
            }
            return (utc + _offset).ToString("yyyy-MM-dd HH:mm:ss", Inv);
        }

        static DateTime? ParseUtc(string text)
        {
            if (string.IsNullOrEmpty(text)) { return null; }
            DateTimeOffset d;
            if (DateTimeOffset.TryParse(text, Inv, DateTimeStyles.AssumeUniversal, out d)) { return d.UtcDateTime; }
            return null;
        }

        static string Str(JsonElement o, string name)
        {
            JsonElement v;
            if (o.ValueKind != JsonValueKind.Object || !o.TryGetProperty(name, out v)) { return ""; }
            switch (v.ValueKind)
            {
                case JsonValueKind.String: return v.GetString() ?? "";
                case JsonValueKind.Null:
                case JsonValueKind.Undefined: return "";
                default: return v.GetRawText();
            }
        }

        static string Bool(JsonElement o, string name)
        {
            JsonElement v;
            if (o.TryGetProperty(name, out v))
            {
                if (v.ValueKind == JsonValueKind.True) { return "Yes"; }
                if (v.ValueKind == JsonValueKind.False) { return "No"; }
            }
            return "";
        }

        static string Title(string text)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
            return char.ToUpperInvariant(text[0]) + text.Substring(1);
        }

        /// <summary>The kind of item, from its @odata.type (absent for a plain message).</summary>
        static string Kind(string type)
        {
            switch (type)
            {
                case "": return "Message";
                case "#microsoft.graph.message": return "Message";
                case "#microsoft.graph.eventMessageRequest": return "Meeting request";
                case "#microsoft.graph.eventMessageResponse": return "Meeting response";
                case "#microsoft.graph.eventMessage": return "Meeting message";
                case "#microsoft.graph.calendarSharingMessage": return "Sharing invitation";
                default:
                    var i = type.LastIndexOf('.');
                    return i >= 0 ? type.Substring(i + 1) : type;
            }
        }

        static string Address(JsonElement o, string property, string part)
        {
            JsonElement r, e, v;
            if (o.TryGetProperty(property, out r) && r.ValueKind == JsonValueKind.Object && r.TryGetProperty("emailAddress", out e) && e.ValueKind == JsonValueKind.Object && e.TryGetProperty(part, out v) && v.ValueKind == JsonValueKind.String)
            {
                return v.GetString() ?? "";
            }
            return "";
        }

        /// <summary>The addresses of a recipient list, separated by "; " (the name when a recipient has no address).</summary>
        static string Recipients(JsonElement o, string property, out int count)
        {
            JsonElement list;
            count = 0;
            if (!o.TryGetProperty(property, out list) || list.ValueKind != JsonValueKind.Array) { return ""; }
            var sb = new StringBuilder();
            foreach (var r in list.EnumerateArray())
            {
                JsonElement e, v;
                string text = "";
                if (r.ValueKind == JsonValueKind.Object && r.TryGetProperty("emailAddress", out e) && e.ValueKind == JsonValueKind.Object)
                {
                    if (e.TryGetProperty("address", out v) && v.ValueKind == JsonValueKind.String) { text = v.GetString() ?? ""; }
                    if (text.Length == 0 && e.TryGetProperty("name", out v) && v.ValueKind == JsonValueKind.String) { text = v.GetString() ?? ""; }
                }
                if (text.Length == 0) { continue; }
                if (sb.Length > 0) { sb.Append("; "); }
                sb.Append(text);
                count++;
            }
            return sb.ToString();
        }

        public void Dispose()
        {
            _json.Dispose();
            _buffer.Dispose();
            _stream.Dispose();
        }
    }

    /// <summary>A CSV file written row by row: UTF-8 with BOM, the header first.</summary>
    public sealed class CsvTarget : IDisposable
    {
        readonly StreamWriter _writer;
        readonly string _delimiter;
        readonly string[] _cells;

        public CsvTarget(string path, string[] columns, string delimiter)
        {
            Path = path;
            _delimiter = delimiter;
            _writer = new StreamWriter(path, false, new UTF8Encoding(true), 65536);
            _cells = new string[columns.Length];
            Write(columns);
        }

        public string Path { get; private set; }
        public long Rows { get; private set; }

        public void Write(string[] cells)
        {
            for (int i = 0; i < _cells.Length; i++) { _cells[i] = Fast.CsvCell(i < cells.Length ? cells[i] : "", _delimiter); }
            _writer.Write(string.Join(_delimiter, _cells));
            _writer.Write("\r\n");
            Rows++;
        }

        public void Dispose() { _writer.Dispose(); }
    }

    /// <summary>The first rows of a merge (JSON arrays as written in the part files), for the HTML report or the window.</summary>
    public sealed class RowBuffer
    {
        public RowBuffer(int max) { Max = Math.Max(0, max); Lines = new List<string>(Math.Min(Max, 100000)); }
        public int Max { get; private set; }
        public List<string> Lines { get; private set; }
        public long Seen { get; private set; }
        public bool Truncated { get { return Seen > Lines.Count; } }

        public void Add(string line)
        {
            Seen++;
            if (Lines.Count < Max) { Lines.Add(line); }
        }

        /// <summary>The rows kept as one JSON array (safe inside a script block of the HTML report).</summary>
        public string ToJson() { return "[" + string.Join(",", Lines) + "]"; }

        /// <summary>The rows kept by another buffer (the sample of one folder), as long as there is room.</summary>
        public void AddFrom(RowBuffer other)
        {
            if (other == null) { return; }
            foreach (var line in other.Lines) { if (Lines.Count >= Max) { break; } Lines.Add(line); }
            Seen += other.Seen;
        }
    }

    /// <summary>
    /// The HTML report of a run or of a mailbox, written as it goes: its template up to {{MESSAGES}}, then the messages
    /// in blocks of columns (JSON, gzip, base64: script blocks the page decompresses), then the rest of the template.
    /// The texts repeated from row to row (folders, subjects, addresses, names, kinds) are stored once, in dictionaries
    /// that each block extends: a report of hundreds of thousands of messages stays small and opens in seconds. The same
    /// design as the HTML report of Purview DLP Report.
    /// Markers ({{NAME}}): given when the report is opened (those before the messages) or closed (those after); every
    /// marker of the template must have a value, and every value a marker.
    /// </summary>
    public sealed class HtmlReport : IDisposable
    {
        public const string MessagesMarker = "{{MESSAGES}}";
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly DateTime Epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        static readonly JsonWriterOptions JsonOptions = new JsonWriterOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, SkipValidation = true };
        static readonly string[] ListSeparator = { "; " };
        readonly StreamWriter _w;
        readonly string _suffix;
        readonly int _blockRows;
        readonly Dictionary<string, string> _values = new Dictionary<string, string>(StringComparer.Ordinal);
        readonly HashSet<string> _used = new HashSet<string>(StringComparer.Ordinal);
        bool _closed;
        // Dictionaries of the file, and their entries not yet written (the next block carries them).
        readonly Dictionary<string, int> _folders = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly List<string[]> _newFolders = new List<string[]>();
        readonly Dictionary<string, int> _subjects = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly List<string> _newSubjects = new List<string>();
        readonly Dictionary<string, int> _people = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        readonly List<string> _newPeople = new List<string>();
        readonly Dictionary<string, int> _names = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly List<string> _newNames = new List<string>();
        readonly Dictionary<string, int> _kinds = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly List<string> _newKinds = new List<string>();
        // The columns of the block in course (one entry per message; ri: the recipients of every message, in a row).
        readonly List<int> _f = new List<int>(), _o = new List<int>(), _j = new List<int>(), _p = new List<int>(), _n = new List<int>(), _s = new List<int>();
        readonly List<int> _tl = new List<int>(), _cl = new List<int>(), _bl = new List<int>(), _ri = new List<int>(), _x = new List<int>(), _k = new List<int>();
        readonly List<long?> _r = new List<long?>(), _d = new List<long?>();
        readonly List<string> _m = new List<string>();

        public HtmlReport(string templatePath, string outputPath, int maxRows, int blockRows, string[] markers, string[] values)
        {
            Path = outputPath;
            Max = Math.Max(0, maxRows);
            _blockRows = Math.Max(1, blockRows);
            var template = File.ReadAllText(templatePath, Encoding.UTF8);
            var at = template.IndexOf(MessagesMarker, StringComparison.Ordinal);
            if (at < 0 || at != template.LastIndexOf(MessagesMarker, StringComparison.Ordinal)) { throw new InvalidDataException("The report template must hold " + MessagesMarker + " exactly once: " + templatePath); }
            Set(markers, values);
            _suffix = template.Substring(at + MessagesMarker.Length);
            var prefix = Replace(template.Substring(0, at));
            _w = new StreamWriter(outputPath, false, new UTF8Encoding(true), 1 << 16);
            _w.Write(prefix);
        }

        public string Path { get; private set; }
        /// <summary>The most messages of this file (Report.HtmlMaxMessages).</summary>
        public int Max { get; private set; }
        /// <summary>The messages written; Seen: the messages given (more than Rows when Max was reached).</summary>
        public long Rows { get; private set; }
        public long Seen { get; private set; }
        public int Blocks { get; private set; }

        void Set(string[] markers, string[] values)
        {
            if (markers == null) { return; }
            for (int i = 0; i < markers.Length; i++) { _values[markers[i]] = values != null && i < values.Length && values[i] != null ? values[i] : ""; }
        }

        // The markers are found in the template first: a value that holds the text of a marker (a subject) is never touched.
        string Replace(string text)
        {
            var sb = new StringBuilder(text.Length + 1024);
            int last = 0;
            foreach (System.Text.RegularExpressions.Match m in System.Text.RegularExpressions.Regex.Matches(text, @"\{\{[A-Z_]+\}\}"))
            {
                string v;
                if (!_values.TryGetValue(m.Value, out v)) { throw new InvalidOperationException("Report template marker not replaced: " + m.Value); }
                if (!_used.Add(m.Value)) { throw new InvalidOperationException("Report template marker found twice: " + m.Value); }
                sb.Append(text, last, m.Index - last).Append(v);
                last = m.Index + m.Length;
            }
            sb.Append(text, last, text.Length - last);
            return sb.ToString();
        }

        static int Index(Dictionary<string, int> map, List<string> added, string value)
        {
            value = value ?? "";
            int i;
            if (!map.TryGetValue(value, out i)) { i = map.Count; map.Add(value, i); added.Add(value); }
            return i;
        }

        /// <summary>A local date and time of a row (yyyy-MM-dd HH:mm:ss, or the UTC one with T and Z) as seconds: the page shows them as they are.</summary>
        static long? Seconds(string text)
        {
            DateTime d;
            if (string.IsNullOrEmpty(text)) { return null; }
            if (!DateTime.TryParseExact(text, new[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ" }, Inv, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out d)) { return null; }
            return (d - Epoch).Ticks / TimeSpan.TicksPerSecond;
        }

        int AddList(string list)
        {
            if (string.IsNullOrEmpty(list)) { return 0; }
            int n = 0;
            foreach (var a in list.Split(ListSeparator, StringSplitOptions.RemoveEmptyEntries)) { _ri.Add(Index(_people, _newPeople, a)); n++; }
            return n;
        }

        /// <summary>One message (the cells of a part file, in the order of Columns.Messages).</summary>
        public void Add(string[] c)
        {
            Seen++;
            if (Rows >= Max) { return; }
            Rows++;
            var recoverable = c[Columns.IRecoverable] == "Yes";
            var key = c[Columns.IMailbox] + "\u0001" + c[Columns.ILocation] + "\u0001" + (recoverable ? "1" : "0") + "\u0001" + c[Columns.IFolderPath];
            int f;
            if (!_folders.TryGetValue(key, out f))
            {
                f = _folders.Count;
                _folders.Add(key, f);
                _newFolders.Add(new[] { c[Columns.IMailbox], c[Columns.IMailboxName], c[Columns.ILocation], recoverable ? "1" : "0", c[Columns.IFolderPath], c[Columns.IFolder] });
            }
            _f.Add(f);
            var received = Seconds(c[Columns.IReceived]);
            var utc = Seconds(c[Columns.IReceivedUtc]);
            _r.Add(received);
            _o.Add(received.HasValue && utc.HasValue ? (int)((received.Value - utc.Value) / 60) : 0);
            var sent = Seconds(c[Columns.ISent]);
            _d.Add(received.HasValue && sent.HasValue ? received.Value - sent.Value : (long?)null);
            _j.Add(Index(_subjects, _newSubjects, c[Columns.ISubject]));
            var from = c[Columns.IFrom];
            _p.Add(from.Length == 0 ? -1 : Index(_people, _newPeople, from));
            var name = c[Columns.IFromName];
            _n.Add(name.Length == 0 || string.Equals(name, from, StringComparison.OrdinalIgnoreCase) ? -1 : Index(_names, _newNames, name));
            var sender = c[Columns.ISender];
            _s.Add(sender.Length == 0 || string.Equals(sender, from, StringComparison.OrdinalIgnoreCase) ? -1 : Index(_people, _newPeople, sender));
            _tl.Add(AddList(c[Columns.ITo]));
            _cl.Add(AddList(c[Columns.ICc]));
            _bl.Add(AddList(c[Columns.IBcc]));
            _m.Add(c[Columns.IMessageId]);
            int x = 0;
            if (c[Columns.IAttachments] == "Yes") { x |= 1; }
            if (c[Columns.IIsRead] == "Yes") { x |= 2; }
            if (c[Columns.IImportance] == "High") { x |= 4; } else if (c[Columns.IImportance] == "Low") { x |= 8; }
            if (c[Columns.IRecipientCount].Length == 0) { x |= 16; }
            _x.Add(x);
            _k.Add(Index(_kinds, _newKinds, c[Columns.IType]));
            if (_f.Count >= _blockRows) { Flush(); }
        }

        static void Strings(Utf8JsonWriter j, string name, List<string> list) { j.WriteStartArray(name); foreach (var v in list) { j.WriteStringValue(v); } j.WriteEndArray(); }
        static void Ints(Utf8JsonWriter j, string name, List<int> list) { j.WriteStartArray(name); foreach (var v in list) { j.WriteNumberValue(v); } j.WriteEndArray(); }
        static void Longs(Utf8JsonWriter j, string name, List<long?> list) { j.WriteStartArray(name); foreach (var v in list) { if (v.HasValue) { j.WriteNumberValue(v.Value); } else { j.WriteNullValue(); } } j.WriteEndArray(); }

        /// <summary>The block in course written: its new dictionary entries and its columns.</summary>
        void Flush()
        {
            if (_f.Count == 0) { return; }
            using (var raw = new MemoryStream())
            {
                using (var gzip = new GZipStream(raw, CompressionLevel.Optimal, true))
                using (var j = new Utf8JsonWriter(gzip, JsonOptions))
                {
                    j.WriteStartObject();
                    j.WriteStartArray("df");
                    foreach (var folder in _newFolders) { j.WriteStartArray(); foreach (var v in folder) { j.WriteStringValue(v); } j.WriteEndArray(); }
                    j.WriteEndArray();
                    Strings(j, "dj", _newSubjects); Strings(j, "dp", _newPeople); Strings(j, "dn", _newNames); Strings(j, "dk", _newKinds);
                    Ints(j, "f", _f); Longs(j, "r", _r); Ints(j, "o", _o); Longs(j, "d", _d); Ints(j, "j", _j); Ints(j, "p", _p); Ints(j, "n", _n); Ints(j, "s", _s);
                    Ints(j, "tl", _tl); Ints(j, "cl", _cl); Ints(j, "bl", _bl); Ints(j, "ri", _ri); Strings(j, "m", _m); Ints(j, "x", _x); Ints(j, "k", _k);
                    j.WriteEndObject();
                }
                _w.Write("<script type=\"application/x-mmr-block\">");
                _w.Write(Convert.ToBase64String(raw.GetBuffer(), 0, (int)raw.Length));
                _w.Write("</script>\n");
            }
            Blocks++;
            _newFolders.Clear(); _newSubjects.Clear(); _newPeople.Clear(); _newNames.Clear(); _newKinds.Clear();
            foreach (var l in new[] { _f, _o, _j, _p, _n, _s, _tl, _cl, _bl, _ri, _x, _k }) { l.Clear(); }
            _r.Clear(); _d.Clear(); _m.Clear();
        }

        /// <summary>The last block, then the rest of the template with the markers given now (summary, mailboxes, folders).</summary>
        public void Close(string[] markers, string[] values)
        {
            if (_closed) { return; }
            Flush();
            Set(markers, values);
            _w.Write(Replace(_suffix));
            foreach (var k in _values.Keys) { if (!_used.Contains(k)) { throw new InvalidOperationException("Report template marker missing: " + k); } }
            _w.Dispose();
            _closed = true;
        }

        public void Dispose()
        {
            if (_closed) { return; }
            _w.Dispose();
            _closed = true;
        }
    }

    public static class Merge
    {
        /// <summary>The rows of a part file appended to CSV files, row buffers and HTML reports, in the order of the file; returns how many rows.</summary>
        public static long AppendPart(string path, CsvTarget[] csv, RowBuffer[] buffers, HtmlReport[] html)
        {
            long n = 0;
            if (!File.Exists(path)) { return 0; }
            bool toCsv = csv != null && csv.Length > 0;
            bool toHtml = html != null && html.Length > 0;
            foreach (var line in File.ReadLines(path, Encoding.UTF8))
            {
                if (line.Length == 0) { continue; }
                n++;
                if (toCsv || toHtml)
                {
                    var cells = Fast.ParseRow(line);
                    if (toCsv) { foreach (var t in csv) { if (t != null) { t.Write(cells); } } }
                    if (toHtml) { foreach (var h in html) { if (h != null) { h.Add(cells); } } }
                }
                if (buffers != null) { foreach (var b in buffers) { if (b != null) { b.Add(line); } } }
            }
            return n;
        }

        public static long AppendPart(string path, CsvTarget[] csv, RowBuffer[] buffers) { return AppendPart(path, csv, buffers, null); }
    }
    /// <summary>A message in the preview of the window (the flat list, the folder view and its reading pane).</summary>
    public sealed class PreviewRow : INotifyPropertyChanged
    {
        string _body = "";
        public string Received { get; set; }
        public string ReceivedShort { get; set; }
        public string Sent { get; set; }
        public string Mailbox { get; set; }
        public string MailboxName { get; set; }
        public string Location { get; set; }
        public string LocationText { get; set; }
        public bool RecoverableItems { get; set; }
        public string FolderPath { get; set; }
        public string Folder { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public string FromDisplay { get; set; }
        public string Sender { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public string Bcc { get; set; }
        public string InternetMessageId { get; set; }
        public bool HasAttachments { get; set; }
        public string AttachmentGlyph { get; set; }
        public string Importance { get; set; }
        public string Type { get; set; }
        public string ItemId { get; set; }
        /// <summary>The folder of the message in the tree: mailbox|location|recoverable|path.</summary>
        public string FolderKey { get; set; }
        /// <summary>The content of the message once read (reading pane), or why it is not.</summary>
        public string Body { get { return _body; } set { if (_body != value) { _body = value ?? ""; Notify("Body"); } } }
        public bool BodyLoaded { get; set; }
        public event PropertyChangedEventHandler PropertyChanged;
        void Notify(string name) { var h = PropertyChanged; if (h != null) { h(this, new PropertyChangedEventArgs(name)); } }

        public static string Key(string mailbox, string location, bool recoverable, string path)
        {
            return (mailbox ?? "").ToLowerInvariant() + "|" + location + "|" + (recoverable ? "1" : "0") + "|" + (path ?? "").ToLowerInvariant();
        }

        /// <summary>One row from the cells of a part file.</summary>
        public static PreviewRow FromCells(string[] c)
        {
            var recoverable = c[Columns.IRecoverable] == "Yes";
            var location = c[Columns.ILocation];
            var from = c[Columns.IFrom];
            var name = c[Columns.IFromName];
            var received = c[Columns.IReceived];
            var fromName = name.Length > 0 ? name : from;
            return new PreviewRow
            {
                Mailbox = c[Columns.IMailbox], MailboxName = c[Columns.IMailboxName], Location = location + (recoverable ? " (RI)" : ""), RecoverableItems = recoverable,
                LocationText = (location == "Archive" ? "Archive" : "Primary mailbox") + (recoverable ? " - Recoverable Items" : ""),
                FolderPath = c[Columns.IFolderPath], Folder = c[Columns.IFolder], Received = received, ReceivedShort = received.Length >= 16 ? received.Substring(0, 16) : received,
                Sent = c[Columns.ISent], Subject = c[Columns.ISubject].Length > 0 ? c[Columns.ISubject] : "(no subject)",
                From = from.Length > 0 ? from : name, FromDisplay = fromName.Length > 0 ? fromName : "(no sender)", Sender = c[Columns.ISender],
                To = c[Columns.ITo], Cc = c[Columns.ICc], Bcc = c[Columns.IBcc], InternetMessageId = c[Columns.IMessageId],
                HasAttachments = c[Columns.IAttachments] == "Yes", AttachmentGlyph = c[Columns.IAttachments] == "Yes" ? "\uE723" : "",
                Importance = c[Columns.IImportance], Type = c[Columns.IType], ItemId = c[Columns.IItemId],
                FolderKey = Key(c[Columns.IMailbox], location, recoverable, c[Columns.IFolderPath])
            };
        }

        /// <summary>The rows of a buffer, for the lists of the window.</summary>
        public static List<object> Build(RowBuffer buffer)
        {
            var rows = new List<object>();
            if (buffer == null) { return rows; }
            foreach (var line in buffer.Lines) { rows.Add(FromCells(Fast.ParseRow(line))); }
            return rows;
        }
    }

    /// <summary>A node of the folder tree of the window: a mailbox, a location (primary mailbox, archive, Recoverable Items) or a folder.</summary>
    public sealed class FolderNode : INotifyPropertyChanged
    {
        bool _expanded;
        bool _selected;
        public FolderNode() { Children = new ObservableCollection<FolderNode>(); Messages = new List<object>(); }
        public string Name { get; set; }
        public string Glyph { get; set; }
        /// <summary>Mailbox | Location | Folder.</summary>
        public string Kind { get; set; }
        public string Key { get; set; }
        public long Found { get; set; }
        public long Items { get; set; }
        public string CountText { get; set; }
        public string Status { get; set; }
        public string Tip { get; set; }
        public bool Dim { get; set; }
        public ObservableCollection<FolderNode> Children { get; private set; }
        /// <summary>The messages of the preview in this folder (a sample: Window.PreviewPerFolder at most).</summary>
        public List<object> Messages { get; private set; }
        public bool IsExpanded { get { return _expanded; } set { if (_expanded != value) { _expanded = value; Notify("IsExpanded"); } } }
        public bool IsSelected { get { return _selected; } set { if (_selected != value) { _selected = value; Notify("IsSelected"); } } }
        public event PropertyChangedEventHandler PropertyChanged;
        void Notify(string name) { var h = PropertyChanged; if (h != null) { h(this, new PropertyChangedEventArgs(name)); } }

        static string Text(object o, string name)
        {
            if (o == null) { return ""; }
            var p = System.Management.Automation.PSObject.AsPSObject(o).Properties[name];
            return p == null || p.Value == null ? "" : Convert.ToString(p.Value, CultureInfo.InvariantCulture);
        }

        static long Number(object o, string name)
        {
            long n;
            return long.TryParse(Text(o, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : 0;
        }

        static string Count(long found) { return found > 0 ? found.ToString("N0", CultureInfo.InvariantCulture) : ""; }

        static string FolderGlyph(string path)
        {
            switch ((path ?? "").ToLowerInvariant())
            {
                case "\\inbox": return "\uE715";
                case "\\sent items": return "\uE724";
                case "\\drafts": return "\uE70F";
                case "\\deleted items": return "\uE74D";
                case "\\junk email": return "\uE7BA";
                case "\\archive": return "\uE7B8";
                default: return "\uE8B7";
            }
        }

        /// <summary>
        /// The tree of the window: one node per mailbox, then its primary mailbox and its archive, then their folders by
        /// path (Recoverable Items in a node of its own); the messages of the preview attached to their folder.
        /// </summary>
        public static List<object> Build(object mailboxes, object folders, IEnumerable preview)
        {
            var roots = new List<object>();
            var byMailbox = new Dictionary<string, FolderNode>(StringComparer.OrdinalIgnoreCase);
            var nodes = new Dictionary<string, FolderNode>(StringComparer.OrdinalIgnoreCase);
            foreach (var m in Each(mailboxes))
            {
                var address = Text(m, "Address");
                var name = Text(m, "DisplayName");
                var state = Text(m, "State");
                var node = new FolderNode { Name = name.Length > 0 ? name : address, Glyph = "\uE77B", Kind = "Mailbox", Key = address.ToLowerInvariant(), Found = Number(m, "Messages"), Tip = address + (state != "Ok" ? " - " + Text(m, "Detail") : ""), Dim = state != "Ok", IsExpanded = true };
                node.CountText = Count(node.Found);
                byMailbox[address] = node;
                roots.Add(node);
            }
            Func<string, string, bool, FolderNode> location = (address, loc, recoverable) =>
            {
                FolderNode mb;
                if (!byMailbox.TryGetValue(address, out mb)) { return null; }
                var key = address.ToLowerInvariant() + "|" + loc + "|" + (recoverable ? "1" : "0");
                FolderNode node;
                if (nodes.TryGetValue(key, out node)) { return node; }
                if (recoverable)
                {
                    var parent = nodes.ContainsKey(address.ToLowerInvariant() + "|" + loc + "|0") ? nodes[address.ToLowerInvariant() + "|" + loc + "|0"] : null;
                    if (parent == null) { return null; }
                    node = new FolderNode { Name = "Recoverable Items", Glyph = "\uE74D", Kind = "Location", Key = key };
                    parent.Children.Add(node);
                }
                else
                {
                    node = new FolderNode { Name = loc == "Archive" ? "Archive" : "Primary mailbox", Glyph = loc == "Archive" ? "\uE7B8" : "\uE715", Kind = "Location", Key = key, IsExpanded = true };
                    mb.Children.Add(node);
                }
                nodes[key] = node;
                return node;
            };
            // Primary and archive before Recoverable Items (which hang under them), whatever the order of the folders.
            var list = Each(folders).ToList();
            foreach (var f in list) { if (Text(f, "RecoverableItems") != "True") { location(Text(f, "Mailbox"), Text(f, "Location"), false); } }
            foreach (var f in list)
            {
                var address = Text(f, "Mailbox");
                var loc = Text(f, "Location");
                var recoverable = Text(f, "RecoverableItems") == "True";
                if (recoverable && location(address, loc, false) == null) { continue; }
                var parent = location(address, loc, recoverable);
                if (parent == null) { continue; }
                var path = Text(f, "Path");
                var parts = path.Trim('\\').Split('\\');
                var start = recoverable && parts.Length > 0 && parts[0] == "Recoverable Items" ? 1 : 0;
                var prefix = recoverable ? "\\Recoverable Items" : "";
                for (int i = start; i < parts.Length; i++)
                {
                    prefix += "\\" + parts[i];
                    var key = PreviewRow.Key(address, loc, recoverable, prefix);
                    FolderNode node;
                    if (!nodes.TryGetValue(key, out node))
                    {
                        node = new FolderNode { Name = parts[i], Glyph = recoverable ? "\uE8B7" : FolderGlyph(prefix), Kind = "Folder", Key = key, Dim = true, Tip = prefix };
                        parent.Children.Add(node);
                        nodes[key] = node;
                    }
                    parent = node;
                }
                var leaf = parent;
                leaf.Found = Number(f, "Messages");
                leaf.Items = Number(f, "TotalItems");
                leaf.Status = Text(f, "Status");
                leaf.CountText = Count(leaf.Found);
                leaf.Dim = leaf.Found == 0;
                var detail = Text(f, "Detail");
                leaf.Tip = path + "  -  " + leaf.Items.ToString("N0", CultureInfo.InvariantCulture) + " item(s), " + leaf.Found.ToString("N0", CultureInfo.InvariantCulture) + " message(s) found" + (leaf.Status.Length > 0 ? ", " + leaf.Status : "") + (detail.Length > 0 ? " (" + detail + ")" : "");
            }
            // Messages found per location, and the messages of the preview in their folder.
            foreach (var node in nodes.Values) { if (node.Kind == "Location") { node.Found = Sum(node); node.CountText = Count(node.Found); } }
            if (preview != null)
            {
                foreach (var o in preview)
                {
                    var r = o as PreviewRow;
                    FolderNode node;
                    if (r != null && nodes.TryGetValue(r.FolderKey, out node)) { node.Messages.Add(r); }
                }
            }
            return roots;
        }

        static long Sum(FolderNode n)
        {
            long s = 0;
            foreach (var c in n.Children) { s += (c.Kind == "Folder" ? c.Found : 0) + Sum(c); }
            return s;
        }

        static IEnumerable<object> Each(object o)
        {
            if (o == null) { yield break; }
            var p = o as System.Management.Automation.PSObject;
            var b = p != null ? p.BaseObject : o;
            var e = b as IEnumerable;
            if (e == null || b is string) { yield return o; yield break; }
            foreach (var x in e) { if (x != null) { yield return x; } }
        }
    }
    /// <summary>A mailbox in the summary list of the window.</summary>
    public sealed class MailboxRow
    {
        public string Mailbox { get; set; }
        public string Name { get; set; }
        public string Archive { get; set; }
        public string Folders { get; set; }
        public long Primary { get; set; }
        public long ArchiveMessages { get; set; }
        public long Recoverable { get; set; }
        public string State { get; set; }
    }

    /// <summary>A list bound to the window, replaced in one go (one refresh instead of one per row).</summary>
    public class BulkCollection : ObservableCollection<object>
    {
        public void ReplaceAll(IEnumerable items)
        {
            Items.Clear();
            if (items != null) { foreach (var i in items) { Items.Add(i); } }
            OnPropertyChanged(new PropertyChangedEventArgs("Count"));
            OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
            OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
}
