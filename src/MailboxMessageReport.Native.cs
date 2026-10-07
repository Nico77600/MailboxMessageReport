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
//   CsvTarget   a CSV file written row by row (UTF-8 with BOM, formula injection neutralised)
//   RowBuffer   the first rows of a merge, kept as JSON for the HTML report and the window
//   Merge       a part file appended to CSV files and row buffers; the HTML report written from its template
//   Fast        dates, CSV cells
//   PreviewRow  a message in the window; BulkCollection, the list bound to the window
//
// Author : Nicolas Fabert
// Version: 1.0.0

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
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
            "From", "FromName", "Sender", "To", "Cc", "Bcc", "InternetMessageId", "HasAttachments", "Importance", "IsRead",
            "Type", "ReceivedUtc", "ItemId" };

        /// <summary>The Microsoft Graph properties read for each message ($select).</summary>
        public const string Select = "id,subject,receivedDateTime,sentDateTime,from,sender,toRecipients,ccRecipients,bccRecipients,internetMessageId,hasAttachments,importance,isRead";
    }

    public static class Fast
    {
        public const string Version = "1.0.0";
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

        /// <summary>A CSV cell: text starting with = + - @ (or tab, CR) prefixed with an apostrophe; quoted when needed.</summary>
        public static string CsvCell(string text, string delimiter)
        {
            if (string.IsNullOrEmpty(text)) { return ""; }
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
            _json.WriteStringValue(Recipients(m, "toRecipients"));
            _json.WriteStringValue(Recipients(m, "ccRecipients"));
            _json.WriteStringValue(Recipients(m, "bccRecipients"));
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
        static string Recipients(JsonElement o, string property)
        {
            JsonElement list;
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
    }

    public static class Merge
    {
        /// <summary>
        /// The HTML report: the template with its markers ({{NAME}}) replaced - the values given, the messages from their
        /// buffer - written as UTF-8 with BOM. The markers are found in the template first: a value that holds the text
        /// of a marker (a subject) is never touched. Every marker of the template must have a value, and the reverse.
        /// </summary>
        public static void WriteHtml(string templatePath, string outputPath, string[] markers, string[] values, RowBuffer messages, string messagesMarker)
        {
            var template = File.ReadAllText(templatePath, Encoding.UTF8);
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            for (int i = 0; i < markers.Length; i++) { map[markers[i]] = values[i] ?? ""; }
            var used = new HashSet<string>(StringComparer.Ordinal);
            var found = System.Text.RegularExpressions.Regex.Matches(template, @"\{\{[A-Z_]+\}\}");
            foreach (System.Text.RegularExpressions.Match m in found)
            {
                if (m.Value != messagesMarker && !map.ContainsKey(m.Value)) { throw new InvalidOperationException("Report template marker not replaced: " + m.Value); }
                if (!used.Add(m.Value)) { throw new InvalidOperationException("Report template marker found twice: " + m.Value); }
            }
            foreach (var k in map.Keys) { if (!used.Contains(k)) { throw new InvalidOperationException("Report template marker missing: " + k); } }
            if (!used.Contains(messagesMarker)) { throw new InvalidOperationException("Report template marker missing: " + messagesMarker); }
            // Written as it goes: the messages (up to hundreds of thousands) are never copied into one string.
            using (var w = new StreamWriter(outputPath, false, new UTF8Encoding(true), 1 << 16))
            {
                int last = 0;
                foreach (System.Text.RegularExpressions.Match m in found)
                {
                    w.Write(template.Substring(last, m.Index - last));
                    if (m.Value == messagesMarker)
                    {
                        w.Write('[');
                        if (messages != null)
                        {
                            for (int i = 0; i < messages.Lines.Count; i++) { if (i > 0) { w.Write(','); } w.Write(messages.Lines[i]); }
                        }
                        w.Write(']');
                    }
                    else { w.Write(map[m.Value]); }
                    last = m.Index + m.Length;
                }
                w.Write(template.Substring(last));
            }
        }
        /// <summary>The rows of a part file appended to CSV files and to row buffers, in the order of the file; returns how many rows.</summary>
        public static long AppendPart(string path, CsvTarget[] csv, RowBuffer[] buffers)
        {
            long n = 0;
            if (!File.Exists(path)) { return 0; }
            bool toCsv = csv != null && csv.Length > 0;
            foreach (var line in File.ReadLines(path, Encoding.UTF8))
            {
                if (line.Length == 0) { continue; }
                n++;
                if (toCsv)
                {
                    var cells = Fast.ParseRow(line);
                    foreach (var t in csv) { if (t != null) { t.Write(cells); } }
                }
                if (buffers != null) { foreach (var b in buffers) { if (b != null) { b.Add(line); } } }
            }
            return n;
        }
    }

    /// <summary>A message in the preview of the window.</summary>
    public sealed class PreviewRow
    {
        public string Received { get; set; }
        public string Mailbox { get; set; }
        public string Location { get; set; }
        public string FolderPath { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public string To { get; set; }

        /// <summary>The rows of a buffer, for the list of the window.</summary>
        public static List<object> Build(RowBuffer buffer)
        {
            var rows = new List<object>();
            if (buffer == null) { return rows; }
            foreach (var line in buffer.Lines)
            {
                var c = Fast.ParseRow(line);
                var location = c[2] + (c[3] == "Yes" ? " (RI)" : "");
                rows.Add(new PreviewRow { Mailbox = c[0], Location = location, FolderPath = c[4], Received = c[6], Subject = c[8], From = c[9].Length > 0 ? c[9] : c[10], To = c[12] });
            }
            return rows;
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
