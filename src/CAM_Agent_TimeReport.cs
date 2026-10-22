using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;

namespace CamAgentReports
{
    // Offline XLSX export from a committed CSV snapshot. This class never writes
    // the time database and never starts or attaches to an Excel process.
    public static class TimeReport
    {
        private static readonly string[] Headers = {
            "ProgramCodeF", "PreviousValidTimeSec", "PreviousValidMeasuredAt", "PreviousValidSource",
            "LastMeasuredTimeSec", "LastMeasuredAt", "LastSource", "ActiveTimeSec",
            "ActiveMeasuredAt", "ActiveSource", "Status", "UpdatedAt"
        };

        private static string ReadCommittedSnapshot(string path)
        {
            string directory = Path.GetDirectoryName(Path.GetFullPath(path));
            string lockPath = Path.Combine(directory, "CAM_Agent_Times_Write.lock");
            string journalPath = Path.Combine(directory, "CAM_Agent_TimesImport_Commit.json");
            for (int attempt = 0; ; attempt++)
            {
                FileStream writeGuard;
                try
                {
                    writeGuard = new FileStream(lockPath, FileMode.OpenOrCreate,
                        FileAccess.ReadWrite, FileShare.None);
                }
                catch (IOException error)
                {
                    if (attempt >= 3) throw new IOException("Базата с времена се използва. Изчакай текущият импорт да приключи или го спри безопасно, след което създай отчета.", error);
                    Thread.Sleep(50 * (attempt + 1));
                    continue;
                }
                using (writeGuard)
                {
                    if (File.Exists(journalPath)) throw new InvalidOperationException("Намерен е незавършен запис. Продължи импорта, за да се възстанови базата, преди да създадеш отчет.");
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read,
                        FileShare.Read | FileShare.Delete))
                    using (var reader = new StreamReader(stream, Encoding.UTF8, true))
                        return reader.ReadToEnd();
                }
            }
        }

        private static List<string[]> ReadCsv(string text)
        {
            var rows = new List<string[]>();
            var row = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (i == 0 && ch == '\uFEFF') continue;
                if (ch == '"')
                {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"') { field.Append('"'); i++; }
                    else if (quoted || field.Length == 0) quoted = !quoted;
                    else field.Append(ch);
                }
                else if (!quoted && ch == ';') { row.Add(field.ToString()); field.Clear(); }
                else if (!quoted && (ch == '\r' || ch == '\n'))
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    row.Add(field.ToString()); field.Clear();
                    if (row.Count > 1 || row[0].Length > 0) rows.Add(row.ToArray());
                    row.Clear();
                }
                else field.Append(ch);
            }
            if (quoted) throw new InvalidDataException("Незавършено текстово поле в базата с времена.");
            if (row.Count > 0 || field.Length > 0) { row.Add(field.ToString()); rows.Add(row.ToArray()); }
            return rows;
        }

        private static bool IsTimeColumn(int column) { return column == 1 || column == 4 || column == 7; }
        private static bool IsDateColumn(int column) { return column == 2 || column == 5 || column == 8 || column == 11; }
        private static string Address(int column, int row) { return ((char)('A' + column)).ToString() + row.ToString(CultureInfo.InvariantCulture); }

        private static XElement TextCell(XNamespace ns, string address, string value, string style)
        {
            // inlineStr prevents Excel auto-dates, lost leading zeros and formula execution.
            return new XElement(ns + "c", new XAttribute("r", address), new XAttribute("s", style),
                new XAttribute("t", "inlineStr"), new XElement(ns + "is",
                    new XElement(ns + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), value)));
        }

        public static int Export(string csvPath, string templatePath, string outputPath)
        {
            if (!File.Exists(templatePath)) throw new FileNotFoundException("Липсва Excel шаблонът за отчета.", templatePath);
            if (File.Exists(outputPath)) throw new IOException("Отчет с това име вече съществува: " + outputPath);
            var rows = ReadCsv(ReadCommittedSnapshot(csvPath));
            if (rows.Count == 0) throw new InvalidDataException("Базата с времена няма заглавен ред.");
            int[] indices = new int[Headers.Length];
            for (int c = 0; c < Headers.Length; c++)
            {
                indices[c] = Array.FindIndex(rows[0], delegate(string h) { return String.Equals(h, Headers[c], StringComparison.Ordinal); });
                if (indices[c] < 0) throw new InvalidDataException("Липсва колона в базата: " + Headers[c]);
            }
            if (rows.Count > 1048576) throw new InvalidDataException("Отчетът надхвърля максималния брой редове в Excel.");
            string temporary = outputPath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            int invalidNumbers = 0;
            try
            {
                File.Copy(templatePath, temporary, false);
                using (var zip = ZipFile.Open(temporary, ZipArchiveMode.Update))
                {
                    var entry = zip.GetEntry("xl/worksheets/sheet1.xml");
                    if (entry == null) throw new InvalidDataException("Excel шаблонът няма лист с времена.");
                    XDocument document;
                    using (var stream = entry.Open()) document = XDocument.Load(stream);
                    XNamespace ns = document.Root.Name.Namespace;
                    var data = document.Root.Element(ns + "sheetData");
                    if (data == null) throw new InvalidDataException("Excel шаблонът няма sheetData.");
                    var header = data.Elements(ns + "row").FirstOrDefault(x => (string)x.Attribute("r") == "1");
                    var sample = data.Elements(ns + "row").FirstOrDefault(x => (string)x.Attribute("r") == "2");
                    if (header == null || sample == null) throw new InvalidDataException("Непълен Excel шаблон.");
                    var styles = new string[12];
                    for (int c = 0; c < 12; c++)
                    {
                        var cell = sample.Elements(ns + "c").FirstOrDefault(x => (string)x.Attribute("r") == Address(c, 2));
                        if (cell == null || cell.Attribute("s") == null) throw new InvalidDataException("Липсва формат за " + Address(c, 2));
                        styles[c] = (string)cell.Attribute("s");
                    }
                    var newData = new XElement(ns + "sheetData", new XElement(header));
                    for (int r = 1; r < rows.Count; r++)
                    {
                        if (rows[r].Length != rows[0].Length) throw new InvalidDataException("Различен брой колони в ред " + (r + 1));
                        var rowElement = new XElement(ns + "row", new XAttribute("r", r + 1),
                            new XAttribute("ht", "30"), new XAttribute("customHeight", "1"));
                        for (int c = 0; c < 12; c++)
                        {
                            string value = rows[r][indices[c]] ?? String.Empty;
                            string address = Address(c, r + 1);
                            if (value.Length == 0)
                            {
                                rowElement.Add(new XElement(ns + "c", new XAttribute("r", address), new XAttribute("s", styles[c])));
                                continue;
                            }
                            double number;
                            DateTime date;
                            if (IsTimeColumn(c) && Double.TryParse(value.Trim().Replace(',', '.'), NumberStyles.Float,
                                CultureInfo.InvariantCulture, out number) && !Double.IsNaN(number) && !Double.IsInfinity(number))
                            {
                                rowElement.Add(new XElement(ns + "c", new XAttribute("r", address), new XAttribute("s", styles[c]),
                                    new XAttribute("t", "n"), new XElement(ns + "v", number.ToString("R", CultureInfo.InvariantCulture))));
                            }
                            else if (IsDateColumn(c) && DateTime.TryParseExact(value.Trim(),
                                new string[] { "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd HH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss", "o" },
                                CultureInfo.InvariantCulture, DateTimeStyles.None, out date) && date.Year >= 1900)
                            {
                                rowElement.Add(new XElement(ns + "c", new XAttribute("r", address), new XAttribute("s", styles[c]),
                                    new XAttribute("t", "n"), new XElement(ns + "v", date.ToOADate().ToString("R", CultureInfo.InvariantCulture))));
                            }
                            else
                            {
                                if (IsTimeColumn(c)) invalidNumbers++;
                                rowElement.Add(TextCell(ns, address, value, IsTimeColumn(c) || IsDateColumn(c) ? styles[0] : styles[c]));
                            }
                        }
                        newData.Add(rowElement);
                    }
                    if (rows.Count == 1) newData.Add(new XElement(sample));
                    data.ReplaceWith(newData);
                    int lastRow = Math.Max(2, rows.Count);
                    var dimension = document.Root.Element(ns + "dimension");
                    if (dimension != null) dimension.SetAttributeValue("ref", "A1:L" + lastRow.ToString(CultureInfo.InvariantCulture));
                    var filter = document.Root.Element(ns + "autoFilter");
                    if (filter == null) { filter = new XElement(ns + "autoFilter"); newData.AddAfterSelf(filter); }
                    filter.SetAttributeValue("ref", "A1:L" + lastRow.ToString(CultureInfo.InvariantCulture));
                    entry.Delete();
                    var newEntry = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
                    using (var stream = newEntry.Open()) document.Save(stream, SaveOptions.DisableFormatting);
                }
                File.Move(temporary, outputPath);
                return invalidNumbers;
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
