using ClosedXML.Excel;
using LHDN.ExcelToXml.WinForms.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace LHDN.ExcelToXml.WinForms.Services
{
    // Reads and parses Excel data
    // Purpose: Converts Excel sheet → C# objects, ready for XML conversion.
    public static class ExcelReader
    {
        public static (int appType, List<Instrument> instruments) LoadFromExcel(string path, Action<string> log)
        {
            var instruments = new List<Instrument>();
            int detectedAppType = 44;

            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheets.First();
            var headerRow = ws.FirstRowUsed();
            var headers = headerRow.CellsUsed().ToDictionary(c => c.Address.ColumnNumber, c => c.GetString().Trim().ToLowerInvariant());

            // Count columns to find where transferee columns start
            int totalCols = headers.Count;
            // Hardcode Excel structure based on actual layout
            // Transferor columns: G1–Y1 (7–25)
            // Transferee columns: Z1–AR1 (26–44)
            int transferorStart = 7;
            int transferorEnd = 25;
            int transfereeStart = 26;
            int transfereeEnd = 44;

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                try
                {
                    // Detect application type from first data row
                    int appType = TryInt(GetValue(row, headers, "applicationtype"));
                    detectedAppType = (appType == 0) ? 44 : appType;

                    // find instrumentDate column index (safe lookup)
                    int instrumentDateCol = headers.FirstOrDefault(h => h.Value.Contains("instrumentdate")).Key;

                    // read raw cell value (if column exists)
                    string rawInstrumentDate = instrumentDateCol > 0 ? row.Cell(instrumentDateCol).GetString() : "";

                    // log the raw value (will reveal hidden chars, dashes, etc.)
                    log?.Invoke($"RAW instrumentDate cell value: '{rawInstrumentDate}'");

                    // Read main instrument fields
                    var inst = new Instrument
                    {
                        RefNo = GetValue(row, headers, "refno"),
                        InstrumentDate = GetDateCellString(row, headers, "instrumentdate", log),
                        InstrumentDateReceive = GetDateCellString(row, headers, "instrumentdatereceive", log),
                        TypeOfInstrument = GetValue(row, headers, "typeofinstrument"),
                        TypeOfInstrumentOthers = GetValue(row, headers, "typeofinstrumentothers"),
                        NoOfCopy = GetValue(row, headers, "noofcopy"),
                        RemissionOrExemption = GetValue(row, headers, "remissionorexemption"),
                        Payment = GetValue(row, headers, "payment"),
                        AggrementInfo = GetValue(row, headers, "aggrementinfo"),
                        AttachmentName = GetValue(row, headers, "attachment name=")
                    };
                    log?.Invoke($"PARSED instrumentDate value: '{inst.InstrumentDate}'");

                    // --- TRANSFEROR (left side) ---
                    var tr = new Party();
                    var keys = headers.Keys.ToList();
                    // --- TRANSFEROR (columns G–Y) ---
                    for (int i = transferorStart; i <= transferorEnd; i++)
                    {
                        string col = headers[i];
                        string val = row.Cell(i).GetString().Trim();

                        if (col == "type") tr.Type = TryInt(val);
                        else if (col == "name") tr.Name = val;
                        else if (col == "nationality") tr.Nationality = val;
                        else if (col == "icno") tr.IcNo = val;
                        else if (col == "pasportno") tr.PassportNo = val;
                        else if (col == "pasportcountry") tr.PassportCountry = val;
                        else if (col == "rocno") tr.RocNo = val;
                        else if (col == "bustype") tr.BusType = val;
                        else if (col == "incometaxno") tr.IncomeTaxNo = val;
                        else if (col == "incometaxbranch") tr.IncomeTaxBranch = val;
                        else if (col == "street1") tr.Street1 = val;
                        else if (col == "street2") tr.Street2 = val;
                        else if (col == "street3") tr.Street3 = val;
                        else if (col == "postcode") tr.Postcode = val;
                        else if (col == "city") tr.City = val;
                        else if (col == "state") tr.State = val;
                        else if (col == "country") tr.Country = val;
                        else if (col == "telno") tr.TelNo = val;
                        else if (col == "email") tr.Email = val;
                    }

                    // --- TRANSFEREE (right side) ---
                    var tf = new Party();
                    // --- TRANSFEREE (columns Z–AR) ---
                    for (int i = transfereeStart; i <= transfereeEnd; i++)
                    {
                        string col = headers.ContainsKey(i) ? headers[i] : "";
                        string val = row.Cell(i).GetString().Trim();

                        if (col == "type") tf.Type = TryInt(val);
                        else if (col == "name") tf.Name = val;
                        else if (col == "nationality") tf.Nationality = val;
                        else if (col == "icno") tf.IcNo = val;
                        else if (col == "pasportno") tf.PassportNo = val;
                        else if (col == "pasportcountry") tf.PassportCountry = val;
                        else if (col == "rocno") tf.RocNo = val;
                        else if (col == "bustype") tf.BusType = val;
                        else if (col == "incometaxno") tf.IncomeTaxNo = val;
                        else if (col == "incometaxbranch") tf.IncomeTaxBranch = val;
                        else if (col == "street1") tf.Street1 = val;
                        else if (col == "street2") tf.Street2 = val;
                        else if (col == "street3") tf.Street3 = val;
                        else if (col == "postcode") tf.Postcode = val;
                        else if (col == "city") tf.City = val;
                        else if (col == "state") tf.State = val;
                        else if (col == "country") tf.Country = val;
                        else if (col == "telno") tf.TelNo = val;
                        else if (col == "email") tf.Email = val;
                    }

                    inst.Transferors.Add(tr);
                    inst.Transferees.Add(tf);

                    // Attachment handling
                    // Load file and convert to Base64
                    string attachPath = GetValue(row, headers, "attachment name=");
                    if (!string.IsNullOrWhiteSpace(attachPath) && File.Exists(attachPath))
                    {
                        inst.AttachmentBase64 = Convert.ToBase64String(File.ReadAllBytes(attachPath));
                        inst.AttachmentName = Path.GetFileName(attachPath);
                    }

                    instruments.Add(inst);
                }
                catch (Exception ex)
                {
                    log?.Invoke($"Row parse error: {ex.Message}");
                }
            }

            return (detectedAppType, instruments);
        }

        // Utility
        private static string GetValue(IXLRow row, Dictionary<int, string> headers, string key)
        {
            var match = headers.FirstOrDefault(h => h.Value.Contains(key.ToLower()));
            return match.Key > 0 ? row.Cell(match.Key).GetString().Trim() : "";
        }

        // Normalize Excel date cells to DD/MM/YYYY format for STAMPS compliance
        private static string GetDateCellString(IXLRow row, Dictionary<int, string> headers, string key, Action<string>? log = null)
        {
            var match = headers.FirstOrDefault(h => h.Value.Contains(key.ToLower()));
            if (match.Key <= 0) return "";

            var cell = row.Cell(match.Key);

            try
            {
                string result = "";

                // 1) If cell is true Excel date type, use it
                if (cell.DataType == XLDataType.DateTime)
                {
                    var dt = cell.GetDateTime();
                    result = dt.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    return result;
                }

                // 2) Read raw string and normalize
                var raw = cell.GetString().Trim();
                if (string.IsNullOrEmpty(raw)) return "";

                // Strip trailing time if any
                var dateCandidate = raw.Split(' ')[0].Trim();

                // Try parsing with common formats
                var formats = new[]
                {
                    "dd/MM/yyyy","d/M/yyyy",
                    "dd-MM-yyyy","d-M-yyyy",
                    "yyyy-MM-dd","yyyy/MM/dd",
                    "M/d/yyyy","MM/dd/yyyy",
                    "d MMM yyyy","dd MMM yyyy"
                };

                if (DateTime.TryParseExact(dateCandidate, formats, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedExact))
                {
                    result = parsedExact.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    return result;
                }

                // Try parse using en-GB culture (day/month/year)
                if (DateTime.TryParse(dateCandidate, new System.Globalization.CultureInfo("en-GB"),
                    System.Globalization.DateTimeStyles.None, out var parsedGb))
                {
                    result = parsedGb.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    return result;
                }

                // If still not parsed, try replacing any dash-like characters with '/'
                // normalize common dash characters to ASCII hyphen then to '/'
                var normalizedCandidate = dateCandidate
                    .Replace('–', '-')  // en-dash
                    .Replace('—', '-')  // em-dash
                    .Replace('\u2011', '-') // non-breaking hyphen
                    .Replace('‐', '-')    // hyphen
                    .Replace('-', '/');

                // try parsing again
                if (DateTime.TryParseExact(normalizedCandidate, formats, System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedNormalizedExact))
                {
                    result = parsedNormalizedExact.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    return result;
                }
                if (DateTime.TryParse(normalizedCandidate, new System.Globalization.CultureInfo("en-GB"),
                    System.Globalization.DateTimeStyles.None, out var parsedNormalizedGb))
                {
                    result = parsedNormalizedGb.ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                    return result;
                }

                // Last resort: return original candidate but with '-' replaced to '/'
                return dateCandidate.Replace('-', '/');
            }
            catch (Exception ex)
            {
                log?.Invoke($"Date parse warning for column '{key}': {ex.Message}");
                return "";
            }
        }

        private static int TryInt(string s)
        {
            return int.TryParse(s, out var i) ? i : 0;
        }
    }
}
