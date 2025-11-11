using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using LHDN.ExcelToXml.WinForms.Models;

namespace LHDN.ExcelToXml.WinForms.Services
{
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
            // Transferor columns end roughly before the second "type" header appears
            int secondTypeCol = headers.FirstOrDefault(h => h.Key > headers.First().Key && h.Value == "type").Key;
            int transferorEnd = (secondTypeCol > 0) ? secondTypeCol - 1 : totalCols;

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                try
                {
                    int appType = TryInt(GetValue(row, headers, "applicationtype"));
                    detectedAppType = (appType == 0) ? 44 : appType;

                    var inst = new Instrument
                    {
                        RefNo = GetValue(row, headers, "refno"),
                        InstrumentDate = GetValue(row, headers, "instrumentdate"),
                        InstrumentDateReceive = GetValue(row, headers, "instrumentdatereceive"),
                        TypeOfInstrument = GetValue(row, headers, "typeofinstrument"),
                        TypeOfInstrumentOthers = GetValue(row, headers, "typeofinstrumentothers"),
                        NoOfCopy = GetValue(row, headers, "noofcopy"),
                        RemissionOrExemption = GetValue(row, headers, "remissionorexemption"),
                        Payment = GetValue(row, headers, "payment"),
                        AggrementInfo = GetValue(row, headers, "aggrementinfo"),
                        AttachmentName = GetValue(row, headers, "attachment name=")
                    };

                    // --- TRANSFEROR (left side) ---
                    var tr = new Party();
                    var keys = headers.Keys.ToList();
                    for (int i = headers.First().Key; i <= transferorEnd; i++)
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
                    for (int i = transferorEnd + 1; i <= totalCols; i++)
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

        private static int TryInt(string s)
        {
            return int.TryParse(s, out var i) ? i : 0;
        }
    }
}
