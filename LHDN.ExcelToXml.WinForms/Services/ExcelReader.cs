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
        // common header candidates (lowercase keys)
        static readonly Dictionary<string, string[]> HeaderAlternatives = new()
        {
            { "applicationType", new[] { "applicationtype", "application type", "application_type" } },
            { "refNo", new[] { "refno", "ref no", "ref" } },
            { "instrumentDate", new[] { "instrumentdate", "instrument date", "date" } },
            { "typeOfInstrumentOthers", new[] { "typeofinstrumentothers", "type of instrument others", "typeOfInstrument" } },
            { "transferorName", new[] { "transferorname", "transferor name", "pihak pertama", "nama pemberi" } },
            { "transfereeName", new[] { "transfereename", "transferee name", "pihak kedua", "nama penerima" } },
            { "attachmentPath", new[] { "attachmentpath", "attachment path", "attachment file", "attachment" } },
            { "attachmentName", new[] { "attachmentname", "attachment name", "attachment name=" } },
            { "applicationType43", new[] { "43" } }
        };

        public static (int appType, List<Instrument> instruments) LoadFromExcel(string path, Action<string> log)
        {
            var instruments = new List<Instrument>();
            int detectedAppType = 44; // default

            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheets.First();

            // read header row and map column letter -> header key (normalized)
            var headerRow = ws.FirstRowUsed();
            var headerMap = new Dictionary<int, string>(); // columnNumber -> headerNormalized

            foreach (var cell in headerRow.CellsUsed())
            {
                var text = cell.GetString().Trim();
                if (string.IsNullOrEmpty(text)) continue;
                var normalized = text.Trim().ToLowerInvariant();
                headerMap[cell.Address.ColumnNumber] = normalized;
            }

            // helper: find column index by any candidate list
            int FindCol(params string[] candidates)
            {
                var candSet = new HashSet<string>(candidates.Select(c => c.ToLowerInvariant()));
                var found = headerMap.FirstOrDefault(kv => candSet.Contains(kv.Value) || candSet.Any(c => kv.Value.Contains(c)));
                return found.Key == 0 ? -1 : found.Key;
            }

            // fallback mapping: try to find common names
            int colApplicationType = headerMap.FirstOrDefault(kv => kv.Value.Contains("applicationtype") || kv.Value.Contains("application type")).Key;
            int colRefNo = headerMap.FirstOrDefault(kv => kv.Value.Contains("ref") && kv.Value.Contains("no") || kv.Value == "refno" || kv.Value == "ref").Key;
            if (colRefNo == 0) colRefNo = headerMap.FirstOrDefault(kv => kv.Value.Contains("ref")).Key;

            // prepare indices (look up many fields by name; common names are used)
            int idxAppType = headerMap.FirstOrDefault(kv => kv.Value.Contains("applicationtype") || kv.Value.Contains("application type")).Key;
            int idxRefNo = headerMap.FirstOrDefault(kv => kv.Value.Contains("refno") || kv.Value.Contains("ref no") || kv.Value == "ref").Key;
            int idxInstrumentDate = headerMap.FirstOrDefault(kv => kv.Value.Contains("instrumentdate") || kv.Value.Contains("instrument date") || kv.Value == "date").Key;
            int idxTypeOfInstrumentOthers = headerMap.FirstOrDefault(kv => kv.Value.Contains("typeofinstrumentothers") || kv.Value.Contains("type of instrument") || kv.Value.Contains("typeofinstrument")).Key;

            // find transferor/transferee name columns
            int idxTransferorName = headerMap.FirstOrDefault(kv => kv.Value.Contains("transferor") && kv.Value.Contains("name") || kv.Value.Contains("pihak pert")).Key;
            int idxTransfereeName = headerMap.FirstOrDefault(kv => kv.Value.Contains("transferee") && kv.Value.Contains("name") || kv.Value.Contains("pihak ked")).Key;

            // fallback: detect "name" columns: find first name not used by others
            if (idxTransferorName == 0) idxTransferorName = headerMap.FirstOrDefault(kv => kv.Value == "name" || kv.Value.Contains("nama")).Key;
            if (idxTransfereeName == 0)
            {
                // pick next "name" column if available
                var nameColumns = headerMap.Where(kv => kv.Value == "name" || kv.Value.Contains("nama") || kv.Value.Contains("name")).Select(kv => kv.Key).ToList();
                if (nameColumns.Count >= 2)
                    idxTransfereeName = nameColumns.ElementAtOrDefault(1);
            }

            // attachment path
            int idxAttachmentPath = headerMap.FirstOrDefault(kv => kv.Value.Contains("attachment") || kv.Value.Contains("attachmentpath") || kv.Value.Contains("file")).Key;

            // now iterate rows after header
            foreach (var row in ws.RowsUsed().Skip(1))
            {
                try
                {
                    int appType = 44;
                    if (idxAppType > 0)
                    {
                        var atxt = row.Cell(idxAppType).GetString().Trim();
                        if (int.TryParse(atxt, out var v)) appType = v;
                    }
                    detectedAppType = appType;

                    var inst = new Instrument
                    {
                        RefNo = (idxRefNo > 0) ? row.Cell(idxRefNo).GetString().Trim() : row.Cell(1).GetString().Trim(),
                        InstrumentDate = (idxInstrumentDate > 0) ? row.Cell(idxInstrumentDate).GetString().Trim() : "",
                        TypeOfInstrumentOthers = (idxTypeOfInstrumentOthers > 0) ? row.Cell(idxTypeOfInstrumentOthers).GetString().Trim() : ""
                    };

                    // Transferor
                    var tr = new Party();
                    if (idxTransferorName > 0) tr.Name = row.Cell(idxTransferorName).GetString().Trim();

                    // Transferee
                    var tf = new Party();
                    if (idxTransfereeName > 0) tf.Name = row.Cell(idxTransfereeName).GetString().Trim();

                    inst.Transferors.Add(tr);
                    inst.Transferees.Add(tf);

                    // if applicationType == 43: try to read some extra fields by common column names
                    if (appType == 43)
                    {
                        // try to read consideration (search header by "consideration" word)
                        var idxConsideration = headerMap.FirstOrDefault(kv => kv.Value.Contains("consideration") || kv.Value.Contains("amount") || kv.Value.Contains("consider")).Key;
                        if (idxConsideration > 0) inst.Consideration = row.Cell(idxConsideration).GetString().Trim();

                        var idxPrincipal = headerMap.FirstOrDefault(kv => kv.Value.Contains("principal")).Key;
                        if (idxPrincipal > 0 && int.TryParse(row.Cell(idxPrincipal).GetString(), out var pr)) inst.Principal = pr;
                    }
                    else
                    {
                        // AM: look for payment or aggrementInfo by header names
                        var idxPayment = headerMap.FirstOrDefault(kv => kv.Value.Contains("payment") || kv.Value.Contains("paymentamount")).Key;
                        if (idxPayment > 0) inst.Payment = row.Cell(idxPayment).GetString().Trim();

                        var idxAgreement = headerMap.FirstOrDefault(kv => kv.Value.Contains("agreement") || kv.Value.Contains("aggrementinfo") || kv.Value.Contains("agreementinfo")).Key;
                        if (idxAgreement > 0) inst.AggrementInfo = row.Cell(idxAgreement).GetString().Trim();
                    }

                    // Attachment: if supplied path exists
                    if (idxAttachmentPath > 0)
                    {
                        var attach = row.Cell(idxAttachmentPath).GetString().Trim();
                        if (!string.IsNullOrEmpty(attach))
                        {
                            if (File.Exists(attach))
                            {
                                inst.AttachmentBase64 = Convert.ToBase64String(File.ReadAllBytes(attach));
                                inst.AttachmentName = Path.GetFileName(attach);
                            }
                            else
                            {
                                log?.Invoke($"Attachment path not found: {attach}");
                                inst.AttachmentName = Path.GetFileName(attach);
                            }
                        }
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
    }
}
