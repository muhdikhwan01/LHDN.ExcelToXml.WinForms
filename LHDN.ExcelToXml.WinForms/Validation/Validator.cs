using System.Collections.Generic;
using LHDN.ExcelToXml.WinForms.Models;

namespace LHDN.ExcelToXml.WinForms.Validation
{
    // Checks for missing/invalid data
    // Purpose: Prevents incomplete XML from being uploaded.
    public static class Validator
    {
        // Return list of validation messages. Empty list => OK
        public static List<string> ValidateInstrument(Instrument inst, int appType)
        {
            var issues = new List<string>();

            if (string.IsNullOrWhiteSpace(inst.RefNo))
                issues.Add("Missing refNo.");

            if (string.IsNullOrWhiteSpace(inst.InstrumentDate))
                issues.Add($"Missing instrumentDate for {inst.RefNo}.");

            if (inst.Transferors.Count == 0 || string.IsNullOrWhiteSpace(inst.Transferors[0].Name))
                issues.Add($"Missing transferor name for {inst.RefNo}.");

            if (inst.Transferees.Count == 0 || string.IsNullOrWhiteSpace(inst.Transferees[0].Name))
                issues.Add($"Missing transferee name for {inst.RefNo}.");

            // Application type specific checks
            if (appType == 43)
            {
                // for Sekuriti, consideration is required
                if (string.IsNullOrWhiteSpace(inst.Consideration))
                    issues.Add($"Sekuriti: consideration required for {inst.RefNo}.");
            }
            else
            {
                // for Am, noOfCopy optional, but log if missing both payment & agreement info
                if (string.IsNullOrWhiteSpace(inst.Payment) && string.IsNullOrWhiteSpace(inst.AggrementInfo))
                    issues.Add($"Am: payment or aggrementInfo missing for {inst.RefNo}.");
            }

            return issues;
        }
    }
}
