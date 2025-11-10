using System.Collections.Generic;

namespace LHDN.ExcelToXml.WinForms.Models
{
    public class Instrument
    {
        public string RefNo { get; set; } = "";
        public string InstrumentDate { get; set; } = "";
        public string InstrumentDateReceive { get; set; } = "";
        public int Principal { get; set; } = -1;
        public string Subsidiary { get; set; } = "";
        public string TypeOfInstrument { get; set; } = "";
        public string TypeOfInstrumentOthers { get; set; } = "";

        public List<Party> Transferors { get; set; } = new();
        public List<Party> Transferees { get; set; } = new();

        // Sekuriti fields
        public string Consideration { get; set; } = "";
        public string Duration { get; set; } = "";
        public string DurationDesc { get; set; } = "";
        // Am fields
        public string NoOfCopy { get; set; } = "";
        public string RemissionOrExemption { get; set; } = "";
        public string Payment { get; set; } = "";
        public string AggrementInfo { get; set; } = "";

        // Attachment
        public string AttachmentName { get; set; } = "";
        public string AttachmentBase64 { get; set; } = "";
    }
}
