using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Xml.Linq;
using LHDN.ExcelToXml.WinForms.Models;

namespace LHDN.ExcelToXml.WinForms.Services
{
    public static class XmlGenerator
    {
        const double MaxMb = 30.0;
        public static List<string> GenerateXmlFiles(int appType, List<Instrument> instruments, string outputDir, Action<string> log)
        {
            // Build XElement instruments individually (for easy batching)
            var instrumentElements = instruments.Select(inst => BuildInstrumentElement(appType, inst)).ToList();

            var outputFiles = new List<string>();
            int part = 1;
            int i = 0;
            while (i < instrumentElements.Count)
            {
                // create a batch and keep adding until size ok
                var batch = new List<XElement>();
                for (int j = i; j < instrumentElements.Count; j++)
                {
                    batch.Add(instrumentElements[j]);

                    // create doc and test size
                    var testDoc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                                    new XElement("bulkstamping",
                                        new XElement("applicationType", appType),
                                        batch
                                    ));
                    // save to memory stream to check length
                    using var ms = new MemoryStream();
                    testDoc.Save(ms);
                    double mb = ms.Length / (1024.0 * 1024.0);
                    if (mb > MaxMb)
                    {
                        // if batch currently > Max, remove last and finalize current batch
                        batch.RemoveAt(batch.Count - 1);
                        break;
                    }
                    else
                    {
                        // continue
                    }
                }

                if (batch.Count == 0)
                {
                    // single instrument too large (unlikely), force include one instrument then error if still >max
                    batch.Add(instrumentElements[i]);
                }

                var doc = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"),
                            new XElement("bulkstamping",
                                new XElement("applicationType", appType),
                                batch
                            ));

                var filename = Path.Combine(outputDir, $"{(appType == 43 ? "output_sekuriti" : "output_am")}_part{part}.xml");
                doc.Save(filename);
                outputFiles.Add(filename);
                log?.Invoke($"Saved {filename} ({new FileInfo(filename).Length / (1024.0 * 1024.0):F2} MB) -- contains {batch.Count} instruments.");

                i += batch.Count;
                part++;
            }

            return outputFiles;
        }

        static XElement BuildInstrumentElement(int appType, Instrument inst)
        {
            XElement[] transferorEls = inst.Transferors.Select(t =>
                new XElement("transferor",
                    new XElement("type", t.Type),
                    new XElement("name", SecurityElement.Escape(t.Name)),
                    new XElement("nationality", SecurityElement.Escape(t.Nationality)),
                    new XElement("icNo", SecurityElement.Escape(t.IcNo)),
                    new XElement("passportNo", SecurityElement.Escape(t.PassportNo)),
                    new XElement("passportCountry", SecurityElement.Escape(t.PassportCountry)),
                    new XElement("rocNo", SecurityElement.Escape(t.RocNo)),
                    new XElement("busType", SecurityElement.Escape(t.BusType)),
                    new XElement("incomeTaxNo", SecurityElement.Escape(t.IncomeTaxNo)),
                    new XElement("incomeTaxBranch", SecurityElement.Escape(t.IncomeTaxBranch)),
                    new XElement("street1", SecurityElement.Escape(t.Street1)),
                    new XElement("street2", SecurityElement.Escape(t.Street2)),
                    new XElement("street3", SecurityElement.Escape(t.Street3)),
                    new XElement("postcode", SecurityElement.Escape(t.Postcode)),
                    new XElement("city", SecurityElement.Escape(t.City)),
                    new XElement("state", SecurityElement.Escape(t.State)),
                    new XElement("country", SecurityElement.Escape(t.Country)),
                    new XElement("telNo", SecurityElement.Escape(t.TelNo)),
                    new XElement("email", SecurityElement.Escape(t.Email))
                )).ToArray();

            XElement[] transfereeEls = inst.Transferees.Select(t =>
                new XElement("transferee",
                    new XElement("type", t.Type),
                    new XElement("name", SecurityElement.Escape(t.Name)),
                    new XElement("nationality", SecurityElement.Escape(t.Nationality)),
                    new XElement("icNo", SecurityElement.Escape(t.IcNo)),
                    new XElement("passportNo", SecurityElement.Escape(t.PassportNo)),
                    new XElement("passportCountry", SecurityElement.Escape(t.PassportCountry)),
                    new XElement("rocNo", SecurityElement.Escape(t.RocNo)),
                    new XElement("busType", SecurityElement.Escape(t.BusType)),
                    new XElement("incomeTaxNo", SecurityElement.Escape(t.IncomeTaxNo)),
                    new XElement("incomeTaxBranch", SecurityElement.Escape(t.IncomeTaxBranch)),
                    new XElement("street1", SecurityElement.Escape(t.Street1)),
                    new XElement("street2", SecurityElement.Escape(t.Street2)),
                    new XElement("street3", SecurityElement.Escape(t.Street3)),
                    new XElement("postcode", SecurityElement.Escape(t.Postcode)),
                    new XElement("city", SecurityElement.Escape(t.City)),
                    new XElement("state", SecurityElement.Escape(t.State)),
                    new XElement("country", SecurityElement.Escape(t.Country)),
                    new XElement("telNo", SecurityElement.Escape(t.TelNo)),
                    new XElement("email", SecurityElement.Escape(t.Email))
                )).ToArray();

            var children = new List<object>
            {
                new XElement("refNo", inst.RefNo),
                new XElement("instrumentDate", inst.InstrumentDate ?? ""),
                new XElement("instrumentDateReceive", inst.InstrumentDateReceive ?? "")
            };

            if (appType == 43)
            {
                children.Add(new XElement("principal", inst.Principal));
                children.Add(new XElement("subsidiary", inst.Subsidiary));
                children.Add(new XElement("typeOfInstrument", inst.TypeOfInstrument ?? ""));
                children.Add(new XElement("typeOfInstrumentOthers", SecurityElement.Escape(inst.TypeOfInstrumentOthers ?? "")));
                children.AddRange(transferorEls);
                children.AddRange(transfereeEls);
                children.Add(new XElement("consideration", inst.Consideration ?? ""));
                children.Add(new XElement("duration", inst.Duration ?? ""));
                children.Add(new XElement("durationDesc", inst.DurationDesc ?? ""));
                // collateral nodes placeholder
                children.Add(new XElement("colLand", ""));
                children.Add(new XElement("colLandDesc", ""));
                children.Add(new XElement("colShare", ""));
                children.Add(new XElement("colDeposit", ""));
                children.Add(new XElement("colOthers", ""));
                children.Add(new XElement("colOthersDesc", ""));
            }
            else
            {
                children.Add(new XElement("typeOfInstrument", inst.TypeOfInstrument ?? ""));
                children.Add(new XElement("typeOfInstrumentOthers", SecurityElement.Escape(inst.TypeOfInstrumentOthers ?? "")));
                children.AddRange(transferorEls);
                children.AddRange(transfereeEls);
                children.Add(new XElement("noOfCopy", inst.NoOfCopy ?? ""));
                children.Add(new XElement("remissionOrExemption", inst.RemissionOrExemption ?? ""));
                children.Add(new XElement("payment", inst.Payment ?? ""));
                children.Add(new XElement("aggrementInfo", SecurityElement.Escape(inst.AggrementInfo ?? "")));
            }

            var attachment = new XElement("attachment", new XAttribute("name", inst.AttachmentName ?? ""), inst.AttachmentBase64 ?? "");
            children.Add(attachment);

            return new XElement("instrument", children.ToArray());
        }
    }
}
