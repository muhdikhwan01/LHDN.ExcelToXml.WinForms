using LHDN.ExcelToXml.WinForms.Services;
using LHDN.ExcelToXml.WinForms.Validation;  // ✅ add this line
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LHDN.ExcelToXml.WinForms
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void Log(string msg)
        {
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
        }

        private void btnAddExcel_Click(object sender, EventArgs e)
        {
            using var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Multiselect = true
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (var file in ofd.FileNames)
                {
                    if (!lstFiles.Items.Contains(file))
                        lstFiles.Items.Add(file);
                }
                Log($"Added {ofd.FileNames.Length} file(s).");
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (lstFiles.Items.Count == 0)
            {
                MessageBox.Show("Please add at least one Excel file.", "No files", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (folderDialog.ShowDialog() != DialogResult.OK)
                return;

            string outputDir = folderDialog.SelectedPath;
            progressBar.Value = 0;
            progressBar.Maximum = lstFiles.Items.Count;

            foreach (string filePath in lstFiles.Items)
            {
                try
                {
                    Log($"Processing file: {filePath}");
                    var (appType, instruments) = ExcelReader.LoadFromExcel(filePath, Log);
                    Log($"Detected applicationType: {appType}. Found {instruments.Count} record(s).");

                    var issues = instruments.SelectMany(i => Validator.ValidateInstrument(i, appType)).ToList();
                    if (issues.Any())
                    {
                        Log("Validation issues:");
                        foreach (var issue in issues) Log(" - " + issue);
                        var choice = MessageBox.Show($"Found {issues.Count} issue(s). Continue anyway?", "Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (choice == DialogResult.No)
                        {
                            Log("Skipped file due to validation.");
                            progressBar.Value++;
                            continue;
                        }
                    }

                    var outputs = XmlGenerator.GenerateXmlFiles(appType, instruments, outputDir, Log);
                    foreach (var outFile in outputs)
                        Log($"✅ Generated: {outFile}");
                }
                catch (Exception ex)
                {
                    Log($"❌ Error: {ex.Message}");
                }

                progressBar.Value++;
            }

            Log("All files processed successfully.");
            MessageBox.Show("Conversion complete! Check logs for output paths.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
