using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using OfficeOpenXml;

namespace NavisExcelExporter
{
    [Plugin("NavisExcelExporter.FullModelExportPlugin", "Omar", DisplayName = "Export Full Model to Excel")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class FullModelExportPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
                
                if (document == null || document.Models.Count == 0)
                {
                    MessageBox.Show("No model is currently open in Navisworks.", "No Model", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }

                using (var progressForm = new ProgressForm())
                {
                    progressForm.Show();
                    System.Windows.Forms.Application.DoEvents();

                    string excelFilePath = ExportModelToExcel(document, progressForm);
                    progressForm.Close();

                    MessageBox.Show($"Model data exported successfully!\n\nFile saved to:\n{excelFilePath}", 
                        "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting model data:\n\n{ex.Message}\n\n{ex.StackTrace}", "Export Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private string ExportModelToExcel(Document document, ProgressForm progressForm)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = Path.Combine(desktopPath, "NavisModelData.xlsx");

            if (File.Exists(excelFilePath))
            {
                File.Delete(excelFilePath);
            }

            // PASS 1: Collect all items and their properties to determine all unique columns
            progressForm.SetProgress(0, 100, "Pass 1: Scanning properties...");
            var allItems = GetAllModelItems(document).ToList();
            int totalItems = allItems.Count;

            var allPropertyKeys = new HashSet<string>();
            var itemDataList = new List<Dictionary<string, object>>();

            for (int i = 0; i < allItems.Count; i++)
            {
                var item = allItems[i];
                var itemData = ExtractItemData(item, allPropertyKeys);
                itemDataList.Add(itemData);

                if (i % 100 == 0)
                {
                    progressForm.SetProgress(i * 100 / totalItems, 100, 
                        $"Pass 1: Scanning item {i + 1} of {totalItems}...");
                    System.Windows.Forms.Application.DoEvents();
                }
            }

            // Sort property keys for consistent column order
            var sortedPropertyKeys = allPropertyKeys.OrderBy(k => k).ToList();

            // PASS 2: Write to Excel with all columns
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Model Data");

                // Create headers
                CreateHeaders(worksheet, sortedPropertyKeys);

                // Write data
                progressForm.SetProgress(0, 100, "Pass 2: Writing to Excel...");
                
                for (int i = 0; i < itemDataList.Count; i++)
                {
                    WriteItemToExcel(itemDataList[i], worksheet, i + 2, sortedPropertyKeys);

                    if (i % 100 == 0)
                    {
                        progressForm.SetProgress(i * 100 / totalItems, 100, 
                            $"Pass 2: Writing item {i + 1} of {totalItems}...");
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                // Auto-fit columns (can be slow for many columns - consider removing if too slow)
                progressForm.SetProgress(95, 100, "Formatting columns...");
                System.Windows.Forms.Application.DoEvents();
                
                // Only autofit first 50 columns to save time
                int columnsToFit = Math.Min(50, worksheet.Dimension.End.Column);
                worksheet.Cells[1, 1, worksheet.Dimension.End.Row, columnsToFit].AutoFitColumns();

                progressForm.SetProgress(98, 100, "Saving file...");
                System.Windows.Forms.Application.DoEvents();
                package.SaveAs(new FileInfo(excelFilePath));
            }

            return excelFilePath;
        }

        private Dictionary<string, object> ExtractItemData(ModelItem item, HashSet<string> allPropertyKeys)
        {
            var data = new Dictionary<string, object>();

            try
            {
                // Basic properties
                data["Element Name"] = item.DisplayName ?? item.ClassDisplayName ?? "Unknown";
                data["Category/Class"] = item.ClassDisplayName ?? "Unknown";

                // Extract GUID from properties (not InstanceGuid which is often all zeros)
                string guid = ExtractGuidFromProperties(item);
                data["GUID"] = guid;

                // Coordinates
                try
                {
                    BoundingBox3D boundingBox = item.BoundingBox();
                    Point3D center = boundingBox.Center;
                    data["X Coordinate"] = center.X;
                    data["Y Coordinate"] = center.Y;
                    data["Z Coordinate"] = center.Z;
                }
                catch
                {
                    data["X Coordinate"] = "";
                    data["Y Coordinate"] = "";
                    data["Z Coordinate"] = "";
                }

                // Add all properties as separate columns
                var propertyCategories = item.PropertyCategories;
                foreach (var category in propertyCategories)
                {
                    string categoryName = category.DisplayName;
                    
                    foreach (var property in category.Properties)
                    {
                        try
                        {
                            string propertyKey = $"{categoryName}.{property.DisplayName}";
                            string propertyValue = property.Value.ToDisplayString();

                            // Add to all keys set for column tracking
                            allPropertyKeys.Add(propertyKey);

                            // Store value
                            if (!string.IsNullOrEmpty(propertyValue))
                            {
                                data[propertyKey] = propertyValue;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                data["Error"] = ex.Message;
            }

            return data;
        }

        private string ExtractGuidFromProperties(ModelItem item)
        {
            try
            {
                var propertyCategories = item.PropertyCategories;
                
                foreach (var category in propertyCategories)
                {
                    foreach (var property in category.Properties)
                    {
                        // Look for GUID property (common names: GUID, Item.GUID, etc.)
                        if (property.DisplayName.Equals("GUID", StringComparison.OrdinalIgnoreCase) ||
                            property.Name.Equals("GUID", StringComparison.OrdinalIgnoreCase))
                        {
                            string guidValue = property.Value.ToDisplayString();
                            if (!string.IsNullOrEmpty(guidValue) && guidValue != "00000000-0000-0000-0000-000000000000")
                            {
                                return guidValue;
                            }
                        }
                    }
                }

                // Fallback to InstanceGuid if no GUID found in properties
                var instanceGuid = item.InstanceGuid.ToString();
                if (instanceGuid != "00000000-0000-0000-0000-000000000000")
                {
                    return instanceGuid;
                }

                return "No GUID Available";
            }
            catch
            {
                return "Error Reading GUID";
            }
        }

        private void CreateHeaders(ExcelWorksheet worksheet, List<string> propertyKeys)
        {
            // Fixed columns first
            var fixedHeaders = new[] { "Element Name", "Category/Class", "GUID", 
                                       "X Coordinate", "Y Coordinate", "Z Coordinate" };
            
            int colIndex = 1;
            foreach (var header in fixedHeaders)
            {
                worksheet.Cells[1, colIndex].Value = header;
                worksheet.Cells[1, colIndex].Style.Font.Bold = true;
                colIndex++;
            }

            // Dynamic property columns
            foreach (var propertyKey in propertyKeys)
            {
                worksheet.Cells[1, colIndex].Value = propertyKey;
                worksheet.Cells[1, colIndex].Style.Font.Bold = true;
                colIndex++;
            }
        }

        private void WriteItemToExcel(Dictionary<string, object> itemData, ExcelWorksheet worksheet, 
                                      int rowIndex, List<string> propertyKeys)
        {
            int colIndex = 1;

            // Fixed columns
            var fixedColumns = new[] { "Element Name", "Category/Class", "GUID", 
                                       "X Coordinate", "Y Coordinate", "Z Coordinate" };
            
            foreach (var colName in fixedColumns)
            {
                if (itemData.ContainsKey(colName))
                {
                    worksheet.Cells[rowIndex, colIndex].Value = itemData[colName];
                }
                colIndex++;
            }

            // Dynamic property columns
            foreach (var propertyKey in propertyKeys)
            {
                if (itemData.ContainsKey(propertyKey))
                {
                    worksheet.Cells[rowIndex, colIndex].Value = itemData[propertyKey];
                }
                colIndex++;
            }
        }

        private IEnumerable<ModelItem> GetAllModelItems(Document document)
        {
            foreach (var model in document.Models)
            {
                var rootItems = model.RootItem.Children;
                foreach (var item in rootItems)
                {
                    yield return item;
                    foreach (var child in GetAllChildren(item))
                    {
                        yield return child;
                    }
                }
            }
        }

        private IEnumerable<ModelItem> GetAllChildren(ModelItem parent)
        {
            foreach (var child in parent.Children)
            {
                yield return child;
                foreach (var grandchild in GetAllChildren(child))
                {
                    yield return grandchild;
                }
            }
        }
    }

    public partial class ProgressForm : Form
    {
        private Label statusLabel;
        private ProgressBar progressBar;

        public ProgressForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Exporting Model Data";
            this.Size = new System.Drawing.Size(400, 150);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.TopMost = true;

            statusLabel = new Label
            {
                Text = "Preparing export...",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(350, 20),
                AutoSize = false
            };

            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 50),
                Size = new System.Drawing.Size(350, 23),
                Style = ProgressBarStyle.Continuous
            };

            this.Controls.Add(statusLabel);
            this.Controls.Add(progressBar);
        }

        public void SetProgress(int current, int total, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new System.Action<int, int, string>(SetProgress), current, total, status);
                return;
            }

            statusLabel.Text = status;
            progressBar.Maximum = total;
            progressBar.Value = Math.Min(current, total);
        }
    }
}