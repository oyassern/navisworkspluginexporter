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
                // Set EPPlus license context to NonCommercial
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Get the current document
                Document document = Autodesk.Navisworks.Api.Application.ActiveDocument;
                
                if (document == null || document.Models.Count == 0)
                {
                    MessageBox.Show("No model is currently open in Navisworks.", "No Model", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                }

                // Show progress dialog
                using (var progressForm = new ProgressForm())
                {
                    progressForm.Show();
                    System.Windows.Forms.Application.DoEvents();

                    // Export model data to Excel
                    string excelFilePath = ExportModelToExcel(document, progressForm);
                    
                    progressForm.Close();

                    // Show completion message
                    MessageBox.Show($"Model data exported successfully!\n\nFile saved to:\n{excelFilePath}", 
                        "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting model data:\n\n{ex.Message}", "Export Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return -1;
            }
        }

        private string ExportModelToExcel(Document document, ProgressForm progressForm)
        {
            // Create Excel file path on desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string excelFilePath = Path.Combine(desktopPath, "NavisModelData.xlsx");

            // Delete existing file if it exists
            if (File.Exists(excelFilePath))
            {
                File.Delete(excelFilePath);
            }

            // Create Excel package
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Model Data");

                // Create headers
                CreateHeaders(worksheet);

                int rowIndex = 2; // Start from row 2 (row 1 is headers)
                
                // Get all model items
                var allItems = GetAllModelItems(document);
                int totalItems = allItems.Count();
                int currentItem = 0;

                progressForm.SetProgress(0, totalItems, "Exporting model data...");

                // Export each item
                foreach (var item in allItems)
                {
                    try
                    {
                        ExportItemToExcel(item, worksheet, rowIndex);
                        rowIndex++;
                        currentItem++;
                        
                        // Update progress
                        if (currentItem % 100 == 0) // Update every 100 items
                        {
                            progressForm.SetProgress(currentItem, totalItems, 
                                $"Exporting item {currentItem} of {totalItems}...");
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error but continue with other items
                        System.Diagnostics.Debug.WriteLine($"Error exporting item: {ex.Message}");
                    }
                }

                // Auto-fit columns
                worksheet.Cells.AutoFitColumns();

                // Save the Excel file
                package.SaveAs(new FileInfo(excelFilePath));
            }

            return excelFilePath;
        }

        private void CreateHeaders(ExcelWorksheet worksheet)
        {
            string[] headers = {
                "Element Name",
                "Category/Class",
                "GUID",
                "X Coordinate",
                "Y Coordinate", 
                "Z Coordinate",
                "Level",
                "Properties"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1].Value = headers[i];
                worksheet.Cells[1, i + 1].Style.Font.Bold = true;
            }
        }

        private IEnumerable<ModelItem> GetAllModelItems(Document document)
        {
            // Get all model items from all models
            foreach (var model in document.Models)
            {
                var rootItems = model.RootItem.Children;
                foreach (var item in rootItems)
                {
                    yield return item;
                    // Recursively get all children
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
                // Recursively get grandchildren
                foreach (var grandchild in GetAllChildren(child))
                {
                    yield return grandchild;
                }
            }
        }

        private void ExportItemToExcel(ModelItem item, ExcelWorksheet worksheet, int rowIndex)
        {
            try
            {
                // Element Name
                worksheet.Cells[rowIndex, 1].Value = item.DisplayName ?? item.ClassDisplayName ?? "Unknown";

                // Category/Class Name
                worksheet.Cells[rowIndex, 2].Value = item.ClassDisplayName ?? "Unknown";

                // GUID
                worksheet.Cells[rowIndex, 3].Value = item.InstanceGuid.ToString();

                // Coordinates (center of bounding box)
                BoundingBox3D boundingBox = item.BoundingBox();
                Point3D center = boundingBox.Center;
                worksheet.Cells[rowIndex, 4].Value = center.X;
                worksheet.Cells[rowIndex, 5].Value = center.Y;
                worksheet.Cells[rowIndex, 6].Value = center.Z;

                // Level (if available)
                string level = GetLevel(item);
                worksheet.Cells[rowIndex, 7].Value = level;

                // Properties (flattened)
                string properties = GetFlattenedProperties(item);
                worksheet.Cells[rowIndex, 8].Value = properties;
            }
            catch (Exception ex)
            {
                // Set error values if something goes wrong
                worksheet.Cells[rowIndex, 1].Value = "Error";
                worksheet.Cells[rowIndex, 8].Value = $"Error: {ex.Message}";
            }
        }

        private string GetLevel(ModelItem item)
        {
            try
            {
                // Try to get level from properties
                var properties = item.PropertyCategories;
                
                foreach (var category in properties)
                {
                    foreach (var property in category.Properties)
                    {
                        if (property.DisplayName.ToLower().Contains("level") ||
                            property.Name.ToLower().Contains("level"))
                        {
                            return property.Value.ToDisplayString();
                        }
                    }
                }
                
                return "Not Available";
            }
            catch
            {
                return "Not Available";
            }
        }

        private string GetFlattenedProperties(ModelItem item)
        {
            try
            {
                var properties = new List<string>();
                var propertyCategories = item.PropertyCategories;

                foreach (var category in propertyCategories)
                {
                    foreach (var property in category.Properties)
                    {
                        try
                        {
                            string propertyName = $"{category.DisplayName}.{property.DisplayName}";
                            string propertyValue = property.Value.ToDisplayString();
                            
                            if (!string.IsNullOrEmpty(propertyValue))
                            {
                                properties.Add($"{propertyName}={propertyValue}");
                            }
                        }
                        catch
                        {
                            // Skip problematic properties
                            continue;
                        }
                    }
                }

                return string.Join("; ", properties);
            }
            catch
            {
                return "Error reading properties";
            }
        }
    }

    // Simple progress form for user feedback
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
