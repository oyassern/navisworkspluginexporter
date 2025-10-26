using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using OfficeOpenXml;

namespace NavisExcelExporter
{
    [Plugin("NavisExcelExporter.FullModelExportPlugin", "Omar", DisplayName = "Export Full Model to Excel")]
    [AddInPlugin(AddInLocation.AddIn)]
    public class FullModelExportPlugin : AddInPlugin
    {
        private static readonly HttpClient _httpClient = new HttpClient();
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

                using (var selectionForm = new SelectionForm(document))
                {
                    var result = selectionForm.ShowDialog();
                    if (result != DialogResult.OK)
                    {
                        return 0; // user cancelled
                    }

                    var selectedItems = selectionForm.GetCheckedItems().ToList();
                    if (selectedItems.Count == 0)
                    {
                        MessageBox.Show("No items selected for export.", "Nothing Selected", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return 0;
                    }

                    using (var progressForm = new ProgressForm())
                    {
                        progressForm.Show();
                        System.Windows.Forms.Application.DoEvents();

                        string excelFilePath = ExportModelToExcel(selectedItems, progressForm);
                        progressForm.Close();

                        // Run AI automation immediately after export if requested
                        if (selectionForm.StartAutomation)
                        {
                            try
                            {
                                SendToN8nAsync(excelFilePath, selectionForm.StartDate).GetAwaiter().GetResult();
                                MessageBox.Show("Excel file uploaded to n8n webhook successfully.", "Automation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception uploadEx)
                            {
                                MessageBox.Show($"Failed to upload to n8n webhook.\n\n{uploadEx.Message}", "Automation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }

                        MessageBox.Show($"Model data exported successfully!\n\nFile saved to:\n{excelFilePath}",
                            "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
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

		private string ExportModelToExcel(IEnumerable<ModelItem> selectedRoots, ProgressForm progressForm)
		{
			string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
			string excelFilePath = Path.Combine(desktopPath, "NavisModelData.xlsx");

            if (File.Exists(excelFilePath))
            {
                File.Delete(excelFilePath);
            }

            // PASS 1: Collect all items (from selection) and their properties to determine all unique columns
            progressForm.SetProgress(0, 100, "Pass 1: Scanning properties...");
            var allItems = GetAllModelItems(selectedRoots).ToList();
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

		private async Task SendToN8nAsync(string excelPath, DateTime? startDate)
		{
			if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
			{
				throw new FileNotFoundException("Excel file not found for upload.", excelPath);
			}

			const string webhookUrl = "http://localhost:5678/webhook-test/nwx-export";

			using (var form = new MultipartFormDataContent())
			using (var stream = File.OpenRead(excelPath))
			using (var fileContent = new StreamContent(stream))
			{
				fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
				form.Add(fileContent, "data", Path.GetFileName(excelPath));

				if (startDate.HasValue)
				{
					// Send date as ISO date (yyyy-MM-dd)
					var dateContent = new StringContent(startDate.Value.ToString("yyyy-MM-dd"));
					form.Add(dateContent, "startDate");
				}

				var response = await _httpClient.PostAsync(webhookUrl, form).ConfigureAwait(false);
				response.EnsureSuccessStatusCode();
			}
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

        private IEnumerable<ModelItem> GetAllModelItems(IEnumerable<ModelItem> roots)
        {
            foreach (var r in roots)
            {
                yield return r;
                foreach (var c in GetAllChildren(r))
                {
                    yield return c;
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

    public class SelectionForm : Form
    {
        private readonly Document _document;
        private readonly TreeView _tree;
        private readonly Button _okButton;
        private readonly Button _cancelButton;
        private readonly CheckBox _automationCheck;
        private readonly DateTimePicker _startDatePicker;
        private readonly Label _startDateLabel;

        public SelectionForm(Document document)
        {
            _document = document;
            Text = "Select Items to Export";
            Size = new System.Drawing.Size(500, 480);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                CheckBoxes = true
            };

            _okButton = new Button { Text = "Export", AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Top };
            _cancelButton = new Button { Text = "Cancel", AutoSize = true, Anchor = AnchorStyles.Right | AnchorStyles.Top };
            _automationCheck = new CheckBox { Text = "Run AI Automation to generate a time plan", AutoSize = true, Checked = false, Anchor = AnchorStyles.Left | AnchorStyles.Top };

            _startDateLabel = new Label { Text = "Project start:", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top };
            _startDatePicker = new DateTimePicker { Format = DateTimePickerFormat.Short, Width = 120, Anchor = AnchorStyles.Left | AnchorStyles.Top };
            _startDatePicker.Value = DateTime.Today;
            _startDatePicker.Enabled = _automationCheck.Checked;

            var bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 76,
                ColumnCount = 2,
                RowCount = 1,
                Padding = new Padding(10),
            };
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // Left side: 2-row table (row0: checkbox, row1: label + picker aligned)
            var leftTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            leftTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            leftTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            leftTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            leftTable.Controls.Add(_automationCheck, 0, 0);
            leftTable.SetColumnSpan(_automationCheck, 2);
            _automationCheck.Margin = new Padding(0, 4, 8, 4);

            _startDateLabel.Margin = new Padding(0, 4, 8, 4);
            _startDatePicker.Margin = new Padding(0, 0, 8, 0);
            leftTable.Controls.Add(_startDateLabel, 0, 1);
            leftTable.Controls.Add(_startDatePicker, 1, 1);

            // Right side: buttons aligned right
            var rightFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            rightFlow.Controls.Add(_okButton);
            rightFlow.Controls.Add(new Label { Text = "  ", AutoSize = true, Width = 8 }); // spacer
            rightFlow.Controls.Add(_cancelButton);

            bottomLayout.Controls.Add(leftTable, 0, 0);
            bottomLayout.Controls.Add(rightFlow, 1, 0);

            Controls.Add(_tree);
            Controls.Add(bottomLayout);

            Load += SelectionForm_Load;
            _okButton.Click += (s, e) => DialogResult = DialogResult.OK;
            _cancelButton.Click += (s, e) => DialogResult = DialogResult.Cancel;
            AcceptButton = _okButton;
            CancelButton = _cancelButton;
            _tree.AfterCheck += Tree_AfterCheck;
            _tree.BeforeExpand += Tree_BeforeExpand;
            _automationCheck.CheckedChanged += (s, e) => { _startDatePicker.Enabled = _automationCheck.Checked; };
            _startDatePicker.Enabled = _automationCheck.Checked;
        }

        public bool StartAutomation => _automationCheck.Checked;
        public DateTime? StartDate => _automationCheck.Checked ? _startDatePicker.Value.Date : (DateTime?)null;

        private void SelectionForm_Load(object sender, EventArgs e)
        {
            BuildTree();
        }

        private void BuildTree()
        {
            _tree.BeginUpdate();
            _tree.Nodes.Clear();
            foreach (var model in _document.Models)
            {
                var modelNode = new TreeNode(model.FileName) { Tag = model.RootItem };
                foreach (var child in model.RootItem.Children)
                {
                    modelNode.Nodes.Add(BuildNodeShallow(child));
                }
                _tree.Nodes.Add(modelNode);
            }
            _tree.EndUpdate();
        }

        private const string PlaceholderTag = "__placeholder__";

        private TreeNode BuildNodeShallow(ModelItem item)
        {
            var node = new TreeNode(item.DisplayName ?? item.ClassDisplayName ?? "(Item)") { Tag = item };
            if (HasChildItems(item))
            {
                node.Nodes.Add(new TreeNode("…") { Tag = PlaceholderTag });
            }
            return node;
        }

        private bool HasChildItems(ModelItem item)
        {
            try { return item.Children != null && item.Children.Any(); }
            catch { return false; }
        }

        private void Tree_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.Nodes.Count == 1 && Equals(e.Node.Nodes[0].Tag, PlaceholderTag) && e.Node.Tag is ModelItem mi)
            {
                e.Node.Nodes.Clear();
                foreach (var child in mi.Children)
                {
                    e.Node.Nodes.Add(BuildNodeShallow(child));
                }
            }
        }

        private bool _suppressAfterCheck = false;
        private void Tree_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_suppressAfterCheck) return;
            try
            {
                _suppressAfterCheck = true;
                // Propagate check state to children for convenience
                foreach (TreeNode child in e.Node.Nodes)
                {
                    child.Checked = e.Node.Checked;
                }
            }
            finally
            {
                _suppressAfterCheck = false;
            }
        }

        public IEnumerable<ModelItem> GetCheckedItems()
        {
            foreach (TreeNode root in _tree.Nodes)
            {
                foreach (var item in GetCheckedItemsRecursive(root))
                {
                    yield return item;
                }
            }
        }

        private IEnumerable<ModelItem> GetCheckedItemsRecursive(TreeNode node)
        {
            if (node.Checked && node.Tag is ModelItem mi)
            {
                // If node is checked, treat it as selected root and do not traverse deeper
                yield return mi;
                yield break;
            }
            foreach (TreeNode child in node.Nodes)
            {
                foreach (var c in GetCheckedItemsRecursive(child))
                {
                    yield return c;
                }
            }
        }
    }
}
