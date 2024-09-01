using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInteropLibrary
{
    public class InsertTextForm : Form
    {
        public string userDeleteNameInput;
        private string addPattern;
        private string deletePattern;

        private System.Windows.Forms.TextBox textBox;
        private System.Windows.Forms.Button submitButton;
        private System.Windows.Forms.TextBox textToDelete;
        private System.Windows.Forms.Button deleteButton;

        private Excel.Application excelApp;

        public string InsertText { get; private set; }

        public InsertTextForm()
        {
            this.Text = "Enter Text to Insert or Delete";
            this.Size = new System.Drawing.Size(400, 200);

            textBox = new System.Windows.Forms.TextBox()
            {
                Location = new System.Drawing.Point(15, 20),
                Width = 250
            };

            //calc position for second textbox+btn
            int padding = 10;
            int secondTextBoxGroupTop = textBox.Top + textBox.Height + padding;
            int firstButtonLeft = textBox.Right + padding;

            submitButton = new System.Windows.Forms.Button()
            {
                Text = "Insert entered text",
                Location = new System.Drawing.Point(firstButtonLeft, 20),
                DialogResult = DialogResult.OK
            };
            submitButton.Click += SubmitButton_Click;

            textToDelete = new System.Windows.Forms.TextBox()
            {
                Location = new System.Drawing.Point(15, secondTextBoxGroupTop),
                Width = 250
            };

            int secondButtonLeft = textToDelete.Right + padding;

            deleteButton = new System.Windows.Forms.Button()
            {
                Text = "Delete the entered text from column",
                Location = new System.Drawing.Point(secondButtonLeft, secondTextBoxGroupTop),
                DialogResult = DialogResult.OK
            };
            deleteButton.Click += DeleteButton_Click;

            this.Controls.Add(textBox);
            this.Controls.Add(submitButton);
            this.Controls.Add(textToDelete);
            this.Controls.Add(deleteButton);

            this.FormClosing += InsertTextForm_FormClosing;

            excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            //Attach new workbookdeactivate event
            excelApp.WorkbookDeactivate += Application_WorkbookDeactivate;
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            string toDelete = textToDelete.Text;
            if (string.IsNullOrWhiteSpace(toDelete) || toDelete.Length < 4)
            {
                MessageBox.Show("Please enter a phrase of > 3 letters to delete.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //else if () <--other edge cases?

            //init delete pattern
            InitializeDeletePattern(toDelete);
            //bool deletionOccurred = Globals.ThisAddIn.PerformDeleteOperation();
            DeleteTextInSelectedColumn(toDelete);
            Close();
        }

        private void SubmitButton_Click(Object sender, EventArgs e)
        {
            string toInsert = textBox.Text;
            if (string.IsNullOrWhiteSpace(toInsert) || toInsert.Length < 4)
            {
                MessageBox.Show("Please enter a phrase of > 3 letters to insert.", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            InitializeInsertPattern();
            InsertTextInSelectedColumn(toInsert);
            Close();
        }

        private void InsertTextForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Globals.ThisAddIn.Application.WorkbookDeactivate -= Application_WorkbookDeactivate;
        }

        private void Application_WorkbookDeactivate(Excel.Workbook Wb)
        {
            if (this.Visible) this.Close();
        }

        public void InitializeInsertPattern()
        {
            addPattern = @"\s*(\;|\,)(?=\S| (?=[A-Z]))";
        }

        public void InitializeDeletePattern(string toDelete)
        {
            deletePattern = $@"(\;|\,)\s*{Regex.Escape(toDelete)}\s*";
        }

        private bool InsertTextInSelectedColumn(string insertText)
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet activeSheet = ((Excel.Worksheet)(excelApp.ActiveSheet));
            Excel.Range selectedRange = (excelApp.Selection) as Excel.Range;

            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range of cells.", "Cell selection error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //only one column at a time for v0.1
            if (selectedRange.Columns.Count > 1)
            {
                MessageBox.Show("Please only select one column at a time!", "Column range > 1 column", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            try
            {
                Regex reg = new Regex(addPattern);

                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value != null)
                    {
                        string cellValue = cell.Value.ToString();

                        if (!cellValue.Contains(insertText))
                        {
                            string newValue = reg.Replace(cellValue, match =>
                            {
                                return match.Value + " " + insertText + "; ";
                            }, 1);

                            if (newValue == cellValue)
                            {
                                newValue = cellValue + "; " + insertText;
                            }

                            cell.Value = newValue;
                        }
                    }
                }
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while reading column", "Column read error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }


        }

        private bool DeleteTextInSelectedColumn(string searchText)
        //this has to delete " ; name" or ",name" or " , name"
        {
            Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            Excel.Worksheet activeSheet = ((Excel.Worksheet)(excelApp.ActiveSheet));
            Excel.Range selectedRange = (excelApp.Selection) as Excel.Range;

            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range of cells.", "Cell selection error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            //only one column at a time for v0.1
            if (selectedRange.Columns.Count > 1)
            {
                MessageBox.Show("Please only select one column at a time!", "Column range > 1 column", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            bool deletionOccurred = false;
            int numInstances = 0;

            try
            {
                Regex reg = new Regex(searchText);

                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value != null)
                    {

                        string cellValue = cell.Value.ToString();
                        string newCellValue = Regex.Replace(cellValue, deletePattern, "");

                        if (newCellValue != cellValue)
                        {
                            cell.Value = newCellValue;
                            numInstances++;
                            deletionOccurred = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while reading column", "Column read error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!deletionOccurred) MessageBox.Show("The entered phrase was not found in the selected column.", "Phrase not found", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else MessageBox.Show($"Deletion Successful: {numInstances} instances of phrase found", "Deletion complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return deletionOccurred;

        }
    }
}
