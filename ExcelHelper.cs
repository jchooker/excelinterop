using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.AxHost;
using System.Text.RegularExpressions;

namespace ExcelInteropLibrary
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExcelHelper
    {
        private Excel.Application excelApp;

        public ExcelHelper() 
        {
            excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");

            excelApp.WorkbookOpen += ExcelApp_WorkbookOpen;
            excelApp.WorkbookBeforeClose += ExcelApp_WorkbookBeforeClose;
        }

        private void ExcelApp_WorkbookOpen(Excel.Workbook Wb)
        {
            // Handle WorkbookOpen event
        }

        private void ExcelApp_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            // Handle WorkbookBeforeClose event
        }

        public void ShowInsertTextForm()
        {
            // Display the InsertTextForm
            InsertTextForm form = new InsertTextForm();
            form.ShowDialog();
        }
    }
}
