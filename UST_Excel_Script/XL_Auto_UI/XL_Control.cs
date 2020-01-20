using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XL_Auto_UI
{
    class XL_Control
    {
        private string _originalFile;
        private string _newFile;

        private Excel.Application xl_App = new Excel.Application();
        private Excel._Workbook xl_Workbook;
        private Excel._Worksheet xl_Worksheet;

        public string Original_File
        {
            get { return _originalFile; }
            set { _originalFile = value; }
        }

        public Excel._Workbook XL_Workbook
        {
            get { return xl_Workbook; }
            set { }
        }

        public Excel._Worksheet XL_Worksheet
        {
            get { return xl_Worksheet; }
            set { }
        }

        public void Open(string file)
        {
            //xl_App.ScreenUpdating = false;
            xl_Workbook = xl_App.Workbooks.Open(file);

        }

        public void SaveAs()
        {
            _newFile = _originalFile.Replace(".xlsx", "_Copy.xlsx");
            xl_Workbook.SaveAs(_newFile);
            xl_Worksheet = xl_Workbook.ActiveSheet;
        }

        public void Save() { }

        //Makes Excel Application appear
        public void Show_App()
        {
            xl_App.Visible = true;
        }

        public int FindColumnToCheckAgainst(string colName, Excel._Worksheet ws)
        {
            xl_Worksheet = ws;
            Excel.Range last = xl_Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string lastColumn = ColumnIndexToColumnLetter(last.Column) + "1";
            Excel.Range range = xl_Worksheet.get_Range("A1", lastColumn);

            int colIndex = 1;

            foreach (Excel.Range r in range)
            {
                if (r.Text == colName)
                {
                    colIndex = r.Column;
                    break;
                }
            }
            return colIndex;
        }

        //Turns column int to Excel column alphanumeric format
        public string ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter;
        }
    }
}
