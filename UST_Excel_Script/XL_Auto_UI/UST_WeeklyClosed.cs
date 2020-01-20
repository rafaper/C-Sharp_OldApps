using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XL_Auto_UI
{
    class UST_WeeklyClosed
    {
        XL_Control _xlControl;

        private Dictionary<string, bool> _relevantColumns;
        private Dictionary<string, bool> _irrelevantRows;

        private Excel.Application xl_App;
        private Excel._Workbook xl_Workbook;
        private Excel._Worksheet xl_Worksheet;

        private void AddRelevantColumns()
        {
            _relevantColumns = new Dictionary<string, bool>
            {
                {"Service Request", false },
                {"Status", false },
                {"Owner Name", false },
                {"Subject", false },
                {"Opened Date", false },
                {"Closed Date", false },
                {"SLA Hours", false },
                {"SLA Days", false },
                {"Customer Country", false },
                {"Vetting Tier_DN", false },
                {"Resolution [dev]", false },
                {"ASDLevel1_DN", false },
                {"ASDLevel2_DN", false },
                {"ASDLevel3_DN", false },
                {"ASDLevel4_DN", false },
                {"ProductCategory (DEVAB)", false },
                {"PUID_DN", false }
            };
        }

        public UST_WeeklyClosed()
        {

        }

        private void AddirrelevantRows()
        {
            _irrelevantRows = new Dictionary<string, bool>
            {
                {"John Doe",false }
            };
        }
        public void Delete_Irrelevant_Columns()
        {
            AddRelevantColumns();

            Excel.Range last = xl_Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            string lastColumn = _xlControl.ColumnIndexToColumnLetter(last.Column) + "1";
            Excel.Range range = xl_Worksheet.get_Range("A1", lastColumn);

            int i;

            // iterate range
            for (i = last.Column; i > 0; i--)
            {
                Excel.Range cell = (Excel.Range)range[1, i];

                // check condition:
                if (!_relevantColumns.ContainsKey(cell.Text))
                {
                    // if not match, delete and shift remaining cells left:
                    cell.EntireColumn.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                }
            }
        }

        public void Delete_Irrelevant_Rows()
        {
            AddirrelevantRows();

            Excel.Range last = xl_Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = xl_Worksheet.UsedRange;

            int i;
            int columnIndex = _xlControl.FindColumnToCheckAgainst("Owner Name", xl_Worksheet);

            // iterate range
            for (i = last.Row; i > 0; i--)
            {
                Excel.Range cell = (Excel.Range)range[i, columnIndex];

                // check condition:
                if (_irrelevantRows.ContainsKey(cell.Text) || string.IsNullOrWhiteSpace(cell.Text))
                {
                    // if match, delete and shift remaining cells up:
                    cell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
            }
        }

        public void DeleteIfOutsideDateRange(DateTime start, DateTime end)
        {
            Excel.Range last = xl_Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = xl_Worksheet.UsedRange;

            int i;
            int columnIndex = _xlControl.FindColumnToCheckAgainst("Closed Date", xl_Worksheet);

            // iterate range
            for (i = last.Row; i > 1; i--)
            {
                Excel.Range cell = (Excel.Range)range[i, columnIndex];

                if (!string.IsNullOrWhiteSpace(cell.Text))
                {
                    double xlDate = double.Parse(cell.Value2.ToString());
                    var convDT = DateTime.FromOADate(xlDate);

                    // check condition:
                    // if match, delete and shift remaining cells up:
                    if (convDT.Date < start.Date)
                    {
                        cell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    }
                    else if (convDT.Date > end.Date)
                    {
                        cell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    }
                }
            }
        }
    }
}
