using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XL_Auto_UI
{
    class UST_WeeklyActive
    {
        XL_Control _xlControl;

        Dictionary<string, int> _partnerCenterStats;
        Dictionary<string, int> _devCenterStats;
        Dictionary<string, int> _sellerDashStats;
        int _blankProductCategory = -1;

        //constants used for Switch method
        private const string _devCenter = "Universal Store";
        private const string _partnerCenter = "Other";
        private const string _sellerDash = "Seller Dashboard";

        private string _fileName;

        //private Excel.Application xl_App;
        private Excel._Workbook xl_Workbook;
        private Excel._Worksheet xl_Worksheet;

        public string FileName
        {
            get { return _fileName; }
            set { _fileName = value; }
        }

        //Adds the keys (Issue Code 4) to the appropiate dictionary and sets the int values to 0
        private void AddPartnerCenterKeys()
        {
            _partnerCenterStats = new Dictionary<string, int>
            {
                {"Email Ownership", 0 },
                {"Domain Validation", 0 },
                {"Business Individual Association Investigation", 0 },
                {"Business Investigation", 0 },
                {"Vetting Status Inquiry", 0 },
                {"Business Status", 0 },
                {"Update User Information", 0 },
                {"Request Validation", 0 },
                {"Business Rime Check", 0 },
                {"Business Market Profile", 0 },
                {"Business Classification", 0 },
                {"Business Government Controlled List", 0 },
                {"Individual Government Controlled List", 0 },
                {"Verisign", 0 },
                {"Symantec Code Signing Cert", 0 },
                {"SIVS Validation", 0 },
                {"Other", 0 },
                {"Blank/Investigate", 0 }
            };
        }

        //Adds the keys (Issue Code 4) to the appropiate dictionary and sets the int values to 0
        private void AddDevCenterKeys()
        {
            _devCenterStats = new Dictionary<string, int>
            {
                {"Email Ownership", 0 },
                {"Domain Validation", 0 },
                {"Business Individual Association Investigation", 0 },
                {"Business Investigation", 0 },
                {"Vetting Status Inquiry", 0 },
                {"Business Status", 0 },
                {"Update User Information", 0 },
                {"Request Validation", 0 },
                {"Business Rime Check", 0 },
                {"Business Market Profile", 0 },
                {"Business Classification", 0 },
                {"Business Government Controlled List", 0 },
                {"Individual Government Controlled List", 0 },
                {"Verisign", 0 },
                {"Symantec Code Signing Cert", 0 },
                {"SIVS Validation", 0 },
                {"Other", 0 },
                {"Blank/Investigate", 0 }
            };
        }

        //Adds the keys (Issue Code 4) to the appropiate dictionary and sets the int values to 0
        private void AddSellerDashKeys()
        {
            _sellerDashStats = new Dictionary<string, int>
            {
                {"Email Ownership", 0 },
                {"Domain Validation", 0 },
                {"Business Individual Association Investigation", 0 },
                {"Business Investigation", 0 },
                {"Vetting Status Inquiry", 0 },
                {"Business Status", 0 },
                {"Update User Information", 0 },
                {"Request Validation", 0 },
                {"Business Rime Check", 0 },
                {"Business Market Profile", 0 },
                {"Business Classification", 0 },
                {"Business Government Controlled List", 0 },
                {"Individual Government Controlled List", 0 },
                {"Verisign", 0 },
                {"Symantec Code Signing Cert", 0 },
                {"SIVS Validation", 0 },
                {"Other", 0 },
                { "PVS", 0 },
                {"Blank/Investigate", 0 }
            };
        }

        public UST_WeeklyActive()
        {
            _xlControl = new XL_Control();

            AddPartnerCenterKeys();
            AddDevCenterKeys();
            AddSellerDashKeys();
        }

        public void GrabData()
        {
            int colIndex = _xlControl.FindColumnToCheckAgainst("ProductCategory (DEVAB)", xl_Worksheet);
            int colIndex2 = _xlControl.FindColumnToCheckAgainst("ASDLevel4_DN", xl_Worksheet);

            Excel.Range last = xl_Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastRow = last.Row;
            Excel.Range range = xl_Worksheet.UsedRange;

            int i;

            for (i = 2; i < lastRow + 1; i++)
            {
                Excel.Range cell = (Excel.Range)range[i, colIndex];
                Excel.Range cell2 = (Excel.Range)range[i, colIndex2];

                if (!string.IsNullOrWhiteSpace(cell.Text))
                {
                    switch (cell.Value2.ToString())
                    {
                        case _partnerCenter:
                            if (_partnerCenterStats.ContainsKey(cell2.Value2.ToString()))
                            {
                                int keyValue = _partnerCenterStats[cell2.Value2.ToString()];
                                keyValue++;
                                _partnerCenterStats[cell2.Value2.ToString()] = keyValue;
                            }
                            else
                            {
                                int keyValue = _partnerCenterStats["Blank/Investigate"];
                                keyValue++;
                                _partnerCenterStats["Blank/Investigate"] = keyValue;
                            }
                            break;
                        case _devCenter:
                            if (_devCenterStats.ContainsKey(cell2.Value2.ToString()))
                            {
                                int keyValue = _devCenterStats[cell2.Value2.ToString()];
                                keyValue++;
                                _devCenterStats[cell2.Value2.ToString()] = keyValue;
                            }
                            else
                            {
                                int keyValue = _devCenterStats["Blank/Investigate"];
                                keyValue++;
                                _devCenterStats["Blank/Investigate"] = keyValue;
                            }
                            break;
                        case _sellerDash:
                            if (_sellerDashStats.ContainsKey(cell2.Value2.ToString()))
                            {
                                int keyValue = _sellerDashStats[cell2.Value2.ToString()];
                                keyValue++;
                                _sellerDashStats[cell2.Value2.ToString()] = keyValue;
                            }
                            else
                            {
                                int keyValue = _sellerDashStats["Blank/Investigate"];
                                keyValue++;
                                _sellerDashStats["Blank/Investigate"] = keyValue;
                            }
                            break;
                    }
                }
                else
                {
                    int keyValue = _blankProductCategory;
                    keyValue++;
                    _blankProductCategory = keyValue;
                }
            }
        }

        public void WriteDataToTxt()
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C: \Users\Rafael\Desktop\UST_Weekly_ActiveStats.txt"))
            {
                if (_blankProductCategory > 0)
                {
                    file.WriteLine("Please note that {0} rows were found with no Product Category value listed." + Environment.NewLine +
                        Environment.NewLine, _blankProductCategory);
                }
                file.WriteLine(Environment.NewLine + "PARTNER CENTER (OTHER)" + Environment.NewLine + "====================");
                foreach (var s in _partnerCenterStats)
                {
                    file.WriteLine("{0}: {1}", s.Key, s.Value);
                }

                file.WriteLine(Environment.NewLine + "DEV CENTER (UST)" + Environment.NewLine + "====================");
                foreach (var s in _devCenterStats)
                {
                    file.WriteLine("{0}: {1}", s.Key, s.Value);
                }

                file.WriteLine(Environment.NewLine + "SELLER DASHBOARD" + Environment.NewLine + "====================");
                foreach (var s in _sellerDashStats)
                {
                    file.WriteLine("{0}: {1}", s.Key, s.Value);
                }
            }
        }

    }
}
