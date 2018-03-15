using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;



namespace MTools2013
{
    [ComVisible(true)]
    public class mtoolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public mtoolsRibbon()
        {
            
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MTools2013.mtoolsRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        
        public void insertRow(Office.IRibbonControl control)
        {
            Excel.Workbook currentWB = Globals.ThisAddIn.GetActiveWorkbook();
            Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
            Excel.Application activeApp = Globals.ThisAddIn.GetActiveApp();

            //activWorkSheet.Columns.ClearFormats();
            //activWorkSheet.Rows.ClearFormats();
    
            int totalRowNum = activWorkSheet.UsedRange.Rows.Count;
            int totalColNum = activWorkSheet.UsedRange.Columns.Count;

            int lastRowNum = activWorkSheet.UsedRange.Row + totalRowNum-1;
            int lastColNum = activWorkSheet.UsedRange.Column + totalColNum-1;

            Excel.Range rangeSelected = (Excel.Range)activeApp.Selection;

           
            try
            {
                Excel.Range ColumnRange = activeApp.InputBox("Please select the \'ID\' column of select range from \'ID\' column: \n ", "Insert Row Below Similar ID", rangeSelected.Address, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

                if (ColumnRange == null)
                {
                    MessageBox.Show("Please select Range from worksheet");
                }else
                {
                    int TargetColno = ColumnRange.Column;

                    int FirstRowofSelection = ColumnRange.Row;
                    int LastRowofSelection = ColumnRange.Rows.Count + FirstRowofSelection - 1;

                    if (LastRowofSelection < lastRowNum)
                    {
                        lastRowNum = LastRowofSelection;
                    }


                    for (int i = FirstRowofSelection; i < lastRowNum; i++)
                    {
                        Excel.Range firstCell = (Excel.Range)activWorkSheet.Cells[i, TargetColno];
                        Excel.Range secondCell = (Excel.Range)activWorkSheet.Cells[i + 1, TargetColno];


                        if (firstCell.Value == null || secondCell.Value == null)
                        {
                            i = i + 1;
                        }
                        else
                        {
                            if (secondCell.Text != firstCell.Text)
                            {
                                if(i==1)
                                {
                                    
                                }
                                else
                                {
                                    Excel.Range newBlankline = (Excel.Range)activWorkSheet.Rows[i + 1];
                                    newBlankline.Insert();

                                    lastRowNum++;
                                    i++;
                                }
                            }

                        }
                    }
                }

                
            }catch
            {
                
            }
            
        }


        public void GetLowestValue(Office.IRibbonControl control)
        {
            MessageBox.Show("Get Lowest Point");
        }

        static string ColumnIndexToColumnLetter(int colIndex)
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

        public void DataInventory(Office.IRibbonControl control)
        {
            Excel.Workbook currentWB = Globals.ThisAddIn.GetActiveWorkbook();
            Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
            Excel.Application activeApp = Globals.ThisAddIn.GetActiveApp();
            
            int totalRowNum = activWorkSheet.UsedRange.Rows.Count;
            int totalColNum = activWorkSheet.UsedRange.Columns.Count;

            int lastRowNum = activWorkSheet.UsedRange.Row + totalRowNum-1;
            int lastColNum = activWorkSheet.UsedRange.Column + totalColNum-1;

            Excel.Range rangeSelected = (Excel.Range)activeApp.Selection;

            List<int> yearList = new List<int>();
            List<int> monthList = new List<int>();
            List<int> inventoryList = new List<int>();
            List<int> dataAvailailityList = new List<int>();

            try
            {
                reselect:

                Excel.Range ColumnRange = activeApp.InputBox("Please select the \'Date\' column and \'Value\' column: \n ", "Select Date and Value", rangeSelected.Address, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                
                if (ColumnRange.Columns.Count == 2)
                {
                    int DateColumn = ColumnRange.Column;
                    int ValueColumn = ColumnRange.Column + 1;

                    int FirstRowofSelection = ColumnRange.Row;
                    int LastRowofSelection = ColumnRange.Rows.Count + FirstRowofSelection - 1;

                    if (LastRowofSelection < lastRowNum)
                    {
                        lastRowNum = LastRowofSelection;
                    }

                    string dateColNameRangeFrom = ColumnIndexToColumnLetter(DateColumn) + FirstRowofSelection.ToString();
                    string dateColNameRangeTo = ColumnIndexToColumnLetter(DateColumn) + lastRowNum.ToString();

                    string valueColNameRangeFrom = ColumnIndexToColumnLetter(ValueColumn) + FirstRowofSelection.ToString();
                    string valueColNameRangeTo = ColumnIndexToColumnLetter(ValueColumn) + lastRowNum.ToString();

                    
                    Excel.Range DateRange = ((Excel.Range)activWorkSheet.Range[dateColNameRangeFrom, dateColNameRangeTo]);
                    DateRange.EntireColumn.NumberFormat = "DD-MMM-YYYY";
                    DateRange.EntireColumn.ColumnWidth = 15;
                    Excel.Range ValueRange = ((Excel.Range)activWorkSheet.Range[valueColNameRangeFrom, valueColNameRangeTo]);
                    ValueRange.EntireColumn.ColumnWidth = 8.43;
                    
                    int currentYear = 0;
                    int currentMonth = 0;
                    int currentDay = 0;
                    int dataAvailaility = 0;

                    for (int row = 1; row <= DateRange.Rows.Count; row++)
                    {
                        DateTime dateValue;
                        bool isDate = DateTime.TryParse(DateRange[row].Text, out dateValue);
                        if (isDate)
                        {
                            currentYear = dateValue.Year;
                            currentMonth = dateValue.Month;
                            currentDay = dateValue.Day;

                            double value;
                            bool isValue = double.TryParse(ValueRange[row].Text, out value);
                            if (isValue)
                            {
                                dataAvailaility = 1;
                            }
                            break;
                        }

                    }

                    for (int row = 1; row <= DateRange.Rows.Count; row++)
                    {
                        DateTime dateValue;
                        bool isDate = DateTime.TryParse(DateRange[row].Text, out dateValue);
                        if (isDate)
                        {

                            if (currentYear == dateValue.Year)
                            {
                                if (currentMonth == dateValue.Month)
                                {

                                    double value;
                                    bool isValue = double.TryParse(ValueRange[row].Text, out value);
                                    if (isValue)
                                    {

                                        if (currentDay != dateValue.Day)
                                        {
                                            dataAvailaility = dataAvailaility + 1;
                                            currentDay = dateValue.Day;
                                        }
                                    }
                                    if (row == DateRange.Rows.Count)
                                    {
                                        yearList.Add(currentYear);
                                        monthList.Add(currentMonth);
                                        dataAvailailityList.Add(dataAvailaility);
                                        if (dataAvailaility > 20)
                                        {
                                            inventoryList.Add(1);
                                        }
                                        else
                                        {
                                            inventoryList.Add(0);
                                        }
                                        
                                    }

                                    
                                }
                                else
                                {
                                    yearList.Add(currentYear);
                                    monthList.Add(currentMonth);
                                    dataAvailailityList.Add(dataAvailaility);
                                    if (dataAvailaility > 20)
                                    {
                                        inventoryList.Add(1);
                                    }
                                    else
                                    {
                                        inventoryList.Add(0);
                                    }

                                    double value;
                                    bool isValue = double.TryParse(ValueRange[row].Text, out value);
                                    if (isValue)
                                    {
                                        dataAvailaility = 1;
                                    }
                                    else
                                    {
                                        dataAvailaility = 0;
                                    }
                                    currentDay = dateValue.Day;
                                    currentMonth = dateValue.Month;

                                    yearList.Add(currentYear);
                                    monthList.Add(currentMonth);
                                    dataAvailailityList.Add(dataAvailaility);
                                    if (dataAvailaility > 20)
                                    {
                                        inventoryList.Add(1);
                                    }
                                    else
                                    {
                                        inventoryList.Add(0);
                                    }

                                }

                            }
                            else
                            {
                                yearList.Add(currentYear);
                                monthList.Add(currentMonth);
                                dataAvailailityList.Add(dataAvailaility);
                                if (dataAvailaility > 20)
                                {
                                    inventoryList.Add(1);
                                }
                                else
                                {
                                    inventoryList.Add(0);
                                }

                                double value;
                                bool isValue = double.TryParse(ValueRange[row].Text, out value);
                                if (isValue)
                                {
                                    dataAvailaility = 1;
                                }
                                else
                                {
                                    dataAvailaility = 0;
                                }
                                currentDay = dateValue.Day;
                                currentMonth = dateValue.Month;
                                currentYear = dateValue.Year;

                                yearList.Add(currentYear);
                                monthList.Add(currentMonth);
                                dataAvailailityList.Add(dataAvailaility);
                                if (dataAvailaility > 20)
                                {
                                    inventoryList.Add(1);
                                }
                                else
                                {
                                    inventoryList.Add(0);
                                }
                            }

                            
                        }

                    }
                    
                    if (yearList.Count > 0)
                    {
                       
                        int noOfYear = 0;
                        List<int> yearIndex = new List<int>();
                        List<int> Years = new List<int>();

                        if (yearList.Count == 1)
                        {
                            yearIndex.Add(0);
                            Years.Add(yearList.ElementAt(0));
                        }
                        else if (yearList.Count > 1)
                        {
                            for (int i = 0; i < yearList.Count-1 ; i++)
                            {
                            
                                if (yearList.ElementAt(i) != yearList.ElementAt(i + 1))
                                {
                                    yearIndex.Add(i);
                                    Years.Add(yearList.ElementAt(i));
                                }
                                
                                if (i+1 == yearList.Count - 1)
                                {
                                    yearIndex.Add(i+1);
                                    Years.Add(yearList.ElementAt(i+1));
                                }
                                //MessageBox.Show(noOfYear.ToString());
                            }
                        }
                        noOfYear = Years.Count;
                        
                        int lengthYear = yearList.Count;
                        int lengthMonth = monthList.Count;
                        int lengthInventory = inventoryList.Count;
                        int lengthDataAvailability = dataAvailailityList.Count;


                        int[,] inventoryArray = new int[noOfYear*12,4];

                        
                        int monthConter = 0;
                            for (int i = 0; i < noOfYear; i++)
                            {
                                for (int j = monthConter; j <= yearIndex.ElementAt(i); j++)
                                {
                                    
                                    inventoryArray[i * 12 + monthList.ElementAt(j) - 1, 0] = yearList.ElementAt(j);
                                    inventoryArray[i * 12 + monthList.ElementAt(j) - 1, 1] = monthList.ElementAt(j);
                                    inventoryArray[i * 12 + monthList.ElementAt(j) - 1, 2] = inventoryList.ElementAt(j);
                                    inventoryArray[i * 12 + monthList.ElementAt(j) - 1, 3] = dataAvailailityList.ElementAt(j);

                                }
                                monthConter = yearIndex.ElementAt(i)+1;
                                
                                
                            }
                       
                        string outputCellName = ColumnIndexToColumnLetter(lastColNum + 2) + (FirstRowofSelection + 1).ToString();

                        Excel.Range outputCell = activeApp.InputBox("Data inventory is ready. Please select a cell to get the output: \n ", "Data Inventory Ready", outputCellName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
                        int outputColumn = outputCell.Column;
                        int outputRow = outputCell.Row;

                        activWorkSheet.Cells[outputRow, outputColumn] = "Year";
                        activWorkSheet.Cells[outputRow + 1, outputColumn] = "Month";
                        activWorkSheet.Cells[outputRow + 2, outputColumn] = "Inventory:";
                        activWorkSheet.Cells[outputRow + 3, outputColumn] = "Availability:";

                        for (int i = 0; i < noOfYear*12; i++)
                        {
                            activWorkSheet.Cells[outputRow, outputColumn + i + 1] = inventoryArray[i,0];
                            activWorkSheet.Cells[outputRow + 1, outputColumn + i + 1] = inventoryArray[i, 1];
                            activWorkSheet.Cells[outputRow + 2, outputColumn + i + 1] = inventoryArray[i, 2];
                            activWorkSheet.Cells[outputRow + 3, outputColumn + i + 1] = inventoryArray[i, 3];
                        }
                        
                        activWorkSheet.Cells[outputRow, outputColumn].EntireColumn.ColumnWidth = 11;

                        string inventoryCellColFrom = ColumnIndexToColumnLetter(outputColumn + 1) + outputRow.ToString();
                        string inventoryCellColTo = ColumnIndexToColumnLetter(outputColumn + noOfYear * 12) + (outputRow + 3).ToString();

                        Excel.Range inventoryRange = ((Excel.Range)activWorkSheet.Range[inventoryCellColFrom, inventoryCellColTo]);
                        inventoryRange.Borders.Color = 0x00000000;
                        inventoryRange.EntireColumn.ColumnWidth = 2.43;
                        inventoryRange.EntireRow.RowHeight = 15;


                        for (int i = 0; i < noOfYear * 12; i++)
                        {
                            int monthFill = (i + 1) % 12;
                            if (monthFill == 0) { monthFill = 12; }
                            if (activWorkSheet.Cells[outputRow + 1, outputColumn + i + 1].Value == 0)
                            {
                                activWorkSheet.Cells[outputRow , outputColumn + i + 1] = Years.ElementAt(i/12);
                                activWorkSheet.Cells[outputRow + 1, outputColumn + i + 1] = monthFill;
                                activWorkSheet.Cells[outputRow + 2, outputColumn + i + 1] = 0;
                                activWorkSheet.Cells[outputRow + 3, outputColumn + i + 1] = 0;
                            }

                        }

                        string inventoryCellColourFrom = ColumnIndexToColumnLetter(outputColumn + 1) + (outputRow + 2).ToString();
                        string inventoryCellColourTo = ColumnIndexToColumnLetter(outputColumn + noOfYear * 12) + (outputRow + 2).ToString();
                        
                        for(int n =  outputColumn + 1; n <= outputColumn + noOfYear * 12; n++)
                        {
                            if (activWorkSheet.Cells[outputRow + 2,  n].Value == 1)
                            {
                                activWorkSheet.Cells[outputRow + 2, n].Interior.Color = 0x00A9A9A9;
                            }
                        }
                        //Excel.FormatCondition format = (Excel.FormatCondition)(activWorkSheet.get_Range(inventoryCellColourFrom, inventoryCellColourTo).FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, 1, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                        //format.Interior.Color = 0x00A9A9A9;



                        for (int i = 0; i < noOfYear; i++)
                        {
                            activeApp.DisplayAlerts = false;
                            Excel.Range yearRange = ((Excel.Range)activWorkSheet.Range[ColumnIndexToColumnLetter(outputColumn + 1 + (i * 12)) + outputRow.ToString(), ColumnIndexToColumnLetter(outputColumn + ((i + 1) * 12)) + outputRow.ToString()]);
                            yearRange.Merge();
                            yearRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No data found for inventory!","No Data Found");
                    }

                    

                }
                else
                {
                    MessageBox.Show("Please select only two columns: First one Date and Second one Value.","Incorrect Input");
                    goto reselect;
                }
            }
            catch
            {

            }

        }


        public void xns11(Office.IRibbonControl control)
        {
            var ownerWindow = new Win32Window(Globals.ThisAddIn.Application.Hwnd);

            Prompt xnsForm = new Prompt();
            xnsForm.TopLevel = true;
            xnsForm.Show(ownerWindow);

            xnsForm.okButtonActionEvent += xnsForm_okButtonActionEvent;

        }

        private void xnsForm_okButtonActionEvent(string m1, string m2, string m3, string m4, string m5, string m6, string m7, string m8)
        {
            Excel.Workbook currentWB = Globals.ThisAddIn.GetActiveWorkbook();
            Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
            Excel.Application activeApp = Globals.ThisAddIn.GetActiveApp();

            try
            {
                Cursor.Current = Cursors.WaitCursor;

                string riverName = m1;
                string topoID = m2;
                Excel.Range secIdRange = (Excel.Range)activWorkSheet.Range[m3];
                Excel.Range distanceRange = (Excel.Range)activWorkSheet.Range[m4];
                Excel.Range chainageRange = (Excel.Range)activWorkSheet.Range[m5];
                Excel.Range rlRange = (Excel.Range)activWorkSheet.Range[m6];
                Excel.Range xRange = null;
                Excel.Range yRange = null;

                int secIdColNo = secIdRange.Column;
                int distanceColNo = distanceRange.Column;
                int chainageColNo = chainageRange.Column;
                int rlColNo = rlRange.Column;
                int XColNo =0;
                int YColNo =0;

                if (m7 != "" & m8 != "")
                {
                    xRange = (Excel.Range)activWorkSheet.Range[m7];
                    yRange = (Excel.Range)activWorkSheet.Range[m8];
                    XColNo = xRange.Column;
                    YColNo = yRange.Column;
                }

                 
                

                int firstRowNum;
                int lastRowNum;
                #region Getting first and last row no of data....

                int totalRowNum = activWorkSheet.UsedRange.Rows.Count;
                lastRowNum = activWorkSheet.UsedRange.Row + totalRowNum - 1;

                int lastRowOfSelection = secIdRange.Row + secIdRange.Rows.Count - 1;
                if (lastRowOfSelection < distanceRange.Row + distanceRange.Rows.Count - 1)
                {
                    lastRowOfSelection = distanceRange.Row + distanceRange.Rows.Count - 1;
                }
                if (lastRowOfSelection < chainageRange.Row + chainageRange.Rows.Count - 1)
                {
                    lastRowOfSelection = chainageRange.Row + chainageRange.Rows.Count - 1;
                }
                if (lastRowOfSelection < rlRange.Row + rlRange.Rows.Count - 1)
                {
                    lastRowOfSelection = rlRange.Row + rlRange.Rows.Count - 1;
                }
                if (m7 != "" & m8 != "")
                {
                    if (lastRowOfSelection < xRange.Row + xRange.Rows.Count - 1)
                    {
                        lastRowOfSelection = xRange.Row + xRange.Rows.Count - 1;
                    }
                    if (lastRowOfSelection < yRange.Row + yRange.Rows.Count - 1)
                    {
                        lastRowOfSelection = yRange.Row + yRange.Rows.Count - 1;
                    }
                }


                if (lastRowNum > lastRowOfSelection)
                {
                    lastRowNum = lastRowOfSelection;
                }

                int firstRowOfSelection = secIdRange.Row;
                if (firstRowOfSelection > distanceRange.Row)
                {
                    firstRowOfSelection = distanceRange.Row;
                }
                if (firstRowOfSelection > chainageRange.Row)
                {
                    firstRowOfSelection = chainageRange.Row;
                }
                if (firstRowOfSelection > rlRange.Row)
                {
                    firstRowOfSelection = rlRange.Row;
                }

                if (m7 != "" & m8 != "")
                {
                    if (firstRowOfSelection > xRange.Row)
                    {
                        firstRowOfSelection = xRange.Row;
                    }
                    if (firstRowOfSelection > yRange.Row)
                    {
                        firstRowOfSelection = yRange.Row;
                    }
                }

                firstRowNum = firstRowOfSelection;
                for (int i = 1; i <= lastRowNum - firstRowNum + 1; i++)
                {
                    double firstDist;
                    bool header = double.TryParse(distanceRange[i].Text, out(firstDist));
                    if (header)
                    {
                        firstRowNum = firstRowNum + i - 1;
                        break;
                    }
                }

                #endregion

                string secColumnName = ColumnIndexToColumnLetter(secIdColNo);
                string distanceColumnName = ColumnIndexToColumnLetter(distanceColNo);
                string chainageColumnName = ColumnIndexToColumnLetter(chainageColNo);
                string rlColumnName = ColumnIndexToColumnLetter(rlColNo);
                string XColumnName;
                string YColumnName;

                string secRangeFrom = secColumnName + firstRowNum.ToString();
                string secRangeTo = secColumnName + lastRowNum.ToString();
                string distanceRangeFrom = distanceColumnName + firstRowNum.ToString();
                string distanceRangeTo = distanceColumnName + lastRowNum.ToString();
                string chainageRangeFrom = chainageColumnName + firstRowNum.ToString();
                string chainageRangeTo = chainageColumnName + lastRowNum.ToString();
                string rlRangeFrom = rlColumnName + firstRowNum.ToString();
                string rlRangeTo = rlColumnName + lastRowNum.ToString();
                string XRangeFrom;
                string XRangeTo;
                string YRangeFrom;
                string YRangeTo;

                Excel.Range secIdRangeNew = (Excel.Range)activWorkSheet.Range[secRangeFrom, secRangeTo];
                Excel.Range distanceRangeNew = (Excel.Range)activWorkSheet.Range[distanceRangeFrom, distanceRangeTo];
                Excel.Range chainageRangeNew = (Excel.Range)activWorkSheet.Range[chainageRangeFrom, chainageRangeTo];
                Excel.Range rlRangeNew = (Excel.Range)activWorkSheet.Range[rlRangeFrom, rlRangeTo];
                Excel.Range xRangeNew = null;
                Excel.Range yRangeNew = null;
                if (m7 != "" & m8 != "")
                {
                    XColumnName = ColumnIndexToColumnLetter(XColNo);
                    YColumnName = ColumnIndexToColumnLetter(YColNo);

                    XRangeFrom = XColumnName + firstRowNum.ToString();
                    XRangeTo = XColumnName + lastRowNum.ToString();
                    YRangeFrom = YColumnName + firstRowNum.ToString();
                    YRangeTo = YColumnName + lastRowNum.ToString();

                    xRangeNew = (Excel.Range)activWorkSheet.Range[XRangeFrom, XRangeTo];
                    yRangeNew = (Excel.Range)activWorkSheet.Range[YRangeFrom, YRangeTo];
                }


                List<string> secIdList = new List<string>();
                List<string> chainageList = new List<string>();
                List<double> distList = new List<double>();
                List<double> rltList = new List<double>();
                List<string> xList = new List<string>();
                List<string> yList = new List<string>();

                List<int> indexHolder = new List<int>();
                //List<string> chainageList = new List<string>();

                int indexCounter = 0;
                for (int i = 1; i <= lastRowNum - firstRowNum + 1; i++)
                {
                    double distValue;
                    bool isValue = double.TryParse(distanceRangeNew[i].Text, out distValue);
                    if (isValue)
                    {
                        secIdList.Add(secIdRangeNew[i].Text);
                        chainageList.Add(chainageRangeNew[i].Text);
                        distList.Add(distanceRangeNew[i].Value);
                        rltList.Add(rlRangeNew[i].Value);
                        if (m7 != "" & m8 != "")
                        {
                            xList.Add(xRangeNew[i].Text);
                            yList.Add(yRangeNew[i].Text);
                        }

                        if (chainageRangeNew[i].Text != chainageRangeNew[i + 1].Text)
                        {
                            indexHolder.Add(indexCounter);
                        }
                        //if (i == lastRowNum - firstRowNum+1)
                        //{
                        //    indexHolder.Add(indexCounter);
                        //}
                        //MessageBox.Show(indexHolder.ElementAt(i - 2).ToString() + "+" + i.ToString());
                        indexCounter++;
                    }
                    //MessageBox.Show(indexHolder.ElementAt(i-2).ToString()+"+"+i.ToString());
                }

                
                

                string xns11Data = null;

                int secCounter = 0;
                for (int i = 0; i < indexHolder.Count; i++)
                {
                    int profileNo = 0;
                    string coordString = null;
                    xns11Data = xns11Data + topoID + System.Environment.NewLine;
                    xns11Data = xns11Data + riverName + System.Environment.NewLine;
                    xns11Data = xns11Data + "            " + chainageList.ElementAt(indexHolder.ElementAt(i)) + System.Environment.NewLine;
                    if (m7 != "" & m8 != "")
                    {

                        if (i == 0)
                        {
                            double xleft = 0;
                            double yleft = 0;
                            double xRight = 0;
                            double yRight = 0;

                            for (int k = i; k <= indexHolder.ElementAt(i); k++ )
                            {
                                bool isXvalue = double.TryParse(xList.ElementAt(k), out xleft);
                                bool isYvalue = double.TryParse(yList.ElementAt(k), out yleft);
                                if(isXvalue == true & isYvalue == true)
                                {
                                    break;
                                }
                                
                            }
                            for (int k = indexHolder.ElementAt(i); k >= i; k--)
                            {
                                bool isXvalue = double.TryParse(xList.ElementAt(k), out xRight);
                                bool isYvalue = double.TryParse(yList.ElementAt(k), out yRight);
                                if (isXvalue == true & isYvalue == true)
                                {
                                    break;
                                }
                            }
                            coordString = "    2  " + xleft + "  " + yleft + " " + xRight + "  " + yRight;
                        }
                        else
                        {
                            double xleft = 0;
                            double yleft = 0;
                            double xRight = 0;
                            double yRight = 0;

                            for (int k = indexHolder.ElementAt(i - 1) + 1; k <= indexHolder.ElementAt(i); k++)
                            {
                                bool isXvalue = double.TryParse(xList.ElementAt(k), out xleft);
                                bool isYvalue = double.TryParse(yList.ElementAt(k), out yleft);
                                if (isXvalue == true & isYvalue == true)
                                {
                                    break;
                                }

                            }
                            for (int k = indexHolder.ElementAt(i); k >= indexHolder.ElementAt(i - 1) + 1; k--)
                            {
                                bool isXvalue = double.TryParse(xList.ElementAt(k), out xRight);
                                bool isYvalue = double.TryParse(yList.ElementAt(k), out yRight);
                                if (isXvalue == true & isYvalue == true)
                                {
                                    break;
                                }
                            }
                            //coordString = "    2  " + xList.ElementAt(indexHolder.ElementAt(i - 1) + 1) + "  " + yList.ElementAt(indexHolder.ElementAt(i - 1) + 1) + " " + xList.ElementAt(indexHolder.ElementAt(i)) + "  " + yList.ElementAt(indexHolder.ElementAt(i));
                            coordString = "    2  " + xleft + "  " + yleft + " " + xRight + "  " + yRight;
                        }
                    }
                    else
                    {
                        coordString = "    0";
                    }

                    xns11Data = xns11Data + "COORDINATES" + System.Environment.NewLine + coordString + System.Environment.NewLine;
                    xns11Data = xns11Data + "FLOW DIRECTION" + System.Environment.NewLine + "    0      " + System.Environment.NewLine + "PROTECT DATA" + System.Environment.NewLine + "    0      " + System.Environment.NewLine + "DATUM" + System.Environment.NewLine + "      0.00" + System.Environment.NewLine + "RADIUS TYPE" + System.Environment.NewLine + "    0" + System.Environment.NewLine + "DIVIDE X-Section" + System.Environment.NewLine + "0" + System.Environment.NewLine + "SECTION ID" + System.Environment.NewLine;
                    xns11Data = xns11Data + "            " + secIdList.ElementAt(indexHolder.ElementAt(i)) + System.Environment.NewLine;
                    xns11Data = xns11Data + "INTERPOLATED" + System.Environment.NewLine + "    0" + System.Environment.NewLine + "ANGLE" + System.Environment.NewLine + "    0.00   0" + System.Environment.NewLine + "RESISTANCE NUMBERS" + System.Environment.NewLine + "   2  0     1.000     1.000     1.000    1.000    1.000" + System.Environment.NewLine;

                    if (i == 0)
                    {
                        profileNo = indexHolder.ElementAt(i) + 1;
                    }
                    else
                    {
                        profileNo = indexHolder.ElementAt(i) - indexHolder.ElementAt(i - 1);
                    }

                    xns11Data = xns11Data + "PROFILE        " + profileNo + System.Environment.NewLine;


                    int bankMarker = 0;

                    for (int j = secCounter; j <= indexHolder.ElementAt(i); j++)
                    {
                        //Here goes Sec info
                        double centrePoint = rltList.ElementAt(secCounter);
                        int clMarkerIndex = secCounter;
                        for (int h = secCounter; h < indexHolder.ElementAt(i); h++)
                        {
                            if (centrePoint > rltList.ElementAt(h))
                            {
                                centrePoint = rltList.ElementAt(h);
                                clMarkerIndex = h;
                            }
                        }

                        double leftBankPoint = rltList.ElementAt(secCounter);
                        int lbMarkerIndex = secCounter;
                        for (int h = secCounter; h < clMarkerIndex; h++)
                        {
                            if (leftBankPoint < rltList.ElementAt(h))
                            {
                                leftBankPoint = rltList.ElementAt(h);
                                lbMarkerIndex = h;
                            }
                        }
                        double rightBankPoint = rltList.ElementAt(clMarkerIndex);
                        int rbMarkerIndex = clMarkerIndex;
                        for (int h = clMarkerIndex; h <= indexHolder.ElementAt(i); h++)
                        {
                            if (rightBankPoint < rltList.ElementAt(h))
                            {
                                rightBankPoint = rltList.ElementAt(h);
                                rbMarkerIndex = h;
                            }
                        }

                        //MessageBox.Show(lbMarkerIndex.ToString() + "+" + clMarkerIndex.ToString() + "+" + rbMarkerIndex.ToString());

                        if (j == lbMarkerIndex)
                        {
                            xns11Data = xns11Data + "    " + distList.ElementAt(j).ToString("F3") + "     " + rltList.ElementAt(j).ToString("F3") + "     1.000     <#1>     0     0.000     0" + System.Environment.NewLine;
                        }
                        else if (j == clMarkerIndex)
                        {
                            xns11Data = xns11Data + "    " + distList.ElementAt(j).ToString("F3") + "     " + rltList.ElementAt(j).ToString("F3") + "     1.000     <#2>     0     0.000     0" + System.Environment.NewLine;
                        }
                        else if (j == rbMarkerIndex)
                        {
                            xns11Data = xns11Data + "    " + distList.ElementAt(j).ToString("F3") + "     " + rltList.ElementAt(j).ToString("F3") + "     1.000     <#4>     0     0.000     0" + System.Environment.NewLine;
                        }
                        else
                        {
                            xns11Data = xns11Data + "    " + distList.ElementAt(j).ToString("F3") + "     " + rltList.ElementAt(j).ToString("F3") + "     1.000     <#0>     0     0.000     0" + System.Environment.NewLine;
                        }

                        bankMarker++;
                    }
                    bankMarker = 0;

                    secCounter = indexHolder.ElementAt(i) + 1;

                    xns11Data = xns11Data + "LEVEL PARAMS" + System.Environment.NewLine + "   0  0    0.000  0    0.000  20" + System.Environment.NewLine + "*******************************" + System.Environment.NewLine;
                }

                Cursor.Current = Cursors.Default;

                #region Writing to file:-----------------------
                if (xns11Data != null)
                {
                    Stream myStream;
                    SaveFileDialog saveXns11 = new SaveFileDialog();

                    saveXns11.Filter = "text files (*.txt)|*.txt";
                    saveXns11.RestoreDirectory = true;
                    saveXns11.FileName = riverName;
                    if (saveXns11.ShowDialog() == DialogResult.OK)
                    {
                        if ((myStream = saveXns11.OpenFile()) != null)
                        {
                            // Code to write the stream goes here.
                            //string filePath= saveXns11.FileName;

                            using (StreamWriter sw = new StreamWriter(myStream))
                            {
                                sw.WriteLine(xns11Data.Trim());
                            }
                            myStream.Close();
                            MessageBox.Show("  Text file exported successfully!  ");
                        }
                    }
                }
                #endregion
            }
            catch
            {
                MessageBox.Show(" Please check if your input is correct. Error may occure for the \n following reasons: \n 1. Input Range did not match as indicated. \n 2. Input Ranges are not equal in dimention.", "Error!");
            }
        }

        public void GenCoordinate(Office.IRibbonControl control)
        {
            var ownerWindow = new Win32Window(Globals.ThisAddIn.Application.Hwnd);
            
            var CoordGenForm = new CoordPrompt();
            CoordGenForm.TopLevel = true;
            CoordGenForm.Show(ownerWindow);
            
            CoordGenForm.okButtonActionEvent += CoordGenForm_okButtonActionEvent;
        }

        private void CoordGenForm_okButtonActionEvent(string m1, string m2, string m3)
        {
            try
            {
                Excel.Workbook currentWB = Globals.ThisAddIn.GetActiveWorkbook();
                Excel.Worksheet activWorkSheet = Globals.ThisAddIn.GetActiveWorksheet();
                Excel.Application activeApp = Globals.ThisAddIn.GetActiveApp();

            refresh:
                
                Cursor.Current = Cursors.WaitCursor;

                Excel.Range xRange = (Excel.Range)activWorkSheet.Range[m1];
                Excel.Range yRange = (Excel.Range)activWorkSheet.Range[m2];
                Excel.Range distanceRange = (Excel.Range)activWorkSheet.Range[m3];

                int firstRowNum;
                int lastRowNum;
                int xColumn = xRange.Column;
                int yColumn = yRange.Column;
                int distColumn = distanceRange.Column;

                #region Getting first and last row no of data....

                int totalRowNum = activWorkSheet.UsedRange.Rows.Count;
                lastRowNum = activWorkSheet.UsedRange.Row + totalRowNum - 1;

                int lastRowOfSelection = xRange.Row + xRange.Rows.Count - 1;
                if (lastRowOfSelection < yRange.Row + yRange.Rows.Count - 1)
                {
                    lastRowOfSelection = yRange.Row + yRange.Rows.Count - 1;

                }
                if (lastRowOfSelection < distanceRange.Row + distanceRange.Rows.Count - 1)
                {
                    lastRowOfSelection = distanceRange.Row + distanceRange.Rows.Count - 1;
                }

                if (lastRowNum > lastRowOfSelection)
                {
                    lastRowNum = lastRowOfSelection;
                }

                int firstRowOfSelection = xRange.Row;
                if (firstRowOfSelection > yRange.Row)
                {
                    firstRowOfSelection = yRange.Row;
                }
                if (firstRowOfSelection > distanceRange.Row)
                {
                    firstRowOfSelection = distanceRange.Row;
                }

                firstRowNum = firstRowOfSelection;

                for (int i = 1; i <= lastRowNum - firstRowNum + 1; i++)
                {
                    double firstDist;
                    bool header = double.TryParse(distanceRange[i].Text, out(firstDist));
                    if (header)
                    {
                        firstRowNum = firstRowNum + i - 1;
                        break;
                    }
                }

                var gapChecker = new List<int>();
                int changeTracker = 0;

                for (int i = 1; i <= lastRowNum - firstRowNum + 1; i++)
                {
                    double distCellValue;
                    bool isDistValue = double.TryParse(distanceRange[i].Text, out(distCellValue));

                    if (isDistValue == false)
                    {
                        gapChecker.Add(i);
                    }
                }
                for (int i = gapChecker.Count - 1; i > 0; i--)
                {
                    if (gapChecker.ElementAt(i) - 1 == gapChecker.ElementAt(i - 1))
                    {
                        ((Excel.Range)activWorkSheet.Rows[xRange.Row + gapChecker.ElementAt(i) - 1, Missing.Value]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        changeTracker = 1;
                    }
                }
                if (changeTracker == 1)
                {
                    goto refresh;
                }

                #endregion

                string xRangeFrom = ColumnIndexToColumnLetter(xColumn) + firstRowNum.ToString();
                string xRangeTo = ColumnIndexToColumnLetter(xColumn) + lastRowNum.ToString();
                string yRangeFrom = ColumnIndexToColumnLetter(yColumn) + firstRowNum.ToString();
                string yRangeTo = ColumnIndexToColumnLetter(yColumn) + lastRowNum.ToString();
                string distRangeFrom = ColumnIndexToColumnLetter(distColumn) + firstRowNum.ToString();
                string distRangeTo = ColumnIndexToColumnLetter(distColumn) + lastRowNum.ToString();

                Excel.Range xRangeNew = (Excel.Range)activWorkSheet.Range[xRangeFrom, xRangeTo];
                Excel.Range yRangeNew = (Excel.Range)activWorkSheet.Range[yRangeFrom, yRangeTo];
                Excel.Range distanceRangeNew = (Excel.Range)activWorkSheet.Range[distRangeFrom, distRangeTo];

                int ColToInsert = xColumn;
                if (ColToInsert < yColumn)
                {
                    ColToInsert = yColumn;
                }
                if (ColToInsert < distColumn)
                {
                    ColToInsert = distColumn;
                }

                ColToInsert = ColToInsert + 1;
                string insertColFrom = ColumnIndexToColumnLetter(ColToInsert) + firstRowNum.ToString();
                string insertColTo = ColumnIndexToColumnLetter(ColToInsert) + lastRowNum.ToString();
                Excel.Range insertCol = (Excel.Range)activWorkSheet.Range[insertColFrom, insertColTo];

                insertCol.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                        Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                insertCol.EntireColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight,
                        Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);

                
                activWorkSheet.Cells[1, ColToInsert] = activWorkSheet.Cells[1, xColumn].Text + "_Gen";
                activWorkSheet.Cells[1, ColToInsert + 1] = activWorkSheet.Cells[1, yColumn].Text + "_Gen";
                

                var DistGapList = new List<int>();

                for (int i = 1; i <= lastRowNum - firstRowNum + 1; i++)
                {
                    double distCellValue;
                    bool isDistValue = double.TryParse(distanceRangeNew[i + 1].Text, out(distCellValue));

                    if (isDistValue == false || distCellValue == 0)
                    {
                        DistGapList.Add(i);
                        i++;
                    }

                }

                int RowCounter = 1;
                string XfromulaStringFixed = null;
                string YfromulaStringFixed = null;
                for (int i = 0; i < DistGapList.Count; i++)
                {

                    for (int j = RowCounter; j <= DistGapList.ElementAt(i); j++)
                    {
                        if (i == 0)
                        {
                            XfromulaStringFixed = "$" + ColumnIndexToColumnLetter(xColumn) + "$" + (xRangeNew.Row).ToString() + ":$" + ColumnIndexToColumnLetter(xColumn) + "$" + (xRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString() + ",$" + ColumnIndexToColumnLetter(distColumn) + "$" + (xRangeNew.Row).ToString() + ":$" + ColumnIndexToColumnLetter(distColumn) + "$" + (xRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString();
                            YfromulaStringFixed = "$" + ColumnIndexToColumnLetter(yColumn) + "$" + (yRangeNew.Row).ToString() + ":$" + ColumnIndexToColumnLetter(yColumn) + "$" + (xRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString() + ",$" + ColumnIndexToColumnLetter(distColumn) + "$" + (yRangeNew.Row).ToString() + ":$" + ColumnIndexToColumnLetter(distColumn) + "$" + (yRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString();
                        }
                        else
                        {
                            XfromulaStringFixed = "$" + ColumnIndexToColumnLetter(xColumn) + "$" + (xRangeNew.Row + RowCounter - 1).ToString() + ":$" + ColumnIndexToColumnLetter(xColumn) + "$" + (xRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString() + ",$" + ColumnIndexToColumnLetter(distColumn) + "$" + (xRangeNew.Row + RowCounter - 1).ToString() + ":$" + ColumnIndexToColumnLetter(distColumn) + "$" + (xRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString();
                            YfromulaStringFixed = "$" + ColumnIndexToColumnLetter(yColumn) + "$" + (yRangeNew.Row + RowCounter - 1).ToString() + ":$" + ColumnIndexToColumnLetter(yColumn) + "$" + (yRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString() + ",$" + ColumnIndexToColumnLetter(distColumn) + "$" + (yRangeNew.Row + RowCounter - 1).ToString() + ":$" + ColumnIndexToColumnLetter(distColumn) + "$" + (yRangeNew.Row + DistGapList.ElementAt(i) - 1).ToString();
                        }

                        activWorkSheet.Cells[xRangeNew.Row + j - 1, ColToInsert] = "=FORECAST(" + ColumnIndexToColumnLetter(distColumn) + (xRangeNew.Row + j - 1).ToString() + "," + XfromulaStringFixed + ")";
                        activWorkSheet.Cells[xRangeNew.Row + j - 1, ColToInsert + 1] = "=FORECAST(" + ColumnIndexToColumnLetter(distColumn) + (xRangeNew.Row + j - 1).ToString() + "," + YfromulaStringFixed + ")";

                    }

                    double distCellValue;
                    bool isDistValue = double.TryParse(distanceRangeNew[DistGapList.ElementAt(i) + 1].Text, out(distCellValue));

                    if (isDistValue == false)
                    {
                        RowCounter = DistGapList.ElementAt(i) + 2;
                    }
                    else if (distCellValue == 0)
                    {
                        RowCounter = DistGapList.ElementAt(i) + 1;
                    }

                }
                
                Cursor.Current = Cursors.Default;
            }
            catch
            {
                MessageBox.Show("  There is problem!  \n ", "Problem!");
            }

         }
        
        public void AboutMe(Office.IRibbonControl control)
        {

            MessageBox.Show("");
            
        }

        public void mtoolsHelp(Office.IRibbonControl control)
        {
            MessageBox.Show("");
        }
        
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
