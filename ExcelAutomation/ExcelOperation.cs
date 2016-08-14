using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel=Microsoft.Office.Interop.Excel;

namespace ExcelAutomation
{
    public class ExcelOperation
    {
        #region Data Member

        private Excel.Application ExcelApp;
        private Excel.Workbook ExcelWorkBook;
        private Excel.Worksheet ExcelWorkSheet;
        private Excel.Range ExcelRange;
        private Excel.Application ExcelApp2;
        private Excel.Workbook ExcelWorkBook2;
        private Excel.Worksheet ExcelWorkSheet2;
        private Excel.Range ExcelRange2;
        private int nRow, nColumn;
        private static int newFileRow=1, newFileColumn=0;
        private object[,] items;
        private string columnMaxName = string.Empty;
        private object misValue = System.Reflection.Missing.Value;
        private ArrayList arrForValues;
        private ArrayList Sheet2ItemsArrayList;
        private List<int> columnIndex;
        private char[] splitCharacter;
        
        #endregion

        #region Private Methods

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public void Quit()
        {
            ExcelWorkBook.Close(false, Type.Missing, Type.Missing);
            ExcelApp.Quit();
        }

        private bool CheckParameter(string [] Parameters)
        {
            for (int i = 0; i < Parameters.Length; i++)
            {
                if (Sheet2ItemsArrayList.Contains(Parameters[i]))
                    return true;
                else
                    continue;
            }
            return false;
        }

        private bool CheckCondition()
        {
            int length = columnIndex.Count;
            splitCharacter=new char[] {';',',','|'};
            string[] AllFirstParameter = new string[10];
            string[] AllSecondParameter = new string[10];
            string secondParameter = string.Empty;
            string firstParameter = string.Empty;
            bool second = false;
            
            try
            {
                firstParameter = (string)items[1, columnIndex[0] + 1];
                firstParameter = firstParameter.ToLower();
                AllFirstParameter = firstParameter.Split(splitCharacter);           
           
                if (length>1)
                 {
                   second = true;
                    secondParameter = (string)items[1, columnIndex[1] + 1];
                    secondParameter = secondParameter.ToLower();
                    AllSecondParameter = secondParameter.Split(splitCharacter);
               
                 }
           
                if (firstParameter.Equals("") && second)
                {
                    if (CheckParameter(AllSecondParameter))
                    {
                        return true;
                    }
                    else
                        return false;
                }
                else if (firstParameter != null)
                {
                    if (CheckParameter(AllFirstParameter))
                        return true;
                    else
                        return false;
                }
                else
                    return false;
            }
            catch (NullReferenceException nre)
            {
                return false;
            }
       }
        private void GetSheet2()
        {
            Sheet2ItemsArrayList = new ArrayList();
            Excel.Worksheet newWorkSheet = ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(2);
            Microsoft.Office.Interop.Excel.Range excelcells = newWorkSheet.UsedRange;
            Microsoft.Office.Interop.Excel.Range newRange = newWorkSheet.get_Range("A1", "A" + excelcells.Rows.Count);
            object[,] sheet2Items = (object[,])newRange.Value2;
            for (int i = 1; i < sheet2Items.Length; i++)
            {
                sheet2Items[i, 1]=sheet2Items[i, 1].ToString().ToLower();
                Sheet2ItemsArrayList.Add(sheet2Items[i, 1]);
            }
        }
        private void SetNewFileNameVariable(string newFileName)
        {
            object misValue = System.Reflection.Missing.Value;
            ExcelApp2 = new Excel.Application();
            ExcelWorkBook2 = ExcelApp2.Workbooks.Add(1);
            //ExcelWorkBook2 = ExcelApp2.Workbooks.Open(newFileName, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", true, false, 0, true, 1, 0);
            ExcelWorkSheet2 = (Excel.Worksheet)ExcelWorkBook2.Worksheets.get_Item(1);
        }

        #endregion

        #region Public Methods

        public void Initialize(string fileName,string newFileName)
        {
            object misValue = System.Reflection.Missing.Value;
            ExcelApp = new Excel.Application();
            ExcelWorkBook = ExcelApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            GetSheet2();
            ExcelWorkSheet = (Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelRange = ExcelWorkSheet.UsedRange;
            nRow = ExcelRange.Rows.Count;
            nColumn = ExcelRange.Columns.Count;
            columnMaxName = GetExcelColumnName(nColumn);
            SetNewFileNameVariable(newFileName);
           
        }

        public List<string> ReadFirstRow(List<string> columnPopulated)
        {
            int rCnt = 1;// We need to read the names of the column;
            for (int cCnt = 1; cCnt <= nColumn; cCnt++)
            {
                string str = (string)(ExcelRange.Cells[rCnt, cCnt] as Excel.Range).Value2;
                columnPopulated.Add(str);    
            }
            return columnPopulated;
        }
        public int getMaxRow()
        {
            return nRow;
        }
        public int getMaxColumn()
        {
            return nColumn;
        }
        public List<int> getColumnUsedForComparison()
        {
            return columnIndex;
        }
        public void setColumnUsedForComparison(List<int> columns)
        {
            columnIndex=columns;
        }
        public bool Operation(int rowNumber)
        {
            arrForValues = new ArrayList();
            items = new string[1, nColumn];
            ExcelRange = ExcelWorkSheet.get_Range("A"+rowNumber,columnMaxName+rowNumber );
            items = (object[,])ExcelRange.Value2;
            if (rowNumber == 1)
            {
                WriteFirstLine(items);
                return false;
            }
            else
            {
                bool answer = CheckCondition();
                if (answer)
                {
                    WritetoFile(items);
                    arrForValues.Clear();
                    return true;
                }
                else
                    return false;
            }
        }

        private void WritetoFile(object[,] arrForValues)
        {
            ExcelRange2 = ExcelWorkSheet2.get_Range("A" + newFileRow, columnMaxName + newFileRow);
            ExcelRange2.Value2 = arrForValues;
            newFileRow++;
            
        }
        private void WriteFirstLine(object[,] arrForValues)
        {
            ExcelRange2= ExcelWorkSheet2.get_Range("A" + newFileRow, columnMaxName + newFileRow);
            ExcelRange2.Value2 = arrForValues;
            newFileRow++;
        }
        public void WriteTheDate()
        {
            string date = DateTime.Today.ToString();
            ExcelWorkSheet2.Cells[newFileRow, 1] = date;
            ((Microsoft.Office.Interop.Excel.Range)ExcelWorkSheet2.Cells[newFileRow, 1]).EntireColumn.ColumnWidth = date.Length;
            //ExcelWorkBook2.SaveAs(@"C:\temp",Type.Missing, null, null, null, null, Excel.XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            ExcelWorkBook2.Close(true, Type.Missing, Type.Missing);
            ExcelApp2.Quit();
        }
        #endregion
        
    }
}
