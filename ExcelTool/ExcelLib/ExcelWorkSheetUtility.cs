using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelABMergeTool
{
    public class ExcelWorkSheetUtility
    {
        /// <summary>
        /// 使用している列の数を取得する
        /// </summary>
        /// <param name="sheet">使用しているExcelシート</param>
        /// <returns>列を返す</returns>
        public static int GetWorkSheetUsedRangeRow(Excel.Worksheet sheet)
        {
            Excel.Range param = sheet.UsedRange;
            //int nRowT = param.Row;
            int BsheetMaxRow = param.Row + param.Rows.Count - 1;
            //int nColumnL = param.Column;
            //int BsheetMaxColumn = param.Column + param.Columns.Count - 1;
            return BsheetMaxRow;
        }

    }
}
