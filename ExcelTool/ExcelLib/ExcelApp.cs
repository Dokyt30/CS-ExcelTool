using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ExcelABMergeTool
{
    public class ExcelApp
    {
        public Excel.Application sExcelApp = null;
        public Excel.Workbooks sWorkBook = null;

        static ExcelApp instance = null;

        public static ExcelApp GetInstance()
        {
            if (instance == null) instance = new ExcelApp();
            return instance;
        }
        public static void Release()
        {
            instance.TermExcelApp();
            instance = null;
        }
        private ExcelApp()
        {
            InitExcelApp();
        }

        /// <summary>
        /// ExcelApplication initialize
        /// </summary>
        void InitExcelApp()
        {
            //Excelシートのインスタンスを作る
            if (sExcelApp == null) sExcelApp = new Excel.Application();
            if (sWorkBook == null) sWorkBook = sExcelApp.Workbooks;
            sExcelApp.DisplayAlerts = false; // 保存しますかダイアログ表示なし (警告非表示)
            // Excel を表示しない
            sExcelApp.Visible = false;

        }

        /// <summary>
        /// ExcelApplication terminate
        /// </summary>
        void TermExcelApp()
        {
            if (sWorkBook != null)
            {
                sWorkBook.Close();
                Marshal.ReleaseComObject(sWorkBook);
                sWorkBook = null;
            }
            if (sExcelApp != null)
            {
                sExcelApp.Quit();
                Marshal.ReleaseComObject(sExcelApp);
                sExcelApp = null;
            }
        }


    }
}
