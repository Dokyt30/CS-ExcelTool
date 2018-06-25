using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

#pragma warning disable 649

namespace ExcelABMergeTool
{

    /// <summary>
    /// エクセルファイル１つのデータ
    /// </summary>
    public  struct ExcelData
    {
        public string filename;
        public Excel.Workbook workbook;
        public Excel.Sheets sheets;
        public Excel.Worksheet worksheet;
        public bool saveflag; // fix warning


        /// <summary>
        /// 初期化
        /// </summary>
        /// <param name="filename">エクセルファイル名</param>
        public void init(string filename)
        {
            this.filename = filename;
            workbook = ExcelApp.GetInstance().sWorkBook.Open(filename);
            sheets = workbook.Sheets;

            int sheetNum = sheets.Count;
            foreach (Excel.Worksheet ws in sheets)
            {
                worksheet = ws;
                break;
#if false
                    WorkSheetDebug(ws);
#endif
            }

        }

        /// <summary>
        /// 終了
        /// </summary>
        public void term()
        {

            if (saveflag)
            {
                workbook.SaveAs(filename); // 保存
            }
            if (workbook != null)
            {
                workbook.Close(); // 閉じる
            }

            //COM解放
            if (worksheet != null)
            {
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }

            if (sheets != null)
            {
                Marshal.ReleaseComObject(sheets);
                sheets = null;
            }

            if (workbook != null)
            {

                Marshal.ReleaseComObject(workbook);
                workbook = null;
            }

        }

        /// <summary>
        /// worksheet設定
        /// </summary>
        public void SetWorkSheet(int sheetid)
        {
            if (sheets.Count < 1) System.Diagnostics.Debug.Assert(false); // 0未満
            if (sheetid > sheets.Count) System.Diagnostics.Debug.Assert(false); // 超えている
            worksheet = sheets[sheetid]; 
        }

        /// ===============================================
        /// static
        /// ===============================================

        /// <summary>
        /// ロード
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static ExcelData LoadExcel(string filename)
        {
            ExcelData edata = new ExcelData();

            try
            {
                edata.init(filename);
            }
            catch (System.IO.FileNotFoundException)
            {
                TermExcel(edata);
            }
            return edata;
        }

        /// <summary>
        /// 終了
        /// </summary>
        /// <param name="edata"></param>
        public static void TermExcel(ExcelData edata)
        {
            edata.term();
            GC.Collect();
        }


    };
}
