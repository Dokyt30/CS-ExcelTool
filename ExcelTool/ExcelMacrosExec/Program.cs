/**
 * 任意のエクセルのマクロを実行する。
 * 
 * 連続して実行したいときに使用する。
 * 
 * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelABMergeTool;
using Microsoft.Office.Interop.Excel;

namespace ExcelMacrosExec
{
    class Program
    {
        /// <summary>
        /// ログ出力
        /// </summary>
        /// <param name="msg"></param>
        static void log(string msg)
        {
            Console.WriteLine(msg);
        }


        /// <summary>
        /// ファイル名取得
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static string getFileName(string[] args)
        {
            log("start args " + args.Count());
            string argscmd = @"temp.xlsm";
            if (args.Count() >= 1)
            {
                argscmd = args[0];
                log("load filename = " + argscmd);
            }
            return Utility.GetFullPathName(argscmd);
        }

        static void Main(string[] args)
        {
            string fileFullePathName = getFileName(args);
            string filename = Utility.GetFileNameFromPath(fileFullePathName);
            // Excel.Application の新しいインスタンスを生成する
            ExcelApp.GetInstance();
            var xlApp = ExcelApp.GetInstance().sExcelApp;
            var xlBooks = ExcelApp.GetInstance().sWorkBook;

            try
            {
                Workbook wb = xlBooks.Open(fileFullePathName);

                // マクロを実行する
                // 標準モジュール内のTestメソッドに "Hello World" を引数で渡し実行
                //xlApp.Run("work.xlsm!Test", "Hello World");
                // Sheet1内のSheetTestメソッドを実行(引数なし)
                xlApp.Run(filename + "test");
                wb.Save();

            }
            finally
            {
                ExcelApp.Release();
            }
        }
    }
}
