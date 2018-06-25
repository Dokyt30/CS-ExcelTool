/**
 * エクセルの任意の列を結合して保存しなおすツール。コピー先の列はコピー元の列に上書きされる。
 
 * コピー元の列、コピー先の列をマージ可能。任意の列を移動させるファイルが多いときに便利。
 * */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelABMergeTool
{
    class Program
    {

        static string replacedir; // 置き換えフォルダ
        static string s_srcFile;    // 入力エクセルファイル
        static string s_destFile;   // 出力エクセルファイル
        
        /// <summary>
        /// ログ出力
        /// </summary>
        /// <param name="msg"></param>
        static void log(string msg){
            Console.WriteLine(msg);
        }

        /// <summary>
        /// debug用
        /// </summary>
        /// <param name="worksheet"></param>
        static void WorkSheetDebug(Excel.Worksheet worksheet)
        {
            // debug
            Excel.Range param;
            param = worksheet.UsedRange;
            //int nRowT = param.Row;
            int nRowB = param.Row + param.Rows.Count - 1;
            //int nColumnL = param.Column;
            int nColumnR = param.Column + param.Columns.Count - 1;
            for (int row = 1; row <= nRowB; row++)
            {
                for (int col = 1; col <= nColumnR; col++)
                {
                    log(worksheet.Cells[row, col].Text.ToString());
                }
            }
        }


        /// <summary>
        /// ロード
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        static ExcelData LoadExcel(string filename)
        {
            return ExcelData.LoadExcel(filename);
        }

        /// <summary>
        /// 終了
        /// </summary>
        /// <param name="edata"></param>
        static void TermExcel(ExcelData edata)
        {
            ExcelData.TermExcel(edata);
        }

        /// <summary>
        /// Excelデータまとめて終了
        /// </summary>
        /// <param name="ebdata"></param>
        /// <param name="ecdata"></param>
        static void TermLoadFile(ExcelData ebdata, ExcelData ecdata)
        {

            TermExcel(ebdata);
            TermExcel(ecdata);
        }

        /// <summary>
        /// Excelデータ作成
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        static ExcelData CreateExcel(string fileName)
        {
            ExcelData ed = new ExcelData();
            log("'" + fileName + "'を作成します。");
            ed.filename = fileName;
            ed.workbook = ExcelApp.GetInstance().sWorkBook.Add();
//            ed.workbook.SaveAs(fileName); // 後でセーブ
//            ed.workbook.Worksheets.Add(Type.Missing);
            ed.sheets = ed.workbook.Sheets;
//            ed.worksheet = ed.sheets[1];
            ed.saveflag = true;
            return ed;
        }

        /// <summary>
        /// R列とS列をマージする
        /// </summary>
        /// <param name="bsheet">出力シート</param>
        /// <param name="csheet">吐き出し先</param>
        static void BCSheetMerge(Excel.Worksheet bsheet, Excel.Worksheet csheet)
        {
            int BsheetMaxRow = ExcelWorkSheetUtility.GetWorkSheetUsedRangeRow(bsheet);
            csheet.get_Range(Const.FirstRangeA + Const.FirstRangeColumn, Const.FirstRangeA + BsheetMaxRow.ToString()).Value2 = bsheet.get_Range(Const.FirstRangeB + Const.FirstRangeColumn, Const.FirstRangeB + BsheetMaxRow.ToString()).Value2;
        }


        /// <summary>
        /// Excelのシートをマージする S列をR列にコピーのみ
        /// </summary>
        /// <param name="ebdata"></param>
        /// <param name="ecdata"></param>
        static void MergeSheet(ExcelData ebdata, ExcelData ecdata)
        {
            const int sheetid = 1; // 固定で1のはず
            ebdata.SetWorkSheet(sheetid);
            ecdata.SetWorkSheet(sheetid);
            BCSheetMerge(ebdata.worksheet, ecdata.worksheet);
        }

        /// <summary>
        /// Excelデータ読み込み
        /// </summary>
        /// <param name="pathExcelA"></param>
        /// <param name="pathExcelB"></param>
        /// <param name="ebdata"></param>
        /// <param name="ecdata"></param>
        static void LoadFile(string pathExcelA, string pathExcelB, ref ExcelData ebdata, ref ExcelData ecdata)
        {

            ebdata = new ExcelData();
            ecdata = new ExcelData();
            try
            {
                string localSrcDir = Utility.GetDirectoryPath(pathExcelA);
                string localDestDir = Utility.GetDirectoryPath(pathExcelB);
                string localMergeDir = localSrcDir + Const.MergeFolderName;
                replacedir = localMergeDir; // 保存
                string filename = Utility.GetFileNameFromPath(pathExcelA);
                string outputfilename = Utility.GetFileNameFromPathWithoutExt(pathExcelA) + ExcelABMergeTool.Const.MergeName + Utility.GetFileNameExtension(pathExcelA); //.xlsx";
                Utility.CreateDirectory(localMergeDir);
                Utility.CreateDirectory(localSrcDir);
                Utility.CreateDirectory(localDestDir);
                {
                    log("入力ファイル:" + pathExcelA);
                    ExcelData eadata = new ExcelData();
                    string file = localSrcDir + filename;
                    eadata = (!Utility.CheckExistFile(file)) ? CreateExcel(file) : LoadExcel(file);
                    Utility.DeleteFile(localMergeDir + outputfilename);
                    // aを別ファイルに保存する。
                    eadata.workbook.SaveCopyAs(localMergeDir + outputfilename);
                    log("入力ファイルをコピー => " + outputfilename);
                    TermExcel(eadata); // eaは閉じる
                    log("入力ファイルは閉じる");
                }

                {
                    log("出力ファイル:" + pathExcelB);
                    string file = localDestDir + filename;
                    ebdata = (!Utility.CheckExistFile(file)) ? CreateExcel(file) : LoadExcel(file);
                    // eb開いたまま
                }

                {
                    string file = localMergeDir + outputfilename;
                    log("mergeファイルを開く:" + outputfilename);
                    ecdata = (!Utility.CheckExistFile(file)) ? CreateExcel(file) : LoadExcel(file);
                    ecdata.saveflag = true;
                }
            }
            finally
            {
            }
        }

        ///==========================================
        
        
        /// cmd の引数文字列を置き換える
        static void SetupCmdArgs(ref string srcfile, ref string destfile)
        {
            //カレントディレクトリを取得する
            string sCurrent = System.IO.Directory.GetCurrentDirectory() + "\\";

            //コマンドライン引数を表示する
            //Console.WriteLine(System.Environment.CommandLine);

            //コマンドライン引数を配列で取得する
            string[] cmds = System.Environment.GetCommandLineArgs();
            //コマンドライン引数を列挙する
            int count = 0;
            foreach (string cmd in cmds)
            {
                //Console.WriteLine(cmd);
                if (count == (int)Const.ArgsId.src)
                {
                    if (!System.IO.Path.IsPathRooted(cmd))
                    {
                        srcfile = sCurrent + cmd;
                    }
                    else
                    {
                        srcfile = cmd;
                    }
                }
                else if (count == (int)Const.ArgsId.dest)
                {
                    if (!System.IO.Path.IsPathRooted(cmd))
                    {
                        destfile = sCurrent + cmd;
                    }
                    else
                    {
                        destfile = cmd;
                    }
                }
                count++;
            }
        }

        ///==========================================

        /// main
        static void Main(string[] args)
        {

            // 入力エクセルファイル
            s_srcFile = "src.xlsm";
            // 出力エクセルファイル
            s_destFile = "dest.xlsx";   
            Const.ExtensionName = Utility.GetFileNameExtension(s_srcFile); // extension 置き換え

            SetupCmdArgs(ref s_srcFile, ref s_destFile);

            ExcelApp.GetInstance(); // instance create

            ExcelData ebdata = new ExcelData();
            ExcelData ecdata = new ExcelData();

            LoadFile(
                s_srcFile                
                , s_destFile
                , ref ebdata
                , ref ecdata
                );
            log("マージ中 ...");
            MergeSheet(ebdata, ecdata);
            log("マージ終了");

            TermLoadFile(ebdata, ecdata);
            ExcelApp.Release();
            Utility.ReplaceFileNames(replacedir, ExcelABMergeTool.Const.MergeName, Const.ExtensionName);
            log("完了! ");
        }
    }
}
