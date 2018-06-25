using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelABMergeTool
{
    public class Utility
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
        /// ファイル存在しているかどうかチェック
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool CheckExistFile(string fileName)
        {
            if (System.IO.File.Exists(fileName))
            {
                return true;
            }
            else
            {
                log("'" + fileName + "'は存在しません。");
                return false;
            }
        }
        /// <summary>
        /// ディレクトリ作成
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static bool CreateDirectory(string path)
        {
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
            return true;
        }


        /// <summary>
        /// Directory パス取得 \\ 付き
        /// </summary>
        /// <param name="sPath">ファイルパス</param>
        /// <returns>Directory返す</returns>
        public static string GetDirectoryPath(string sPath)
        {
            //// ファイル名と拡張子
            //Console.WriteLine(Path.GetFileName(sPath));

            //// ディレクトリ名 (最後に'\\'は付かない)
            //Console.WriteLine(Path.GetDirectoryName(sPath));

            //// ファイル名のみ 
            //Console.WriteLine(Path.GetFileNameWithoutExtension(sPath));

            //// 拡張子のみ(1文字目は'.')
            //Console.WriteLine(Path.GetExtension(sPath));

            return Path.GetDirectoryName(sPath) + "\\";
        }

        /// <summary>
        /// ファイル名取得
        /// </summary>
        /// <param name="sPath">ファイルパス</param>
        /// <returns>ファイル名返す</returns>
        public static string GetFileNameFromPath(string sPath)
        {
            // ファイル名と拡張子
            return Path.GetFileName(sPath);
        }


        /// <summary>
        /// ファイル名取得 拡張子なし
        /// </summary>
        /// <param name="sPath">ファイルパス</param>
        /// <returns>ファイル名返す</returns>
        public static string GetFileNameFromPathWithoutExt(string sPath)
        {
            // ファイル名と拡張子
            return Path.GetFileNameWithoutExtension(sPath);
        }

        /// <summary>
        /// ファイル名取得 拡張子なし
        /// </summary>
        /// <param name="sPath">ファイルパス</param>
        /// <returns>ファイル名返す</returns>
        public static string GetFileNameExtension(string sPath)
        {
            // ファイル名と拡張子
            return Path.GetExtension(sPath);
        }

        /// <summary>
        /// ファイル削除
        /// </summary>
        public static void DeleteFile(string f)
        {
            if (CheckExistFile(f)) System.IO.File.Delete(f);
        }

        /// <summary>
        /// 名前を置換する
        /// </summary>
        /// <param name="directorypath"></param>
        public static void ReplaceFileNames(string directorypath, string replacestr, string end_extension)
        {
            string endwith = replacestr + end_extension;
            var directory = new DirectoryInfo(directorypath);

            foreach (var file in directory.EnumerateFiles().Where(f => f.Name.EndsWith(endwith)))
            {
                string repfilename = file.FullName.Replace(replacestr, "");
                if(CheckExistFile(repfilename)) DeleteFile(repfilename); // ファイルがある場合は削除
                file.MoveTo(repfilename);
            }
        }

        /// <summary>
        /// 現在位置からのFullpathを入れる
        /// </summary>
        /// <param name="srcfile">現在ディレクトリからの相対パス</param>
        /// <returns></returns>
        public static string GetFullPathName(string srcfile)
        {
            string s;
            //カレントディレクトリを取得する
            string sCurrent = System.IO.Directory.GetCurrentDirectory() + "\\";
            if (!System.IO.Path.IsPathRooted(srcfile))
            {
                s = sCurrent + srcfile;
            }
            else
            {
                s = srcfile;
            }
            return s;
        }
    }
}
