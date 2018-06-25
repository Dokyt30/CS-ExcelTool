using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelABMergeTool
{
    public class Const
    {
        public const string MergeName = "_merge";
        public static string ExtensionName = ".xlsx"; // extension 置き換えのためconst外し
        public const string MergeFolderName = "..\\merge\\";
        public const string FirstRangeA = "R"; // R列
        public const string FirstRangeB = "S"; // S列
        public const int FirstRangeColumn = 2; // 開始行

        public enum ArgsId{
            src = 1,
            dest,
        };
    }
}
