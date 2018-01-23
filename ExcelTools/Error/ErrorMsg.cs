using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class ErrorMsg
    {
        public static string Error( string table, string reason )
        {
            return string.Format( "在表[{0}]中：[{1}]", table, reason );
        }

        public static string FormatError( string table, string cell, string reason )
        {
            return string.Format( "在表[{0}]的[{1}]处出错：[{2}]", table, cell, reason );
        }

        public static string OpenError( string table, string sheet, string reason )
        {
            return string.Format( "打开表[{0}][{1}]出错：[{2}]", table, sheet, reason );
        }

        public static string CustomError( string table, string sheet, string cell, string reason )
        {
            return string.Format( "订制[{0}]的[{1}]版本时在[{2}]处出错：[{3}]", table, sheet, cell, reason );
        }

        public static string CustomError( string table, string sheet, int rowId, string reason )
        {
            return string.Format( "订制[{0}]的[{1}]版本时在操作行Id[{2}]时出错：[{3}]", table, sheet, rowId, reason );
        }

        public static string CustomError( string table, string sheet, int rowId, int colId, string reason )
        {
            return string.Format( "订制[{0}]的[{1}]版本时在操作行Id[{2}] 列Id[{3}]时出错：[{4}]", table, sheet, rowId, colId, reason );
        }
    }
}
