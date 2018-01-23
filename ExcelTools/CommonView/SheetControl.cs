using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class SheetControl
    {
        // convention
        public const int m_wTblIdLen = 2;
        public const int m_wColIdLen = 3;
        public const int m_wRowIdLen = 4;
        public const string m_strTextColPrefix = "Text_";

        // if id is ascending, binary search
        public int GetRowByIndex( YYExcel inExcel, int RowIndex, int StartRow, int EndRow, int IdColumn )
        {
            while ( inExcel.getCellValue( EndRow, IdColumn ) == "" )
                EndRow--;
            int st = StartRow, ed = EndRow + 1;
            while ( st < ed )
            {
                int mid = ( st + ed ) / 2;
                string strCell = inExcel.getCellValue( mid, IdColumn );
                int    wCell   = Int32.Parse( strCell );
                if ( wCell > RowIndex ) ed = mid;
                else if ( wCell < RowIndex ) st = mid + 1;
                else
                {
                    st = mid;
                    break;
                }
            }

            return st;
        }

    }
}
