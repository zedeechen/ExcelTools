using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System;
using System.Collections.Generic;

namespace ExcelTools
{
    class YYExcel
    {
        public const string copyRight = "Powered By WAR10CK @ Gamed9";
        public const string version = "Version 1.2.0 alpha";

        public enum Authority
        {
            A_READ_ONLY = 1,
            A_READ_AND_WRITE = 2,
        };

        private bool m_disposed;

        private OleDbConnection m_conn;
        private System.Data.DataTable m_table;// cell index start from 0

        private Microsoft.Office.Interop.Excel.Application m_app;
        private Microsoft.Office.Interop.Excel.Workbook m_book;
        private Microsoft.Office.Interop.Excel.Worksheet m_sheet;// cell index start from 1

        public string path { get; set; }

        public YYExcel() { }

        public Authority Power
        {
            get;
            protected set;
        }

        protected void OpenExcel( string strExcelPath, Authority _power )
        {
            this.Power = _power;
            this.path = strExcelPath;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    System.String strConn = "Provider=Microsoft.Ace.OleDb.12.0;" +
                                            "Data Source=" + strExcelPath + ";" +
                                            "Extended Properties=\'Excel 12.0;HDR=No;IMEX=1\'";
                    m_conn = new OleDbConnection( strConn );
                    m_conn.Open();
                    break;
                case Authority.A_READ_AND_WRITE:
                    m_app = new Microsoft.Office.Interop.Excel.Application();
                    m_book = m_app.Workbooks.Open( strExcelPath );
                    break;
            }
        }

        public void Open( string strExcelPath, string strSheetName, Authority _power )
        {
            this.Power = _power;
            OpenExcel( strExcelPath, _power );
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    OleDbDataAdapter cmd = new OleDbDataAdapter( "SELECT * FROM [" + strSheetName + "$]", m_conn );
                    DataSet dataSet = new DataSet();
                    cmd.Fill( dataSet, "[" + strSheetName + "$]" );
                    m_table = dataSet.Tables[0];
                    // 提前关闭oledbconnection
                    m_conn.Close();
                    break;
                case Authority.A_READ_AND_WRITE:
                    m_sheet = m_app.Worksheets.get_Item( strSheetName );
                    break;
            }
        }

        public void Add( string strExcelPath, string strSheetName, Authority _power )
        {
            this.Power = _power;
            OpenExcel( strExcelPath, _power );
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    break;
                case Authority.A_READ_AND_WRITE:
                    //m_book.Sheets.Add( After: m_book.Sheets[m_book.Sheets.Count] ); 
                    m_sheet = m_app.Worksheets.Add();
                    m_sheet.Name = strSheetName;
                    break;
            }
        }

        public void Close()
        {
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_conn != null )
                    {
                        //m_conn.Close();
                        m_conn.Dispose();
                    }
                    if ( m_table != null )
                        m_table.Dispose();
                    m_conn = null;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_app != null)
                        m_app.DisplayAlerts = false;
                    m_sheet = null;
                    if ( m_book != null )
                        m_book.Close();
                    m_book = null;
                    if ( m_app != null )
                        m_app.Quit();
                    m_app = null;
                    break;
            }

            // Manual Clear CLR
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        //public void Dispose()
        //{
        //    Dispose( true );
        //    GC.SuppressFinalize( this );
        //}

        //protected virtual void Dispose( bool disposing )
        //{
        //    if ( !m_disposed )
        //    {
        //        if ( disposing )
        //        {
        //            switch ( Power )
        //            {
        //                case Authority.A_READ_ONLY:
        //                    m_conn.Dispose();
        //                    break;
        //                case Authority.A_READ_AND_WRITE:
        //                    m_app.DisplayAlerts = false;
        //                    if ( m_book != null )
        //                        m_book.Close();
        //                    if ( m_app != null )
        //                        m_app.Quit();
        //                    break;
        //            }
        //        }
        //        switch ( Power )
        //        {
        //            case Authority.A_READ_ONLY:
        //                m_conn = null;
        //                m_table.Dispose();
        //                break;
        //            case Authority.A_READ_AND_WRITE:
        //                m_sheet = null;
        //                m_book = null;
        //                m_app = null;
        //                break;
        //        }
        //        m_disposed = true;
        //    }
        //}

        //~YYExcel()
        //{
        //    Dispose( false );
        //}

        public bool Create( string strExcelPath, string strSheetName, Authority _power )
        {
            this.Power = _power;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    return false;
                case Authority.A_READ_AND_WRITE:
                    if ( m_app != null || m_book != null || m_sheet != null )
                    {
                        return false;
                    }
                    m_app = new Microsoft.Office.Interop.Excel.Application();

                    m_book = this.m_app.Workbooks.Add( XlWBATemplate.xlWBATWorksheet );

                    Sheets sheets = m_book.Sheets;
                    m_sheet = this.m_app.Application.Worksheets.get_Item( 1 );
                    m_sheet.Name = strSheetName;
                    return true;
            }
            return true;
        }

        public void GetSheetsName( string strExcelPath, out List<string> sheets, Authority _power )
        {
            sheets = new List<string>();
            HashSet<string> hash = new HashSet<string>();
            OpenExcel( strExcelPath, _power );
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_conn != null )
                    {
                        System.Data.DataTable dtSheet = m_conn.GetOleDbSchemaTable( OleDbSchemaGuid.Tables, null );
                        foreach ( DataRow drSheet in dtSheet.Rows )
                        {
                            string tblName = drSheet["TABLE_NAME"].ToString();
                            //checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                            if ( tblName.Contains( "$" ) )
                            {
                                tblName = tblName.Substring( 0, tblName.LastIndexOf( "$" ) );
                                if ( !hash.Contains( tblName ) )
                                {
                                    sheets.Add( tblName );
                                    hash.Add( tblName );
                                }
                            }
                        }
                        // 提前关闭oledbconnection
                        m_conn.Close();
                    }
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_book != null )
                    {
                        Sheets shs = m_book.Sheets;
                        foreach ( _Worksheet _wsh in shs )
                        {
                            if ( !hash.Contains( _wsh.Name ) )
                            {
                                sheets.Add( _wsh.Name );
                                hash.Add( _wsh.Name );
                            }
                        }
                    }
                    break;
            }
            Close();
        }

        public int GetRowsCount()
        {
            int rowCount = -1;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_table != null )
                        rowCount = m_table.Rows.Count;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet != null )
                    {
                        Range last = m_sheet.Cells.SpecialCells( XlCellType.xlCellTypeLastCell, Type.Missing );
                        Range range = m_sheet.get_Range( "A1", last );
                        rowCount = range.Rows.Count;
                        //m_sheet.Rows.ClearFormats();
                        //rowCount = m_sheet.UsedRange.Rows.Count;
                    }
                    break;
            }
            return rowCount;
        }

        public int GetColumnsCount()
        {
            int colCount = -1;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_table != null )
                        colCount = m_table.Columns.Count;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet != null )
                    {
                        Range last = m_sheet.Cells.SpecialCells( XlCellType.xlCellTypeLastCell, Type.Missing );
                        Range range = m_sheet.get_Range( "A1", last );
                        colCount = range.Columns.Count;
                        //m_sheet.Columns.ClearFormats();
                        //colCount = m_sheet.UsedRange.Columns.Count;
                    }
                    break;
            }
            return colCount;
        }
        // cell index start from 1
        public string getCellValue( int row, int column )
        {
            string strCell = "";
            if ( row < 1 || column < 1 || row > GetRowsCount() || column > GetColumnsCount() )
                return strCell;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_table != null )
                        strCell = m_table.Rows[row - 1][column - 1].ToString();
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet != null )
                        strCell = m_sheet.Cells[row, column].Text;
                    break;
            }
            return strCell;
        }
        // cell index start from 1
        public bool setCellValue( int row, int column, string value )
        {
            if ( row < 1 || column < 1 )
                return false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    return false;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet != null )
                        m_sheet.Cells[row, column].Value2 = value;
                    break;
            }
            return true;
        }
        /// <summary>
        /// get range value to matrix
        /// </summary>
        /// <param name="rowCell1"></param>
        /// <param name="colCell1"></param>
        /// <param name="rowCell2"></param>
        /// <param name="colCell2"></param>
        /// <param name="matrix"></param>
        /// <returns></returns>
        public bool getRangeValue( int rowCell1, int colCell1, int rowCell2, int colCell2, ref object[,] values)
        {
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( rowCell1 > rowCell2 || colCell1 > colCell2 ) return false;
                    for ( int i = 0; i <= rowCell2 - rowCell1; i++ )
                    {
                        for ( int j = 0; j <= colCell2 - colCell1; j++ )
                        {
                            values[i, j] = getCellValue( rowCell1 + i, colCell1 + j );
                        }
                    }
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    Range range = m_sheet.get_Range( getCellAddress( rowCell1, colCell1 ), getCellAddress( rowCell2, colCell2 ) );
                    values = (object[,])range.Value2;
                    break;
            }
            return true;
        }
        /// <summary>
        /// set range value from matrix
        /// </summary>
        /// <param name="row"> top-left row idex</param>
        /// <param name="column"> top-left column index </param>
        /// <param name="values"> input data </param>
        /// <returns> success or not </returns>
        public bool setRangeValue( int row, int column, object[,] values )
        {
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    return false;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    int rowCount = values.GetLength( 0 );
                    int columnCount = values.GetLength( 1 );
                    if ( rowCount <= 0 || columnCount <= 0 ) return false;
                    Range range = (Range)m_sheet.Cells[row, column];
                    range = range.get_Resize( rowCount, columnCount );
                    range.set_Value( XlRangeValueDataType.xlRangeValueDefault, values );
                    break;
            }
            return true;
        }

        public bool Save()
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_book == null ) return false;
                    m_app.DisplayAlerts = false;
                    m_app.AlertBeforeOverwriting = false;
                    m_book.Save();
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool SaveAs( string strPath )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_book == null ) return false;
                    m_app.DisplayAlerts = false;
                    m_app.AlertBeforeOverwriting = false;
                    m_book.SaveAs( strPath,
                                   Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                                   Type.Missing, Type.Missing, false, false,
                                   Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing
                                 );
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setCellAlignCenter( int row, int column )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    ( (Range)m_sheet.Cells[row, column] ).HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setRangeAlignCenter( int rowCell1, int colCell1, int rowCell2, int colCell2 )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).HorizontalAlignment = XlVAlign.xlVAlignCenter;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setCellInteriorColor( int row, int column, int colorIndex )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    //( (Range)m_sheet.Cells[row, column] ).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;
                    ( (Range)m_sheet.Cells[row, column] ).Interior.ColorIndex = colorIndex;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setRangeInteriorColor( int rowCell1, int colCell1, int rowCell2, int colCell2, int colorIndex )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    //m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                    //                   (Range)m_sheet.Cells[rowCell2, colCell2]
                    //                 ).Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).Interior.ColorIndex = colorIndex;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setCellBorder( int row, int column )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    ( (Range)m_sheet.Cells[row, column] ).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                    ( (Range)m_sheet.Cells[row, column] ).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                    ( (Range)m_sheet.Cells[row, column] ).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
                    ( (Range)m_sheet.Cells[row, column] ).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool setRangeBorder( int rowCell1, int colCell1, int rowCell2, int colCell2 )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThin;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThin;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThin;
                    m_sheet.get_Range( (Range)m_sheet.Cells[rowCell1, colCell1],
                                       (Range)m_sheet.Cells[rowCell2, colCell2]
                                     ).Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThin;
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool InsertRow( int row )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    ( (Range)m_sheet.Rows[row, Type.Missing] ).Insert( XlInsertFormatOrigin.xlFormatFromLeftOrAbove, Type.Missing );
                    ret = true;
                    break;
            }
            return ret;
        }

        public bool InsertColumn( int col )
        {
            bool ret = false;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    ret = false;
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return false;
                    m_sheet.get_Range( (Range)m_sheet.Cells[1, col], (Range)m_sheet.Cells[m_sheet.Rows.Count, col] ).EntireColumn.Insert( XlInsertFormatOrigin.xlFormatFromLeftOrAbove, Type.Missing );
                    ret = true;
                    break;
            }
            return ret;
        }

        /// <summary>
        /// Convert to Excel address format, e.g. Row=5, Col=3 --> C5
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public string getCellAddress( int row, int col )
        {
            string ret = "";
            if ( row < 1 && col < 1 ) return ret;
            switch ( Power )
            {
                case Authority.A_READ_ONLY:
                    if ( m_table == null || row > GetRowsCount() || col > GetColumnsCount() ) return "";
                    ret = ColumnIndexToName( col ) + ( row );
                    break;
                case Authority.A_READ_AND_WRITE:
                    if ( m_sheet == null ) return "";
                    ret = ( (Range)m_sheet.Cells[row, col] ).get_Address( false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing ); 
                    break;
            }
            return ret;
        }

        /// <summary>
        /// Convert column number to letters/name, e.g. 1=A, 2=B, ... 27=AA ...
        /// </summary>
        /// <param name="col">Column index</param>
        /// <returns></returns>
        public static string ColumnIndexToName( int col )
        {
            col--; // 1-indexed --> 0-indexed

            // Determine row ref length L, then how many row refs
            // for lengths 1 to length L-1 -- this is the number
            // of combinations that must be skipped over before doing
            // Base-26 conversion

            // A-Z     (1..26)               = 26       
            // --> 0..(26-1)
            // AA-ZZ   (1..26)(1..26)        = 26*26
            // --> 0..(26*26-1) start from 26
            // --> 26..702
            // AAA-ZZZ (1..26)(1..26)(1..26) = 26*26*26
            // --> 0..(26*26*26-1) start from 702
            // --> 702..18277

            int len = 1;
            int colCountForThisLen = 26;
            int lastColForLen = 25;
            int skip = 0; // Cumulative no. of columns for previous lengths

            while ( col > lastColForLen )
            {
                colCountForThisLen *= 26;
                skip = lastColForLen + 1;
                lastColForLen += colCountForThisLen;
                len++;
            }

            col -= skip;

            char[] colRefChars = new char[len];
            for ( var idx = 0; idx < len; idx++ )
            {
                colRefChars[len - idx - 1] = (char)( 'A' + (int)( col % 26 ) );
                col /= 26;
            }
            return new string( colRefChars );
        }
    }
}
