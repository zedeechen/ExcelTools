using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTools
{
    class FunctionSheetControl : SheetControl
    {
        private const int m_wTitleChsRow = 1;
        private const int m_wServerOnlyRow = 2;
        private const int m_wTitleIdxRow = 3;
        private const int m_wTitleEngRow = 4;

        private void GetHeader( YYExcel inExcel, Dictionary<int, FunctionSheet.TitleConfig> header,
                                out bool bWithTxtCol, ref List<string> lstErrorMsg,
                                int titleChsNameRow, int serverOnlyRow, int titleIndexRow, int titleNameRow = 0 )
        {
            bWithTxtCol = false;
            if ( inExcel == null ) return;

            string excelName = Path.GetFileName( inExcel.path );
            int wRowCount = inExcel.GetRowsCount();
            int wColCount = inExcel.GetColumnsCount();

            if ( titleChsNameRow <= 0 || titleChsNameRow > wRowCount || 
                 titleIndexRow <= 0 || titleIndexRow > wRowCount ||
                 titleNameRow < 0 || titleNameRow > wRowCount )
                return;

            Dictionary<string, int> dictColNames = new Dictionary<string, int>();

            for ( int i = 1; i <= wColCount; i++ )
            {
                string address = inExcel.getCellAddress( titleIndexRow, i );
                string strNameChs = inExcel.getCellValue( titleChsNameRow, i );
                string strServer  = inExcel.getCellValue( serverOnlyRow, i );
                string strIndex   = inExcel.getCellValue( titleIndexRow, i );
                string strName = "";
                if ( titleNameRow != 0 )
                    strName = inExcel.getCellValue( titleNameRow, i );

                if ( strIndex == "0" ) continue;
                if ( strIndex == "" ) continue;

                int    wIndex;
                bool   res        = Int32.TryParse( strIndex, out wIndex );

                // 列Id检查
                if ( !res || strIndex.Length != m_wColIdLen )
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "列Id应为" + m_wColIdLen + "位整数" ) );
                    continue;
                }
                
                // 标记检查
                bool bIsServerOnly = false;
                if ( strServer == "s" || strServer == "S" )
                    bIsServerOnly = true;
                else if ( strServer != string.Empty)
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "第" + serverOnlyRow + "行仅支持 空 或标记 s 或 S " ) );
                    continue;
                }

                // 列名非空检查
                if ( titleNameRow != 0 && strName == string.Empty )
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "列名不能为空" ) );
                    continue;
                }

                // 列名重复检查
                if ( dictColNames.ContainsKey( strName ) )
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "列名不能重复" ) );
                    continue;
                }
                dictColNames.Add( strName, 1 );

                FunctionSheet.TitleConfig config;
                if ( titleNameRow != 0 )
                {
                    if ( Regex.IsMatch( strName, m_strTextColPrefix + ".*" ) )
                    {
                        bWithTxtCol = true;
                        config = new FunctionSheet.TitleConfig( wIndex, i, strName, strNameChs, bIsServerOnly, true );
                    }
                    else
                        config = new FunctionSheet.TitleConfig( wIndex, i, strName, strNameChs, bIsServerOnly, false );
                }
                else
                    config = new FunctionSheet.TitleConfig( wIndex, i, strNameChs, bIsServerOnly );

                if ( !header.ContainsKey( wIndex ) )
                    header.Add( wIndex, config );
                else
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "列ID重复" ) );
                }
            }
        }

        private void GetHeader( YYExcel inExcel, Dictionary<int, FunctionSheet.TitleConfig> header,
                                ref List<string> lstErrorMsg,
                                int titleChsNameRow, int serverOnlyRow, int titleIndexRow, int titleNameRow = 0 )
        {
            bool bWithTxtCol;
            GetHeader( inExcel, header, out bWithTxtCol, ref lstErrorMsg, titleChsNameRow, serverOnlyRow, titleIndexRow, titleNameRow );
        }

        private void GetItemsID( YYExcel inExcel, bool bWithTxtCol, SortedDictionary<int, int> item,
                                 out bool bIdAscending, ref List<string> lstErrorMsg,
                                 int wFirstValueRow, int wItemIdCol = 1 )
        {
            bIdAscending = true;
            if ( inExcel == null ) return;

            string excelName = Path.GetFileName( inExcel.path );
            int wRowCount = inExcel.GetRowsCount();
            int wColCount = inExcel.GetColumnsCount();

            while ( inExcel.getCellValue( wRowCount, wItemIdCol ) == "" && wRowCount > 0 )
                wRowCount--;

            if ( wItemIdCol <= 0 || wItemIdCol > wColCount )
                return;

            int preRowIdex = 0;
            for ( int i = wFirstValueRow; i <= wRowCount; i++ )
            {
                string address = inExcel.getCellAddress( i, wItemIdCol );
                string strRowId = inExcel.getCellValue( i, wItemIdCol );
                if ( strRowId == "" )
                {
                    bIdAscending = false;
                    continue;
                }
                
                int wRowId;
                bool res = Int32.TryParse( strRowId, out wRowId );

                if ( !res || ( bWithTxtCol && strRowId.Length > m_wRowIdLen ) )
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "行Id应为" + m_wRowIdLen + "位整数" ) );
                    continue;
                }

                if ( !item.ContainsKey( wRowId ) )
                    item.Add( wRowId, i );
                else
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "行ID重复" ) );
                }

                if ( wRowId > preRowIdex )
                {
                    preRowIdex = wRowId;
                }
                else
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "行ID未递增" ) );
                    bIdAscending = false;
                }
            }
        }

        private void GetItemsID( YYExcel inExcel, SortedDictionary<int, int> item,
                                 out bool bIdAscending, ref List<string> lstErrorMsg,
                                 int wFirstValueRow, int wItemIdCol = 1 )
        {
            GetItemsID( inExcel, false, item, out bIdAscending, ref lstErrorMsg, wFirstValueRow, wItemIdCol );
        }

        private void GetValues( YYExcel inExcel, Dictionary<int, FunctionSheet.TitleConfig> header, SortedDictionary<int, int> item,
                                SortedDictionary<int, Dictionary<int, Node>> values, ref List<string> lstErrorMsg )
        {
            string name = Path.GetFileName( inExcel.path );
            foreach ( int rowID in item.Keys )
            {
                int i = item[rowID];
                if ( !values.ContainsKey( rowID ) )
                {
                    Dictionary<int, Node> dictTemp = new Dictionary<int, Node>();
                    values.Add( rowID,dictTemp );
                }
                else
                {
                    lstErrorMsg.Add( ErrorMsg.Error( name, "行Id" + rowID.ToString() + "重复" ) );
                    continue;
                }
                foreach( int colID in header.Keys )
                {
                    int j = header[colID].titlePos;
                    string cell = inExcel.getCellValue( i, j );
                    int result = 0;

                    Node node = new Node( i, j, cell );

                    if ( Regex.IsMatch( header[colID].titleName, "^i_" ) &&
                        cell != "" && !Int32.TryParse( cell, out result ) )
                    {
                        // type int check
                        string address = inExcel.getCellAddress( i, j );

                        lstErrorMsg.Add( ErrorMsg.FormatError( name, address, "单元格应为整数" ) );
                    }

                    if ( !values[rowID].ContainsKey( colID ) )
                    {
                        values[rowID].Add( colID, node );
                    }
                    else
                    {
                        lstErrorMsg.Add( ErrorMsg.Error( name, "列Id" + colID.ToString() + "重复" ) );
                        continue;
                    }
                }
            }
        }

        public bool Check( string strExcelPath, string strSheetName, bool bExistLineFour,
                           out bool bWithTxtCol, out bool bIdAscending, out FunctionSheet funcSheet, out List<string> lstErrorMsg
                         )
        {
            bWithTxtCol = false;
            bIdAscending = true;
            funcSheet = new FunctionSheet();
            lstErrorMsg = new List<string>();

            string excelName = Path.GetFileName( strExcelPath );

            YYExcel inExcel = new YYExcel();
            try
            {
                inExcel.Open( strExcelPath, strSheetName, YYExcel.Authority.A_READ_ONLY );
            }
            catch (System.Exception ex)
            {
                lstErrorMsg.Add( ErrorMsg.OpenError( excelName ,strSheetName, ex.Message) );
                return false;
            }

            if ( bExistLineFour )
                GetHeader( inExcel, funcSheet.headers, out bWithTxtCol, ref lstErrorMsg, m_wTitleChsRow, m_wServerOnlyRow, m_wTitleIdxRow, m_wTitleEngRow );
            else
                GetHeader( inExcel, funcSheet.headers, out bWithTxtCol, ref lstErrorMsg, m_wTitleChsRow, m_wServerOnlyRow, m_wTitleIdxRow );

            GetItemsID( inExcel, bWithTxtCol, funcSheet.itemPos, out bIdAscending, ref lstErrorMsg, bExistLineFour ? 5 : 4, 1 );

            GetValues( inExcel, funcSheet.headers, funcSheet.itemPos, funcSheet.cells, ref lstErrorMsg );

            inExcel.Close();
            return lstErrorMsg.Count == 0;
        }

        public bool Check( string strExcelPath, string strSheetName, bool bExistLineFour,
                           out bool bWithTxtCol, out bool bIdAscending, out List<string> lstErrorMsg
                         )
        {
            FunctionSheet funcSheet;
            return Check( strExcelPath, strSheetName, bExistLineFour, out bWithTxtCol, out bIdAscending, out funcSheet, out lstErrorMsg );
        }

        public bool Check( string strExcelPath, string strSheetName, bool bExistLineFour,
                           out bool bIdAscending, out List<string> lstErrorMsg
                         )
        {
            bool bWithTxtCol;
            return Check( strExcelPath, strSheetName, bExistLineFour, out bWithTxtCol, out bIdAscending, out lstErrorMsg );
        }

        public bool Read( string strExcelPath, string strSheetName, bool bExistLineFour, out FunctionSheet funcSheet, out List<string> lstErrorMsg )
        {
            funcSheet = new FunctionSheet();
            lstErrorMsg = new List<string>();

            string excelName = Path.GetFileName( strExcelPath );

            YYExcel inExcel = new YYExcel();
            try
            {
                inExcel.Open( strExcelPath, strSheetName, YYExcel.Authority.A_READ_ONLY );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ErrorMsg.OpenError( excelName, strSheetName, ex.Message ) );
                return false;
            }

            if ( bExistLineFour )
                GetHeader( inExcel, funcSheet.headers, ref lstErrorMsg, m_wTitleChsRow, m_wServerOnlyRow, m_wTitleIdxRow, m_wTitleEngRow );
            else
                GetHeader( inExcel, funcSheet.headers, ref lstErrorMsg, m_wTitleChsRow, m_wServerOnlyRow, m_wTitleIdxRow );

            GetItemsID( inExcel, funcSheet.itemPos, out funcSheet.bIsAscending, ref lstErrorMsg, bExistLineFour ? 5 : 4, 1 );

            GetValues( inExcel, funcSheet.headers, funcSheet.itemPos, funcSheet.cells, ref lstErrorMsg );

            inExcel.Close();

            return lstErrorMsg.Count == 0;
        }

    }
}
