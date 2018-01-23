using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelTools
{
    class CustomSheetControl
    {
        public bool Create( string strExcelPath, string strSheetName, out List<string> lstErrorMsg )
        {
            lstErrorMsg = new List<string>();
            YYExcel outExcel = new YYExcel();
            try
            {
                outExcel.Create( strExcelPath, strSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ex.Message );
                return false;
            }

            outExcel.setCellValue( 1, 1, "表名" );
            outExcel.setCellValue( 1, 2, "修改(1)/新增(2)" );
            outExcel.setCellValue( 1, 3, "行Id" );
            for ( int i = 4; i < 100; i++ )
            {
                if ( i % 2 == 0 )
                    outExcel.setCellValue( 1, i, "列名" );
                else
                    outExcel.setCellValue( 1, i, "内容" );
            }
            outExcel.setRangeAlignCenter( 1, 1, 1, 100 );
            outExcel.setRangeInteriorColor( 1, 1, 1, 100, 37 );
            for ( int i = 1; i < 100; i++ )
                outExcel.setCellBorder( 1, i );

            try
            {
                outExcel.SaveAs( strExcelPath );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ex.Message );
                outExcel.Close();
                return false;
            }

            outExcel.Close();
            return true;
        }
        public bool Read( string strExcelPath, string strSheetName, Dictionary<string ,string> dictNameHash, out CustomSheet customSheet, out List<string> lstErrorMsg )
        {
            string excelName = Path.GetFileName( strExcelPath );
            lstErrorMsg = new List<string>();
            customSheet = new CustomSheet();
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

            int wRowCount = inExcel.GetRowsCount();
            int wColCount = inExcel.GetColumnsCount();

            for ( int i = 2; i <= wRowCount; i++ )
            {
                string table = inExcel.getCellValue( i, 1 );
                string opCode = inExcel.getCellValue( i, 2 );
                string rowId = inExcel.getCellValue( i, 3 );

                bool bFoundError = false;
                if ( !dictNameHash.ContainsKey( table ) )
                {
                    string address = inExcel.getCellAddress( i, 1 );
                    lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, table + "不存在" ) );
                    bFoundError = true;
                }
                if ( opCode != "1" && opCode != "2" )
                {
                    string address = inExcel.getCellAddress( i, 2 );
                    lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, "操作不合法" ) );
                    bFoundError = true;
                }

                int wRowId;
                bool isDigital = Int32.TryParse( rowId, out wRowId );
                if ( !isDigital )
                {
                    string address = inExcel.getCellAddress( i, 3 );
                    lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, "不是合法数字，不能作为行Id" ) );
                    bFoundError = true;
                }

                if ( bFoundError ) continue;

                if ( !customSheet.dictCustomItem.ContainsKey( table ) )
                {
                    Dictionary<int, CustomSheet.Item> tmpDict = new Dictionary<int, CustomSheet.Item>();
                    customSheet.dictCustomItem.Add( table, tmpDict );
                }
                if ( !customSheet.dictCustomItem[table].ContainsKey( wRowId ) )
                {
                    CustomSheet.Item item = new CustomSheet.Item( Int32.Parse( opCode ) );
                    int j = 4;
                    string cell = inExcel.getCellValue( i, j );
                    while ( cell != string.Empty )
                    {
                        int wCol;
                        bool cellOk = Int32.TryParse( cell, out wCol );
                        if ( !cellOk )
                        {
                            string address = inExcel.getCellAddress( i, j );
                            lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, "不是合法数字，不能作为列Id" ) );
                        }
                        j++;
                        cell = inExcel.getCellValue( i, j );
                        if ( item.dictOpCell.ContainsKey( wCol ) )
                        {
                            string address = inExcel.getCellAddress( i, j );
                            lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, "对列Id重复操作" ) );
                        }
                        else
                        {
                            item.dictOpCell.Add( wCol, cell );
                        }
                        j++;
                        cell = inExcel.getCellValue( i, j );
                    }
                    customSheet.dictCustomItem[table].Add( wRowId, item );
                }
                else
                {
                    string address = inExcel.getCellAddress( i, 3 );
                    lstErrorMsg.Add( ErrorMsg.CustomError( excelName, strSheetName, address, "对行Id重复操作" ) );
                }
            }
            inExcel.Close();
            return lstErrorMsg.Count == 0;
        }
    }
}
