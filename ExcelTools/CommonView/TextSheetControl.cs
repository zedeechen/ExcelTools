using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelTools
{
    class TextSheetControl : SheetControl
    {
        public bool Create( string strExcelPath, string strSheetName, bool bExistLineFour, out List<string> lstErrorMsg, string colId, string colName )
        {
            lstErrorMsg = new List<string>();
            YYExcel outExcel = new YYExcel();
            try
            {
                outExcel.Create( strExcelPath, strSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            }
            catch (System.Exception ex)
            {
                lstErrorMsg.Add( ex.Message );
                return false;
            }

            int row = 3;
            outExcel.setCellValue( 1, 1, "ID" );
            outExcel.setCellValue( 1, 2, "名称" );
            outExcel.setCellValue( 3, 1, "101" );
            outExcel.setCellValue( 3, 2, "102" );
            if ( bExistLineFour )
            {
                outExcel.setCellValue( 4, 1, colId );
                outExcel.setCellValue( 4, 2, colName );
                row++;
            }
            outExcel.setRangeAlignCenter( 1, 1, row, 2 );
            outExcel.setRangeInteriorColor( 1, 1, row, 2, 37 );
            for ( int i = 1; i <= row; i++ )
                for ( int j = 1; j <= 2; j++ )
                    outExcel.setCellBorder( i, j );

            try
            {
                outExcel.SaveAs( strExcelPath );
            }
            catch (System.Exception ex)
            {
                lstErrorMsg.Add( ex.Message );
                outExcel.Close();
                return false;
            }
            
            outExcel.Close();
            return true;
        }

        public void Create( string strExcelPath, string strSheetName, bool bExistLineFour, string colId = "Id", string colName = "Name" )
        {
            List<string> lstErrorMsg;
            Create( strExcelPath, strSheetName, bExistLineFour, out lstErrorMsg, colId, colName );
        }

        public bool Update( string strExcelPath, string strSheetName, bool bExistLineFour, TextSheet textSheet, out List<string> lstErrorMsg )
        {
            lstErrorMsg = new List<string>();
            object[,] values = textSheet.ToObject();

            if ( values == null ) return false;

            YYExcel outExcel = new YYExcel();
            try
            {
                outExcel.Open( strExcelPath, strSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            }
            catch (System.Exception ex)
            {
                lstErrorMsg.Add( ex.Message );
                return false;
            }
            
            outExcel.setRangeValue( bExistLineFour ? 5 : 4, 1, values );

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

        public bool Read( string strExcelPath, string strSheetName, bool bExistLineFour, out TextSheet textSheet, out List<string> lstErrorMsg )
        {
            textSheet = new TextSheet();
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
            int wRowCount = inExcel.GetRowsCount();
            
            int startRow = bExistLineFour ? 5 : 4;

            int txtIdxLen = m_wTblIdLen + m_wColIdLen + m_wRowIdLen;
            for ( int i = startRow; i <= wRowCount; i++ )
            {
                string address = inExcel.getCellAddress( i, 1 );
                string strIndex = inExcel.getCellValue( i, 1 );
                string strName = inExcel.getCellValue( i, 2 );
                int wIndex;

                if ( strIndex == string.Empty ) continue;
                bool   res        = Int32.TryParse( strIndex, out wIndex );

                // Id检查
                if ( !res || strIndex.Length > txtIdxLen )
                {
                    lstErrorMsg.Add( ErrorMsg.FormatError( excelName, address, "列Id应为小于" + Util.TenPow( txtIdxLen ).ToString() + "的整数" ) );
                    continue;
                }

                if ( !textSheet.dictText.ContainsKey( wIndex ) )
                    textSheet.dictText.Add( wIndex, strName );
                else
                    lstErrorMsg.Add( ErrorMsg.Error( excelName, "列Id重复" ) );
            }

            inExcel.Close();

            return lstErrorMsg.Count == 0;
        }

        public bool SetItem( YYExcel outExcel, int row, int id, string name, int color )
        {
            outExcel.setCellValue( row, 1, id.ToString() );
            outExcel.setCellValue( row, 2, name );

            outExcel.setRangeInteriorColor( row, 1, row, 2, color );
            return true;
        }

        public bool AddItem( YYExcel outExcel, int row, int id, string name, int color )
        {
            outExcel.InsertRow( row );

            outExcel.setCellValue( row, 1, id.ToString() );
            outExcel.setCellValue( row, 2, name );

            outExcel.setRangeInteriorColor( row, 1, row, 2, color );
            return true;
        }

    }
}
