using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelTools
{
    public partial class FunctionSheetDiff
    {
        void PushNewProcess()
        {
            FunctionSheetControl funcControl = new FunctionSheetControl();
            YYExcel outExcel = new YYExcel();
            outExcel.Open( m_strNewExcelPath, m_strSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            foreach ( int index in m_lstSelectedIndexs )
            {
                int rowIndex = Int32.Parse( m_oldExcelCache[index].Name );
                bool IsAscending = m_sheetDiffInfo.newSheet.bIsAscending;
                int row;

                if ( IsAscending )
                    row = funcControl.GetRowByIndex( outExcel, rowIndex, m_bExsistLineFour ? 5 : 4, outExcel.GetRowsCount(), 1 );
                else
                {
                    if ( m_sheetDiffInfo.newSheet.itemPos.ContainsKey( rowIndex ) )
                    {
                        row = m_sheetDiffInfo.newSheet.itemPos[rowIndex];
                    }
                    else
                    {
                        row = outExcel.GetRowsCount();
                        while ( outExcel.getCellValue( row, 1 ) == "" )
                            row--;
                        row++;
                    }
                }

                if ( !m_sheetDiffInfo.newSheet.itemPos.ContainsKey( rowIndex ) )
                    outExcel.InsertRow( row );

                {
                    object[,] values = new object[1, outExcel.GetColumnsCount()];
                    outExcel.getRangeValue( row, 1, row, outExcel.GetColumnsCount(), ref values );
                    foreach ( int colIndex in m_sheetDiffInfo.oldSheet.headers.Keys )
                    {
                        if ( m_sheetDiffInfo.newSheet.headers.ContainsKey( colIndex ) )
                        {
                            int col = m_sheetDiffInfo.newSheet.headers[colIndex].titlePos;
                            values[0, col - 1] = m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value;
                        }
                    }

                    outExcel.setRangeValue( row, 1, values );
                }

                {
                    lvwOldExcelDiff.EnsureVisible( index );
                    lvwNewExcelDiff.EnsureVisible( index );
                    m_newExcelCache[index] = m_oldExcelCache[index];
                    m_oldExcelCache[index].UseItemStyleForSubItems = true;
                    m_oldExcelCache[index].BackColor = cPushNew;
                    m_oldExcelCache[index].ForeColor = Color.White;
                    m_newExcelCache[index].UseItemStyleForSubItems = true;
                    m_newExcelCache[index].BackColor = cPushNew;
                    m_newExcelCache[index].ForeColor = Color.White;                    
                }
            }

            outExcel.SaveAs( m_strNewExcelPath );
            outExcel.Close();
            
            //this.Invoke( (UpdateButtonStateDelegate)delegate()
            //{
                btnPushNew.Enabled = true;
                btnPushOld.Enabled = true;
            //} );
        }
    }
}
