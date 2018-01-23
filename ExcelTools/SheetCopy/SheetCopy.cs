using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace ExcelTools
{
    partial class Main
    {
        void SheetCopyMain()
        {
            string[] oldExcels = Directory.GetFiles( m_strSheetCopyOldExcelDirPath, "*.xlsx", SearchOption.TopDirectoryOnly );
            string[] newExcels = Directory.GetFiles( m_strSheetCopyNewExcelDirPath, "*.xlsx", SearchOption.TopDirectoryOnly );

            Dictionary<string, string> dictExcels = new Dictionary<string, string>();
            //for ( int i = 0; i < oldExcels.Length; i++ )
            //{
            //    string path = oldExcels[i];
            //    string name = Path.GetFileName( path );
            //    dictExcels.Add( name, path );
            //}

            for ( int i = 0; i < newExcels.Length; i++ )
            {
                string path = newExcels[i];
                string name = Path.GetFileName( path );
                if ( !dictExcels.ContainsKey( name ) )
                    dictExcels.Add( name, path );
            }

            string[] excels = dictExcels.Values.ToArray();

            // 提示已打开的输入文件
            string ret = Util.GetOpenedExcelList( excels );
            if ( ret != string.Empty )
            {
                MessageBox.Show( ret + "\n如有需要，请保存后确定", "已打开Excel列表" );
            }
            // 从Array中移除已打开文件的副本
            excels = Array.FindAll( excels, Util.IsExcelOpened );

            lvwSheetCopyResult.Items.Clear();
            lvwSheetCopyResult.BeginUpdate();

            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                string name = Path.GetFileName( path );
                ListViewItem lvi = new ListViewItem( ( i + 1 ).ToString() );
                lvi.Name = name;
                lvi.SubItems.Add( name );
                lvi.SubItems.Add( "未处理" );
                lvwSheetCopyResult.Items.Add( lvi );
            }

            lvwSheetCopyResult.EndUpdate();

            btnCopySheet.Enabled = false;
            Thread processThread = new Thread( new ParameterizedThreadStart( SheetCopyProcess ) );
            processThread.Start( excels );
        }

        private void SheetCopyProcess( object o )
        {
            string[] excels = o as string[];

            List<string> errorMsgs;
            // 拷贝
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                CopySheet( path );
            }
            
            this.Invoke( (UpdateButtonStateDelegate)delegate()
            { 
                btnCopySheet.Enabled = true;
            } );
        }

        public void CopySheet( string strExcelPath )
        {
            string name = Path.GetFileName( strExcelPath );
            string oldPath = m_strSheetCopyOldExcelDirPath + "\\" + name;
            string newPath = m_strSheetCopyNewExcelDirPath + "\\" + name;
            int wStatus = 0; // 0 不存在需要拷贝的Sheet 1 不存在需要拷贝的Excel文件 2 修改成功

            if ( !File.Exists( oldPath ) )
            {
                wStatus = 1;
            }
            else
            {
                if ( SheetCopy( oldPath, newPath, m_strSheetCopySheetName ) )
                {
                    wStatus = 2;
                }
            }

            this.Invoke( (UpdateFunctionVersionDiffResultDelegate)delegate( string tbl, int res )
            {
                switch ( res )
                {
                    //case 0:
                    //    //lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].SubItems[2].Text = "相同";
                    //    lvwSheetCopyResult.Items.RemoveByKey( tbl );
                    //    break;
                    case 0:
                        lvwSheetCopyResult.Items.Find( tbl, false )[0].SubItems[2].Text = "不存在Sheet";
                        break;
                    case 1:
                        lvwSheetCopyResult.Items.Find( tbl, false )[0].SubItems[2].Text = "不存在Excel文件";
                        break;
                    case 2:
                        lvwSheetCopyResult.Items.Find( tbl, false )[0].SubItems[2].Text = "已拷贝";
                        break;
                    default:
                        break;
                }
                if ( lvwSheetCopyResult.Items.Find( tbl, false ).Count() != 0 )
                    lvwSheetCopyResult.Items.Find( tbl, false )[0].EnsureVisible();
            }, name, wStatus );
        }

        public bool SheetCopy( string path_in, string path_out, string sheet_name )
        {
            ExcelTools.YYExcel inExcel = new ExcelTools.YYExcel();
            ExcelTools.YYExcel outExcel = new ExcelTools.YYExcel();
            try
            {
                inExcel.Open( path_in, sheet_name, ExcelTools.YYExcel.Authority.A_READ_ONLY );
            }
            catch ( System.Exception ex )
            {
                //lstErrorMsg.Add( ErrorMsg.OpenError( excelName, strSheetName, ex.Message ) );
                return false;
            }

            List<string> lstSheet = new List<string>();
            if ( !File.Exists( path_out ) )
            {
                return false;
            }

            try
            {
                outExcel.GetSheetsName( path_out, out lstSheet, ExcelTools.YYExcel.Authority.A_READ_ONLY );
            }
            catch ( System.Exception ex )
            {
                //lstErrorMsg.Add( ErrorMsg.OpenError( excelName, strSheetName, ex.Message ) );
                return false;
            }

            bool bHasSheet = false;
            foreach ( var sheet in lstSheet )
            {
                if ( sheet == sheet_name )
                {
                    bHasSheet = true;
                }
            }

            if ( !bHasSheet )
            {
                try
                {
                    outExcel.Add( path_out, sheet_name, ExcelTools.YYExcel.Authority.A_READ_AND_WRITE );
                }
                catch ( System.Exception ex )
                {
                    //lstErrorMsg.Add( ErrorMsg.OpenError( excelName, strSheetName, ex.Message ) );
                    return false;
                }
            }
            else
            {
                try
                {
                    outExcel.Open( path_out, sheet_name, ExcelTools.YYExcel.Authority.A_READ_AND_WRITE );
                }
                catch ( System.Exception ex )
                {
                    //lstErrorMsg.Add( ErrorMsg.OpenError( excelName, strSheetName, ex.Message ) );
                    return false;
                }
            }
            
            int row = inExcel.GetRowsCount();
            int col = inExcel.GetColumnsCount();
            Console.WriteLine( row );
            Console.WriteLine( col );

            object[,] values = new object[row, col];
            inExcel.getRangeValue( 1, 1, row, col, ref values );

            outExcel.setRangeValue( 1, 1, values );

            inExcel.Close();
            outExcel.SaveAs( path_out );
            //outExcel.Save();
            outExcel.Close();

            return true;
        }
    }
}
