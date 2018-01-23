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
        void FuncVersionDiffMain()
        {
            string[] oldExcels = Directory.GetFiles( m_strFuncVersionDiffOldExcelDirPath, "*.xlsx", SearchOption.TopDirectoryOnly );
            string[] newExcels = Directory.GetFiles( m_strFuncVersionDiffNewExcelDirPath, "*.xlsx", SearchOption.TopDirectoryOnly );

            Dictionary<string, string> dictExcels = new Dictionary<string, string>();
            for ( int i = 0; i < oldExcels.Length; i++ )
            {
                string path = oldExcels[i];
                string name = Path.GetFileName( path );
                dictExcels.Add( name, path );
            }

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

            lvwFuncVersionDiffResult.Items.Clear();
            lvwFuncVersionDiffResult.BeginUpdate();

            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                string name = Path.GetFileName( path );
                ListViewItem lvi = new ListViewItem( ( i + 1 ).ToString() );
                lvi.Name = name;
                lvi.SubItems.Add( name );
                lvi.SubItems.Add( "未处理" );
                lvwFuncVersionDiffResult.Items.Add( lvi );
            }

            lvwFuncVersionDiffResult.EndUpdate();

            btnFuncVersionDiff.Enabled = false;
            Thread processThread = new Thread( new ParameterizedThreadStart( FuncVersionDiffProcess ) );
            processThread.Start( excels );
        }

        private void FuncVersionDiffProcess( object o )
        {
            string[] excels = o as string[];

            List<string> errorMsgs;
            // 检查
            bool chkPass = FuncVersionDiffCheck( excels, out errorMsgs );
            if ( !chkPass )
            {
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
            }
            else
            {
                // 比较
                m_dictFuncVersionDiff = new Dictionary<string, SheetDiffInfo>();
                for ( int i = 0; i < excels.Length; i++ )
                {
                    string path = excels[i];
                    FuncVersionDiffSheetDiff( path );
                }
                this.Invoke( (UpdateFunctionVersionLvwMouseDoubleClickEvent)delegate()
                {
                    this.lvwFuncVersionDiffResult.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler( this.lvwFuncVersionDiffResult_MouseDoubleClick );
                } );
            }
            
            this.Invoke( (UpdateButtonStateDelegate)delegate()
            { 
                btnFuncVersionDiff.Enabled = true;
            } );
        }

        public bool FuncVersionDiffCheck( string[] excels, out List<string> errorMsgs )
        {
            errorMsgs = new List<string>();
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                string name = Path.GetFileName( path );
                bool bWithTxtCol = false;
                bool bIsAscending = false;
                List<string> lstErrorMsg;

                FunctionSheetControl funcControl = new FunctionSheetControl();
                bool chkPass = funcControl.Check( path, m_strFuncVersionDiffSheetName, m_bFuncVersionDiffExistLineFour, out bWithTxtCol, out bIsAscending, out lstErrorMsg );
                errorMsgs.AddRange( lstErrorMsg );
                this.Invoke( (UpdateFunctionCSVConvertResultDelegate)delegate( int idx, bool res )
                {
                    switch ( res )
                    {
                        case true:
                            lvwFuncVersionDiffResult.Items[idx].SubItems[2].Text = "检查通过";
                            break;

                        default:
                            lvwFuncVersionDiffResult.Items[idx].SubItems[2].Text = "检查未通过";
                            break;
                    }

                    lvwFuncVersionDiffResult.EnsureVisible( idx );
                }, i, chkPass && bIsAscending );

                if ( !bIsAscending )
                {
                    errorMsgs.Add( ErrorMsg.Error( name, "行Id未严格递增" ) );
                }
            }
            return errorMsgs.Count == 0;
        }

        public void FuncVersionDiffSheetDiff( string strExcelPath )
        {
            string name = Path.GetFileName( strExcelPath );
            string oldPath = m_strFuncVersionDiffOldExcelDirPath + "\\" + name;
            string newPath = m_strFuncVersionDiffNewExcelDirPath + "\\" + name;
            int wStatus = 0; // 0 相同 1 新建 2 修改 3 删除

            if ( !File.Exists( oldPath ) )
            {
                wStatus = 1;
            }
            else if ( !File.Exists( newPath ) )
            {
                wStatus = 3;
            }
            else
            {
                if ( !FuncSheetEquals( oldPath, newPath, m_strFuncVersionDiffSheetName ) )
                {
                    wStatus = 2;
                }
            }

            this.Invoke( (UpdateFunctionVersionDiffResultDelegate)delegate( string tbl, int res )
            {
                switch ( res )
                {
                    case 0:
                        //lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].SubItems[2].Text = "相同";
                        lvwFuncVersionDiffResult.Items.RemoveByKey( tbl );
                        break;
                    case 1:
                        lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].SubItems[2].Text = "新建";
                        break;
                    case 2:
                        lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].SubItems[2].Text = "修改";
                        break;
                    case 3:
                        lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].SubItems[2].Text = "删除";
                        break;
                    default:
                        break;
                }
                if ( lvwFuncVersionDiffResult.Items.Find( tbl, false ).Count() != 0 )
                    lvwFuncVersionDiffResult.Items.Find( tbl, false )[0].EnsureVisible();
            }, name, wStatus );
        }

        public bool FuncSheetEquals( string strOldExcelPath, string strNewExcelPath, string strSheetName )
        {
            SheetDiffInfo sheetDiff = new SheetDiffInfo();
            FunctionSheetControl funcControl = new FunctionSheetControl();
            FunctionSheet oldFuncSheet;
            FunctionSheet newFuncSheet;

            string name = Path.GetFileName( strOldExcelPath );

            List<string> lstErrorMsg;
            funcControl.Read( strOldExcelPath, strSheetName, m_bFuncVersionDiffExistLineFour, out oldFuncSheet, out lstErrorMsg );
            funcControl.Read( strNewExcelPath, strSheetName, m_bFuncVersionDiffExistLineFour, out newFuncSheet, out lstErrorMsg );

            sheetDiff.oldSheet = oldFuncSheet;
            sheetDiff.newSheet = newFuncSheet;
            Dictionary<int, FunctionSheet.TitleConfig> headers = sheetDiff.headers;
            Dictionary<int, int> items = sheetDiff.items;
            SortedDictionary<int, int> diffItems = sheetDiff.diffItems;

            foreach ( int ColIndex in oldFuncSheet.headers.Keys )
            {
                headers.Add( ColIndex, oldFuncSheet.headers[ColIndex] );
            }
            foreach ( int ColIndex in newFuncSheet.headers.Keys )
            {
                if ( !headers.ContainsKey( ColIndex ) )
                    headers.Add( ColIndex, newFuncSheet.headers[ColIndex] );
            }

            foreach ( int RowIndex in oldFuncSheet.itemPos.Keys )
            {
                items.Add( RowIndex, oldFuncSheet.itemPos[RowIndex] );
            }
            foreach ( int RowIndex in newFuncSheet.itemPos.Keys )
            {
                if ( !items.ContainsKey( RowIndex ) )
                    items.Add( RowIndex, newFuncSheet.itemPos[RowIndex] );
            }

            if ( !m_dictFuncVersionDiff.ContainsKey( name ) )
                m_dictFuncVersionDiff.Add( name, sheetDiff );
            else
                m_dictFuncVersionDiff[name] = sheetDiff;

            // diff compare

            foreach ( int RowIndex in items.Keys )
            {
                if ( !oldFuncSheet.itemPos.ContainsKey( RowIndex ) || !newFuncSheet.itemPos.ContainsKey( RowIndex ) )
                {
                    diffItems.Add( RowIndex, items[RowIndex] );
                    continue;
                }
                foreach ( int ColIndex in headers.Keys )
                {
                    if ( !oldFuncSheet.headers.ContainsKey( ColIndex ) || !newFuncSheet.headers.ContainsKey( ColIndex ) )
                    {
                        diffItems.Add( RowIndex, items[RowIndex] );
                        break;
                    }
                    if ( oldFuncSheet.cells[RowIndex][ColIndex].value != newFuncSheet.cells[RowIndex][ColIndex].value )
                    {
                        diffItems.Add( RowIndex, items[RowIndex] );
                        break;
                    }
                }
            }

            return diffItems.Count == 0;
        }
    }
}
