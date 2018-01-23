using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;

namespace ExcelTools
{
    partial class Main
    {
        void FuncCustomizerMain()
        {
            string[] fileNames = Directory.GetFiles( m_strFuncCustomizerInputFunctionDirPath, "*.xlsx", SearchOption.TopDirectoryOnly );

            // 提示已打开的输入文件
            string ret = Util.GetOpenedExcelList( fileNames );
            if ( ret != string.Empty )
            {
                MessageBox.Show( ret + "\n如有需要，请保存后确定", "已打开Excel列表" );
            }
            // 从Array中移除已打开文件的副本
            fileNames = Array.FindAll( fileNames, Util.IsExcelOpened );

            // 结果列表初始化
            lvwFuncCustomizerResult.Items.Clear();
            lvwFuncCustomizerResult.BeginUpdate();
            int fLength = fileNames.Length;
            for ( int i = 0; i < fLength; i++ )
            {
                string name = Path.GetFileName( fileNames[i] );

                string path = Path.Combine( m_strFuncCustomizerOutputFunctionDirPath, name );
                while ( Util.IsFileInUse( path ) )
                {
                    MessageBox.Show( "请先关闭" + path );
                }
               
                ListViewItem lvi = new ListViewItem( ( i + 1 ).ToString() );
                lvi.SubItems.Add( name );
                lvi.SubItems.Add( "尚未订制" );
                lvwFuncCustomizerResult.Items.Add( lvi );
            }
            lvwFuncCustomizerResult.EndUpdate();

            lblFuncCustomizerReadResult.Text = "准备。。";
            btnFuncCustomize.Enabled = false;

            // 处理开始
            Thread processThread = new Thread( new ParameterizedThreadStart( FuncCustomizeProcess ) );
            processThread.Start( fileNames );
        }

        private void FuncCustomizeProcess( object o )
        {
            List<string> errorMsgs = new List<string>();
            List<string> lstErrorMsg;
            string[] excels = o as string[];
            //////////////////////////////////////////////////////////////////////////
            // File Hash
            Dictionary<string, string> dictNameHash = new Dictionary<string, string>();
            FuncCustomFileHash( excels, ref dictNameHash );

            //////////////////////////////////////////////////////////////////////////
            // 读订制表

            CustomSheetControl customControl = new CustomSheetControl();
            CustomSheet customSheet = new CustomSheet();
            
            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncCustomizerReadResult.Text = "读取订制表中。。";
            } );

            bool customOK = customControl.Read( m_strFuncCustomizerCustomPath, m_strFuncCustomizerCustomSheetName, dictNameHash, out customSheet, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );
            this.Invoke( (UpdateLabelStateDelegateWithResult)delegate( bool res )
            {
                switch ( res )
                {
                    case true:
                        lblFuncCustomizerReadResult.Text = "订制表读取成功";
                        break;
                    default:
                        lblFuncCustomizerReadResult.Text = "订制表读取失败";
                        break;
                }
            }, customOK );

            //////////////////////////////////////////////////////////////////////////
            // 检查
            List<bool> lstIsAscending;
            bool chkPass = FuncCustomCheck( excels, customSheet, out lstIsAscending, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );

            if ( errorMsgs.Count != 0 )
            {
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
                this.Invoke( (UpdateButtonStateDelegate)delegate()
                {
                    btnFuncCustomize.Enabled = true;
                } );
                return;
            }

            //////////////////////////////////////////////////////////////////////////
            // 订制
            if ( m_strFuncCustomizerOldFunctionDirPath != string.Empty && Directory.Exists( m_strFuncCustomizerOldFunctionDirPath )
                       && m_strFuncCustomizerCustomClashPath != string.Empty )
            {
                if ( File.Exists( m_strFuncCustomizerCustomClashPath ) && Util.IsFileInUse( m_strFuncCustomizerCustomClashPath ) )
                {
                    this.Invoke( (CreateMessageBoxDelegate)delegate()
                    {
                        while ( Util.IsFileInUse( m_strFuncCustomizerCustomClashPath ) )
                        {
                            MessageBox.Show( "请关闭" + m_strFuncCustomizerCustomClashPath );
                        }
                    } );
                }
                customControl.Create( m_strFuncCustomizerCustomClashPath, m_strFuncCustomizerCustomSheetName, out lstErrorMsg );
            }

            bool customPass = FuncCustomize( excels, customSheet, lstIsAscending, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );

            if ( errorMsgs.Count != 0 )
            {
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
            }

            this.Invoke( (UpdateButtonStateDelegate)delegate()
            {
                btnFuncCustomize.Enabled = true;
            } );
        }

        private void FuncCustomFileHash( string[] excels, ref Dictionary<string, string> dict )
        {
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                string name = Path.GetFileName( path );
                string raw  = name;

                if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
                {
                    raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
                }
                raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

                dict.Add( raw, name );
            }
        }

        private bool FuncCustomCheck( string[] excels, CustomSheet customSheet, out List<bool> lstIsAscending, out List<string> errorMsgs )
        {
            errorMsgs = new List<string>();
            lstIsAscending = new List<bool>();
            FunctionSheetControl funcControl = new FunctionSheetControl();
            List<string> lstErrorMsg;
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                string name = Path.GetFileName( path );
                bool bFoundError = false;

                string raw  = name;
                if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
                {
                    raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
                }
                raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

                bool bIsAscending;
                bool chkPass = funcControl.Check( path, m_strFuncCustomizerFuncSheetName, m_bFuncCustomizerExistLineFour, out bIsAscending, out lstErrorMsg );
                if ( !bIsAscending )
                {
                    errorMsgs.Add( ErrorMsg.Error( name, "行Id未严格递增" ) );
                }
                lstIsAscending.Add( bIsAscending );
                errorMsgs.AddRange( lstErrorMsg );
                
                do 
                {
                    if ( !chkPass || !bIsAscending )
                    {
                        bFoundError = true;
                        break;
                    }
                    if ( !customSheet.dictCustomItem.ContainsKey( raw ) ) break;

                    FunctionSheet funcSheet;
                    funcControl.Read( path, m_strFuncCustomizerFuncSheetName, m_bFuncCustomizerExistLineFour, out funcSheet, out lstErrorMsg );
                    errorMsgs.AddRange( lstErrorMsg );
                    foreach ( int rowId in customSheet.dictCustomItem[raw].Keys )
                    {
                        if ( customSheet.dictCustomItem[raw][rowId].status == 1 ) // 修改 
                        {
                            if ( !funcSheet.itemPos.ContainsKey( rowId ) )
                            {
                                bFoundError = true;
                                errorMsgs.Add( ErrorMsg.CustomError( name, m_strFuncCustomizerCustomSheetName, rowId, "修改项不存在" ) );
                            }
                        }
                        else // 新增
                        {
                            if ( funcSheet.itemPos.ContainsKey( rowId ) )
                            {
                                bFoundError = true;
                                errorMsgs.Add( ErrorMsg.CustomError( name, m_strFuncCustomizerCustomSheetName, rowId, "新增项不存在" ) );
                            }
                        }
                        foreach ( int colId in customSheet.dictCustomItem[raw][rowId].dictOpCell.Keys )
                        {
                            if ( !funcSheet.headers.ContainsKey( colId ) )
                            {
                                bFoundError = true;
                                errorMsgs.Add( ErrorMsg.CustomError( name, m_strFuncCustomizerCustomSheetName, rowId, colId, "列Id不存在" ) );
                            }
                        }
                    }
                    
                } while ( false );

                this.Invoke( (UpdateFunctionCustomResultDelegate)delegate( int idx, bool res )
                {
                    switch ( res )
                    {
                        case true:
                            lvwFuncCustomizerResult.Items[idx].SubItems[2].Text = "检查通过";
                            break;
                        default:
                            lvwFuncCustomizerResult.Items[idx].SubItems[2].Text = "检查未通过";
                            break;
                    }

                    lvwFuncCustomizerResult.Items[idx].EnsureVisible();
                }, i, !bFoundError );
            }
            if ( errorMsgs.Count == 0 )
                return true;
            else
                return false;
        }

        private bool FuncCustomize( string[] excels, CustomSheet customSheet, List<bool> lstIsAscending, out List<string> errorMsgs )
        {
            errorMsgs = new List<string>();

            for ( int i = 0; i < excels.Length; i++ )
            {
                List<string> lstErrorMsg;
                string path = excels[i];
                string name = Path.GetFileName( path );
                string raw  = name;
                if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
                {
                    raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
                }
                raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

                bool updateOk = true;
                bool checkOk = true;

                if ( !DiffCheck( customSheet, m_strFuncCustomizerInputFunctionDirPath + "\\" + name,
                    m_strFuncCustomizerOutputFunctionDirPath + "\\" + name, m_strFuncCustomizerFuncSheetName, lstIsAscending[i], out lstErrorMsg ) )
                {
                    File.Copy( m_strFuncCustomizerInputFunctionDirPath + "\\" + name, m_strFuncCustomizerOutputFunctionDirPath + "\\" + name, true );

                    if ( customSheet.dictCustomItem.ContainsKey( raw ) )
                    {
                        if ( m_strFuncCustomizerOldFunctionDirPath != string.Empty && Directory.Exists( m_strFuncCustomizerOldFunctionDirPath )
                            && m_strFuncCustomizerCustomClashPath != string.Empty )
                        {
                            checkOk = ClashCheck( customSheet,
                                        m_strFuncCustomizerOldFunctionDirPath + "\\" + name, m_strFuncCustomizerInputFunctionDirPath + "\\" + name, m_strFuncCustomizerFuncSheetName,
                                        m_strFuncCustomizerCustomClashPath, m_strFuncCustomizerCustomSheetName, out lstErrorMsg );
                            errorMsgs.AddRange( lstErrorMsg );
                        }
                        updateOk = FuncCustomUpdateFunc( customSheet, m_strFuncCustomizerOutputFunctionDirPath + "\\" + name, m_strFuncCustomizerFuncSheetName, lstIsAscending[i], out lstErrorMsg );
                        errorMsgs.AddRange( lstErrorMsg );
                    }
                }

                this.Invoke( (UpdateFunctionCustomResultDelegate)delegate( int idx, bool res )
                {
                    switch ( res )
                    {
                        case true:
                            lvwFuncCustomizerResult.Items[idx].SubItems[2].Text = "完成";
                            break;
                        default:
                            lvwFuncCustomizerResult.Items[idx].SubItems[2].Text = "未完成";
                            break;
                    }

                    lvwFuncCustomizerResult.Items[idx].EnsureVisible();
                }, i, updateOk );
            }

            if ( errorMsgs.Count == 0 )
                return true;
            else
                return false;
        }

        private bool FuncCustomUpdateFunc( CustomSheet customSheet, string strExcelPath, string strSheetName, bool IsAscending, out List<string> lstErrorMsg )
        {
            lstErrorMsg = new List<string>();
            string name = Path.GetFileName( strExcelPath );
            string raw  = name;
            if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
            {
                raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
            }
            raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

            FunctionSheetControl funcControl = new FunctionSheetControl();
            FunctionSheet funcSheet = new FunctionSheet();
            funcControl.Read( strExcelPath, strSheetName, m_bFuncCustomizerExistLineFour, out funcSheet, out lstErrorMsg );

            YYExcel outExcel = new YYExcel();
            try
            {
                outExcel.Open( strExcelPath, strSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ErrorMsg.OpenError( name, strSheetName, ex.Message ) );
                return false;
            }

            foreach ( int rowIndex in customSheet.dictCustomItem[raw].Keys )
            {
                CustomSheet.Item item = customSheet.dictCustomItem[raw][rowIndex];

                int rowPos;

                if ( IsAscending )
                    rowPos = funcControl.GetRowByIndex( outExcel, rowIndex, m_bFuncCustomizerExistLineFour ? 5 : 4, outExcel.GetRowsCount(), 1 );
                else
                {
                    if ( !funcSheet.itemPos.ContainsKey(rowIndex) )
                    {
                        rowPos = outExcel.GetRowsCount();
                        while ( outExcel.getCellValue( rowPos, 1 ) == "" )
                            rowPos--;
                        rowPos++;
                    }
                    else
                        rowPos = funcSheet.itemPos[rowIndex];
                }

                // 新增插入行
                if ( item.status == 2 )
                {
                    outExcel.InsertRow( rowPos );
                    outExcel.setCellValue( rowPos, 1, rowIndex.ToString() );
                    outExcel.setRangeInteriorColor( rowPos, 1, rowPos, outExcel.GetColumnsCount(), Convert.ToInt32( m_dFuncCustomizerNewItemColor ) );
                }
                foreach ( int colIndex in item.dictOpCell.Keys )
                {
                    try
                    {
                        outExcel.setCellValue( rowPos, funcSheet.headers[colIndex].titlePos, item.dictOpCell[colIndex] );
                    }
                    catch ( System.Exception ex )
                    {
                        string address = outExcel.getCellAddress( rowPos, colIndex );
                        lstErrorMsg.Add( ErrorMsg.FormatError( raw, address, "该单元格不存在:" + ex.ToString() ) );
                        continue;
                    }

                    if ( item.status == 1 )
                    {
                        outExcel.setCellInteriorColor( rowPos, funcSheet.headers[colIndex].titlePos, Convert.ToInt32( m_dFuncCustomizerOldItemColor ) );
                    }
                }
            }

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

        // return true if no diff
        private bool DiffCheck( CustomSheet customSheet, string strInExcelPath, string strOutExcelPath, string strSheetName, bool IsAscending, out List<string> lstErrorMsg )
        {
            lstErrorMsg = new List<string>();
            YYExcel inExcel = new YYExcel();
            YYExcel outExcel = new YYExcel();

            string name = Path.GetFileName( strInExcelPath );
            string raw  = name;
            if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
            {
                raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
            }
            raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

            if ( !File.Exists( strOutExcelPath ) )
            {
                return false;
            }

            try
            {
                outExcel.Open( strInExcelPath, strSheetName, YYExcel.Authority.A_READ_ONLY );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ErrorMsg.OpenError( name, strSheetName, ex.Message ) );
                return false;
            }

            try
            {
                outExcel.Open( strOutExcelPath, strSheetName, YYExcel.Authority.A_READ_ONLY );
            }
            catch ( System.Exception ex )
            {
                lstErrorMsg.Add( ErrorMsg.OpenError( name, strSheetName, ex.Message ) );
                return false;
            }

            FunctionSheetControl funcControl = new FunctionSheetControl();
            FunctionSheet funcInSheet = new FunctionSheet();
            FunctionSheet funcOutSheet = new FunctionSheet();
            FunctionSheet diffFuncSheet = new FunctionSheet();
            funcControl.Read( strInExcelPath, strSheetName, m_bFuncCustomizerExistLineFour, out funcInSheet, out lstErrorMsg );
            funcControl.Read( strOutExcelPath, strSheetName, m_bFuncCustomizerExistLineFour, out funcOutSheet, out lstErrorMsg );

            foreach ( int rowId in funcOutSheet.cells.Keys )
            {
                foreach ( int colId in funcOutSheet.cells[rowId].Keys )
                {
                    bool bNewItem = false;
                    bool bEdtItem = false;
                    if ( !funcInSheet.cells.ContainsKey( rowId ) )
                    {
                        bNewItem = true;
                    }
                    else
                    {
                        if ( !funcInSheet.cells[rowId].ContainsKey( colId ) )
                        {
                            bEdtItem = true;
                        }
                        else
                        {
                            if ( funcInSheet.cells[rowId][colId].value != funcOutSheet.cells[rowId][colId].value )
                            {
                                bEdtItem = true;
                            }
                        }
                    }

                    if ( bNewItem )
                    {
                        if ( !diffFuncSheet.cells.ContainsKey( rowId ) )
                        {
                            Dictionary<int, Node> dict = new Dictionary<int, Node>();
                            diffFuncSheet.cells.Add( rowId, dict );
                        }
                        diffFuncSheet.cells[rowId].Add( colId, new Node( 2, funcOutSheet.cells[rowId][colId].value ) );
                    }
                    if ( bEdtItem )
                    {
                        if ( !diffFuncSheet.cells.ContainsKey( rowId ) )
                        {
                            Dictionary<int, Node> dict = new Dictionary<int, Node>();
                            diffFuncSheet.cells.Add( rowId, dict );
                        }
                        diffFuncSheet.cells[rowId].Add( colId, new Node( 1, funcOutSheet.cells[rowId][colId].value ) );
                    }
                }
            }

            bool bRet = true;
            if ( customSheet.dictCustomItem.ContainsKey( raw ) )
            {
                foreach ( int rowIndex in customSheet.dictCustomItem[raw].Keys )
                {
                    foreach ( int colIndex in customSheet.dictCustomItem[raw][rowIndex].dictOpCell.Keys )
                    {
                        if ( diffFuncSheet.cells.ContainsKey( rowIndex ) && diffFuncSheet.cells[rowIndex].ContainsKey( colIndex ) )
                        {
                            if ( customSheet.dictCustomItem[raw][rowIndex].status == 2 && diffFuncSheet.cells[rowIndex][colIndex].status == 2
                                && customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] != diffFuncSheet.cells[rowIndex][colIndex].value )
                            {
                                bRet = false;
                            }
                            else if ( customSheet.dictCustomItem[raw][rowIndex].status == 1 && diffFuncSheet.cells[rowIndex][colIndex].status == 1
                                        && customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] != diffFuncSheet.cells[rowIndex][colIndex].value )
                            {
                                bRet = false;
                            }
                        }
                        else if ( funcOutSheet.cells.ContainsKey( rowIndex ) && funcOutSheet.cells[rowIndex].ContainsKey( colIndex ) &&
                            funcOutSheet.cells[rowIndex][colIndex].value != customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] )
                        {
                            bRet = false;
                        }
                    }
                }

                foreach ( int rowIndex in diffFuncSheet.cells.Keys )
                {
                    foreach ( int colIndex in diffFuncSheet.cells[rowIndex].Keys )
                    {
                        if ( customSheet.dictCustomItem[raw].ContainsKey( rowIndex ) && customSheet.dictCustomItem[raw][rowIndex].dictOpCell.ContainsKey( colIndex ) )
                        {
                            if ( customSheet.dictCustomItem[raw][rowIndex].status == 2 && diffFuncSheet.cells[rowIndex][colIndex].status == 2
                                && customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] != diffFuncSheet.cells[rowIndex][colIndex].value )
                            {
                                bRet = false;
                            }
                            else if ( customSheet.dictCustomItem[raw][rowIndex].status == 1 && diffFuncSheet.cells[rowIndex][colIndex].status == 1
                                        && customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] != diffFuncSheet.cells[rowIndex][colIndex].value )
                            {
                                bRet = false;
                            }
                        }
                        else
                        {
                            bRet = false;
                        }
                    }
                }
            }
            else if ( diffFuncSheet.cells.Count > 0 )
            {
                bRet = false;
            }
        
            inExcel.Close();
            outExcel.Close();
            return bRet;
        }

        private bool ClashCheck( CustomSheet customSheet,
                                string strOldFuncExcelPath, string strNewFuncExcelPath, string strFuncSheetName,
                                string strCustomizeClashPath, string strClashSheetName, out List<string> errorMsgs )
        {
            errorMsgs = new List<string>();
            List<string> lstErrorMsg;

            FunctionSheetControl funcControl = new FunctionSheetControl();
            FunctionSheet oldFuncSheet = new FunctionSheet();
            FunctionSheet newFuncSheet = new FunctionSheet();
            FunctionSheet diffFuncSheet = new FunctionSheet();

            funcControl.Read( strOldFuncExcelPath, strFuncSheetName, m_bFuncCustomizerExistLineFour, out oldFuncSheet, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );
            funcControl.Read( strNewFuncExcelPath, strFuncSheetName, m_bFuncCustomizerExistLineFour, out newFuncSheet, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );

            if ( errorMsgs.Count != 0 )
                return false;
            
            // 得到升级差异
            foreach ( int rowId in newFuncSheet.cells.Keys )
            {
                foreach ( int colId in newFuncSheet.cells[rowId].Keys )
                {
                    bool bNewItem = false;
                    bool bEdtItem = false;
                    if ( !oldFuncSheet.cells.ContainsKey( rowId ) )
                    {
                        bNewItem = true;
                    }
                    else
                    {
                        if ( !oldFuncSheet.cells[rowId].ContainsKey( colId ) )
                        {
                            bEdtItem = true;
                        }
                        else
                        {
                            if ( oldFuncSheet.cells[rowId][colId].value != newFuncSheet.cells[rowId][colId].value )
                            {
                                bEdtItem = true;
                            }
                        }
                    }

                    if ( bNewItem )
                    {
                        if ( !diffFuncSheet.cells.ContainsKey( rowId ) )
                        {
                            Dictionary<int, Node> dict = new Dictionary<int, Node>();
                            diffFuncSheet.cells.Add( rowId, dict );
                        }
                        diffFuncSheet.cells[rowId].Add( colId, new Node( 2, newFuncSheet.cells[rowId][colId].value ) );
                    }
                    if ( bEdtItem )
                    {
                        if ( !diffFuncSheet.cells.ContainsKey( rowId ) )
                        {
                            Dictionary<int, Node> dict = new Dictionary<int, Node>();
                            diffFuncSheet.cells.Add( rowId, dict );
                        }
                        diffFuncSheet.cells[rowId].Add( colId, new Node( 1, newFuncSheet.cells[rowId][colId].value ) );
                    }
                }
            }

            // 统计冲突
            CustomSheet clashSheet = new CustomSheet();

            string name = Path.GetFileName( strNewFuncExcelPath );
            string raw  = name;
            if ( Regex.IsMatch( raw, "[0-9]+_.*.xlsx" ) )
            {
                raw = raw.Substring( raw.IndexOf( "_" ) + 1 );
            }
            raw = raw.Substring( 0, raw.LastIndexOf( "." ) );

            Dictionary<int, CustomSheet.Item> clashItems = new Dictionary<int, CustomSheet.Item>();
            foreach ( int rowIndex in customSheet.dictCustomItem[raw].Keys )
            {
                CustomSheet.Item clashItem = new CustomSheet.Item( customSheet.dictCustomItem[raw][rowIndex].status );
                foreach ( int colIndex in customSheet.dictCustomItem[raw][rowIndex].dictOpCell.Keys )
                {
                    if ( diffFuncSheet.cells.ContainsKey(rowIndex) && diffFuncSheet.cells[rowIndex].ContainsKey(colIndex) )
                    {
                        if ( diffFuncSheet.cells[rowIndex][colIndex].status == 2 )
                        {
                            clashItem.dictOpCell.Add( colIndex, diffFuncSheet.cells[rowIndex][colIndex].value );
                        }
                        else if ( customSheet.dictCustomItem[raw][rowIndex].status == 1 && diffFuncSheet.cells[rowIndex][colIndex].status == 1
                                    && customSheet.dictCustomItem[raw][rowIndex].dictOpCell[colIndex] != diffFuncSheet.cells[rowIndex][colIndex].value )
                        {
                            clashItem.dictOpCell.Add( colIndex, diffFuncSheet.cells[rowIndex][colIndex].value );
                        }
                    }
                }

                if ( clashItem.dictOpCell.Count != 0 )
                {
                    clashItems.Add( rowIndex, clashItem );
                }
            }
            if ( clashItems.Count != 0 )
                clashSheet.dictCustomItem.Add( raw, clashItems );

            // 写入冲突
            YYExcel outExcel = new YYExcel();
            try
            {
                outExcel.Open( strCustomizeClashPath, strClashSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            }
            catch (System.Exception ex)
            {
                errorMsgs.Add( ex.Message );
                return false;
            }
            
            foreach ( string table in clashSheet.dictCustomItem.Keys )
            {
                int row = outExcel.GetRowsCount() + 1;
                
                foreach ( int rowIndex in clashSheet.dictCustomItem[table].Keys )
                {
                    int col = 3;
                    outExcel.setCellValue( row, 1, table );
                    outExcel.setCellValue( row, 2, clashSheet.dictCustomItem[table][rowIndex].status.ToString() );
                    outExcel.setCellValue( row, 3, rowIndex.ToString() );

                    foreach( int colIndex in clashSheet.dictCustomItem[table][rowIndex].dictOpCell.Keys )
                    {
                        col++;
                        outExcel.setCellValue( row, col, colIndex.ToString() );
                        col++;
                        outExcel.setCellValue( row, col, clashSheet.dictCustomItem[table][rowIndex].dictOpCell[colIndex] );
                    }
                }
               
            }

            try
            {
                outExcel.SaveAs( strCustomizeClashPath );
            }
            catch ( System.Exception ex )
            {
                errorMsgs.Add( ex.Message );
                outExcel.Close();
                return false;
            }
            outExcel.Close();

            return errorMsgs.Count == 0;
        }
    }
}
