using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ExcelTools.FunctionCSVConverter.ConvererMgr;
using System.Diagnostics;

namespace ExcelTools
{
    partial class Main : Form
    {
        private void FuncCSVConverterMain()
        {
            //MessageBox.Show( lvwFuncCSVConverterResult.SelectedItems.Count.ToString() );
            List<string> lstExcelPath = new List<string>();
            lstExcelPath.Clear();

            bool bFountText = false;
            lvwFuncCSVConverterResult.BeginUpdate();
            foreach ( ListViewItem lvi in lvwFuncCSVConverterResult.Items )
            {
                lvi.SubItems[2].Text = "尚未检查";

                if ( lvi.SubItems[1].Text == m_strFuncCSVConverterTextName )
                {
                    bFountText = true;
                    // Text交换至最后并选中
                    lvi.SubItems[1].Text = lvwFuncCSVConverterResult.Items[lvwFuncCSVConverterResult.Items.Count - 1].SubItems[1].Text;
                    lvwFuncCSVConverterResult.Items[lvwFuncCSVConverterResult.Items.Count - 1].SubItems[1].Text = m_strFuncCSVConverterTextName;
                    bool bTemp = lvi.Selected;
                    lvi.Selected = lvwFuncCSVConverterResult.Items[lvwFuncCSVConverterResult.Items.Count - 1].Selected;
                    lvwFuncCSVConverterResult.Items[lvwFuncCSVConverterResult.Items.Count - 1].Selected = bTemp;
                }

                if ( lvi.Selected )
                    lstExcelPath.Add( m_strFuncCSVConverterExcelDirPath + "\\" + lvi.SubItems[1].Text );
                else
                {
                    lstExcelPath.Add( "" );
                }
            }
            lvwFuncCSVConverterResult.EndUpdate();
            if ( m_bFuncCSVConverterProduceText )
            {
                while ( Util.IsFileInUse( m_strFuncCSVConverterTextPath ) )
                {
                    MessageBox.Show( "请先关闭" + m_strFuncCSVConverterTextPath );
                }
                if ( !File.Exists( m_strFuncCSVConverterTextPath ) )
                {
                    TextSheetControl sheet = new TextSheetControl();
                    sheet.Create( m_strFuncCSVConverterTextPath, m_strFuncCSVConverterSheetName, m_bFuncCSVConverterExistLineFour,
                        m_strFuncCSVConverterTextColId, m_strFuncCSVConverterTextColName );
                }
                if ( !bFountText )
                {
                    lvwFuncCSVConverterResult.BeginUpdate();
                    ListViewItem lvi = new ListViewItem( ( lvwFuncCSVConverterResult.Items.Count + 1 ).ToString() );
                    lvi.SubItems.Add( m_strFuncCSVConverterTextName );
                    lvi.SubItems.Add( "尚未检查" );
                    lvi.Selected = true;
                    lvwFuncCSVConverterResult.Items.Add( lvi );
                    lvwFuncCSVConverterResult.EndUpdate();
                    lstExcelPath.Add( m_strFuncCSVConverterTextPath );
                }
            }

            string[] fileNames = lstExcelPath.ToArray();
            int fLength = fileNames.Length;
            for ( int i = 0; i < fLength; i++ )
            {
                string name = Path.GetFileName( fileNames[i] );

                // 把Text表放到最后处理
                if ( name == m_strFuncCSVConverterTextName && i != fLength - 1 )
                {
                    Util.Swap( ref fileNames[i], ref fileNames[fLength - 1] );
                    name = Path.GetFileName( fileNames[i] );
                }
            }

            // 按钮锁定
            btnFuncConvertTextAndCSV.Enabled = false;
            btnFuncConvertCSVOnly.Enabled = false;

            // 处理开始
            Thread processThread = new Thread( new ParameterizedThreadStart( FuncCSVConvertProcess ) );
            processThread.Start( fileNames );
        }

        private void FuncCSVConvertProcess( object o )
        {
            string[]  excels = o as string[];

            List<ExcelItem> newExcels;
            Dictionary<string, FunctionSheet> dicFuncSheet;
            int maxTblId;

            int chkRes = FuncCSVConvertCheck( excels, out newExcels, out maxTblId, out dicFuncSheet );
            if ( chkRes != 0 )
            {
                this.Invoke( (UpdateButtonStateDelegate)delegate()
                {
                    //MessageBox.Show( "请修改不规范处后重试" );
                    btnFuncConvertTextAndCSV.Enabled = true;
                    btnFuncConvertCSVOnly.Enabled = true;
                } );
                return;
            }

            foreach ( ExcelItem excel in newExcels )
            {
                while ( Util.IsFileInUse( excel.path ) )
                {
                    MessageBox.Show( "请先关闭" + excel.path );
                }
            }

            if ( RenameNewFuncExcel( excels, newExcels, maxTblId ) )
            {
                if ( m_bFuncCSVConverterProduceText )
                {
                    while ( Util.IsFileInUse( m_strFuncCSVConverterTextPath ) )
                    {
                        MessageBox.Show( "请先关闭" + m_strFuncCSVConverterTextPath );
                    }
                    TextSheetControl sheet = new TextSheetControl();
                    sheet.Create( m_strFuncCSVConverterTextPath, m_strFuncCSVConverterSheetName, m_bFuncCSVConverterExistLineFour,
                        m_strFuncCSVConverterTextColId, m_strFuncCSVConverterTextColName );
                }

                FuncCSVConvert( excels, ref dicFuncSheet );
                DoTSConvert( excels );
                DoProtobufConvert( excels );
            }
            else
            {
                this.Invoke( (CreateMessageBoxDelegate)delegate()
                {
                    MessageBox.Show( "含中文的功能表个数超过限制" );
                } );
            }

            this.Invoke( (UpdateButtonStateDelegate)delegate()
            {
                btnFuncConvertTextAndCSV.Enabled = true;
                btnFuncConvertCSVOnly.Enabled = true;
            } );
        }

        private int FuncCSVConvertCheck( string[] excels, out List<ExcelItem> newExcels, out int maxTblId, out Dictionary<string, FunctionSheet> dictExcels )
        {
            newExcels = new List<ExcelItem>();
            dictExcels = new Dictionary<string, FunctionSheet>();
            List<string> errorMsgs = new List<string>();
            maxTblId = 0;
            int ret = 0;
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                if ( path == "" ) continue;
                string name = Path.GetFileName( path );
                bool bWithTxtCol = false;
                bool bIsAscending = false;
                List<string> lstErrorMsg;

                FunctionSheetControl funcControl = new FunctionSheetControl();
                FunctionSheet funcSheet = new FunctionSheet();
                bool chkPass = funcControl.Check( path, m_strFuncCSVConverterSheetName, m_bFuncCSVConverterExistLineFour, out bWithTxtCol, out bIsAscending, out funcSheet, out lstErrorMsg );
                this.Invoke( (UpdateFunctionCSVConvertResultDelegate)delegate( int idx, bool res )
                {
                    switch ( res )
                    {
                        case true:
                            lvwFuncCSVConverterResult.Items[idx].SubItems[2].Text = "检查通过";
                            break;

                        default:
                            lvwFuncCSVConverterResult.Items[idx].SubItems[2].Text = "检查未通过";
                            errorMsgs.AddRange( lstErrorMsg );
                            ret = 1;
                            break;
                    }

                    lvwFuncCSVConverterResult.EnsureVisible( idx );
                }, i, chkPass && bIsAscending );

                if ( !bIsAscending )
                {
                    errorMsgs.Add( ErrorMsg.Error( name, "行Id未严格递增" ) );
                }

                // 标记新的有文字列excel
                if ( bWithTxtCol )
                {
                    string pure_name = name;
                    if ( !Regex.IsMatch( name, "^[0-9]{" + FunctionSheetControl.m_wTblIdLen + "}_" ) )
                    {
                        ExcelItem item = new ExcelItem( i, path );
                        newExcels.Add( item );
                    }
                    else
                    {
                        string id = name.Substring( 0, name.IndexOf( "_" ) );
                        maxTblId = Math.Max( maxTblId, Int32.Parse( id ) );

                        pure_name = name.Substring( name.IndexOf( "_" ) + 1 );
                    }
                    if ( dictExcels.ContainsKey( pure_name ) )
                    {
                        errorMsgs.Add( ErrorMsg.Error( name, "表名" + pure_name + "重复" ) );
                        ret = 1;
                    }
                    else
                    {
                        dictExcels.Add( Path.GetFileNameWithoutExtension( pure_name ), funcSheet );
                    }
                }
            }

            if ( ret == 1 )
            {
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
            }
            return ret;
        }

        private bool RenameNewFuncExcel( string[] excels, List<ExcelItem> newExcels, int maxTblId )
        {
            // 修改有文字列的excel
            if ( maxTblId + newExcels.Count >= Util.TenPow( FunctionSheetControl.m_wTblIdLen ) )
            {
                return false;
            }
            foreach ( ExcelItem item in newExcels )
            {
                string oldName = item.path.Substring( item.path.LastIndexOf( "\\" ) + 1 );
                //oldName = oldName.Substring( oldName.IndexOf( "_" ) + 1 );
                string newId = ( maxTblId + 1 ).ToString();
                while ( newId.Length < FunctionSheetControl.m_wTblIdLen )
                    newId = "0" + newId;
                string newName = newId + "_" + oldName;
                string newPath = m_strFuncCSVConverterExcelDirPath + "\\" + newName;
                File.Move( item.path, newPath );
                excels[item.id] = newPath;
                maxTblId++;
            }
            return true;
        }

        private void DoTSConvert( string[] excels )
        {
            if ( !m_bFuncCSVConverterUseTypescriptChecked )
                return;

            List<string> errorMsgs = new List<string>();
            TextSheet textSheet = new TextSheet();
            FunctionSheetControl funcControl = new FunctionSheetControl();

            FunctionCSVConverter.ConvererMgr.TSHeaderMgr tsHeadMgr = new FunctionCSVConverter.ConvererMgr.TSHeaderMgr();
            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                if ( path == "" ) continue;

                string name = path.Substring( path.LastIndexOf( '\\' ) + 1 );
                if ( name == "text.xlsx" ) continue;
                FunctionSheet funSheet = new FunctionSheet();
                List<string> lstErrorMsg;

                if ( Regex.IsMatch( name, "^[0-9]+_.*" ) )
                    name = name.Substring( name.IndexOf( "_" ) + 1 );
                name = name.Substring( 0, name.LastIndexOf( '.' ) );

                funcControl.Read( path,
                    m_strFuncCSVConverterSheetName,
                    m_bFuncCSVConverterExistLineFour,
                    out funSheet,
                    out lstErrorMsg );

                // Typescript header file
                tsHeadMgr.AddNewSheet( name + "db", name + "DB", funSheet );

                // Json file
                JsonMgr jsonMgr = new JsonMgr();
                jsonMgr.loadSheet( funSheet );
                jsonMgr.BuildContent();
                this.CreateNewFile(
                    name + "db.json",
                    m_strFuncCSVConverterDirPathTypescript,
                    jsonMgr.GetResult() );

                // 
            }

            // Typescript header file
            tsHeadMgr.BuildContent();
            this.CreateNewFile(
                tsHeadMgr.GetExportFileName(),
                m_strFuncCSVConverterDirPathTypescript,
                tsHeadMgr.GetResult()
                );
        }

        private void DoProtobufConvert( string[] excels )
        {
            if ( !m_bFuncCSVConverterUseProtobufferChecked )
                return;

            List<string> errorMsgs = new List<string>();
            TextSheet textSheet = new TextSheet();
            FunctionSheetControl funcControl = new FunctionSheetControl();

            string strCurDirectory = System.Environment.CurrentDirectory;
            if ( File.Exists( strCurDirectory + "\\Google.ProtocolBuffers.dll" ) )
                File.Copy( strCurDirectory + "\\Google.ProtocolBuffers.dll", m_strFuncCSVConverterDirPathProtobuf + "\\Google.ProtocolBuffers.dll" );
            if ( File.Exists( strCurDirectory + "\\protoc.exe" ) )
                File.Copy( strCurDirectory + "\\protoc.exe", m_strFuncCSVConverterDirPathProtobuf + "\\protoc.exe" );
            if ( File.Exists( strCurDirectory + "\\ProtoGen.exe" ) )
                File.Copy( strCurDirectory + "\\ProtoGen.exe", m_strFuncCSVConverterDirPathProtobuf + "\\ProtoGen.exe" );

            StringBuilder cmdBuilder = new StringBuilder();
            cmdBuilder.Append( "@echo off\r\n" );
            cmdBuilder.Append( "setlocal EnableDelayedExpansion\r\n" );
            cmdBuilder.Append( "\r\n" );
            string output = null;

            for ( int i = 0; i < excels.Length; i++ )
            {
                string path = excels[i];
                if ( path == "" ) continue;

                string name = path.Substring( path.LastIndexOf( '\\' ) + 1 );
                FunctionSheet funcSheet = new FunctionSheet();
                List<string> lstErrorMsg;

                if ( Regex.IsMatch( name, "^[0-9]+_.*" ) )
                    name = name.Substring( name.IndexOf( "_" ) + 1 );
                name = name.Substring( 0, name.LastIndexOf( '.' ) );

                funcControl.Read( path,
                    m_strFuncCSVConverterSheetName,
                    m_bFuncCSVConverterExistLineFour,
                    out funcSheet,
                    out lstErrorMsg );

                StringBuilder protoBuilder = new StringBuilder();
                protoBuilder.Append( "using System.IO;\r\n" );
                protoBuilder.Append( "using Google.ProtocolBuffers;\r\n" );
                protoBuilder.Append( "using System;\r\n\r\n" );
                protoBuilder.Append( "class Program\r\n" );
                protoBuilder.Append( "{\r\n" );
                protoBuilder.Append( "\tstatic void addData( out " + name + " data, string[] arr )\r\n" );
                protoBuilder.Append( "\t{\r\n" );
                protoBuilder.Append( "\t\t" + name + ".Builder itemBuilder = " + name + ".CreateBuilder();\r\n" );

                int cnt = 0;
                foreach ( int colId in funcSheet.headers.Keys )
                {
                    string propertyName = funcSheet.headers[colId].titleName;
                    if ( Regex.IsMatch( propertyName, "FIX_.*" ) || Regex.IsMatch( propertyName, "INT_" ) )
                    {
                        propertyName = propertyName.Substring( propertyName.IndexOf( "_" ) + 1 );
                        protoBuilder.Append( "\t\tif ( arr[" + cnt + "] != \"\" ) itemBuilder.Set" + Util.CamelName( propertyName ) + "( Convert.ToInt32( arr[" + cnt + "] ) );\r\n" );
                    }
                    else if ( Regex.IsMatch( propertyName, "FLT_" ) )
                    {
                        propertyName = propertyName.Substring( propertyName.IndexOf( "_" ) + 1 );
                        protoBuilder.Append( "\t\tif ( arr[" + cnt + "] != \"\" ) itemBuilder.Set" + Util.CamelName( propertyName ) + "( Convert.ToSingle( arr[" + cnt + "] ) );\r\n" );
                    }
                    else
                    {
                        protoBuilder.Append( "\t\titemBuilder.Set" + Util.CamelName( propertyName ) + "( arr[" + cnt + "] );\r\n" );
                    }
                    
                    cnt++;
                }

                protoBuilder.Append( "\t\tdata = itemBuilder.Build();\r\n" );
                protoBuilder.Append( "\t}\r\n\r\n" );

                protoBuilder.Append( "\tstatic void Main( string[] args )\r\n" );
                protoBuilder.Append( "\t{\r\n" );
                protoBuilder.Append( "\t\t" + name + "Table.Builder tableBuilder = " + name + "Table.CreateBuilder();\r\n" );
                protoBuilder.Append( "\t\ttableBuilder.SetTname( \"" + name + "\" );\r\n" );

                string txtPath = m_strFuncCSVConverterDirPathProtobuf + "\\" + name + ".txt";
                protoBuilder.Append( "\t\tusing ( StreamReader reader = new StreamReader( File.OpenRead( \"" + txtPath.Replace( "\\", "\\\\" ) + "\" ) ) )\r\n" );
                protoBuilder.Append( "\t\t{\r\n" );
                protoBuilder.Append( "\t\t\tstring strLine = null;\r\n" );
                protoBuilder.Append( "\t\t\twhile ( ( strLine = reader.ReadLine() ) != null )\r\n" );
                protoBuilder.Append( "\t\t\t{\r\n" );
                protoBuilder.Append( "\t\t\t\tstring[] arrCell;\r\n" );
                protoBuilder.Append( "\t\t\t\tarrCell = strLine.Split( '\\t' );\r\n" );
                protoBuilder.Append( "\t\t\t\t" + name + " _data;\r\n" );
                protoBuilder.Append( "\t\t\t\taddData( out _data, arrCell );\r\n" );
                protoBuilder.Append( "\t\t\t\ttableBuilder.AddTlist( _data );\r\n" );
                protoBuilder.Append( "\t\t\t}\r\n" );
                protoBuilder.Append( "\t\t}\r\n\r\n" );
                protoBuilder.Append( "\t\t" + name + "Table tbl = tableBuilder.Build();\r\n\r\n" );

                string dbpPath = m_strFuncCSVConverterDirPathProtobuf + "\\" + name + ".dbp";
                protoBuilder.Append( "\t\tusing ( FileStream stream = new FileStream( \"" + dbpPath.Replace( "\\", "\\\\" ) + "\", FileMode.Create ) )\r\n" );
                protoBuilder.Append( "\t\t{\r\n" );
                protoBuilder.Append( "\t\t\ttbl.WriteTo( stream );\r\n" );
                protoBuilder.Append( "\t\t\tstream.Flush();\r\n" );
                protoBuilder.Append( "\t\t\tstream.Close();\r\n" );
                protoBuilder.Append( "\t\t}\r\n" );
                protoBuilder.Append( "\t}\r\n" );
                protoBuilder.Append( "}\r\n" );

                string protoCS = protoBuilder.ToString();
                byte[] contentCS = new UTF8Encoding( true ).GetBytes( protoCS );
                FileStream fsCS = File.Create( m_strFuncCSVConverterDirPathProtobuf + "\\" + name + "Pro.cs" );
                fsCS.Write( contentCS, 0, contentCS.Length );
                fsCS.Flush();
                fsCS.Close();

                string protoPath = name + ".proto";
                string protobinPath = name + ".protobin";
               
                // protoc
                cmdBuilder.Append( ".\\protoc.exe --descriptor_set_out=" + name + ".protobin --include_imports " + name + ".proto\r\n" );

                // protogen
                cmdBuilder.Append( ".\\ProtoGen.exe " + name + ".protobin\r\n" );

                // csc
                cmdBuilder.Append( "csc.exe /r:Google.ProtocolBuffers.dll /out:" + name + ".exe " + Util.CamelName(name) + ".cs " + name + "Pro.cs\r\n" );

                // exe
                cmdBuilder.Append( ".\\" + name + ".exe\r\n" );

                cmdBuilder.Append( "\r\n" );
            }

            cmdBuilder.Append( "del *.dll\r\n" );
            cmdBuilder.Append( "del *.exe\r\n" );
            cmdBuilder.Append( "del *.protobin\r\n" );
            cmdBuilder.Append( "del *.cs\r\n" );
            cmdBuilder.Append( "del *.txt\r\n" );
            cmdBuilder.Append( "del %0%\r\n" );

            string protoBAT = cmdBuilder.ToString();
            byte[] contentBAT = new UTF8Encoding( true ).GetBytes( protoBAT );
            FileStream fsBAT = File.Create( m_strFuncCSVConverterDirPathProtobuf + "\\Proto.bat" );
            fsBAT.Write( contentBAT, 0, contentBAT.Length );
            fsBAT.Flush();
            fsBAT.Close();

            Util.ConsoleRun( m_strFuncCSVConverterDirPathProtobuf, ".\\Proto.bat", out output );              
        }

        private void CreateNewFile( string filename, string path, string content )
        {
            byte[] bContent = new UTF8Encoding( true ).GetBytes( content );
            FileStream fs = File.Create( path + "\\" + filename );
            fs.Write( bContent, 0, bContent.Length );
            fs.Flush();
            fs.Close();
        }

        private void FuncCSVConvert( string[] excels, ref Dictionary<string, FunctionSheet> dicFuncSheet )
        {
            List<string> errorMsgs = new List<string>();
            TextSheet textSheet = new TextSheet();
            FunctionCSVConverter.ConvererMgr.TSHeaderMgr tsHeadMgr = new FunctionCSVConverter.ConvererMgr.TSHeaderMgr();
            int ret = 0;
            for ( int i = 0; i < excels.Length; i++ )
            {
                try
                {
                    ret |= FuncCSVConvert( excels[i], i, ref errorMsgs, ref textSheet, ref dicFuncSheet );
                }
                catch (System.Exception ex)
                {
                    errorMsgs.Add( ErrorMsg.Error( excels[i], ex.ToString() ) );
                }
            }

            if ( ret == 1 )
            {
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
            }
        }

        private int FuncCSVConvert( string excelPath, int i, ref List<string> errorMsgs, ref TextSheet textSheet, ref Dictionary<string, FunctionSheet> dicFuncSheet )
        {
            int ret = 0;

            string path = excelPath;
            if ( path == "" ) return 0;
            string name = path.Substring( path.LastIndexOf( '\\' ) + 1 );
            string strCsvContentServer = "";
            string strCsvContentClient = "";
            string strJSContentCommonJS = "";
            string strJSContentTypescript = "";
            string strJSContentClient  = "";
            string strProtobuf = "";
            string strProtobufText = "";
            List<string> lstErrorMsg;

            bool updatePass = true;

            if ( name != m_strFuncCSVConverterTextName && Regex.IsMatch( name, "^[0-9]+_.*.xlsx" ) )
            {
                name = name.Substring( name.IndexOf( "_" ) + 1 );
            }

            if ( m_bFuncCSVConverterProduceText && name == m_strFuncCSVConverterTextName )
            {
                TextSheetControl textControl = new TextSheetControl();
                updatePass = textControl.Update( m_strFuncCSVConverterTextPath, m_strFuncCSVConverterSheetName, m_bFuncCSVConverterExistLineFour, textSheet, out lstErrorMsg );
                errorMsgs.AddRange( lstErrorMsg );
            }

            bool convertPass = FuncCSVConvertBuilder( path, m_strFuncCSVConverterSheetName,
                                                      out strCsvContentServer, out strCsvContentClient,
                                                      out strJSContentCommonJS, out strJSContentClient, out strJSContentTypescript,
                                                      out strProtobuf, out strProtobufText,
                                                      ref textSheet, ref dicFuncSheet,
                                                      out lstErrorMsg
                                                    );
            if ( !convertPass )
            {
                ret = 1;
                errorMsgs.AddRange( lstErrorMsg );
            }

            name = name.Substring( 0, name.LastIndexOf( '.' ) );
            if ( updatePass && convertPass && m_bFuncCSVConverterUseCSVChecked )
            {
                byte[] content = new UTF8Encoding( true ).GetBytes( strCsvContentServer );

                if ( content.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出csv文件空" ) );
                }

				string nameZh = name;

				if ( nameZh == m_strFuncCSVConverterTextName.Substring( 0, m_strFuncCSVConverterTextName.LastIndexOf( "." ) ) )
					nameZh += "_zh";

				FileStream fs = File.Create( m_strFuncCSVConverterDirPathServer + "\\" + name + ".csv" );
                fs.Write( content, 0, content.Length );
                fs.Flush();
                fs.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseCSVStringChecked )
            {
                byte[] content = new UTF8Encoding( true ).GetBytes( strCsvContentClient );

                if ( content.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出csv(string)文件空" ) );
                }

				string nameZh = name;

				if ( nameZh == m_strFuncCSVConverterTextName.Substring( 0, m_strFuncCSVConverterTextName.LastIndexOf( "." ) ) )
					nameZh += "_zh";

				FileStream fs = File.Create( m_strFuncCSVConverterDirPathCSVString + "\\" + name + ".csv" );
                fs.Write( content, 0, content.Length );
                fs.Flush();
                fs.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseJSChecked )
            {
                byte[] contentJS = new UTF8Encoding( true ).GetBytes( strJSContentClient );

                if ( contentJS.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出js文件空" ) );
                }

                FileStream fsJS = File.Create( m_strFuncCSVConverterDirPathClient + "\\" + name + "db.js" );
                fsJS.Write( contentJS, 0, contentJS.Length );
                fsJS.Flush();
                fsJS.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseCommonJSChecked )
            {
                byte[] contentJS = new UTF8Encoding( true ).GetBytes( strJSContentCommonJS );

                if ( contentJS.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出commonjs文件空" ) );
                }

                FileStream fsJS = File.Create( m_strFuncCSVConverterDirPathCommonJS + "\\" + name + "db.js" );
                fsJS.Write( contentJS, 0, contentJS.Length );
                fsJS.Flush();
                fsJS.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseTypescriptChecked )
            {
                byte[] contentJS = new UTF8Encoding( true ).GetBytes( strJSContentTypescript );

                if ( contentJS.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出ts文件空" ) );
                }

                FileStream fsJS = File.Create( m_strFuncCSVConverterDirPathTypescript + "\\" + name + "db.ts" );
                fsJS.Write( contentJS, 0, contentJS.Length );
                fsJS.Flush();
                fsJS.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseProtobufferChecked )
            {
                byte[] contentProto = new UTF8Encoding( true ).GetBytes( strProtobuf );

                if ( contentProto.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出proto文件空" ) );
                }

                FileStream fsProto = File.Create( m_strFuncCSVConverterDirPathProtobuf + "\\" + name + ".proto" );
                fsProto.Write( contentProto, 0, contentProto.Length );
                fsProto.Flush();
                fsProto.Close();

                byte[] contentProtoText = new UTF8Encoding( true ).GetBytes( strProtobufText );

                if ( contentProtoText.Length == 0 )
                {
                    ret = 1;
                    errorMsgs.Add( ErrorMsg.Error( name, "导出txt文件空" ) );
                }

                FileStream fsPotoText = File.Create( m_strFuncCSVConverterDirPathProtobuf + "\\" + name + ".txt" );
                fsPotoText.Write( contentProtoText, 0, contentProtoText.Length );
                fsPotoText.Flush();
                fsPotoText.Close();
            }

            if ( updatePass && convertPass && m_bFuncCSVConverterUseTextChecked && m_bFuncCSVConverterProduceText &&
                   name == m_strFuncCSVConverterTextName.Substring( 0, m_strFuncCSVConverterTextName.LastIndexOf( '.' ) ) )
            {
                byte[] contentJS = new UTF8Encoding( true ).GetBytes( strJSContentClient );

                if ( contentJS.Length == 0 )
                {
                    ret = 1;
                    lstErrorMsg.Add( ErrorMsg.Error( name, "导出textdb.js文件空" ) );
                }

                FileStream fsJS = File.Create( m_strFuncCSVConverterDirPathText + "\\" + name + "db.js" );
                fsJS.Write( contentJS, 0, contentJS.Length );
                fsJS.Flush();
                fsJS.Close();

                File.Delete( m_strFuncCSVConverterDirPathClient + "\\" + name + "db.js" );
            }

            this.Invoke( (UpdateFunctionCSVConvertResultDelegate)delegate( int idx, bool res )
            {
                switch ( res )
                {
                    case true:
                        lvwFuncCSVConverterResult.Items[idx].SubItems[2].Text = "完成";
                        break;

                    default:
                        lvwFuncCSVConverterResult.Items[idx].SubItems[2].Text = "未完成";
                        ret = 1;
                        break;

                }

                lvwFuncCSVConverterResult.EnsureVisible( idx );
            }, i, updatePass && convertPass );

            return ret;
        }

        string ObjectDefinePropertyString( string name, string property, string comment, string value, string enumerable )
        {
            string str = "Object.defineProperty(" + name + ", \"" + property + "\",{\t//" + comment + "\r\n"
                        + "\tvalue\t:" + value + ",\r\n"
                        + "\tenumerable\t: " + enumerable + "\r\n"
                        + "});";
            return str;
        }

        bool FuncCSVConvertBuilder( String strExcelFilePath, String strSheetName,
                                    out String strResultServer, out String strResultClient,
                                    out String strJSResultCommonJS, out String strJSResultClient, out String strJSContentTypescript,
                                    out String strProtobuf, out String strProtobufText,
                                    ref TextSheet textSheet,
                                    ref Dictionary<string, FunctionSheet> dicFuncSheet,
                                    out List<string> lstErrorMsg
                                  )
        {
            strResultServer = "";
            strResultClient = "";
            strJSResultCommonJS = "";
            strJSResultClient = "";
            strJSContentTypescript = "";
            strProtobuf = "";
            strProtobufText = "";
            lstErrorMsg = new List<string>();

            bool result = true;
            int cnt = 0;

            StringBuilder builderServer   = new StringBuilder();
            StringBuilder builderClient   = new StringBuilder();
            StringBuilder jsBuilderClient = new StringBuilder();
            StringBuilder jsBuilderCommonJS = new StringBuilder();
            StringBuilder jsBuilderTypescript = new StringBuilder();
            StringBuilder protobufBuilder = new StringBuilder();
            StringBuilder protobufTextBuilder = new StringBuilder();

            // common variable
            string name = strExcelFilePath.Substring( strExcelFilePath.LastIndexOf( '\\' ) + 1 );
            if ( Regex.IsMatch( name, "^[0-9]+_.*" ) )
                name = name.Substring( name.IndexOf( "_" ) + 1 );
            name = name.Substring( 0, name.LastIndexOf( '.' ) );
            string pure_name = name;
            name = name + "DB";
            string nameItemClient = name + "Config";
            string nameItemCommonJS = name + "_$item";
            string nameItemTypescript = name + "Config";

            string declare = "var " + name + " = {};";
            string declareItemClient = "var " + nameItemClient + " = {};";
            string declareItemCommonJS = "var " + nameItemCommonJS + " = {};";

            FunctionSheet funSheet = null;

            if ( dicFuncSheet.ContainsKey( pure_name) )
            {
                funSheet = dicFuncSheet[pure_name];
            }
            else
            {
                funSheet = new FunctionSheet();
                FunctionSheetControl funcControl = new FunctionSheetControl();
                funcControl.Read( strExcelFilePath, m_strFuncCSVConverterSheetName, m_bFuncCSVConverterExistLineFour, out funSheet, out lstErrorMsg );
            }

            // javascript
            jsBuilderClient.Append( declare + "\r\n\r\n" );
            jsBuilderClient.Append( declareItemClient + "\t// just for " + nameItemClient + ".get(id)\r\n\r\n" );
            jsBuilderClient.Append( "/**\r\n" +
                                    " * @param id\r\n" +
                                    " * @returns {" + name + "}\r\n" +
                                    " */\r\n" +
                                    nameItemClient + ".get = function (id) {\r\n" +
                                    "    return " + name + "[id];\r\n" +
                                    "};\r\n\r\n" );

            // nodejs
            jsBuilderCommonJS.Append( declare + "\r\n" );
            jsBuilderCommonJS.Append( declareItemCommonJS + "\t// just for Object.defineProperty\r\n\r\n" );

            jsBuilderCommonJS.Append( "/**\r\n" +
                                    " * @param id\r\n" +
                                    " * @returns {" + declareItemCommonJS + "}\r\n" +
                                    " */\r\n" +
                                    "module.exports.get = function (id) {\r\n" +
                                    "    return " + name + "[id];\r\n" +
                                    "};\r\n\r\n" +
                                    "/**\r\n" +
                                    "* @type {" + name + "}\r\n" +
                                    "*/\r\n" +
                                    "module.exports.config = " + name + ";\r\n\r\n" );

            // typescript
            string tsDeclare = "var " + nameItemTypescript + " : {[ID:number]: " + name + "} = {};";
            jsBuilderTypescript.Append( tsDeclare + "\r\n\r\n" );
            jsBuilderTypescript.Append( "export class " + name + " {\r\n" );
            int maxKeyLength = 0;
            foreach ( int colId in funSheet.headers.Keys )
            {
                string propertyName = funSheet.headers[colId].titleName;
                if ( maxKeyLength < propertyName.Length )
                {
                    maxKeyLength = propertyName.Length;
                }
            }
            int tabCount = ( maxKeyLength + 6 ) / 4 + 1;
            foreach ( int colId in funSheet.headers.Keys )
            {
                string propertyName = funSheet.headers[colId].titleName;
                string chineseName = funSheet.headers[colId].titleChineseName;
                jsBuilderTypescript.Append( "\t" + propertyName + ": any;" );
                // for comments align
                int tempTabCount = ( propertyName.Length + 6 ) / 4;
                for ( int i = 0; i < tabCount - tempTabCount; i++ )
                {
                    jsBuilderTypescript.Append( "\t" );
                }
                jsBuilderTypescript.Append( "//" + chineseName + "\r\n" );
            }
            jsBuilderTypescript.Append( "}\r\n\r\n" );
            // export function get
            jsBuilderTypescript.Append( "export function get(ID:number):" + name + " {\r\n" );
            jsBuilderTypescript.Append( "\treturn " + nameItemTypescript + "[ID];\r\n" );
            jsBuilderTypescript.Append( "}\r\n\r\n" );
            // export function getAll
            jsBuilderTypescript.Append( "export function getAll():{[ID:number]: " + name + "} {\r\n" );
            jsBuilderTypescript.Append( "\treturn " + nameItemTypescript + ";\r\n" );
            jsBuilderTypescript.Append( "}\r\n\r\n" );

            // protobuffer
            protobufBuilder.Append( "message " + pure_name + "Table\r\n{\r\n" );
            protobufBuilder.Append( "\toptional string tname = 1;\r\n" );
            protobufBuilder.Append( "\trepeated " + pure_name + " tlist = 2;\r\n" );
            protobufBuilder.Append( "}\r\n" );
            protobufBuilder.Append( "\r\n" );
            protobufBuilder.Append( "message " + pure_name + "\r\n{\r\n" );

            //todo @lyx  add js api...

            string rowDefineProperty = ObjectDefinePropertyString( name, "rowLength", "数据行数", funSheet.GetDataRowCount().ToString(), "false" );
            string colDefineProperty = ObjectDefinePropertyString( name, "colLength", "数据列数", funSheet.GetDataColCount().ToString(), "false" );
            jsBuilderClient.Append( rowDefineProperty + "\r\n" );
            jsBuilderClient.Append( colDefineProperty + "\r\n" );
            jsBuilderClient.Append( "\r\n" );
            jsBuilderCommonJS.Append( rowDefineProperty + "\r\n" );
            jsBuilderCommonJS.Append( colDefineProperty + "\r\n" );
            jsBuilderCommonJS.Append( "\r\n" );

            cnt = 0;
            foreach ( int colId in funSheet.headers.Keys )
            {
                cnt++;
                if ( m_bFuncCSVConverterUseCSVChecked )
                {
                    AppendCellForCSV( builderServer, colId.ToString() );
                }
                if ( m_bFuncCSVConverterUseCSVStringChecked )
                {
                    AppendCellForCSV( builderClient, funSheet.headers[colId].titleName );
                }

                if ( cnt != funSheet.headers.Count )
                {
                    builderServer.Append( "," );
                    builderClient.Append( "," );
                }

                string propertyName = funSheet.headers[colId].titleName;
                string chineseName  = funSheet.headers[colId].titleChineseName;
                //string objectDefineProperty = "Object.defineProperty(" + name + ", \"" + propertyName + "\",{\t//" + chineseName + "\r\n"
                //                        + "\tvalue\t:0,\r\n"
                //                        + "\tenumerable\t: false\r\n"
                //                        + "});\r\n";
                if ( !funSheet.headers[colId].bIsServerOnly )
                {
                    jsBuilderClient.Append( ObjectDefinePropertyString( name, propertyName, chineseName, "0", "false" ) + "\r\n" );
                }

                jsBuilderCommonJS.Append( ObjectDefinePropertyString( nameItemCommonJS, propertyName, chineseName, "0", "false" ) + "\r\n" );

                AppendHeadForProto( protobufBuilder, propertyName, colId.ToString() );
            }

            if ( funSheet.headers.Count != 0 )
            {
                builderServer.Append( "\r\n" );
                builderClient.Append( "\r\n" );
            }

            protobufBuilder.Append( "}\r\n" );

            foreach ( int rowId in funSheet.itemPos.Keys )
            {
                cnt = 0;

                foreach ( int colId in funSheet.headers.Keys )
                {
                    cnt++;
                    string cell = funSheet.cells[rowId][colId].value;

                    // 文字列cell替换为Text表Id，并做记录
                    if ( funSheet.headers[colId].bIsTxtCol && cell != string.Empty )
                    {
                        string tempName = Path.GetFileName( strExcelFilePath );
                        int    wTblId   = Int32.Parse( tempName.Substring( 0, tempName.IndexOf( "_" ) ) );
                        int    wTxtId   = wTblId * Util.TenPow( FunctionSheetControl.m_wColIdLen + FunctionSheetControl.m_wRowIdLen ) +
                                          colId * Util.TenPow( FunctionSheetControl.m_wRowIdLen ) + rowId;

                        textSheet.dictText.Add( wTxtId, cell );
                        cell = wTxtId.ToString();
                    }

                    AppendCellForCSV( builderServer, cell );
                    AppendCellForCSV( builderClient, funSheet.headers[colId].bIsServerOnly ? "" : cell );

                    try
                    {
                        AppendCellForProto( protobufTextBuilder, funSheet.headers[colId].titleName, cell );
                    }
                    catch ( Exception ex ) //some other exception
                    {
                        lstErrorMsg.Add( ErrorMsg.FormatError( name, rowId.ToString() + "行," + colId.ToString() + "列", ex.ToString() ) );
                        result = false;
                    }

                    try
                    {
                        if ( !funSheet.headers[colId].bIsServerOnly )
                        {
                            AppendCellForJS( name, jsBuilderClient, cell, funSheet.headers[colId].titleName, cnt == 1, cnt == funSheet.headers.Count );
                        }
                        else if ( cnt == 1 )
                        {
                            jsBuilderClient.Append( "\r\n" );
                            jsBuilderClient.Append( name + "[\"" + cell + "\"] = \r\n{\r\n" );
                        }
                        else if ( cnt == funSheet.headers.Count )
                        {
                            jsBuilderClient.Append( "};\r\n" );
                        }
                        AppendCellForJS( name, jsBuilderCommonJS, cell, funSheet.headers[colId].titleName, cnt == 1, cnt == funSheet.headers.Count );
                        AppendCellForJS( nameItemTypescript, jsBuilderTypescript, cell, funSheet.headers[colId].titleName, cnt == 1, cnt == funSheet.headers.Count );                       
                    }
                    catch ( JsonReaderException jex )
                    {
                        //Exception in parsing json
                        lstErrorMsg.Add( ErrorMsg.FormatError( name, rowId.ToString() + "行," + colId.ToString() + "列", jex.Message.ToString() ) );
                        result = false;
                    }
                    catch ( Exception ex ) //some other exception
                    {
                        lstErrorMsg.Add( ErrorMsg.Error( name, ex.ToString() ) );
                        result = false;
                    }

                    if ( cnt != funSheet.headers.Count )
                    {
                        builderServer.Append( "," );
                        builderClient.Append( "," );
                        protobufTextBuilder.Append( "\t" );
                    }
                }

                if ( funSheet.headers.Count != 0 )
                {
                    builderServer.Append( "\r\n" );
                    builderClient.Append( "\r\n" );
                    protobufTextBuilder.Append( "\r\n" );
                }
            }

            strResultServer = builderServer.ToString();
            strResultClient = builderClient.ToString();
            strJSResultCommonJS = jsBuilderCommonJS.ToString();
            strJSResultClient = jsBuilderClient.ToString();
            strJSContentTypescript = jsBuilderTypescript.ToString();
            strProtobuf = protobufBuilder.ToString();
            strProtobufText = protobufTextBuilder.ToString();

            return result;
        }

        private void AppendCellForCSV( StringBuilder builder, String strCell )
        {
            bool bWithQuota = ( strCell.IndexOf( ',' ) != -1 || strCell.IndexOf( '\"' ) != -1 );

            if ( bWithQuota )
            {
                builder.Append( "\"" );
            }

            foreach ( char c in strCell )
            {
                if ( c == '\"' )
                {
                    builder.Append( "\"" );
                }
                builder.Append( c );
            }

            if ( bWithQuota )
            {
                builder.Append( "\"" );
            }
        }

        private void AppendCellForJS( string dbName, StringBuilder builder, String strCell, string propertyName, bool bFirst, bool bLast )
        {
            if ( bFirst )
            {
                builder.Append( "\r\n" );
                builder.Append( dbName + "[\"" + strCell + "\"] = \r\n{\r\n" );
            }
            StringBuilder propertyContent = new StringBuilder();
            bool isJSON = Regex.IsMatch( propertyName, "^JSON_.*$" );
            if ( !strCell.Equals( "" ) )
            {
                if ( isJSON )
                {
                    List<string> errorMsgs = new List<string>();
                    try
                    {
                        var obj = JToken.Parse( strCell );
                    }
                    catch ( JsonReaderException jex )
                    {
                        //Exception in parsing json
                        Console.WriteLine( jex.Message );
                        throw jex;
                    }
                    catch ( Exception ex ) //some other exception
                    {
                        Console.WriteLine( ex.ToString() );
                        throw ex;
                    }
                    propertyContent.Append( strCell );
                }
                else
                {
                    bool bWithQuota = ( ( !Regex.IsMatch( strCell, "^-?[0-9]+$" ) && !Regex.IsMatch( strCell, "^(-?\\d+)(\\.\\d+)?$" ) ) || strCell.IndexOf( '\"' ) != -1 );

                    if ( bWithQuota )
                    {
                        propertyContent.Append( "\"" );
                    }

                    foreach ( char c in strCell )
                    {
                        if ( c == '\"' )
                        {
                            propertyContent.Append( "\\" );
                        }
                        propertyContent.Append( c );
                    }

                    if ( bWithQuota )
                    {
                        propertyContent.Append( "\"" );
                    }
                }
            }
            else
            {
                propertyContent.Append( "0" );
            }

            builder.Append( "\t" + propertyName + "\t:\t" + propertyContent.ToString() );

            if ( bLast )
            {
                builder.Append( "\r\n};\r\n" );
            }
            else
            {
                builder.Append( ",\r\n" );
            }
        }

        private void AppendHeadForProto( StringBuilder builder, string propertyName, string colId )
        {
            if ( Regex.IsMatch( propertyName, "FIX_.*" ) )
            {
                propertyName = propertyName.Substring( propertyName.IndexOf( "_" ) + 1 );
                builder.Append( "\toptional fixed32 " + propertyName + "=" + colId + ";\r\n" );
            }
            else if ( Regex.IsMatch( propertyName, "INT_" ) )
            {
                propertyName = propertyName.Substring( propertyName.IndexOf( "_" ) + 1 );
                builder.Append( "\toptional int32 " + propertyName + "=" + colId + ";\r\n" );
            }
            else if ( Regex.IsMatch( propertyName, "FLT_" ) )
            {
                propertyName = propertyName.Substring( propertyName.IndexOf( "_" ) + 1 );
                builder.Append( "\toptional float " + propertyName + "=" + colId + ";\r\n" );
            }
            else
            {
                builder.Append( "\toptional string " + propertyName + "=" + colId + ";\r\n" );
            }
        }

        private void AppendCellForProto( StringBuilder builder, string propertyName, string cell )
        {
            if ( cell != "" )
            {
                if ( Regex.IsMatch( propertyName, "FIX_.*" ) )
                {
                    Int32.Parse( cell );
                }
                else if ( Regex.IsMatch( propertyName, "INT_" ) )
                {
                    Int32.Parse( cell );
                }
                else if ( Regex.IsMatch( propertyName, "FLT_" ) )
                {
                    Single.Parse( cell );
                }
            }            

            builder.Append( cell );
        }
    }
}
