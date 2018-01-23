using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using System.Collections;

namespace ExcelTools
{
    public partial class Main : Form
    {
        delegate void CreateFormErrorResultDelegate();
        delegate void CreateMessageBoxDelegate();
        delegate void UpdateLabelStateDelegate();
        delegate void UpdateLabelStateDelegateWithResult( bool res );
        delegate void UpdateButtonStateDelegate();
        delegate void UpdateFunctionCSVConvertResultDelegate( int idx, bool res );
        delegate void UpdateFunctionCustomResultDelegate( int idx, bool res );
        delegate void UpdateFunctionVersionDiffResultDelegate( string key, int res );
        delegate void UpdateFunctionVersionLvwMouseDoubleClickEvent();

        public Main()
        {
            InitializeComponent();
            InitFuncCSVConverterDefaultValue();
            InitFuncTextDiffDefaultValue();
            InitFuncCustomizerDefaultValue();
            InitFuncVersionDiffDefaultValue();
            InitUISceneTextDefaultValue();
            InitSheetCopyDefaultValue();
        }

        //////////////////////////////////////////////////////////////////////////
        // FunctionCSVConverter

        private string m_strFuncCSVConverterExcelDirPath;
        private string m_strFuncCSVConverterDirPathServer;
        private string m_strFuncCSVConverterDirPathCSVString;
        private string m_strFuncCSVConverterDirPathClient;
        private string m_strFuncCSVConverterDirPathCommonJS;
        private string m_strFuncCSVConverterDirPathTypescript;
        private string m_strFuncCSVConverterDirPathProtobuf;
        private string m_strFuncCSVConverterDirPathText;

        private string m_strFuncCSVConverterTextName;
        private string m_strFuncCSVConverterTextPath;
        private string m_strFuncCSVConverterSheetName;
        private string m_strFuncCSVConverterTextColId;
        private string m_strFuncCSVConverterTextColName;

        private bool m_bFuncCSVConverterUseCSVChecked;
        private bool m_bFuncCSVConverterUseCSVStringChecked;
        private bool m_bFuncCSVConverterUseJSChecked;
        private bool m_bFuncCSVConverterUseCommonJSChecked;
        private bool m_bFuncCSVConverterUseTypescriptChecked;
        private bool m_bFuncCSVConverterUseProtobufferChecked;
        private bool m_bFuncCSVConverterUseTextChecked;

        private bool m_bFuncCSVConverterExistLineFour;
        private bool m_bFuncCSVConverterProduceText;

        private void InitFuncCSVConverterDefaultValue()
        {
            txtFuncCSVConverterExcelDirPath.Text = Properties.Settings.Default.FuncCSVConverterExcelDirPath;
            txtFuncCSVConverterDirPathServer.Text = Properties.Settings.Default.FuncCSVConverterDirPathServer;
            txtFuncCSVConverterDirPathClient.Text = Properties.Settings.Default.FuncCSVConverterDirPathClient;
            txtFuncCSVConverterDirPathCSVString.Text = Properties.Settings.Default.FuncCSVConverterDirPathCSVString;
            txtFuncCSVConverterDirPathCommonJS.Text = Properties.Settings.Default.FuncCSVConverterDirPathCommonJS;
            txtFuncCSVConverterDirPathTypescript.Text = Properties.Settings.Default.FuncCSVConverterDirPathTypescript;
            txtFuncCSVConverterDirPathProtobuffer.Text = Properties.Settings.Default.FuncCSVConverterDirPathProtobuffer;
            txtFuncCSVConverterDirPathText.Text = Properties.Settings.Default.FuncCSVConverterDirPathText;
            txtFuncCSVConverterTextName.Text = Properties.Settings.Default.FuncCSVConverterTextName;
            txtFuncCSVConverterSheetName.Text = Properties.Settings.Default.FuncCSVConverterSheetName;
            chkFuncCSVConverterUseCSV.Checked = Properties.Settings.Default.FuncCSVConverterUseCSV;
            chkFuncCSVConverterUseCSVString.Checked = Properties.Settings.Default.FuncCSVConverterUseCSVString;
            chkFuncCSVConverterUseJS.Checked = Properties.Settings.Default.FuncCSVConverterUseJS;
            chkFuncCSVConverterUseCommonJS.Checked = Properties.Settings.Default.FuncCSVConverterUseCommonJS;
            chkFuncCSVConverterUseTypescript.Checked = Properties.Settings.Default.FuncCSVConverterUseTypescript;
            chkFuncCSVConverterUseProtobuffer.Checked = Properties.Settings.Default.FuncCSVConverterUseProtobuffer;
            chkFuncCSVConverterUseText.Checked = Properties.Settings.Default.FuncCSVConverterUseText;
            chkFuncCSVConverterExistLineFour.Checked = Properties.Settings.Default.FuncCSVConverterExistLineFour;
            txtFuncCSVConverterTextColId.Text = Properties.Settings.Default.FuncCSVConverterTextColId;
            txtFuncCSVConverterTextColName.Text = Properties.Settings.Default.FuncCSVConverterTextColName;
        }

        private void SaveFuncCSVConverterCurrentValue()
        {
            Properties.Settings.Default.FuncCSVConverterExcelDirPath = txtFuncCSVConverterExcelDirPath.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathServer = txtFuncCSVConverterDirPathServer.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathClient = txtFuncCSVConverterDirPathClient.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathCSVString = txtFuncCSVConverterDirPathCSVString.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathCommonJS = txtFuncCSVConverterDirPathCommonJS.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathTypescript = txtFuncCSVConverterDirPathTypescript.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathProtobuffer = txtFuncCSVConverterDirPathProtobuffer.Text;
            Properties.Settings.Default.FuncCSVConverterDirPathText = txtFuncCSVConverterDirPathText.Text;
            Properties.Settings.Default.FuncCSVConverterTextName = txtFuncCSVConverterTextName.Text;
            Properties.Settings.Default.FuncCSVConverterSheetName = txtFuncCSVConverterSheetName.Text;
            Properties.Settings.Default.FuncCSVConverterUseCSV = chkFuncCSVConverterUseCSV.Checked;
            Properties.Settings.Default.FuncCSVConverterUseCSVString = chkFuncCSVConverterUseCSVString.Checked;
            Properties.Settings.Default.FuncCSVConverterUseJS = chkFuncCSVConverterUseJS.Checked;
            Properties.Settings.Default.FuncCSVConverterUseCommonJS = chkFuncCSVConverterUseCommonJS.Checked;
            Properties.Settings.Default.FuncCSVConverterUseTypescript = chkFuncCSVConverterUseTypescript.Checked;
            Properties.Settings.Default.FuncCSVConverterUseProtobuffer = chkFuncCSVConverterUseProtobuffer.Checked;
            Properties.Settings.Default.FuncCSVConverterUseText = chkFuncCSVConverterUseText.Checked;
            Properties.Settings.Default.FuncCSVConverterExistLineFour = chkFuncCSVConverterExistLineFour.Checked;
            Properties.Settings.Default.FuncCSVConverterTextColId = txtFuncCSVConverterTextColId.Text;
            Properties.Settings.Default.FuncCSVConverterTextColName = txtFuncCSVConverterTextColName.Text;

            Properties.Settings.Default.Save();
        }

        private int FuncCSVConvertInitialize( bool bRefresh = false )
        {
            m_strFuncCSVConverterExcelDirPath = txtFuncCSVConverterExcelDirPath.Text;
            m_strFuncCSVConverterDirPathServer = txtFuncCSVConverterDirPathServer.Text;
            m_strFuncCSVConverterDirPathCSVString = txtFuncCSVConverterDirPathCSVString.Text;
            m_strFuncCSVConverterDirPathClient = txtFuncCSVConverterDirPathClient.Text;
            m_strFuncCSVConverterDirPathCommonJS = txtFuncCSVConverterDirPathCommonJS.Text;
            m_strFuncCSVConverterDirPathTypescript = txtFuncCSVConverterDirPathTypescript.Text;
            m_strFuncCSVConverterDirPathProtobuf = txtFuncCSVConverterDirPathProtobuffer.Text;
            m_strFuncCSVConverterDirPathText = txtFuncCSVConverterDirPathText.Text;

            m_strFuncCSVConverterTextName = txtFuncCSVConverterTextName.Text;
            m_strFuncCSVConverterTextPath = txtFuncCSVConverterExcelDirPath.Text + "\\" + txtFuncCSVConverterTextName.Text;
            m_strFuncCSVConverterSheetName = txtFuncCSVConverterSheetName.Text;

            m_strFuncCSVConverterTextColId = txtFuncCSVConverterTextColId.Text;
            m_strFuncCSVConverterTextColName = txtFuncCSVConverterTextColName.Text;

            m_bFuncCSVConverterExistLineFour = chkFuncCSVConverterExistLineFour.Checked;
            m_bFuncCSVConverterUseCSVChecked = chkFuncCSVConverterUseCSV.Checked;
            m_bFuncCSVConverterUseCSVStringChecked = chkFuncCSVConverterUseCSVString.Checked;
            m_bFuncCSVConverterUseJSChecked = chkFuncCSVConverterUseJS.Checked;
            m_bFuncCSVConverterUseCommonJSChecked = chkFuncCSVConverterUseCommonJS.Checked;
            m_bFuncCSVConverterUseTypescriptChecked = chkFuncCSVConverterUseTypescript.Checked;
            m_bFuncCSVConverterUseProtobufferChecked = chkFuncCSVConverterUseProtobuffer.Checked;
            m_bFuncCSVConverterUseTextChecked = chkFuncCSVConverterUseText.Checked;
            m_bFuncCSVConverterUseTextChecked = chkFuncCSVConverterUseText.Checked;

            if ( m_strFuncCSVConverterExcelDirPath == string.Empty || !Directory.Exists( m_strFuncCSVConverterExcelDirPath ) )
            {
                MessageBox.Show( "Excel目录不存在" );
                return 1;
            }

            if ( m_bFuncCSVConverterUseCSVChecked && ( m_strFuncCSVConverterDirPathServer == string.Empty ||
                !Directory.Exists( m_strFuncCSVConverterDirPathServer ) ) )
            {
                MessageBox.Show( "CSV（数字）导出目录不存在" );
                return 1;
            }

            if ( m_bFuncCSVConverterUseCSVStringChecked && ( m_strFuncCSVConverterDirPathCSVString == string.Empty ||
                !Directory.Exists( m_strFuncCSVConverterDirPathCSVString ) ) )
            {
                MessageBox.Show( "CSV(字符串)导出目录不存在" );
                return 1;
            }

            if ( m_bFuncCSVConverterUseJSChecked && ( m_strFuncCSVConverterDirPathClient == string.Empty 
                || !Directory.Exists( m_strFuncCSVConverterDirPathClient ) ) )
            {
                MessageBox.Show( "JS导出目录不存在" );
                return 1;
            }

            if ( m_bFuncCSVConverterUseCommonJSChecked && (m_strFuncCSVConverterDirPathCommonJS == string.Empty 
                || !Directory.Exists(m_strFuncCSVConverterDirPathCommonJS)))
            {
                MessageBox.Show("CommonJS导出目录不存在");
                return 1;
            }

            if (m_bFuncCSVConverterUseTypescriptChecked && (m_strFuncCSVConverterDirPathTypescript == string.Empty
                || !Directory.Exists(m_strFuncCSVConverterDirPathTypescript)))
            {
                MessageBox.Show("Typescript导出目录不存在");
                return 1;
            }

            if ( m_bFuncCSVConverterUseProtobufferChecked && ( m_strFuncCSVConverterDirPathProtobuf == string.Empty
                || !Directory.Exists( m_strFuncCSVConverterDirPathProtobuf ) ) )
            {
                MessageBox.Show( "Protobuffer导出目录不存在" );
                return 1;
            }

            if ( m_bFuncCSVConverterUseTextChecked && ( m_strFuncCSVConverterDirPathText == string.Empty
                || !Directory.Exists( m_strFuncCSVConverterDirPathText ) ) )
            {
                MessageBox.Show( "textdb.js导出目录不存在" ); 
                return 1;
            }

            if ( m_strFuncCSVConverterSheetName == string.Empty )
            {
                MessageBox.Show( "Sheet名称未指定" );
                return 1;
            }

            if ( m_strFuncCSVConverterTextName == string.Empty )
            {
                MessageBox.Show( "Text表名称未指定" );
                return 1;
            }

            if ( !Regex.IsMatch( m_strFuncCSVConverterTextName, ".*.xlsx" ) )
            {
                MessageBox.Show( "Text表名称不规范，后缀应为xlsx" );
                return 1;
            }

			if ( bRefresh ) lvwFuncCSVConverterResult_Refresh();
            return 0;
        }

        private void btnFuncCSVConvert_Click( object sender, EventArgs e )
        {
            do 
            {
                if ( FuncCSVConvertInitialize( true ) != 0 )
                    break;
                m_bFuncCSVConverterProduceText = true;
                FuncCSVConverterMain();
            } while ( false );

            SaveFuncCSVConverterCurrentValue();
        }

        private void btnFuncCSVConvertOnly_Click( object sender, EventArgs e )
        {
            do
            {
                if ( FuncCSVConvertInitialize() != 0 )
                    break;
                m_bFuncCSVConverterProduceText = false;
                FuncCSVConverterMain();
            } while ( false );

            SaveFuncCSVConverterCurrentValue();
        }

        private void btnChangeFuncCSVConverterExcelDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择功能表导入目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterExcelDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathServer_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择CSV(数字)导出目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterDirPathServer.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathCSVString_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择CSV(字符串)导出目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterDirPathCSVString.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathClient_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择JS导出目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterDirPathClient.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathCommonJS_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择CommonJS导出目录";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtFuncCSVConverterDirPathCommonJS.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathTypescript_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择Typescript导出目录";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtFuncCSVConverterDirPathTypescript.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathProtobuffer_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择Protobuffer导出目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterDirPathProtobuffer.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCSVConverterDirPathText_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择textdb.js导出目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCSVConverterDirPathText.Text = dlg.SelectedPath;
            }
        }

        private void txtFuncCSVConverterExcelDirPath_TextChanged( object sender, EventArgs e )
        {
            lvwFuncCSVConverterResult_Refresh();
        }

        private void lvwFuncCSVConverterResult_Refresh()
        {
            if ( !Directory.Exists( txtFuncCSVConverterExcelDirPath.Text ) )
            {
                lvwFuncCSVConverterResult.Items.Clear();
                return;
            }
            string[] fileNames = Directory.GetFiles( txtFuncCSVConverterExcelDirPath.Text, "*.xlsx", SearchOption.TopDirectoryOnly );

            // 提示已打开文件
            string ret = Util.GetOpenedExcelList( fileNames );
            if ( ret != string.Empty )
            {
                MessageBox.Show( ret + "\n如有需要，请保存后确定", "已打开Excel列表" );
            }
            // 从Array中移除已打开文件的副本
            fileNames = Array.FindAll( fileNames, Util.IsExcelOpened );

            // 结果列表初始化
            lvwFuncCSVConverterResult.Items.Clear();
            lvwFuncCSVConverterResult.BeginUpdate();
            int fLength = fileNames.Length;
            for ( int i = 0; i < fLength; i++ )
            {
                string name = Path.GetFileName( fileNames[i] );

                ListViewItem lvi = new ListViewItem( ( i + 1 ).ToString() );
                lvi.SubItems.Add( name );
                lvi.SubItems.Add( "尚未检查" );
                lvi.Selected = true;
                lvwFuncCSVConverterResult.Items.Add( lvi );
            }
            lvwFuncCSVConverterResult.EndUpdate();
        }

        private void lvwFuncCSVConverterResult_SelectedIndexChanged( object sender, EventArgs e )
        {
            lvwFuncCSVConverterResult.BeginUpdate();
            //foreach ( ListViewItem lvi in lvwFuncCSVConverterResult.Items )
            //{
            //    if ( Regex.IsMatch( lvi.SubItems[1].Text, "[0-9]+_.*" ) )
            //    {
            //        lvi.Selected = true;
            //    }
            //}
            lvwFuncCSVConverterResult.EndUpdate();
        }

        //////////////////////////////////////////////////////////////////////////
        // FunctionTextDiff

        private string m_strFuncTextDiffOldTextPath;
        private string m_strFuncTextDiffNewTextPath;
        private string m_strFuncTextDiffTextDifferencePath;
        private string m_strFuncTextDiffSheetName;
        private decimal m_dFuncTextDiffOldItemColor;
        private decimal m_dFuncTextDiffNewItemColor;
        private bool m_bFuncTextDiffExistLineFour;
        
        private void InitFuncTextDiffDefaultValue()
        {
            txtFuncTextDiffOldTextPath.Text = Properties.Settings.Default.FuncTextDiffOldTextPath;
            txtFuncTextDiffNewTextPath.Text = Properties.Settings.Default.FuncTextDiffNewTextPath;
            txtFuncTextDiffTextDifferencePath.Text = Properties.Settings.Default.FuncTextDiffTextDifferencePath;
            txtFuncTextDiffSheetName.Text = Properties.Settings.Default.FuncTextDiffSheetName;
            nudFuncTextDiffOldItemColor.Value = Properties.Settings.Default.FuncTextDiffOldItemColor;
            nudFuncTextDiffNewItemColor.Value = Properties.Settings.Default.FuncTextDiffNewItemColor;
            chkFuncTextDiffExistLineFour.Checked = Properties.Settings.Default.FuncTextDiffExistLineFour;
        }

        private void SaveFuncTextDiffCurrentValue()
        {
            Properties.Settings.Default.FuncTextDiffOldTextPath = txtFuncTextDiffOldTextPath.Text;
            Properties.Settings.Default.FuncTextDiffNewTextPath = txtFuncTextDiffNewTextPath.Text;
            Properties.Settings.Default.FuncTextDiffTextDifferencePath = txtFuncTextDiffTextDifferencePath.Text;
            Properties.Settings.Default.FuncTextDiffSheetName = txtFuncTextDiffSheetName.Text;
            Properties.Settings.Default.FuncTextDiffOldItemColor = nudFuncTextDiffOldItemColor.Value;
            Properties.Settings.Default.FuncTextDiffNewItemColor = nudFuncTextDiffNewItemColor.Value;
            Properties.Settings.Default.FuncTextDiffExistLineFour = chkFuncTextDiffExistLineFour.Checked;
            Properties.Settings.Default.Save();
        }

        private int FuncTextDiffInitialize()
        {
            m_strFuncTextDiffOldTextPath = txtFuncTextDiffOldTextPath.Text;
            m_strFuncTextDiffNewTextPath = txtFuncTextDiffNewTextPath.Text;
            m_strFuncTextDiffTextDifferencePath = txtFuncTextDiffTextDifferencePath.Text;
            m_strFuncTextDiffSheetName = txtFuncTextDiffSheetName.Text;
            m_dFuncTextDiffOldItemColor = nudFuncTextDiffOldItemColor.Value;
            m_dFuncTextDiffNewItemColor = nudFuncTextDiffNewItemColor.Value;
            m_bFuncTextDiffExistLineFour = chkFuncTextDiffExistLineFour.Checked;

            if ( m_strFuncTextDiffOldTextPath == string.Empty || !File.Exists( m_strFuncTextDiffOldTextPath ) )
            {
                MessageBox.Show( "旧版TEXT文件不存在" );
                return 1;
            }

            if ( rboFuncTextDiffPrintDifference.Checked && ( m_strFuncTextDiffNewTextPath == string.Empty || !File.Exists( m_strFuncTextDiffNewTextPath ) ) ) 
            {
                MessageBox.Show( "新版TEXT文件不存在" );
                return 1;
            }

            if ( rboFuncTextDiffConvertText.Checked && ( m_strFuncTextDiffTextDifferencePath == string.Empty || !File.Exists( m_strFuncTextDiffTextDifferencePath ) ) )
            {
                MessageBox.Show( "差异表文件不存在" );
                return 1;
            }

            if ( m_strFuncTextDiffSheetName == string.Empty )
            {
                MessageBox.Show( "Sheet名称未指定" );
                return 1;
            }

            return 0;
        }

        private void btnFuncTextDiffPrintDiff_Click( object sender, EventArgs e )
        {
            do
            {
                if ( FuncTextDiffInitialize() != 0 )
                    break;
                FuncTextDiffPrintDifferenceMain();
            } while ( false );

            SaveFuncTextDiffCurrentValue();
        }

        private void btnFuncTextDiffConvertText_Click( object sender, EventArgs e )
        {
            do
            {
                if ( FuncTextDiffInitialize() != 0 )
                    break;
                FuncTextDiffConvertTextMain();
            } while ( false );

            SaveFuncTextDiffCurrentValue();
        }

        private void rboFuncTextDiffPrintDifference_CheckedChanged( object sender, EventArgs e )
        {
            if ( rboFuncTextDiffPrintDifference.Checked )
            {
                lblFuncTextDiffProcessDescription.Text = "通过比较配表工具生成的新旧TEXT表生成差异表\n\n" +
                                                         "旧版本TEXT表(中文)+新版本TEXT表（中文）->TEXT差异表（中文）";
                lblFuncTextDiffOldTextPath.Text = "旧TEXT文件(输入)";
                lblFuncTextDiffNewTextPath.Text = "新TEXT文件(输入)";
                lblFuncTextDiffTextDifferencePath.Text = "差异表路径(输出)";
                gbxFuncTextDiffPrintDifferenceResult.Enabled = true;
                btnFuncTextDiffPrintDifference.Enabled = true;
                gbxFuncTextDiffConvertTextResult.Enabled = false;
                btnFuncTextDiffConvertText.Enabled = false;
            }
            else
            {
                lblFuncTextDiffProcessDescription.Text = "通过中文差异表翻译过来的外文差异表和旧版外文Text表生成新版外\n文TEXT表\n\n" +
                                                         "TEXT差异表（外语）+旧版本TEXT表（外语）->新版本TEXT表（外语）";
                lblFuncTextDiffOldTextPath.Text = "旧TEXT文件(输入)";
                lblFuncTextDiffNewTextPath.Text = "新TEXT文件(输出)";
                lblFuncTextDiffTextDifferencePath.Text = "差异表路径(输入)";
                gbxFuncTextDiffPrintDifferenceResult.Enabled = false;
                btnFuncTextDiffPrintDifference.Enabled = false;
                gbxFuncTextDiffConvertTextResult.Enabled = true;
                btnFuncTextDiffConvertText.Enabled = true;
            }
        }

        private void btnChangeFuncTextDiffOldTextPath_Click( object sender, EventArgs e )
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.Title = "旧版TEXT表打开路径";
            dlg.Filter = "|*.xlsx";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncTextDiffOldTextPath.Text = dlg.FileName;
            }
        }

        private void btnChangeFuncTextDiffNewTextPath_Click( object sender, EventArgs e )
        {
            if ( rboFuncTextDiffPrintDifference.Checked )
            {
                OpenFileDialog dlg = new OpenFileDialog();

                dlg.Title = "新版TEXT表打开路径";
                dlg.Filter = "|*.xlsx";

                if ( dlg.ShowDialog() == DialogResult.OK )
                {
                    txtFuncTextDiffNewTextPath.Text = dlg.FileName;
                }
            }
            else
            {
                SaveFileDialog dlg = new SaveFileDialog();

                dlg.Title = "新版TEXT表保存路径";
                dlg.Filter = "|*.xlsx";

                if ( dlg.ShowDialog() == DialogResult.OK )
                {
                    txtFuncTextDiffNewTextPath.Text = dlg.FileName;
                }
            }
        }

        private void btnChangeFuncTextDiffTextDifferencePath_Click( object sender, EventArgs e )
        {
            if ( rboFuncTextDiffPrintDifference.Checked )
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.Title = "差异表保存路径";
                dlg.Filter = "|*.xlsx";

                if ( dlg.ShowDialog() == DialogResult.OK )
                {
                    txtFuncTextDiffTextDifferencePath.Text = dlg.FileName;
                }
            }
            else
            {
                OpenFileDialog dlg = new OpenFileDialog();

                dlg.Title = "差异表打开路径";
                dlg.Filter = "|*.xlsx";

                if ( dlg.ShowDialog() == DialogResult.OK )
                {
                    txtFuncTextDiffTextDifferencePath.Text = dlg.FileName;
                }
            }
        }

        private void nudFuncTextDiffOldItemColor_ValueChanged( object sender, EventArgs e )
        {
            int colorIndex = Convert.ToInt32( nudFuncTextDiffOldItemColor.Value );
            if ( colorIndex > 56 || colorIndex < 1 )
                return;
            txtFuncTextDiffOldItemColor.BackColor = ColorTranslator.FromHtml( ColorPreview.color[colorIndex - 1] );
        }

        private void nudFuncTextDiffNewItemColor_ValueChanged( object sender, EventArgs e )
        {
            int colorIndex = Convert.ToInt32( nudFuncTextDiffNewItemColor.Value );
            if ( colorIndex > 56 || colorIndex < 1 )
                return;
            txtFuncTextDiffNewItemColor.BackColor = ColorTranslator.FromHtml( ColorPreview.color[colorIndex - 1] );
        }

        //////////////////////////////////////////////////////////////////////////
        // FuncCustomizer

        private string m_strFuncCustomizerCustomPath;
        private string m_strFuncCustomizerCustomSheetName;
        private string m_strFuncCustomizerInputFunctionDirPath;
        private string m_strFuncCustomizerOutputFunctionDirPath;
        private string m_strFuncCustomizerOldFunctionDirPath;
        private string m_strFuncCustomizerCustomClashPath;
        private string m_strFuncCustomizerFuncSheetName;
        private decimal m_dFuncCustomizerOldItemColor;
        private decimal m_dFuncCustomizerNewItemColor;
        private bool m_bFuncCustomizerExistLineFour;

        private void InitFuncCustomizerDefaultValue()
        {
            txtFuncCustomizerCustomPath.Text = Properties.Settings.Default.FuncCustomizerCustomPath;
            cboFuncCustomizerCustomSheetName.Text = Properties.Settings.Default.FuncCustomizerCustomSheetName;
            txtFuncCustomizerInputFunctionDirPath.Text = Properties.Settings.Default.FuncCustomizerInputFunctionDirPath;
            txtFuncCustomizerOutputFunctionDirPath.Text = Properties.Settings.Default.FuncCustomizerOutputFunctionDirPath;
            txtFuncCustomizerOldFunctionDirPath.Text = Properties.Settings.Default.FuncCustomizerOldFunctionDirPath;
            txtFuncCustomizerCustomClashPath.Text = Properties.Settings.Default.FuncCustomizerCustomClashPath;
            txtFuncCustomizerFuncSheetName.Text = Properties.Settings.Default.FuncCustomizerFuncSheetName;
            nudFuncCustomizerOldItemColor.Value = Properties.Settings.Default.FuncCustomizerOldItemColor;
            nudFuncCustomizerNewItemColor.Value = Properties.Settings.Default.FuncCustomizerNewItemColor;
            chkFuncCustomizerExistLineFour.Checked = Properties.Settings.Default.FuncCustomizerExistLineFour;
        }

        private void SaveFuncCustomizerCurrentValue()
        {
            Properties.Settings.Default.FuncCustomizerCustomPath = txtFuncCustomizerCustomPath.Text;
            Properties.Settings.Default.FuncCustomizerCustomSheetName = cboFuncCustomizerCustomSheetName.Text;
            Properties.Settings.Default.FuncCustomizerInputFunctionDirPath = txtFuncCustomizerInputFunctionDirPath.Text;
            Properties.Settings.Default.FuncCustomizerOutputFunctionDirPath = txtFuncCustomizerOutputFunctionDirPath.Text;
            Properties.Settings.Default.FuncCustomizerOldFunctionDirPath = txtFuncCustomizerOldFunctionDirPath.Text;
            Properties.Settings.Default.FuncCustomizerCustomClashPath = txtFuncCustomizerCustomClashPath.Text;
            Properties.Settings.Default.FuncCustomizerFuncSheetName = txtFuncCustomizerFuncSheetName.Text;
            Properties.Settings.Default.FuncCustomizerOldItemColor = nudFuncCustomizerOldItemColor.Value;
            Properties.Settings.Default.FuncCustomizerNewItemColor = nudFuncCustomizerNewItemColor.Value;
            Properties.Settings.Default.FuncCustomizerExistLineFour = chkFuncCustomizerExistLineFour.Checked;
            Properties.Settings.Default.Save();
        }

        private int FuncCustomizeInitialize()
        {
            m_strFuncCustomizerCustomPath = txtFuncCustomizerCustomPath.Text;
            m_strFuncCustomizerCustomSheetName = cboFuncCustomizerCustomSheetName.Text;
            m_strFuncCustomizerInputFunctionDirPath = txtFuncCustomizerInputFunctionDirPath.Text;
            m_strFuncCustomizerOutputFunctionDirPath = txtFuncCustomizerOutputFunctionDirPath.Text;
            m_strFuncCustomizerOldFunctionDirPath = txtFuncCustomizerOldFunctionDirPath.Text;
            m_strFuncCustomizerCustomClashPath = txtFuncCustomizerCustomClashPath.Text;
            m_strFuncCustomizerFuncSheetName = txtFuncCustomizerFuncSheetName.Text;
            m_dFuncCustomizerOldItemColor = nudFuncCustomizerOldItemColor.Value;
            m_dFuncCustomizerNewItemColor = nudFuncCustomizerNewItemColor.Value;
            m_bFuncCustomizerExistLineFour = chkFuncCustomizerExistLineFour.Checked;

            if ( m_strFuncCustomizerCustomPath == string.Empty || !File.Exists( m_strFuncCustomizerCustomPath ) )
            {
                MessageBox.Show( "订制表文件不存在" );
                return 1;
            }

            if ( m_strFuncCustomizerCustomSheetName == null || m_strFuncCustomizerCustomSheetName == string.Empty )
            {
                MessageBox.Show( "订制表Sheet未指定" );
                return 1;
            }

            if ( m_strFuncCustomizerInputFunctionDirPath == string.Empty || !Directory.Exists( m_strFuncCustomizerInputFunctionDirPath ) )
            {
                MessageBox.Show( "功能表路径（输入）不存在" );
                return 1;
            }

            if ( m_strFuncCustomizerOutputFunctionDirPath == string.Empty || !Directory.Exists( m_strFuncCustomizerOutputFunctionDirPath ) )
            {
                MessageBox.Show( "功能表路径（输出）不存在" );
                return 1;
            }

            if ( m_strFuncCustomizerFuncSheetName == string.Empty )
            {
                MessageBox.Show( "功能表Sheet名称未指定" );
                return 1;
            }

            if ( m_strFuncCustomizerInputFunctionDirPath == m_strFuncCustomizerOutputFunctionDirPath )
            {
                MessageBox.Show( "功能表输入输出目录不能相同" );
                return 1;
            }

            return 0;
        }

        private void btnFuncCustomize_Click( object sender, EventArgs e )
        {
            do
            {
                if ( FuncCustomizeInitialize() != 0 )
                    break;
                FuncCustomizerMain();
            } while ( false );
            SaveFuncCustomizerCurrentValue();
        }

        private void btnChangeFuncCustomizerCustomPath_Click( object sender, EventArgs e )
        {
            OpenFileDialog dlg = new OpenFileDialog();

            dlg.Title = "多语言订制表打开路径";
            dlg.Filter = "|*.xlsx";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCustomizerCustomPath.Text = dlg.FileName;
            }

            cboFuncCustomizerCustomSheetName_DropDown( sender, e );
        }

        private void cboFuncCustomizerCustomSheetName_DropDown( object sender, EventArgs e )
        {
            m_strFuncCustomizerCustomPath = txtFuncCustomizerCustomPath.Text;

            if ( m_strFuncCustomizerCustomPath == string.Empty || !File.Exists( m_strFuncCustomizerCustomPath ) )
            {
                MessageBox.Show( "订制表文件不存在" );
                return;
            }

            cboFuncCustomizerCustomSheetName.BeginUpdate();
            cboFuncCustomizerCustomSheetName.Items.Clear();
            List<string> sheets;
            YYExcel excel = new YYExcel();
            excel.GetSheetsName( m_strFuncCustomizerCustomPath, out sheets, YYExcel.Authority.A_READ_ONLY );
            if ( sheets == null )
            {
                MessageBox.Show( "订制表读取失败" );
                return;
            }
            Dictionary<string, int> dict = new Dictionary<string, int>();
            foreach ( string item in sheets )
            {
                if ( dict.ContainsKey( item ) )
                {
                    continue;
                }
                dict.Add( item, 1 );
                cboFuncCustomizerCustomSheetName.Items.Add( item );
            }
            if ( cboFuncCustomizerCustomSheetName.Items.Count != 0 )
                cboFuncCustomizerCustomSheetName.Text = cboFuncCustomizerCustomSheetName.Items[0].ToString();
            cboFuncCustomizerCustomSheetName.EndUpdate();
            SaveFuncCustomizerCurrentValue();
        }

        private void btnChangeFuncCustomizerInputFunctionDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择中文功能表目录（输入）";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCustomizerInputFunctionDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCustomizerOutputFunctionDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择外文功能表目录（输出）";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCustomizerOutputFunctionDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCustomizerOldFunctionDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择旧版本功能表目录（可选）";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCustomizerOldFunctionDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeFuncCustomizerCustomClashPath_Click( object sender, EventArgs e )
        {
            SaveFileDialog dlg = new SaveFileDialog();

            dlg.Title = "订制冲突表保存路径";
            dlg.Filter = "|*.xlsx";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncCustomizerCustomClashPath.Text = dlg.FileName;
            }
        }

        private void nudFuncCustomizerOldItemColor_ValueChanged( object sender, EventArgs e )
        {
            int colorIndex = Convert.ToInt32( nudFuncCustomizerOldItemColor.Value );
            if ( colorIndex > 56 || colorIndex < 1 )
                return;
            txtFuncCustomizerOldItemColor.BackColor = ColorTranslator.FromHtml( ColorPreview.color[colorIndex - 1] );
        }

        private void nudFuncCustomizerNewItemColor_ValueChanged( object sender, EventArgs e )
        {
            int colorIndex = Convert.ToInt32( nudFuncCustomizerNewItemColor.Value );
            if ( colorIndex > 56 || colorIndex < 1 )
                return;
            txtFuncCustomizerNewItemColor.BackColor = ColorTranslator.FromHtml( ColorPreview.color[colorIndex - 1] );
        }

        //////////////////////////////////////////////////////////////////////////
        // FunctionVersionDiff

        private string m_strFuncVersionDiffOldExcelDirPath;
        private string m_strFuncVersionDiffNewExcelDirPath;
        private string m_strFuncVersionDiffSheetName;
        private bool m_bFuncVersionDiffExistLineFour;
        private Dictionary<string, SheetDiffInfo> m_dictFuncVersionDiff;

        private void InitFuncVersionDiffDefaultValue()
        {
            txtFuncVersionDiffOldExcelDirPath.Text = Properties.Settings.Default.FuncVersionDiffOldExcelDirPath;
            txtFuncVersionDiffNewExcelDirPath.Text = Properties.Settings.Default.FuncVersionDiffNewExcelDirPath;
            txtFuncVersionDiffSheetName.Text = Properties.Settings.Default.FuncVersionDiffSheetName;
            chkFuncVersionDiffExistLineFour.Checked = Properties.Settings.Default.FuncVersionDiffExistLineFour;
        }

        private void SaveFuncVersionDiffCurrentValue()
        {
            Properties.Settings.Default.FuncVersionDiffOldExcelDirPath = txtFuncVersionDiffOldExcelDirPath.Text;
            Properties.Settings.Default.FuncVersionDiffNewExcelDirPath = txtFuncVersionDiffNewExcelDirPath.Text;
            Properties.Settings.Default.FuncVersionDiffSheetName = txtFuncVersionDiffSheetName.Text;
            Properties.Settings.Default.FuncVersionDiffExistLineFour = chkFuncVersionDiffExistLineFour.Checked;
            Properties.Settings.Default.Save();
        }

        private int FuncVersionDiffInitialize()
        {
            m_strFuncVersionDiffOldExcelDirPath = txtFuncVersionDiffOldExcelDirPath.Text;
            m_strFuncVersionDiffNewExcelDirPath = txtFuncVersionDiffNewExcelDirPath.Text;
            m_strFuncVersionDiffSheetName = txtFuncVersionDiffSheetName.Text;
            m_bFuncVersionDiffExistLineFour = chkFuncVersionDiffExistLineFour.Checked;

            if ( m_strFuncVersionDiffOldExcelDirPath == string.Empty || !Directory.Exists( m_strFuncVersionDiffOldExcelDirPath ) )
            {
                MessageBox.Show( "旧版本目录不合法" );
                return 1;
            }

            if ( m_strFuncVersionDiffNewExcelDirPath == string.Empty || !Directory.Exists( m_strFuncVersionDiffNewExcelDirPath ) )
            {
                MessageBox.Show( "新版本目录不合法" );
                return 1;
            }

            if ( m_strFuncVersionDiffSheetName == string.Empty )
            {
                MessageBox.Show( "Sheet名称未指定" );
                return 1;
            }

            if ( m_strFuncVersionDiffOldExcelDirPath == m_strFuncVersionDiffNewExcelDirPath )
            {
                MessageBox.Show( "新旧版本目录不能相同" );
                return 1;
            }
            return 0;
        }

        private void btnFuncVersionDiff_Click( object sender, EventArgs e )
        {
            do
            {
                if ( FuncVersionDiffInitialize() != 0 )
                    break;
                FuncVersionDiffMain();
            } while ( false );
            SaveFuncVersionDiffCurrentValue();
        }

        private void btnChangeFuncVersionDiffOldExcelDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择老版本功能表目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncVersionDiffOldExcelDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangelFuncVersionDiffNewExcelDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择老版本功能表目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtFuncVersionDiffNewExcelDirPath.Text = dlg.SelectedPath;
            }
        }

        private void lvwFuncVersionDiffResult_MouseDoubleClick( object sender, MouseEventArgs e )
        {
            if ( lvwFuncVersionDiffResult.SelectedItems.Count != 0 )
            {
                if ( lvwFuncVersionDiffResult.SelectedItems[0].SubItems[2].Text == "修改" )
                {
                    int index = lvwFuncVersionDiffResult.SelectedItems[0].Index;
                    string name = lvwFuncVersionDiffResult.SelectedItems[0].SubItems[1].Text;
                    string oldPath = m_strFuncVersionDiffOldExcelDirPath + "\\" + name;
                    string newPath = m_strFuncVersionDiffNewExcelDirPath + "\\" + name;
                    Form form = new FunctionSheetDiff( this, oldPath, newPath, m_strFuncVersionDiffSheetName, m_bFuncVersionDiffExistLineFour, m_dictFuncVersionDiff[name] );
                    form.Show();
                }
            }
        }

        private int m_sortColumn;
        private void lvwFuncVersionDiffResult_ColumnClick( object sender, ColumnClickEventArgs e )
        {
            // Determine whether the column is the same as the last column clicked.
            if ( e.Column != m_sortColumn )
            {
                // Set the sort column to the new column.
                m_sortColumn = e.Column;
                // Set the sort order to ascending by default.
                lvwFuncVersionDiffResult.Sorting = SortOrder.Ascending;
            }
            else
            {
                // Determine what the last sort order was and change it.
                if ( lvwFuncVersionDiffResult.Sorting == SortOrder.Ascending )
                    lvwFuncVersionDiffResult.Sorting = SortOrder.Descending;
                else
                    lvwFuncVersionDiffResult.Sorting = SortOrder.Ascending;
            }
            // Call the sort method to manually sort.
            lvwFuncVersionDiffResult.Sort();
            // Set the ListViewItemSorter property to a new ListViewItemComparer
            // object.
            this.lvwFuncVersionDiffResult.ListViewItemSorter = new ListViewItemComparer( e.Column,
                                                              lvwFuncVersionDiffResult.Sorting );
        }

        //////////////////////////////////////////////////////////////////////////
        // UI Scene Text

        private string m_strUISceneTextCsvPath;
        private string m_strUISceneTextNewCsvPath;

        private void InitUISceneTextDefaultValue()
        {
            txtUISceneTextCsvPath.Text = Properties.Settings.Default.UISceneTextCsvPath;
            txtUISceneTextNewCsvPath.Text = Properties.Settings.Default.UISceneTextNewCsvPath;
        }

        private void SaveUISceneTextCurrentValue()
        {
            Properties.Settings.Default.UISceneTextCsvPath = txtUISceneTextCsvPath.Text;
            Properties.Settings.Default.UISceneTextNewCsvPath = txtUISceneTextNewCsvPath.Text;
            Properties.Settings.Default.Save();
        }      
                     
        private int UISceneTextProduceInitialize()
        {
            m_strUISceneTextCsvPath = txtUISceneTextCsvPath.Text;
            m_strUISceneTextNewCsvPath = txtUISceneTextNewCsvPath.Text;

            if ( !File.Exists( m_strUISceneTextCsvPath ) )
            {
                MessageBox.Show( "旧版CSV路径不存在" );
                return 1;
            }

            if ( !File.Exists( m_strUISceneTextNewCsvPath ) )
            {
                MessageBox.Show( "新版CSV路径不存在" );
                return 1;
            }

            return 0;
        }

        private void btnUISceneTextProduce_Click( object sender, EventArgs e )
        {
            do
            {
                if ( UISceneTextProduceInitialize() != 0 )
                    break;
                UISceneTextProduceMain();
            } while ( false );
            SaveUISceneTextCurrentValue();
        }

        private void btnChangeUISceneTextCsvPath_Click( object sender, EventArgs e )
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "旧版CSV打开路径";
            dlg.Filter = "|*.csv";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtUISceneTextCsvPath.Text = dlg.FileName;
            }
        }

        private void btnChangeUISceneTextNewCsvPath_Click( object sender, EventArgs e )
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "新版CSV打开路径";
            dlg.Filter = "|*.csv";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtUISceneTextNewCsvPath.Text = dlg.FileName;
            }
        }

        private void chkFuncCSVConverterUseCSV_CheckedChanged(object sender, EventArgs e)
        {
            if ( !chkFuncCSVConverterUseCSV.Checked )
            {
                lblFuncCSVConverterDirPathServer.Enabled = false;
                txtFuncCSVConverterDirPathServer.Enabled = false;
                btnChangeFuncCSVConverterDirPathServer.Enabled = false;
                m_bFuncCSVConverterUseCSVChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathServer.Enabled = true;
                txtFuncCSVConverterDirPathServer.Enabled = true;
                btnChangeFuncCSVConverterDirPathServer.Enabled = true;
                m_bFuncCSVConverterUseCSVChecked = true;
            }
        }

        private void chkFuncCSVConverterUseCSVString_CheckedChanged( object sender, EventArgs e )
        {
            if ( !chkFuncCSVConverterUseCSVString.Checked )
            {
                lblFuncCSVConverterDirPathCSVString.Enabled = false;
                txtFuncCSVConverterDirPathCSVString.Enabled = false;
                btnChangeFuncCSVConverterDirPathCSVString.Enabled = false;
                m_bFuncCSVConverterUseCSVStringChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathCSVString.Enabled = true;
                txtFuncCSVConverterDirPathCSVString.Enabled = true;
                btnChangeFuncCSVConverterDirPathCSVString.Enabled = true;
                m_bFuncCSVConverterUseCSVStringChecked = true;
            }
        }

        private void chkFuncCSVConverterUseCommonJS_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkFuncCSVConverterUseCommonJS.Checked)
            {
                lblFuncCSVConverterDirPathCommonJS.Enabled = false;
                txtFuncCSVConverterDirPathCommonJS.Enabled = false;
                btnChangeFuncCSVConverterDirPathCommonJS.Enabled = false;
                m_bFuncCSVConverterUseCommonJSChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathCommonJS.Enabled = true;
                txtFuncCSVConverterDirPathCommonJS.Enabled = true;
                btnChangeFuncCSVConverterDirPathCommonJS.Enabled = true;
                m_bFuncCSVConverterUseCommonJSChecked = true;
            }
        }

        private void chkFuncCSVConverterUseJS_CheckedChanged(object sender, EventArgs e)
        {
            if (!chkFuncCSVConverterUseJS.Checked)
            {
                lblFuncCSVConverterDirPathClient.Enabled = false;
                txtFuncCSVConverterDirPathClient.Enabled = false;
                btnChangeFuncCSVConverterDirPathClient.Enabled = false;
                m_bFuncCSVConverterUseJSChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathClient.Enabled = true;
                txtFuncCSVConverterDirPathClient.Enabled = true;
                btnChangeFuncCSVConverterDirPathClient.Enabled = true;
                m_bFuncCSVConverterUseJSChecked = true;
            }
        }

        private void chkFuncCSVConverterUseTypescritp_CheckedChanged( object sender, EventArgs e )
        {
            if ( !chkFuncCSVConverterUseTypescript.Checked )
            {
                lblFuncCSVConverterDirPathTypescript.Enabled = false;
                txtFuncCSVConverterDirPathTypescript.Enabled = false;
                btnChangeFuncCSVConverterDirPathTypescript.Enabled = false;
                m_bFuncCSVConverterUseTypescriptChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathTypescript.Enabled = true;
                txtFuncCSVConverterDirPathTypescript.Enabled = true;
                btnChangeFuncCSVConverterDirPathTypescript.Enabled = true;
                m_bFuncCSVConverterUseTypescriptChecked = true;
            }
        }

        private void chkFuncCSVConverterUseProtobuffer_CheckedChanged( object sender, EventArgs e )
        {
            if ( !chkFuncCSVConverterUseProtobuffer.Checked )
            {
                lblFuncCSVConverterDirPathProtobuffer.Enabled = false;
                txtFuncCSVConverterDirPathProtobuffer.Enabled = false;
                btnChangeFuncCSVConverterDirPathProtobuffer.Enabled = false;
                m_bFuncCSVConverterUseProtobufferChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathProtobuffer.Enabled = true;
                txtFuncCSVConverterDirPathProtobuffer.Enabled = true;
                btnChangeFuncCSVConverterDirPathProtobuffer.Enabled = true;
                m_bFuncCSVConverterUseProtobufferChecked = true;
            }
        }

        private void chkFuncCSVConverterUseText_CheckedChanged( object sender, EventArgs e )
        {
            if ( !chkFuncCSVConverterUseText.Checked )
            {
                lblFuncCSVConverterDirPathText.Enabled = false;
                txtFuncCSVConverterDirPathText.Enabled = false;
                btnChangeFuncCSVConverterDirPathText.Enabled = false;
                m_bFuncCSVConverterUseTextChecked = false;
            }
            else
            {
                lblFuncCSVConverterDirPathText.Enabled = true;
                txtFuncCSVConverterDirPathText.Enabled = true;
                btnChangeFuncCSVConverterDirPathText.Enabled = true;
                m_bFuncCSVConverterUseTextChecked = true;
            }
        }


        //////////////////////////////////////////////////////////////////////////
        // SheetCopy

        private string m_strSheetCopyOldExcelDirPath;
        private string m_strSheetCopyNewExcelDirPath;
        private string m_strSheetCopySheetName;
        private bool m_bSheetCopyExistLineFour;

        private void InitSheetCopyDefaultValue()
        {
            txtSheetCopyOldExcelDirPath.Text = Properties.Settings.Default.SheetCopyOldExcelDirPath;
            txtSheetCopyNewExcelDirPath.Text = Properties.Settings.Default.SheetCopyNewExcelDirPath;
            txtSheetCopySheetName.Text = Properties.Settings.Default.SheetCopySheetName;
            chkSheetCopyExistLineFour.Checked = Properties.Settings.Default.SheetCopyExistLineFour;
        }

        private void SaveSheetCopyCurrentValue()
        {
            Properties.Settings.Default.SheetCopyOldExcelDirPath = txtSheetCopyOldExcelDirPath.Text;
            Properties.Settings.Default.SheetCopyNewExcelDirPath = txtSheetCopyNewExcelDirPath.Text;
            Properties.Settings.Default.SheetCopySheetName = txtSheetCopySheetName.Text;
            Properties.Settings.Default.SheetCopyExistLineFour = chkSheetCopyExistLineFour.Checked;
            Properties.Settings.Default.Save();
        }

        private int SheetCopyInitialize()
        {
            m_strSheetCopyOldExcelDirPath = txtSheetCopyOldExcelDirPath.Text;
            m_strSheetCopyNewExcelDirPath = txtSheetCopyNewExcelDirPath.Text;
            m_strSheetCopySheetName = txtSheetCopySheetName.Text;
            m_bSheetCopyExistLineFour = chkSheetCopyExistLineFour.Checked;

            if ( m_strSheetCopyOldExcelDirPath == string.Empty || !Directory.Exists( m_strSheetCopyOldExcelDirPath ) )
            {
                MessageBox.Show( "旧版本目录不合法" );
                return 1;
            }

            if ( m_strSheetCopyNewExcelDirPath == string.Empty || !Directory.Exists( m_strSheetCopyNewExcelDirPath ) )
            {
                MessageBox.Show( "新版本目录不合法" );
                return 1;
            }

            if ( m_strSheetCopySheetName == string.Empty )
            {
                MessageBox.Show( "Sheet名称未指定" );
                return 1;
            }

            if ( m_strSheetCopyOldExcelDirPath == m_strSheetCopyNewExcelDirPath )
            {
                MessageBox.Show( "新旧版本目录不能相同" );
                return 1;
            }
            return 0;
        }

        private void btnCopySheet_Click( object sender, EventArgs e )
        {
            do
            {
                if ( SheetCopyInitialize() != 0 )
                    break;
                SheetCopyMain();
            } while ( false );
            SaveSheetCopyCurrentValue();
        }

        private void btnChangeSheetCopyOldExcelDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择老版本功能表目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtSheetCopyOldExcelDirPath.Text = dlg.SelectedPath;
            }
        }

        private void btnChangeSheetCopyNewExcelDirPath_Click( object sender, EventArgs e )
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            dlg.ShowNewFolderButton = false;
            dlg.Description = "选择新版本功能表目录";

            if ( dlg.ShowDialog() == DialogResult.OK )
            {
                txtSheetCopyNewExcelDirPath.Text = dlg.SelectedPath;
            }
        }

        private void lvwSheetCopyResult_ColumnClick( object sender, ColumnClickEventArgs e )
        {
            // Determine whether the column is the same as the last column clicked.
            if ( e.Column != m_sortColumn )
            {
                // Set the sort column to the new column.
                m_sortColumn = e.Column;
                // Set the sort order to ascending by default.
                lvwSheetCopyResult.Sorting = SortOrder.Ascending;
            }
            else
            {
                // Determine what the last sort order was and change it.
                if ( lvwSheetCopyResult.Sorting == SortOrder.Ascending )
                    lvwSheetCopyResult.Sorting = SortOrder.Descending;
                else
                    lvwSheetCopyResult.Sorting = SortOrder.Ascending;
            }
            // Call the sort method to manually sort.
            lvwSheetCopyResult.Sort();
            // Set the ListViewItemSorter property to a new ListViewItemComparer
            // object.
            this.lvwSheetCopyResult.ListViewItemSorter = new ListViewItemComparer( e.Column,
                                                              lvwSheetCopyResult.Sorting );
        }
    }
}
