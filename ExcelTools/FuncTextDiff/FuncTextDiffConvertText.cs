using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;

namespace ExcelTools
{
    partial class Main
    {
        void FuncTextDiffConvertTextMain()
        {
            if ( m_strFuncTextDiffOldTextPath == m_strFuncTextDiffNewTextPath || m_strFuncTextDiffTextDifferencePath == m_strFuncTextDiffNewTextPath )
            {
                MessageBox.Show( "输入和输出路径不能相同" );
                return;
            }

            gbxFuncTextDiffIOSetup.Enabled = false;
            lblFuncTextDiffConvertTextResult.Text = "未处理";
            btnFuncTextDiffConvertText.Enabled = false;           

            while ( Util.IsFileInUse( m_strFuncTextDiffNewTextPath ) )
            {
                MessageBox.Show( "请先关闭" + m_strFuncTextDiffNewTextPath );
            }

            Thread processThread = new Thread( new ParameterizedThreadStart( FuncTextDiffConvertTextProcess ) );
            processThread.Start();
        }

        void FuncTextDiffConvertTextProcess( object o )
        {
            List<string> errorMsgs = new List<string>();
            TextSheetControl textControl = new TextSheetControl();
            
            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffConvertTextResult.Text = "读取中。。";
            } );

            TextSheet oldText = new TextSheet();
            TextSheet diffText = new TextSheet();
            
            List<string> lstErrorMsg;
            textControl.Read( m_strFuncTextDiffOldTextPath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour, out oldText, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );
            textControl.Read( m_strFuncTextDiffTextDifferencePath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour, out diffText, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );

            if ( errorMsgs.Count != 0 )
            {
                this.Invoke( (UpdateLabelStateDelegate)delegate()
                {
                    lblFuncTextDiffConvertTextResult.Text = "读取出错";
                } );
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
                this.Invoke( (UpdateButtonStateDelegate)delegate()
                {
                    gbxFuncTextDiffIOSetup.Enabled = true;
                    btnFuncTextDiffConvertText.Enabled = true;
                } );
                return;
            }

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffConvertTextResult.Text = "写入中。。";
            } );

            textControl.Create( m_strFuncTextDiffNewTextPath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour );

            TextSheet newText = new TextSheet();
            foreach ( var item in diffText.dictText )
                newText.dictText.Add( item.Key, item.Value );
            foreach ( var item in oldText.dictText )
                if ( !diffText.dictText.ContainsKey( item.Key ) )
                    newText.dictText.Add( item.Key, item.Value );

            textControl.Update( m_strFuncTextDiffNewTextPath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour, newText, out lstErrorMsg );

            YYExcel outExcel = new YYExcel();
            outExcel.Open( m_strFuncTextDiffNewTextPath, m_strFuncTextDiffSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            int wOldColor = Convert.ToInt32( m_dFuncTextDiffOldItemColor );
            int wNewColor = Convert.ToInt32( m_dFuncTextDiffNewItemColor );

            int row = m_bFuncTextDiffExistLineFour ? 5 : 4;
            foreach ( var item in newText.dictText )
            {
                if ( diffText.dictText.ContainsKey( item.Key ) )
                {
                    if ( oldText.dictText.ContainsKey( item.Key ) )
                        outExcel.setRangeInteriorColor( row, 1, row, 2, wOldColor );
                    else
                        outExcel.setRangeInteriorColor( row, 1, row, 2, wNewColor );
                }
                row++;
            }

            outExcel.SaveAs( m_strFuncTextDiffNewTextPath );
            outExcel.Close();

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffConvertTextResult.Text = "完成";
            } );

            this.Invoke( (UpdateButtonStateDelegate)delegate()
            {
                gbxFuncTextDiffIOSetup.Enabled = true;
                btnFuncTextDiffConvertText.Enabled = true;
            } );
        }
        
    }
}
