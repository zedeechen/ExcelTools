using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ExcelTools
{
    partial class Main
    {
        void FuncTextDiffPrintDifferenceMain()
        {
            if ( m_strFuncTextDiffOldTextPath == m_strFuncTextDiffTextDifferencePath || m_strFuncTextDiffNewTextPath == m_strFuncTextDiffTextDifferencePath )
            {
                MessageBox.Show( "输入和输出路径不能相同" );
                return;
            }

            gbxFuncTextDiffIOSetup.Enabled = false;
            lblFuncTextDiffPrintDifferenceResult.Text = "未处理";
            btnFuncTextDiffPrintDifference.Enabled = false;

            while ( Util.IsFileInUse( m_strFuncTextDiffTextDifferencePath ) )
            {
                MessageBox.Show( "请先关闭" + m_strFuncTextDiffTextDifferencePath );
            }

            Thread processThread = new Thread( new ParameterizedThreadStart( FuncTextDiffPrintDifferenceProcess ) );
            processThread.Start();
        }

        void FuncTextDiffPrintDifferenceProcess( object o )
        {
            List<string> errorMsgs =  new List<string>();
            TextSheetControl textControl = new TextSheetControl();
            
            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffPrintDifferenceResult.Text = "读取中。。";
            } );

            TextSheet oldText = new TextSheet();
            TextSheet newText = new TextSheet();
            
            List<string> lstErrorMsg;
            textControl.Read( m_strFuncTextDiffOldTextPath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour, out oldText, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );
            textControl.Read( m_strFuncTextDiffNewTextPath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour, out newText, out lstErrorMsg );
            errorMsgs.AddRange( lstErrorMsg );

            if ( errorMsgs.Count != 0 )
            {
                this.Invoke( (UpdateLabelStateDelegate)delegate()
                {
                    lblFuncTextDiffPrintDifferenceResult.Text = "读取出错";
                } );
                this.Invoke( (CreateFormErrorResultDelegate)delegate()
                {
                    Form form = new ErrorResult( errorMsgs );
                    form.ShowDialog();
                } );
                this.Invoke( (UpdateButtonStateDelegate)delegate()
                {
                    gbxFuncTextDiffIOSetup.Enabled = true;
                    btnFuncTextDiffPrintDifference.Enabled = true;
                } );
                return;
            }

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffPrintDifferenceResult.Text = "写入中。。";
            } );

            textControl.Create( m_strFuncTextDiffTextDifferencePath, m_strFuncTextDiffSheetName, m_bFuncTextDiffExistLineFour );
            YYExcel outExcel = new YYExcel();
            outExcel.Open( m_strFuncTextDiffTextDifferencePath, m_strFuncTextDiffSheetName, YYExcel.Authority.A_READ_AND_WRITE );
            int wOldColor = Convert.ToInt32( m_dFuncTextDiffOldItemColor );
            int wNewColor = Convert.ToInt32( m_dFuncTextDiffNewItemColor );

            foreach ( var item in newText.dictText )
            {
                if ( oldText.dictText.ContainsKey( item.Key ) )
                {
                    if ( item.Value != oldText.dictText[item.Key] )
                    {
                        textControl.SetItem( outExcel, outExcel.GetRowsCount() + 1, item.Key, item.Value, wOldColor );
                    }
                }
                else
                {
                    textControl.SetItem( outExcel, outExcel.GetRowsCount() + 1, item.Key, item.Value, wNewColor );
                }
            }

            outExcel.SaveAs( m_strFuncTextDiffTextDifferencePath );
            outExcel.Close();

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblFuncTextDiffPrintDifferenceResult.Text = "完成";
            } );

            this.Invoke( (UpdateButtonStateDelegate)delegate()
            {
                gbxFuncTextDiffIOSetup.Enabled = true;
                btnFuncTextDiffPrintDifference.Enabled = true;
            } );
        }
    }
}
