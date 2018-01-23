using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ExcelTools
{
    partial class Main
    {
        // type 0 初始版转text 1 初始版转json 2后期版做diff 3后期版生成外文json
        private void UISceneTextProduceMain()
        {
            btnUISceneTextProduce.Enabled = false;

            while ( Util.IsFileInUse( m_strUISceneTextCsvPath ) )
            {
                MessageBox.Show( "请关闭" + m_strUISceneTextCsvPath );
            }

            while ( Util.IsFileInUse( m_strUISceneTextNewCsvPath ) )
            {
                MessageBox.Show( "请关闭" + m_strUISceneTextNewCsvPath );
            }

            // 处理开始
            Thread processThread = new Thread( new ParameterizedThreadStart( UISceneTextProduceProcess ) );
            processThread.Start();
        }

        private void UISceneTextProduceProcess( object o )
        {
            List<string> errorMsgs = new List<string>();

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblUISceneTextProduceResult.Text = "处理中";
            } );

            UISceneTextCsvControl csv = new UISceneTextCsvControl();
            UISceneTextCsv oldCsv = new UISceneTextCsv();
            UISceneTextCsv newCsv = new UISceneTextCsv();
            UISceneTextCsv resCsv = new UISceneTextCsv();
            csv.Read( m_strUISceneTextCsvPath, out oldCsv );
            csv.Read( m_strUISceneTextNewCsvPath, out newCsv );
            foreach ( var item in newCsv.dictText )
            {
                if ( oldCsv.dictText.ContainsKey( item.Key ) )
                    resCsv.dictText[item.Key] = oldCsv.dictText[item.Key];
                else
                    resCsv.dictText[item.Key] = newCsv.dictText[item.Key];
            }
            csv.Create( m_strUISceneTextNewCsvPath, resCsv );

            this.Invoke( (UpdateLabelStateDelegate)delegate()
            {
                lblUISceneTextProduceResult.Text = "完成";
            } );

            this.Invoke( (UpdateButtonStateDelegate)delegate()
            {
                btnUISceneTextProduce.Enabled = true;
            } );
        }
    }
}
