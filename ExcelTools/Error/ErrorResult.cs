using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelTools
{
    public partial class ErrorResult : Form
    {
        public ErrorResult( List<string> errors )
        {
            InitializeComponent();
            lvwErrorResult.Items.Clear();
            lvwErrorResult.BeginUpdate();
            int cnt = 0;
            foreach ( string strError in errors )
            {
                ++cnt;
                ListViewItem lvi = new ListViewItem( cnt.ToString() );
                lvi.SubItems.Add( strError );
                lvwErrorResult.Items.Add( lvi );
            }
            lvwErrorResult.EndUpdate();
        }
    }
}
