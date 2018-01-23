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
    public partial class ColorPreview : Form
    {
        // Color
        public static string[] color = { "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF",
                                         "#FFFF00", "#FF00FF", "#00FFFF", "#800000", "#008000",
                                         "#000080", "#808000", "#800080", "#008080", "#C0C0C0",
                                         "#808080", "#9999FF", "#993366", "#FFFFCC", "#CCFFFF",
                                         "#660066", "#FF8080", "#0066CC", "#CCCCFF", "#000080",
                                         "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000",
                                         "#008080", "#0000FF", "#00CCFF", "#CCFFFF", "#CCFFCC",
                                         "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
                                         "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900",
                                         "#FF6600", "#666699", "#969696", "#003366", "#339966",
                                         "#003300", "#333300", "#993300", "#993366", "#333399",
                                         "#333333" };

        public ColorPreview()
        {
            InitializeComponent();

            this.listView1.View = View.LargeIcon;

            this.listView1.BeginUpdate();

            for ( int i = 1; i <= 56; i++ )
            {
                ListViewItem lvi = new ListViewItem( " " + i.ToString() + " " );

                lvi.BackColor = ColorTranslator.FromHtml( color[i - 1] );
                lvi.ForeColor = Color.Gray;

                this.listView1.Items.Add( lvi );
            }

            this.listView1.EndUpdate();
        }
    }
}
