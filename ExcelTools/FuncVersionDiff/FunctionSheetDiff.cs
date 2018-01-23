using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace ExcelTools
{
    public partial class FunctionSheetDiff : Form
    {
        public static Color cOldItem = Color.Red;
        public static Color cNewItem = Color.Yellow;

        public static Color cPushNew = Color.Green;
        public static Color cPushOld = Color.Green;

        private Main m_main;
        private string m_strOldExcelPath;
        private string m_strNewExcelPath;
        private string m_strSheetName;
        private string m_strExcelName;
        private bool m_bExsistLineFour;
        private SheetDiffInfo m_sheetDiffInfo;
        private ListViewItem[] m_oldExcelCache; //array to cache items for the virtual list
        private ListViewItem[] m_newExcelCache; //array to cache items for the virtual list

        private bool m_bSelectFromOld;
        private List<int> m_lstSelectedIndexs;

        delegate void UpdateButtonStateDelegate();
        delegate void RemoveListViewItemsDeltegate( int index );      

        public FunctionSheetDiff( Main main, string oldExcelPath, string newExcelPath, string sheetName, bool bExistLineFour, SheetDiffInfo diffInfo )
        {
            InitializeComponent();
            lvwOldExcelDiff.AddLinkedListView( lvwNewExcelDiff );

            m_main = main;
            m_strOldExcelPath = oldExcelPath;
            m_strNewExcelPath = newExcelPath;
            m_strSheetName = sheetName;
            m_bExsistLineFour = bExistLineFour;
            m_strExcelName = Path.GetFileName( m_strOldExcelPath );

            m_sheetDiffInfo = diffInfo;

            m_lstSelectedIndexs = new List<int>();
            m_oldExcelCache = new ListViewItem[m_sheetDiffInfo.diffItems.Count];
            m_newExcelCache = new ListViewItem[m_sheetDiffInfo.diffItems.Count];

            lblOldExcelPath.Text = m_strOldExcelPath;
            lblNewExcelPath.Text = m_strNewExcelPath;

            lvwOldExcelDiff.Clear();
            lvwNewExcelDiff.Clear();
            lvwOldExcelDiff.VirtualMode = true;
            lvwNewExcelDiff.VirtualMode = true;
            lvwOldExcelDiff.VirtualListSize = m_sheetDiffInfo.diffItems.Count;
            lvwNewExcelDiff.VirtualListSize = m_sheetDiffInfo.diffItems.Count;

            lvwOldExcelDiff.Columns.Add( "ID" );
            lvwOldExcelDiff.Columns[0].Width = 0;
            lvwNewExcelDiff.Columns.Add( "ID" );
            lvwNewExcelDiff.Columns[0].Width = 0;
            foreach ( var col in m_sheetDiffInfo.headers )
            {
                string tmp = col.Value.titleChineseName + "(" + col.Value.titleIndex + ")";
                lvwOldExcelDiff.Columns.Add( tmp, tmp.Length * 15, HorizontalAlignment.Center );
                lvwNewExcelDiff.Columns.Add( tmp, tmp.Length * 15, HorizontalAlignment.Center );
            }

            FunctionSheetDiffProcess();
        }

        private void FunctionSheetDiffProcess()
        {
            int rowCnt = -1;
            foreach ( int rowIndex in m_sheetDiffInfo.diffItems.Keys )
            {
                rowCnt++;
                m_oldExcelCache[rowCnt] = new ListViewItem( rowIndex.ToString() );
                m_newExcelCache[rowCnt] = new ListViewItem( rowIndex.ToString() );
                m_oldExcelCache[rowCnt].Name = rowIndex.ToString();
                m_newExcelCache[rowCnt].Name = rowIndex.ToString();
                m_oldExcelCache[rowCnt].UseItemStyleForSubItems = false;
                m_newExcelCache[rowCnt].UseItemStyleForSubItems = false;

                int colCnt = 0;
                foreach ( int colIndex in m_sheetDiffInfo.headers.Keys )
                {
                    colCnt++;
                    if ( !m_sheetDiffInfo.oldSheet.itemPos.ContainsKey( rowIndex ) )
                    {
                        m_oldExcelCache[rowCnt].SubItems.Add( "" );
                        m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        if ( !m_sheetDiffInfo.newSheet.headers.ContainsKey( colIndex ) )
                        {
                            m_newExcelCache[rowCnt].SubItems.Add( "" );
                            m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        }
                        else
                        {
                            m_newExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.newSheet.cells[rowIndex][colIndex].value );
                            m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        }
                        continue;
                    }

                    if ( !m_sheetDiffInfo.newSheet.itemPos.ContainsKey( rowIndex ) )
                    {
                        m_newExcelCache[rowCnt].SubItems.Add( "" );
                        m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        if ( !m_sheetDiffInfo.oldSheet.headers.ContainsKey( colIndex ) )
                        {
                            m_oldExcelCache[rowCnt].SubItems.Add( "" );
                            m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        }
                        else
                        {
                            m_oldExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value );
                            m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        }
                        continue;
                    }

                    if ( !m_sheetDiffInfo.oldSheet.headers.ContainsKey( colIndex ) )
                    {
                        m_oldExcelCache[rowCnt].SubItems.Add( "" );
                        m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        m_newExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.newSheet.cells[rowIndex][colIndex].value );
                        m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        continue;
                    }

                    if ( !m_sheetDiffInfo.newSheet.headers.ContainsKey( colIndex ) )
                    {
                        m_newExcelCache[rowCnt].SubItems.Add( "" );
                        m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        m_oldExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value );
                        m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cNewItem;
                        continue;
                    }

                    if ( m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value != m_sheetDiffInfo.newSheet.cells[rowIndex][colIndex].value )
                    {
                        m_oldExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value );
                        m_oldExcelCache[rowCnt].SubItems[colCnt].BackColor = cOldItem;
                        m_newExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.newSheet.cells[rowIndex][colIndex].value );
                        m_newExcelCache[rowCnt].SubItems[colCnt].BackColor = cOldItem;
                        continue;
                    }
                    m_oldExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.oldSheet.cells[rowIndex][colIndex].value );
                    m_newExcelCache[rowCnt].SubItems.Add( m_sheetDiffInfo.newSheet.cells[rowIndex][colIndex].value );
                }
            }
        }

        private void FunctionSheetDiff_FormClosing( object sender, FormClosingEventArgs e )
        {
            m_main.FuncVersionDiffSheetDiff( m_strExcelName );
        }

        private void lvwOldExcelDiff_RetrieveVirtualItem( object sender, RetrieveVirtualItemEventArgs e )
        {
            e.Item = m_oldExcelCache[e.ItemIndex];
        }

        private void lvwNewExcelDiff_RetrieveVirtualItem( object sender, RetrieveVirtualItemEventArgs e )
        {
            e.Item = m_newExcelCache[e.ItemIndex];
        }

        private void lvwOldExcelDiff_VirtualItemsSelectionRangeChanged( object sender, ListViewVirtualItemsSelectionRangeChangedEventArgs e )
        {
            if ( m_bSelectFromOld )
            {
                m_bSelectFromOld = true;
                m_lstSelectedIndexs.Clear();
                lvwNewExcelDiff.SelectedIndices.Clear();
                foreach ( int index in lvwOldExcelDiff.SelectedIndices )
                {
                    m_lstSelectedIndexs.Add( index );
                    lvwNewExcelDiff.SelectedIndices.Add( index );
                    //lvwNewExcelDiff.EnsureVisible( index );
                }
            }
        }

        private void lvwOldExcelDiff_SelectedIndexChanged( object sender, EventArgs e )
        {
            if ( m_bSelectFromOld )
            {
                m_bSelectFromOld = true;
                m_lstSelectedIndexs.Clear();
                lvwNewExcelDiff.SelectedIndices.Clear();
                foreach ( int index in lvwOldExcelDiff.SelectedIndices )
                {
                    m_lstSelectedIndexs.Add( index );
                    lvwNewExcelDiff.SelectedIndices.Add( index );
                    //lvwNewExcelDiff.EnsureVisible( index );
                }
            }
        }

        private void lvwNewExcelDiff_VirtualItemsSelectionRangeChanged( object sender, ListViewVirtualItemsSelectionRangeChangedEventArgs e )
        {
            if ( !m_bSelectFromOld )
            {
                m_bSelectFromOld = false;
                m_lstSelectedIndexs.Clear();
                lvwOldExcelDiff.SelectedIndices.Clear();
                foreach ( int index in lvwNewExcelDiff.SelectedIndices )
                {
                    m_lstSelectedIndexs.Add( index );
                    lvwOldExcelDiff.SelectedIndices.Add( index );
                    //lvwOldExcelDiff.EnsureVisible( index );
                }
            }
        }

        private void lvwNewExcelDiff_SelectedIndexChanged( object sender, EventArgs e )
        {
            if ( !m_bSelectFromOld )
            {
                m_bSelectFromOld = false;
                m_lstSelectedIndexs.Clear();
                lvwOldExcelDiff.SelectedIndices.Clear();
                foreach ( int index in lvwNewExcelDiff.SelectedIndices )
                {
                    m_lstSelectedIndexs.Add( index );
                    lvwOldExcelDiff.SelectedIndices.Add( index );
                    //lvwOldExcelDiff.EnsureVisible( index );
                }
            }
        }

        private void lvwOldExcelDiff_MouseDown( object sender, MouseEventArgs e )
        {
            m_bSelectFromOld = true;
        }

        private void lvwNewExcelDiff_MouseDown( object sender, MouseEventArgs e )
        {
            m_bSelectFromOld = false;
        }

        private void btnPushNew_Click( object sender, EventArgs e )
        {
            while ( Util.IsFileInUse( m_strNewExcelPath ) )
            {
                MessageBox.Show( "请关闭" + m_strNewExcelPath );
            }
            foreach ( int index in m_lstSelectedIndexs )
            {
                string oldName = m_oldExcelCache[index].SubItems[1].Text;
                string newName = m_newExcelCache[index].SubItems[1].Text;
                if ( oldName == string.Empty )
                {
                    MessageBox.Show( "==>时，左侧选择项不能有空行" );
                    return;
                }
                if ( m_oldExcelCache[index].UseItemStyleForSubItems == true )
                {
                    MessageBox.Show( "不支持重复操作" );
                    return;
                }

            }
            btnPushNew.Enabled = false;
            btnPushOld.Enabled = false;
            //Thread processThread = new Thread( new ParameterizedThreadStart( PushNewProcess ) );
            //processThread.Start();
            PushNewProcess();
        }

        private void btnPushOld_Click( object sender, EventArgs e )
        {
            while ( Util.IsFileInUse( m_strOldExcelPath ) )
            {
                MessageBox.Show( "请关闭" + m_strOldExcelPath );
            }
            foreach ( int index in m_lstSelectedIndexs )
            {
                string oldName = m_oldExcelCache[index].SubItems[1].Text;
                string newName = m_newExcelCache[index].SubItems[1].Text;
                if ( newName == string.Empty )
                {
                    MessageBox.Show( "<==时，右侧选择项不能有空行" );
                    return;
                }
                if ( m_newExcelCache[index].UseItemStyleForSubItems == true )
                {
                    MessageBox.Show( "不支持重复操作" );
                    return;
                }
                m_oldExcelCache[index].UseItemStyleForSubItems = true;
                m_oldExcelCache[index].BackColor = cPushOld;
                m_oldExcelCache[index].ForeColor = Color.White;
                m_newExcelCache[index].UseItemStyleForSubItems = true;
                m_newExcelCache[index].BackColor = cPushOld;
                m_newExcelCache[index].ForeColor = Color.White;
            }
            btnPushNew.Enabled = false;
            btnPushOld.Enabled = false;
            //Thread processThread = new Thread( new ParameterizedThreadStart( PushOldProcess ) );
            //processThread.Start();
            PushOldProcess();
        }

        private void btnSelectOld_Click( object sender, EventArgs e )
        {
            m_lstSelectedIndexs.Clear();
            lvwOldExcelDiff.SelectedIndices.Clear();
            lvwNewExcelDiff.SelectedIndices.Clear();
            m_bSelectFromOld = true;

            int rowCnt = -1;
            foreach ( int rowIndex in m_sheetDiffInfo.diffItems.Keys )
            {
                rowCnt++;
                string oldName = m_oldExcelCache[rowCnt].SubItems[1].Text;
                string newName = m_newExcelCache[rowCnt].SubItems[1].Text;
                if ( oldName == string.Empty )
                {
                    m_lstSelectedIndexs.Add( rowCnt );
                    lvwOldExcelDiff.SelectedIndices.Add( rowCnt );
                    lvwOldExcelDiff.EnsureVisible( rowCnt );
                }
            }
        }

        private void btnSelectNew_Click( object sender, EventArgs e )
        {
            m_lstSelectedIndexs.Clear();
            lvwOldExcelDiff.SelectedIndices.Clear();
            lvwNewExcelDiff.SelectedIndices.Clear();
            m_bSelectFromOld = false;

            int rowCnt = -1;
            foreach ( int rowIndex in m_sheetDiffInfo.diffItems.Keys )
            {
                rowCnt++;
                string oldName = m_oldExcelCache[rowCnt].SubItems[1].Text;
                string newName = m_newExcelCache[rowCnt].SubItems[1].Text;
                if ( newName == string.Empty )
                {
                    m_lstSelectedIndexs.Add( rowCnt );
                    lvwNewExcelDiff.SelectedIndices.Add( rowCnt );
                    lvwNewExcelDiff.EnsureVisible( rowCnt );
                }
            }
        }

        private void lvwOldExcelDiff_TopItemChanged( object sender, EventArgs e )
        {
            lvwNewExcelDiff.EnsureVisible( lvwOldExcelDiff.TopItem.Index );
            lvwNewExcelDiff.TopItem = m_newExcelCache[lvwOldExcelDiff.TopItem.Index];
            
        }

        private void lvwNewExcelDiff_TopItemChanged( object sender, EventArgs e )
        {
            lvwOldExcelDiff.EnsureVisible( lvwNewExcelDiff.TopItem.Index );
            lvwOldExcelDiff.TopItem = m_oldExcelCache[lvwNewExcelDiff.TopItem.Index];
        }

        private void lvwOldExcelDiff_ColumnClick( object sender, ColumnClickEventArgs e )
        {
            lvwNewExcelDiff.EnsureVisible( lvwNewExcelDiff.TopItem, e.Column );
        }

        private void lvwNewExcelDiff_ColumnClick( object sender, ColumnClickEventArgs e )
        {
            lvwOldExcelDiff.EnsureVisible( lvwOldExcelDiff.TopItem, e.Column );
        }

        private void lvwOldExcelDiff_ColumnWidthChanging( object sender, ColumnWidthChangingEventArgs e )
        {
            lvwNewExcelDiff.Columns[e.ColumnIndex].Width = e.NewWidth;
        }

        private void lvwNewExcelDiff_ColumnWidthChanging( object sender, ColumnWidthChangingEventArgs e )
        {
            lvwOldExcelDiff.Columns[e.ColumnIndex].Width = e.NewWidth;
        }   
    }
}
