using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Drawing;

namespace ExcelTools
{
    class SyncListView : ListView
    {
        public SyncListView()
            : base()
        {
        }

        private List<SyncListView> linkedListViews = new List<SyncListView>();
        public event EventHandler TopItemChanged;

        protected virtual void OnTopItemChanged( EventArgs e )
        {
            var handler = TopItemChanged;
            if ( handler != null ) handler( this, e );
        }

        /// <summary>
        /// Links the specified list view to this list view.  Whenever either list view
        /// scrolls, the other will scroll too.
        /// </summary>
        /// <param name="listView">The ListView to link.</param>
        public void AddLinkedListView( SyncListView listView )
        {
            if ( listView == this )
                throw new ArgumentException( "Cannot link a ListView to itself!", "listView" );

            if ( !linkedListViews.Contains( listView ) )
            {
                // add the list view to our list of linked list views
                linkedListViews.Add( listView );
                // add this to the list view's list of linked list views
                listView.AddLinkedListView( this );

                // make sure the ListView is linked to all of the other ListViews that this ListView is linked to
                for ( int i = 0; i < linkedListViews.Count; i++ )
                {
                    // get the linked list view
                    var linkedListView = linkedListViews[i];
                    // link the list views together
                    if ( linkedListView != listView )
                        linkedListView.AddLinkedListView( listView );
                }
            }
        }

        /// <summary>
        /// Ensure visible of a ListViewItem and SubItem Index.
        /// </summary>
        /// <param name="item">item to visible</param> 
        /// <param name="subItemIndex">idex of subItem to visible</param>
        public void EnsureVisible( ListViewItem item, int subItemIndex )
        {
            if ( item == null || subItemIndex > item.SubItems.Count - 1 )
            {
                throw new ArgumentException();
            }

            // scroll to the item row.
            item.EnsureVisible();
            Rectangle bounds = item.SubItems[subItemIndex].Bounds;

            // need to set width from columnheader, first subitem includes
            // all subitems.
            bounds.Width = this.Columns[subItemIndex].Width;

            ScrollToRectangle( bounds );
        }

        /// <summary>
        /// Scrolls the listview.
        /// </summary>
        /// <param name="bounds"></param>
        private void ScrollToRectangle( Rectangle bounds )
        {
            int scrollToLeft = bounds.X + bounds.Width + 20;
            if ( scrollToLeft > this.Bounds.Width )
            {
                this.ScrollHorizontal( scrollToLeft - this.Bounds.Width );
            }
            else
            {
                int scrollToRight = bounds.X - 20;
                if ( scrollToRight < 0 )
                {
                    this.ScrollHorizontal( scrollToRight );
                }
            }
        }

        private void ScrollHorizontal( int pixelsToScroll )
        {
            User32.SendMessage( this.Handle, User32.LVM_SCROLL, (IntPtr)pixelsToScroll, IntPtr.Zero );
        }

        /// <summary>
        /// Sets the destination's scroll positions to that of the source.
        /// </summary>
        /// <param name="source">The source of the scroll positions.</param>
        /// <param name="dest">The destinations to set the scroll positions for.</param>
        private void SetScrollPositions( SyncListView source, SyncListView dest )
        {
            // get the scroll positions of the source
            int horizontal = User32.GetScrollPos( source.Handle, Orientation.Horizontal );
            int vertical = User32.GetScrollPos( source.Handle, Orientation.Vertical );
            // set the scroll positions of the destination
            User32.SetScrollPos( dest.Handle, Orientation.Horizontal, horizontal, false );
            User32.SetScrollPos( dest.Handle, Orientation.Vertical, vertical, false );

            // convert the position to the windows message equivalent
            //IntPtr msgHPosition = new IntPtr((horizontal << 16) + 4);
            //IntPtr msgVPosition = new IntPtr((vertical << 16) + 4);
            //User32.SendMessage( dest.Handle, User32.WM_HSCROLL, msgHPosition, IntPtr.Zero );
            //User32.SendMessage( dest.Handle, User32.WM_VSCROLL, msgVPosition, IntPtr.Zero );
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="m"></param>
        protected override void WndProc( ref Message m )
        {
            // Trap LVN_ENDSCROLL, delivered with a WM_REFLECT + WM_NOTIFY message
            if ( m.Msg == User32.LVN_ENDSCROLL )
            {
                var notify = (NMHDR)Marshal.PtrToStructure( m.LParam, typeof( NMHDR ) );
                if ( notify.code == -181 )
                {
                    if ( !this.TopItem.Equals( lastTopItem ) )
                    {
                        OnTopItemChanged( EventArgs.Empty );
                        lastTopItem = this.TopItem;
                    }

                    //if ( horizonScrollPixel != User32.GetScrollPos( this.Handle, Orientation.Horizontal ) )
                    //{
                    //    horizonScrollPixel = User32.GetScrollPos( this.Handle, Orientation.Horizontal );
                    //    //MessageBox.Show( horizonScrollPixel.ToString() );
                    //    foreach ( var linkedListView in linkedListViews )
                    //    {

                    //    }
                    //}
                }
            }

            // process the message
            base.WndProc( ref m );

            //// pass scroll messages onto any linked views
            //if ( m.Msg == User32.WM_VSCROLL || m.Msg == User32.WM_HSCROLL || m.Msg == User32.WM_MOUSEWHEEL )
            //{
            //    foreach ( var linkedListView in linkedListViews )
            //    {
            //        // set the scroll positions of the linked list view
            //        SetScrollPositions( this, linkedListView );

            //        // copy the windows message
            //        Message copy = new Message
            //        {
            //            HWnd = linkedListView.Handle,
            //            LParam = m.LParam,
            //            Msg = m.Msg,
            //            Result = m.Result,
            //            WParam = m.WParam
            //        };
            //        // pass the message onto the linked list view
            //        linkedListView.RecieveWndProc( ref copy );
            //    }
            //}
        }

        /// <summary>
        /// Receives a WndProc message without passing it onto any linked list views.  This is useful to avoid infinite loops.
        /// </summary>
        /// <param name="m">The windows message.</param>
        private void RecieveWndProc( ref Message m )
        {
            base.WndProc( ref m );
        }

        private ListViewItem lastTopItem = null;

        private struct NMHDR
        {
            public IntPtr hwndFrom;
            public IntPtr idFrom;
            public int code;
        }

        /// <summary>
        /// Imported functions from the User32.dll
        /// </summary>
        private class User32
        {
            // scroll horizon
            public const int WM_HSCROLL = 0x114;
            // scroll vertical
            public const int WM_VSCROLL = 0x115;
            // mouse wheel
            public const int WM_MOUSEWHEEL = 0x020A;
            // end scroll
            public const int LVN_ENDSCROLL = 0x204e;
            // scroll first
            public const Int32 LVM_FIRST = 0x1000;
            // scroll step
            public const Int32 LVM_SCROLL = LVM_FIRST + 20;

            [DllImport( "user32.dll", CharSet = CharSet.Auto )]
            public static extern int GetScrollPos( IntPtr hWnd, Orientation nBar );

            [DllImport( "user32.dll" )]
            public static extern int SetScrollPos( IntPtr hWnd, Orientation nBar, int nPos, bool bRedraw );

            [DllImport( "user32.dll", EntryPoint = "SendMessage" )]
            public static extern int SendMessage( IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam );
        }
    }
}
