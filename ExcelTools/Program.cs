using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;

namespace ExcelTools
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [DllImport( "kernel32.dll" )]
        static extern IntPtr GetConsoleWindow();

        [DllImport( "user32.dll" )]
        static extern bool ShowWindow( IntPtr hWnd, int nCmdShow );

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        [STAThread]
        static void Main()
        {
            var handle = GetConsoleWindow();
            // Hide
            ShowWindow( handle, SW_HIDE );
            string[] CLA = Environment.GetCommandLineArgs();
            if ( CLA.Length < 2 )
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault( false );
                Application.Run( new Main() );
            }
            else if ( CLA.Length == 4 )
            {
                // Show
                ShowWindow( handle, SW_SHOW );
                new UISceneConverter( CLA[1], CLA[2], CLA[3] );
            }
            else
            {// Show
                ShowWindow( handle, SW_SHOW );
                Console.WriteLine( "args error!" );
            }
            //new UISceneConverter( @"E:\a.ui", @"E:\a.ui", @"E:\ui_text.csv" );
        }
    }
}
