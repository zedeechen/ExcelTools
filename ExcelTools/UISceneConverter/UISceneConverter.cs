using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelTools
{
    class UISceneConverter
    {
        public UISceneConverter( string jsonDirPath, string outDirPath, string configPath )
        {
            if ( !Directory.Exists( jsonDirPath ) )
            {
                Console.WriteLine( "json导入目录不存在" );
                return;
            }
            if ( jsonDirPath != outDirPath )
                Util.DirectoryCopy( jsonDirPath, outDirPath, true );
            //else
            //{
            //    MessageBox.Show( "输入输出目录相同" );
            //}
            string[] fileNames = Directory.GetFiles( outDirPath, "*.json", SearchOption.AllDirectories );

            JsonControl json = new JsonControl();
            UISceneTextCsvControl csv = new UISceneTextCsvControl();
            UISceneTextCsv sheet = new UISceneTextCsv();
            Dictionary< string, int > words = new Dictionary<string, int>();
            Dictionary< int, bool > used = new Dictionary<int, bool>();
            int cnt = 0;

            if ( File.Exists( configPath ) )
            {
                if ( !csv.Read( configPath, out sheet ) )
                {
                    Console.WriteLine( configPath + "读取失败" );
                    return;
                }
            }
            foreach ( var item in sheet.dictText )
            {
                words[item.Value] = item.Key;
                cnt = cnt > item.Key ? cnt : item.Key;
            }
            
            for ( int i = 0; i < fileNames.Length; i++ )
            {
                string path = fileNames[i];

                json.Convert( path, ref words, ref used, ref cnt );
            }
            foreach ( var item in sheet.dictText )
            {
                if ( used.ContainsKey( item.Key ) )
                {
                    words[item.Value] = item.Key;
                }
            }
            sheet.dictText.Clear();
            foreach ( var item in words )
            {
                if ( used.ContainsKey( item.Value ) )
                    sheet.dictText[item.Value] = item.Key;
            }
            if ( !csv.Create( configPath, sheet ) )
            {
                Console.WriteLine( "config write error" );
                return;
            }
            Console.WriteLine( "Done!" );
        }
    }
}
