using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ExcelTools
{
    class UISceneTextCsvControl
    {
        public bool Create( string strPath, UISceneTextCsv textSheet )
        {
            StringBuilder builder = new StringBuilder();

            builder.Append( "101,102\r\n" );

            foreach ( var item in textSheet.dictText )
            {
                AppendCellForCSV( builder, item.Key.ToString() );
                builder.Append( "," );
                AppendCellForCSV( builder, item.Value );
                builder.Append( "\r\n" );
            }

            byte[] content = new UTF8Encoding( true ).GetBytes( builder.ToString() );

            string dirPath = Path.GetDirectoryName( strPath );
            if ( !Directory.Exists( dirPath ) )
                Directory.CreateDirectory( dirPath );
            FileStream fs = File.Create( strPath );
            fs.Write( content, 0, content.Length );
            fs.Flush();
            fs.Close();
            return true;
        }

        public bool Read( string strPath, out UISceneTextCsv textSheet )
        {
            textSheet = new UISceneTextCsv();

            var reader = new StreamReader( File.OpenRead( strPath ) );

            int cnt = 0;
            while ( !reader.EndOfStream )
            {
                cnt++;
                string row = reader.ReadLine();
                // 跳过第一行
                if ( cnt == 1 ) continue;
                int pos = 0;
                while ( row[pos] != ',' && pos < row.Length )
                    pos++;
                if ( pos >= row.Length )
                    return false;
                string col1 = row.Substring( 0, pos );
                string col2 =  row.Substring( pos + 1 );
                int id;
                if ( !int.TryParse( col1, out id ) )
                {
                    Console.WriteLine( "第" + cnt + "行: [" + row + "] Error: 第一列不是数字" );
                    return false;
                }

                StringBuilder name = new StringBuilder();
                if ( col2.Length != 0 )
                {
                    int st = 0, ed = col2.Length - 1;
                    if ( col2[ed] == ',' ) ed--;
                    if ( col2[st] == '"' && col2[ed] == '"' )
                    {
                        st++; ed--;
                    }
                   
                    bool preQuote = false;
                    for ( int i = st; i <= ed; i++ )
                    {
                        if ( col2[i] == '"' )
                        {
                            if ( st == 0 || preQuote )
                            {
                                name.Append( col2[i] );
                                preQuote = false;
                            }
                            else
                            {
                                preQuote = true;
                            }
                        }
                        else
                        {
                            name.Append( col2[i] );
                            preQuote = false;
                        }
                    }
                }
                else
                {
                    Console.WriteLine( "第" + cnt + "行: [" + row + "] Error: 第二列不可为空" );
                    return false;
                }
                
                if ( !textSheet.dictText.ContainsKey( id ) )
                    textSheet.dictText.Add( id, name.ToString() );
            }
            reader.Close();
            reader.Dispose();
            return true;
        }

        private void AppendCellForCSV( StringBuilder builder, String strCell )
        {
            bool bWithQuota = ( strCell.IndexOf( ',' ) != -1 || strCell.IndexOf( '\"' ) != -1 );

            if ( bWithQuota )
            {
                builder.Append( "\"" );
            }

            foreach ( char c in strCell )
            {
                if ( c == '\"' )
                {
                    builder.Append( "\"" );
                }
                builder.Append( c );
            }

            if ( bWithQuota )
            {
                builder.Append( "\"" );
            }
        }

    }
}
