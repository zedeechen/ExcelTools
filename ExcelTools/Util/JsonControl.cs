using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelTools
{
    class JsonControl
    {
        private void Replace( ref string res, ref Dictionary<string, int> words, ref Dictionary<int, bool> used, ref int cnt, ref int id )
        {
            if ( words.ContainsKey( res ) )
            {
                id = words[res];
                used[id] = true;
                res = "#" + id.ToString();
            }
            else
            {
                cnt++;
                used[cnt] = true;
                words.Add( res, cnt );
                res = "#" + cnt.ToString();
            }
        }

        private bool ConvertInside( string res, out string inside, ref Dictionary<string, int> words, ref Dictionary<int, bool> used, ref int cnt )
        {
            StringBuilder result = new StringBuilder();

            for ( int i = 0; i < res.Length; )
            {
                if ( i < res.Length - 1 && ( res[i] == '\\' && res[i + 1] == '"' ) )
                {
                    i += 2;
                    StringBuilder pattern = new StringBuilder();
                    while ( i < res.Length - 1 && ( res[i] != '\\' || res[i + 1] != '"' ) )
                    {
                        pattern.Append( res[i] );
                        i++;
                    }
                    i += 2;
                    string tmp = pattern.ToString();
                    Regex rx = new Regex( ".*[\u4e00-\u9fa5]+.*" );//中文字符unicode范围  
                    int id = 0;
                    if ( rx.IsMatch( tmp ) )
                    {
                        if ( tmp != string.Empty && tmp != "微软雅黑" )
                        {
                            Replace( ref tmp, ref words, ref used, ref cnt, ref id );
                        }
                    }
                    else if ( tmp.Length > 1 && tmp[0] == '#' && int.TryParse( tmp.Substring( 1 ), out id ) )
                    {
                        used[id] = true;
                    }
                    tmp = "\\\"" + tmp + "\\\"";
                    result.Append( tmp );
                }
                else
                {
                    result.Append( res[i++] );
                }
            }

            inside = result.ToString();
            return true;
        }

        public bool Convert( string path, ref Dictionary<string, int> words, ref Dictionary<int, bool> used, ref int cnt )
        {
            string file = File.ReadAllText( path );
            StringBuilder result = new StringBuilder();
            Regex rx = new Regex( ".*[\u4e00-\u9fa5]+.*" );//中文字符unicode范围  
            for ( int i = 0; i < file.Length; )
            {
                if ( file[i] == '"' )
                {
                    ++i;
                    StringBuilder pattern = new StringBuilder();
                    while ( i < file.Length && file[i] != '"' )
                    {
                        pattern.Append( file[i] );
                        i++;
                        if ( file[i - 1] == '\\' )
                        {
                            pattern.Append( file[i] );
                            i++;
                        }
                    }
                    ++i;
                    int id = 0;
                    string res = pattern.ToString();
                    if ( rx.IsMatch( res ) )
                    {
                        if ( res != string.Empty && res != "微软雅黑" )
                        {
                            Regex rx2 = new Regex( "\".*[\u4e00-\u9fa5]+.*\"" );
                            if ( rx2.IsMatch( res ) )
                            {
                                string inside;
                                ConvertInside( res, out inside, ref words, ref used, ref cnt );
                                res = inside;
                            }
                            else
                            {
                                Replace( ref res, ref words, ref used, ref cnt, ref id );
                            }
                        }
                    }
                    else if ( res.Length > 1 && res[0] == '#' && int.TryParse( res.Substring( 1 ), out id ) )
                    {
                        used[id] = true;
                    }
                    res = "\"" + res + "\"";
                    result.Append( res );
                }
                else
                {
                    result.Append( file[i++] );
                }
            }

            byte[] content = new UTF8Encoding( true ).GetBytes( result.ToString() );
            FileStream fs = File.Create( path );
            fs.Write( content, 0, content.Length );
            fs.Flush();
            fs.Close();
            return true;
        }
    }
}
