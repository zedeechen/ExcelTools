using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelTools
{
    class Util
    {
        public static bool IsFileInUse( string fileName, out String strErrorMsg )
        {
            bool inUse = true;
            strErrorMsg = "";
            if ( File.Exists( fileName ) )
            {
                FileStream fs = null;
                try
                {
                    fs = new FileStream( fileName, FileMode.Open, FileAccess.Read, FileShare.None );
                    inUse = false;
                }
                catch ( Exception e )
                {
                    strErrorMsg = e.Message.ToString();
                }
                finally
                {
                    if ( fs != null )
                    {
                        fs.Close();
                    }
                }
                return inUse;           //true表示正在使用,false没有使用
            }
            else
            {
                return false;           //文件不存在则一定没有被使用
            }
        }
        public static bool IsFileInUse( string fileName )
        {
            string strErrorMsg = "";
            if ( IsFileInUse( fileName, out strErrorMsg ) )
                return true;           //true表示正在使用,false没有使用
            else
                return false;           //文件不存在则一定没有被使用
        }
        public static bool IsExcelOpened( string fileName )
        {
            string name = fileName.Substring( fileName.LastIndexOf( '\\' ) + 1 );
            if ( Regex.IsMatch( name, "^~.*" ) )
            {
                return false;
            }
            return true;
        }

        public static string GetOpenedExcelList( string[] fileNames )
        {
            int fLength = fileNames.Length;
            String ret = "";
            for ( int i = 0; i < fLength; i++ )
            {
                string name = fileNames[i].Substring( fileNames[i].LastIndexOf( '\\' ) + 1 );
                string dir = fileNames[i].Substring( 0, fileNames[i].LastIndexOf( '\\' ) );

                if ( Regex.IsMatch( name, "^~.*" ) )
                {
                    ret += dir + "\\" + name.Substring( name.IndexOf("$") + 1 )  + "\n";
                }
            }
            return ret;
        }
        public static void Swap<T>( ref T my, ref T other )
        {
            T temp = my;
            my = other;
            other = temp;
        }
        public static int TenPow( int cnt )
        {
            if ( cnt > 9 ) return -1;
            int ret = 1;
            for ( int i = 0; i < cnt; i++ )
                ret *= 10;
            return ret;
        }
        public static void DirectoryCopy( string sourceDirName, string destDirName, bool copySubDirs )
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo( sourceDirName );
            DirectoryInfo[] dirs = dir.GetDirectories();

            if ( !dir.Exists )
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName );
            }

            // If the destination directory doesn't exist, create it. 
            if ( !Directory.Exists( destDirName ) )
            {
                Directory.CreateDirectory( destDirName );
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach ( FileInfo file in files )
            {
                string temppath = Path.Combine( destDirName, file.Name );
                file.CopyTo( temppath, true );
            }

            // If copying subdirectories, copy them and their contents to new location. 
            if ( copySubDirs )
            {
                foreach ( DirectoryInfo subdir in dirs )
                {
                    string temppath = Path.Combine( destDirName, subdir.Name );
                    DirectoryCopy( subdir.FullName, temppath, copySubDirs );
                }
            }
        }

        public static void ConsoleRun( string strWorkingDirectory, string command, out string output )
        {
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo.WorkingDirectory = strWorkingDirectory; //运行初始目录
            p.StartInfo.FileName = "cmd.exe";
            p.StartInfo.UseShellExecute = false;    //是否使用操作系统shell启动
            p.StartInfo.RedirectStandardInput = true;//接受来自调用程序的输入信息
            p.StartInfo.RedirectStandardOutput = true;//由调用程序获取输出信息
            p.StartInfo.RedirectStandardError = true;//重定向标准错误输出
            p.StartInfo.CreateNoWindow = true;//不显示程序窗口
            p.Start();//启动程序

            //向cmd窗口发送输入信息
            p.StandardInput.WriteLine( command + "&exit" );
            p.StandardInput.AutoFlush = true;

            //p.StandardInput.WriteLine("exit");
            //向标准输入写入要执行的命令。这里使用&是批处理命令的符号，表示前面一个命令不管是否执行成功都执行后面(exit)命令，如果不执行exit命令，后面调用ReadToEnd()方法会假死
            //同类的符号还有&&和||前者表示必须前一个命令执行成功才会执行后面的命令，后者表示必须前一个命令执行失败才会执行后面的命令

            //获取cmd窗口的输出信息
            output = p.StandardOutput.ReadToEnd();

            //StreamReader reader = p.StandardOutput;
            //string line=reader.ReadLine();
            //while (!reader.EndOfStream)
            //{
            //    str += line + "  ";
            //    line = reader.ReadLine();
            //}

            p.WaitForExit();//等待程序执行完退出进程
            p.Close();
        }

        /**
         * 将驼峰式命名的字符串转换为下划线大写方式。如果转换前的驼峰式命名的字符串为空，则返回空字符串。
         * 例如：HelloWorld->HELLO_WORLD
         * @param name 转换前的驼峰式命名的字符串
         * @return 转换后下划线大写方式命名的字符串
         */
        public static String UnderScoreName( String name )
        {
            StringBuilder result = new StringBuilder();
            if ( name != null && name.Length > 0 )
            {
                result.Append( name.Substring( 0, 1 ).ToUpper() );
                for ( int i = 1; i < name.Length; i++ )
                {
                    String s = name.Substring( i, i + 1 );
                    if ( s.Equals( s.ToUpper() ) && !Char.IsDigit( s[0] ) )
                    {
                        result.Append( "_" );
                    }
                    result.Append( s.ToUpper() );
                }
            }
            return result.ToString();
        }

        /**
         * 将下划线大写方式命名的字符串转换为驼峰式。如果转换前的下划线大写方式命名的字符串为空，则返回空字符串。
         * 例如：HELLO_WORLD->HelloWorld
         * @param name 转换前的下划线大写方式命名的字符串
         * @return 转换后的驼峰式命名的字符串
         */
        public static String CamelName( String name )
        {
            StringBuilder result = new StringBuilder();
            if ( name == null || name == "" )
            {
                return "";
            }
            else if ( !name.Contains( "_" ) )
            {
                return name.Substring( 0, 1 ).ToUpper() + name.Substring( 1 );
            }
            String[] spliter = { "_" };
            String[] camels = name.Split( spliter, StringSplitOptions.None );
            foreach ( String camel in camels )
            {
                if ( camel == "" )
                {
                    continue;
                }
                result.Append( camel.Substring( 0, 1 ).ToUpper() );
                result.Append( camel.Substring( 1 ) );
            }
            return result.ToString();
        }
    }
}
