using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    public class Node
    {
        public int x;
        public int y;
        public string value;
        public int status; // 1表示修改，2表示新增
        public Node( int posx, int posy, string val )
        {
            this.x = posx;
            this.y = posy;
            this.value = val;
        }
        public Node( int sta, string val )
        {
            this.value = val;
            this.status = sta;
        }
    }
}
