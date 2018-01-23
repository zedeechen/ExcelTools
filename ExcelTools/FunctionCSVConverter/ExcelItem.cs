using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    class ExcelItem
    {
        public int id;
        public string path;
        public ExcelItem( int idx, string pathx )
        {
            this.id = idx;
            this.path = pathx;
        }
    };
}
