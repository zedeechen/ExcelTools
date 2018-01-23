using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTools
{
    public class SheetDiffInfo
    {
        public FunctionSheet oldSheet;
        public FunctionSheet newSheet;
        public Dictionary<int, FunctionSheet.TitleConfig> headers;
        public Dictionary<int, int> items;
        public SortedDictionary<int, int> diffItems;

        public SheetDiffInfo()
        {
            oldSheet = new FunctionSheet();
            newSheet = new FunctionSheet();
            headers = new Dictionary<int, FunctionSheet.TitleConfig>();
            items = new Dictionary<int, int>();
            diffItems = new SortedDictionary<int, int>();
        }
    }
}
