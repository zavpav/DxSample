using System;

namespace DyDocTestSS.Bl.DataGetters
{
    public class DataGetterBr : IDataGetter
    {
        public Type ColumnDataBindInfoGetterType { get{ return typeof(ColumnDataBindInfoGetterBr); } }
        
        public decimal GetValueDecimal(string pbsCode, ColumnDataBindInfoGetter bindGetterInfo)
        {
            if (string.IsNullOrWhiteSpace(pbsCode))
            {
                return 0;
            }
            return int.Parse(pbsCode) * 10000;
        }
    }
}