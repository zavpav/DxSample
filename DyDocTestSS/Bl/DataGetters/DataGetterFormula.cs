using System;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl.DataGetters
{
    public class DataGetterFormula : IDataGetter
    {
        public DataGetterFormula([NotNull] IDataCalculatorInternal dataCalculator)
        {
            this.DataCalculator = dataCalculator;
        }

        [NotNull]
        private IDataCalculatorInternal DataCalculator { get; set; }

        public Type ColumnDataBindInfoGetterType { get { return typeof(ColumnDataBindInfoGetterSomeFormula); } }

        public decimal GetValueDecimal(string pbsCode, ColumnDataBindInfoGetter bindGetterInfo)
        {
            var formulaInfo = bindGetterInfo as ColumnDataBindInfoGetterSomeFormula;
            if (formulaInfo == null)
                throw new NotSupportedException("Плохой биндинг для ColumnDataBindInfoGetter " + bindGetterInfo.GetType());
            if (formulaInfo.FinalFunction == null)
                throw new NotSupportedException("Пустая формула на этапе получения данных");

            return formulaInfo.FinalFunction(this.DataCalculator, pbsCode);
        }
    }
}