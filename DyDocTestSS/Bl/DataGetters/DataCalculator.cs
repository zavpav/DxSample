using System;
using System.Collections.Generic;
using System.Linq;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl.DataGetters
{
    /// <summary> Собственно, вычислятель данных колонки для ПБС </summary>
    public interface IDataCalculator
    {
        /// <summary> Вычислить значение для ПБС </summary>
        decimal GetDecimal([NotNull] string pbsCode, [NotNull] ColumnDataBind columnDataBind);
        
    }

    public interface IDataCalculatorInternal
    {
        decimal GetDecimal([NotNull] string pbsCode, [NotNull] ColumnDataBindInfoGetter columnDataGetter);
    }

    public class DataCalculator : IDataCalculator, IDataCalculatorInternal
    {
        /// <summary> Список геттеров для данных </summary>
        private List<IDataGetter> Getters { get; set; }

        public DataCalculator()
        {
            //TODO должны быть в конструкторе через DI
            this.Getters = new List<IDataGetter>
            {
                new DataGetterAksiok(),
                new DataGetterBr(),
                new DataGetterFormula(this)
            };

        }

        /// <summary> Получить числовое значение по описанию биндинга </summary>
        public decimal GetDecimal(string pbsCode, ColumnDataBind columnDataBind)
        {
            if (columnDataBind.Columns.Count() != 1)
                throw new NotSupportedException("Ошибка получения значения для множественных колонок");

            var singleCol = columnDataBind.Columns.Single();

            var getter = this.Getters.SingleOrDefault(x => x.ColumnDataBindInfoGetterType == singleCol.GetType());
            if (getter == null)
                throw new NotSupportedException("Не нашли getter " + singleCol.GetType());

            return this.GetDecimal(pbsCode, singleCol);
        }

        /// <summary> Получить значение по описанию геттера </summary>
        public decimal GetDecimal(string pbsCode, ColumnDataBindInfoGetter columnDataGetter)
        {
            var getter = this.Getters.SingleOrDefault(x => x.ColumnDataBindInfoGetterType == columnDataGetter.GetType());
            if (getter == null)
                throw new NotSupportedException("Не нашли getter " + columnDataGetter.GetType());

            return getter.GetValueDecimal(pbsCode, columnDataGetter);
        }
    }
}