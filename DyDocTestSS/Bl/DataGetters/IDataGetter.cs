using System;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl.DataGetters
{
    public interface IDataGetter
    {
        /// <summary> Тип вида парамтеров для получения данных </summary>
        Type ColumnDataBindInfoGetterType { get; }

        /// <summary> Собственно, получение данных для bindInfo </summary>
        decimal GetValueDecimal([NotNull] string pbsCode, [NotNull] ColumnDataBindInfoGetter bindGetterInfo);
    }
}