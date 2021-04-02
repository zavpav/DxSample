using System;

namespace DyDocTestSS.Domain
{
    /// <summary> Типы областей шаблонов </summary>
    [Flags]
    public enum EnumDyDocSsRanges
    {
        /// <summary> Системная информация по листу </summary>
        System = 1 << 1,

        /// <summary> Область красивых заголовков </summary>
        Header = 1 << 2,

        /// <summary> Область красивого подвала </summary>
        Footer = 1 << 3,

        /// <summary> Область коэффициентов </summary>
        Coeff = 1 << 4,

        /// <summary> Собственно, область данных ПБС </summary>
        Data = 1 << 5,

        /// <summary> Область наименований ПБС (входит в Data) </summary>
        DataPbsName = 1 << 6,

        /// <summary> Область кодов ПБС (входит в Data) </summary>
        DataPbsCode = 1 << 7,

        /// <summary> Область итоговых данных (сумма по годам) (входит в Data) </summary>
        TotalYearSum = 1 << 8,

        /// <summary> Строка основных формул по колонке (входит в Data) </summary>
        MainFormula = 1 << 11,

        /// <summary> Область "к распределению" (входит в Data) </summary>
        ForDistrib = 1 << 12,

        /// <summary> Область данных "нераспределенный остаток" (входит в Data) </summary>
        Residue = 1 << 13,

        /// <summary> Область привязки строк </summary>
        RowBind = 1 << 14,

        /// <summary> Область системных наименований колонок (для восстановления формул) </summary>
        SysColumnNames = 1 << 15,

        /// <summary> Область "итого" </summary>
        TotalSum = 1 << 16,

        /// <summary> Область "первой колонки". Т.е. куда мы можем добавить первую колонку. </summary>
        FirstInsertColumn = 1 << 17,

        /// <summary> Область заголовков колонок </summary>
        ColumnHeaders = 1 << 18
    }
}