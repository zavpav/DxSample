using System;
using System.Collections.Generic;
using System.Linq;
using DevExpress.Spreadsheet;
using JetBrains.Annotations;

namespace Sphaera.Bp.Bl.Excel
{
    /// <summary> Значение ячейки отчёта </summary>
    [PublicAPI]
    public struct CellData
    {
        /// <summary> Спец значения ячеек </summary>
        public enum EnumSpecCellData
        {
            /// <summary> Пропускать ячейку при обработке </summary>
            Skip
        }

        /// <summary>
        /// "Видимое" ячейки
        /// </summary>
        [NotNull]
        public readonly object Val;

        /// <summary> Описание сумм в ячейке </summary>
        [CanBeNull]
        public readonly List<ICellDataSumInfo> SumInfos;

        /// <summary> Форматирование ячейки </summary>
        [CanBeNull]
        public readonly Action<Cell> Format;


        /// <summary> Комментарий, который записывается в ячейку </summary>
        [CanBeNull]
        public readonly string CellComment;

        /// <summary> Создание Спецописания </summary>
        [PublicAPI]
        public CellData(EnumSpecCellData specCode)
        {
            this.Val = specCode;
            this.SumInfos = null;
            this.Format = null;
            this.CellComment = null;
        }

        /// <summary> Создание описания значения ячейки из строки </summary>
        [PublicAPI]
        public CellData([NotNull] string strVal, [CanBeNull] IEnumerable<ICellDataSumInfo> sumInfos = null, [CanBeNull] Action<Cell> addFormat = null, [CanBeNull] string cellComment = null)
        {
            this.Val = strVal;
            this.SumInfos = sumInfos != null ? sumInfos.ToList() : null;
            this.Format = addFormat;
            this.CellComment = cellComment;
        }

        ///// <summary> Создание описания значения ячейки по описаниям сумм </summary>
        //private CellData([NotNull] List<CellSumInfo> sumInfos, [CanBeNull] Action<Cell> addFormat = null, [CanBeNull] string cellComment = null)
        //{
        //    this.SumInfos = sumInfos;
        //    this.Val = sumInfos.Sum(x => x.Sm);
        //    this.Format = addFormat;
        //    this.CellComment = cellComment;
        //}

        /// <summary> Создание описания значения ячейки из числа с отдельными описателями сумм </summary>
        public CellData(decimal smChange, [CanBeNull] IEnumerable<ICellDataSumInfo> sumInfos = null, [CanBeNull] Action<Cell> addFormat = null, [CanBeNull] string cellComment = null)
        {
            this.Val = smChange;
            this.SumInfos = sumInfos != null ? sumInfos.ToList() : null;
            this.Format = addFormat;
            this.CellComment = cellComment;
        }

        /// <summary> Создание CellData из строки </summary>
        public static implicit operator CellData(EnumSpecCellData specCode)
        {
            return new CellData(specCode);
        }

        /// <summary> Создание CellData из строки </summary>
        public static implicit operator CellData([NotNull] string strVal)
        {
            return new CellData(strVal);
        }

        /// <summary> Создание CellData из строки </summary>
        public static implicit operator CellData(decimal decimalVal)
        {
            return new CellData(decimalVal);
        }

        ///// <summary> Создание CellData из описаний сумм </summary>
        //public static implicit operator CellData([NotNull] List<CellSumInfo> sumInfos)
        //{
        //    return new CellData(sumInfos);
        //}
    }
}