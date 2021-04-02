using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DyDocTestSS.DyTemplates;
using JetBrains.Annotations;
using Sphaera.Bp.Bl.Excel;
using Sphaera.Bp.Services.Log;

namespace DyDocTestSS.Domain
{
    /// <summary> Динамические листы на spreadsheet </summary>
    public class DyDocSs
    {
        /// <summary> Основной рабочий документ </summary>
        [NotNull]
        public IWorkbook Wb { get; private set; }

        [NotNull]
        public List<DyDocSsSheetInfo> SystemSheetInfos { get; private set; }

        /// <summary> Дубликат документа на момент загрузки документа </summary>
        /// <remarks>Является шаблоном для восстановления ширины колонок и высоты скрываемых строк</remarks>
        [CanBeNull]
        public IWorkbook WbTemplateInfo { get; set; }

        public DyDocSs([NotNull] IWorkbook wb)
        {
            this.Wb = wb;
            this.SystemSheetInfos = new List<DyDocSsSheetInfo>(5);
        }

        /// <summary> Информация по областям листа </summary>
        public class DyDocSsSheetInfo
        {
            public DyDocSsSheetInfo([NotNull] Worksheet sheet)
            {
                this.Sheet = sheet;
                
                var sysN = sheet.Workbook
                               .DefinedNames
                               .Where(x => x.Name.StartsWith("System_"))
                               .SingleOrDefault(x => x.Range.Worksheet == sheet)
                           ??  throw new NotSupportedException("В листе не объявлена область System");
                this.SystemRange = sysN;
                
                this.PostfixName = sysN.Name.Replace("System_", "");
                if (this.PostfixName.Length == 0)
                    throw new NotSupportedException("Пустой постфикс");

                this.DataRange = this.GetDefinedRange(EnumDyDocSsRanges.Data);

                this.SubType = EnumSheetSubType.Undef;
                this.RowType = EnumSheetRowType.Static;
            }

            /// <summary> Подтип листа. (а надо ли?) </summary>
            public enum EnumSheetSubType
            {
                Undef,
                Calc,
                Count,
                Summ
            }

            public enum EnumSheetRowType
            {
                Static,
                Dynamic
            }

            /// <summary> Подтип листа </summary>
            public EnumSheetSubType SubType { get; set; }

            /// <summary> Порядок формирования строк (сейчас статика (все строки в наличии сразу) и динамика (считываем с росписи)) </summary>
            public EnumSheetRowType RowType { get; set; }


            /// <summary> Ссылка на шит </summary>
            [NotNull]
            public Worksheet Sheet { get; set; }
            
            /// <summary> Системная информация по листу </summary>
            [NotNull]
            public DefinedName SystemRange { get; set; }

            /// <summary> Собственно, область данных ПБС </summary>
            [NotNull]
            public DefinedName DataRange { get; set; }

            /// <summary> Постфикс листа </summary>
            [NotNull]
            public string PostfixName { get; set; }

            ///// <summary> Информация по привязке колонки к данным </summary>
            //private Dictionary<int, DyTemplateColumnBind> _columnBind = new Dictionary<int, DyTemplateColumnBind>();

            ///// <summary> Добавить информацию по привязке колонки </summary>
            //public void AddColumnBindInfo(int cellIndex, [NotNull] DyTemplateColumnBind columnBindData)
            //{
            //    this._columnBind.Add(cellIndex, columnBindData);
            //}

            ///// <summary> Получение информации по привязке колонки. Null если почему-то ничего не привязано. </summary>
            //[CanBeNull]
            //public DyTemplateColumnBind GetColumnBind(int cellIndex)
            //{
            //    DyTemplateColumnBind bnd;
            //    if (this._columnBind.TryGetValue(cellIndex, out bnd))
            //        return bnd;
            //    return null;
            //}

            /// <summary> Параметра добавления колонки </summary>
            public class AddColumnParam
            {
                public AddColumnParam()
                {
                    this.IsEditable = true;
                }

                /// <summary> Заголовок </summary>
                public string Caption { get; set; }
                
                /// <summary> Размерность </summary>
                public int Precision { get; set; }

                /// <summary> Системное имя (используется для раскрутки параметров АКСИОКа) </summary>
                [CanBeNull] 
                public string SystemName { get; set; }
                
                /// <summary> Привязка колонки к данным (используется для раскрутки параметров АКСИОКа) </summary>
                [CanBeNull]
                public string ColumnBinding { get; set; }

                /// <summary> Можно ли редактировать (используется для раскрутки параметров АКСИОКа) </summary>
                public bool IsEditable { get; set; }
            }

            public void AddColumnBefore(int beforeColumn, [NotNull] AddColumnParam addColumnParam)
            {
                this.ClearStructCache();

                //var tmpBnd = this._columnBind;
                //this._columnBind = new Dictionary<int, DyTemplateColumnBind>();

                //foreach (var unchangeBnd in tmpBnd.Where(x => x.Key < beforeColumn))
                //    this._columnBind.Add(unchangeBnd.Key, unchangeBnd.Value);

                //foreach (var unchangeBnd in tmpBnd.Where(x => x.Key >= beforeColumn))
                //    this._columnBind.Add(unchangeBnd.Key + 1, unchangeBnd.Value);


                this.Sheet.Workbook.BeginUpdate();
                var columnSysName = addColumnParam.SystemName ?? "Пользователь\n" + DateTime.Now.Ticks;
                //this._columnBind.Add(beforeColumn, new DyTemplateColumnBind(columnSysName));

                var funcTop = new Func<EnumDyDocSsRanges, int>(rngNm =>
                {
                    var rng = this.TryGetDefinedRange(rngNm);
                    if (rng == null)
                        return int.MaxValue;
                    return rng.Range.TopRowIndex;
                });
                var funcBttm = new Func<EnumDyDocSsRanges, int>(rngNm =>
                {
                    var rng = this.TryGetDefinedRange(rngNm);
                    if (rng == null)
                        return int.MinValue;
                    return rng.Range.BottomRowIndex;
                });
                var copyTopBottom = new[]
                    {
                        EnumDyDocSsRanges.ColumnHeaders,
                        EnumDyDocSsRanges.SysColumnNames,
                        EnumDyDocSsRanges.RowBind,
                        EnumDyDocSsRanges.Data,
                        EnumDyDocSsRanges.TotalSum
                    }
                    .Select(x => new {Top = funcTop(x), Bottom = funcBttm(x)})
                    .ToList();
                var topRow = copyTopBottom.Min(x => x.Top);
                var bottomRow = copyTopBottom.Max(x => x.Bottom);
                var copyRange = this.Sheet.Range.FromLTRB(beforeColumn, topRow, beforeColumn, bottomRow);
                this.Sheet.InsertCells(copyRange, InsertCellsMode.ShiftCellsRight);

                var exlColumn = this.Sheet.Columns[beforeColumn];

                var mainDigitFormat = "#,##0" + (addColumnParam.Precision > 0 ? "." + new string('0', addColumnParam.Precision) : "");

                if (addColumnParam.IsEditable)
                {
                    // Расширяем область ввода только если данные редактируемы (т.е. пользовательские)
                    var firstInsertColumn = this.GetDefinedRange(EnumDyDocSsRanges.FirstInsertColumn);
                    if (firstInsertColumn.Range.LeftColumnIndex > exlColumn.LeftColumnIndex)
                        firstInsertColumn.Range = this.Sheet.Range.FromLTRB(exlColumn.LeftColumnIndex, 0,
                            firstInsertColumn.Range.RightColumnIndex, firstInsertColumn.Range.BottomRowIndex);
                }

                // Прописываем системные имена
                var sysColumnNamesRange = this.TryGetDefinedRange(EnumDyDocSsRanges.SysColumnNames);
                if (sysColumnNamesRange != null) 
                    exlColumn[sysColumnNamesRange.Range.TopRowIndex].Value = columnSysName;

                // Прописываем пустой биндинг что б ошибок не вылезало
                var rowBindRange = this.TryGetDefinedRange(EnumDyDocSsRanges.RowBind);
                if (rowBindRange != null) 
                    exlColumn[rowBindRange.Range.TopRowIndex].Value = addColumnParam.ColumnBinding ?? "Пусто";

                // Рисуем футер
                var ttlRange = this.TryGetDefinedRange(EnumDyDocSsRanges.TotalSum);
                if (ttlRange != null)
                {
                    var footerCell = exlColumn[ttlRange.Range.TopRowIndex];
                    var nextToCellFormat = this.Sheet.Range.FromLTRB(exlColumn.LeftColumnIndex+1, ttlRange.Range.TopRowIndex,
                            exlColumn.RightColumnIndex + 1, ttlRange.Range.TopRowIndex);
                    footerCell.CopyFrom(nextToCellFormat, PasteSpecial.Borders | PasteSpecial.Formats); // Копируем всё из соседней ячейки. И потом заменяем формулу и часть формата
                    footerCell.SetFormulaSumRange(exlColumn[this.DataRange.Range.TopRowIndex], exlColumn[this.DataRange.Range.BottomRowIndex]);
                    footerCell.NumberFormat = mainDigitFormat;
                    //footerCell.Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    //footerCell.Font.Bold = true;
                }

                // Рисуем заголовок
                var columnHeaders = this.TryGetDefinedRange(EnumDyDocSsRanges.ColumnHeaders);
                if (columnHeaders != null)
                {
                    var usrColumnHeaderCell = exlColumn[columnHeaders.Range.BottomRowIndex];
                    Range usrColumnHeader;
                    if (usrColumnHeaderCell.IsMerged)
                    {
                        //TODO На данный момент не отрабатывает. При простой вставке колонки ячейки не объеденены.
                        // По хорошему - надо бы скопировать объединение колонок с "первой колонки" Но пока вроде нет необходимости.
                        var rngs = usrColumnHeaderCell.GetMergedRanges().ToList();
                        usrColumnHeader = usrColumnHeaderCell.Worksheet.Range.FromLTRB(
                                    rngs.Min(x => x.LeftColumnIndex),
                                    rngs.Min(x => x.TopRowIndex),
                                    rngs.Max(x => x.RightColumnIndex),
                                    rngs.Max(x => x.BottomRowIndex));
                    }
                    else
                    {
                        usrColumnHeader = usrColumnHeaderCell.Worksheet.Range.FromLTRB(
                                    usrColumnHeaderCell.LeftColumnIndex,
                                    usrColumnHeaderCell.TopRowIndex,
                                    usrColumnHeaderCell.RightColumnIndex,
                                    usrColumnHeaderCell.BottomRowIndex);
                    }

                    var nextToCellFormat = this.Sheet.Range.FromLTRB(exlColumn.LeftColumnIndex + 1, usrColumnHeader.TopRowIndex,
                        exlColumn.RightColumnIndex + 1, usrColumnHeader.TopRowIndex);
                    usrColumnHeader.CopyFrom(nextToCellFormat, PasteSpecial.Borders | PasteSpecial.Formats); // Копируем всё из соседней ячейки. И потом заменяем формулу и часть формата

                    usrColumnHeader.Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    usrColumnHeader.Value = addColumnParam.Caption ?? "";

                    if (addColumnParam.IsEditable)
                    {
                        // Разрешаем ввод данных пользователем
                        usrColumnHeader.Fill.BackgroundColor = Color.Yellow;
                        usrColumnHeader.Protection.Locked = false;
                    }
                }

                // Форматируем колонку с собственно данными
                for (var rI = this.DataRange.Range.TopRowIndex; rI <= this.DataRange.Range.BottomRowIndex; rI++)
                {
                    if (addColumnParam.IsEditable)
                    {
                        // Разрешаем ввод данных пользователем
                        exlColumn[rI].Fill.BackgroundColor = Color.Yellow;
                        exlColumn[rI].Protection.Locked = false;
                    }
                    exlColumn[rI].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);
                    exlColumn[rI].NumberFormat = mainDigitFormat;
                }

                // Рисуем "к распределению" и нераспределенный остаток, если есть.
                var pbsName = this.TryGetDefinedRange(EnumDyDocSsRanges.DataPbsName);
                Debug.Assert(pbsName != null, nameof(pbsName) + " != null");

                var forDistrib = this.TryGetDefinedRange(EnumDyDocSsRanges.ForDistrib);
                if (forDistrib != null)
                {
                    exlColumn[forDistrib.Range.TopRowIndex].CopyFrom(this.Sheet.Range.FromLTRB(pbsName.Range.LeftColumnIndex, 
                        forDistrib.Range.TopRowIndex,
                        pbsName.Range.LeftColumnIndex,
                        forDistrib.Range.TopRowIndex), 
                        PasteSpecial.Formats);
                }

                var residue = this.TryGetDefinedRange(EnumDyDocSsRanges.Residue);
                if (residue != null)
                {
                    exlColumn[residue.Range.TopRowIndex].CopyFrom(this.Sheet.Range.FromLTRB(pbsName.Range.LeftColumnIndex,
                            residue.Range.TopRowIndex,
                            pbsName.Range.LeftColumnIndex,
                            residue.Range.TopRowIndex),
                        PasteSpecial.Formats);
                }

                this.Sheet.Workbook.EndUpdate();
            }


            /// <summary> Получить список ВСЕХ ячеек, входящий в область Data </summary>
            public List<Cell> GetAllDataCells()
            {
                var cells = new List<Cell>();
                foreach (var cell in this.DataRange.Range.ExistingCells)
                    cells.Add(cell);
                return cells;
            }

            /// <summary> Удалить колонку </summary>
            public void RemoveColumn(int columnIndex)
            {
                this.ClearStructCache();

                //var bindCol = this.GetColumnBind(columnIndex);
                //if (bindCol == null)
                //{
                //    LoggerS.WarnMessage("Ошибка поиска биндиной колонки " + columnIndex);
                //    return;
                //}

                // Проверка наличия ссылок на удаляемую ячейку
                var exlColumn = this.Sheet.Columns[columnIndex];
                var chkCells = new List<Cell>();
                for (var rI = this.DataRange.Range.TopRowIndex; rI < this.DataRange.Range.BottomRowIndex; rI++)
                    chkCells.Add(exlColumn[rI]);
                var possibleRefCells = this.GetAllDataCells();
                chkCells.ForEach(x => possibleRefCells.Remove(x));

                var formulaEngine = this.Sheet.Workbook.FormulaEngine;
                foreach (var refCell in possibleRefCells)
                {
                    if (string.IsNullOrWhiteSpace(refCell.Formula))
                        continue;

                    var parsedExpression = formulaEngine.Parse(refCell.Formula);
                    if (parsedExpression
                            .GetRanges()
                            .SelectMany(x => x.ExistingCells)
                            .Any(x => chkCells.Any(xx => Equals(xx, x))))
                    {
                        throw new NotSupportedException("Есть внешняя ссылка из " + refCell.GetReferenceA1());
                    }
                }


                //// Собственно удаление
                //var tmpBnd = this._columnBind;
                //this._columnBind = new Dictionary<int, DyTemplateColumnBind>();

                //foreach (var unchangeBnd in tmpBnd.Where(x => x.Key < columnIndex))
                //    this._columnBind.Add(unchangeBnd.Key, unchangeBnd.Value);

                //foreach (var unchangeBnd in tmpBnd.Where(x => x.Key > columnIndex))
                //    this._columnBind.Add(unchangeBnd.Key - 1, unchangeBnd.Value);

                this.Sheet.Workbook.BeginUpdate();
                this.Sheet.Columns.Remove(columnIndex);
                this.Sheet.Workbook.EndUpdate();
            }


            /// <summary> Получить именованную область </summary>
            [CanBeNull]
            public DefinedName TryGetDefinedRange(EnumDyDocSsRanges enumRangeName)
            {
                return this.Sheet.Workbook.DefinedNames.SingleOrDefault(x => x.Name == enumRangeName + "_" + this.PostfixName);
            }

            /// <summary> Получить именованную область </summary>
            [NotNull]
            public DefinedName GetDefinedRange(EnumDyDocSsRanges enumRangeName)
            {
                return this.Sheet.Workbook.DefinedNames.SingleOrDefault(x => x.Name == enumRangeName + "_" + this.PostfixName)
                    ?? throw new NotSupportedException("Ошибка получения области " + enumRangeName);
            }

            #region Работа с инвариантами формул
            /// <summary> Попытка вычислить основную формулу колонки </summary>
            /// <param name="columnIndex">Колонка в абсолютных значениях</param>
            /// <returns>Формула</returns>
            /// <remarks>
            /// Формула считается основным инвариантом, если этот инвариант БОЛЬШЕ чем в половине строк.
            /// В остальных случаях инвариант нет.
            /// </remarks>
            [NotNull]
            public string GetMainFormulaInvariantForColumn(int columnIndex)
            {
                if (this.DataRange.Range.LeftColumnIndex > columnIndex
                    || this.DataRange.Range.RightColumnIndex < columnIndex)
                {
                    throw new NotSupportedException("Попытка разобраться ");
                }

                var invariantFormulas = new List<string>();
                for (var rIdx = this.DataRange.Range.TopRowIndex; rIdx <= this.DataRange.Range.BottomRowIndex; rIdx++)
                {
                    var cell = this.Sheet.Cells[rIdx, columnIndex];
                    if (string.IsNullOrWhiteSpace(cell.FormulaInvariant))
                        continue;

                    invariantFormulas.Add(this.GetCellFormulaInvariant(cell, cell.FormulaInvariant));
                }

                var rowsCount = this.DataRange.Range.BottomRowIndex - this.DataRange.Range.TopRowIndex;

                var invFrm = invariantFormulas
                            .GroupBy(x => x)
                            .FirstOrDefault(x => !string.IsNullOrEmpty(x.Key) && x.Count() > rowsCount / 2);

                if (invFrm != null)
                    return invFrm.Key;

                return string.Empty;
            }

            /// <summary> Объём кеша </summary>
            private const int FormulaInvariantCacheCapacity = 100000;

            /// <summary> Кеш инвариантов формул </summary>
            private Dictionary<string, Tuple<int, string, string>> _formulaInvariantCache = new Dictionary<string, Tuple<int, string, string>>(FormulaInvariantCacheCapacity);

            /// <summary> Индекс, что б всё время увеличивался </summary>
            private int _formulaInvariantCacheIndex = 0;

            /// <summary> Получить инвариант формулы независимо от строки. </summary>
            /// <param name="cell">Обрабатываемая ячейка</param>
            /// <param name="cellFormula">Формула для этой ячейки (отдельно формула, что б можно было обрабатывать вводимую, но ещё не сохранённую ячейку)</param>
            /// <returns></returns>
            /// <remarks>
            /// Инвариантом формулы относительно строки является некая формула, где вся информация по строкам заменяется на инвариант по колонкам.
            /// Т.е. типа
            /// =SUM(B2:Z2) -> =SUM(B1:Z1)
            /// =SUM(B10:Z10) -> =SUM(B1:Z1)
            /// что б можно было сравнить формулы в двух разных строках без учёта строк
            /// </remarks>>
            [NotNull]
            public string GetCellFormulaInvariant([NotNull] Cell cell, [NotNull] string cellFormula)
            {
                var cellRefStr = cell.GetReferenceA1(ReferenceElement.ColumnAbsolute);

                Tuple<int, string, string> cacheValue;
                if (this._formulaInvariantCache.TryGetValue(cellRefStr, out cacheValue))
                {
                    if (cacheValue.Item2 == cellFormula) // Если формула поменялась - тогда инвариант надо переделать
                        return cacheValue.Item3;
                    else
                        this._formulaInvariantCache.Remove(cellRefStr);
                }

                if (string.IsNullOrWhiteSpace(cellFormula))
                {
                    return "";
                }

                var invFormula = this.GetCellFormulaInvariantCalc(cell, cellFormula);

                cacheValue = Tuple.Create(++this._formulaInvariantCacheIndex, cellFormula, invFormula);
                if (this._formulaInvariantCache.Count > FormulaInvariantCacheCapacity - 2)
                {
                    // Если у нас кончился кеш - просто выкидываем первую тыщщу значений
                    var clearIndex = this._formulaInvariantCacheIndex - FormulaInvariantCacheCapacity + 1000;
                    foreach (var cacheItem in this._formulaInvariantCache
                        .Where(x => x.Value.Item1 < clearIndex).ToList())
                    {
                        this._formulaInvariantCache.Remove(cacheItem.Key);
                    }
                }
                this._formulaInvariantCache.Add(cellRefStr, cacheValue);

                return invFormula;
            }

            /// <summary> Вычисление инвариантной формулы </summary>
            /// <param name="cell">Опорная ячейка</param>
            /// <param name="formulaInvariantCulture">Формула в инвариантной культуре</param>
            /// <returns>Инвариантная формула</returns>
            [NotNull]
            public string GetCellFormulaInvariantCalc([NotNull] Cell cell, [NotNull] string formulaInvariantCulture)
            {
                var formulaEngine = this.Sheet.Workbook.FormulaEngine;
                var culture = this.Sheet.Workbook.Options.Culture;
                this.Sheet.Workbook.Options.Culture = CultureInfo.InvariantCulture;
                var parsedExpression = formulaEngine.Parse(formulaInvariantCulture);
                this.Sheet.Workbook.Options.Culture = culture;
                parsedExpression.Expression.Visit(new InvariantVisitor(cell));
                return parsedExpression.ToString();
            }

            /// <summary> Вычисление оригинальной формулы для чейки по её инварианту </summary>
            /// <param name="cell">Опорная ячейка</param>
            /// <param name="invariantFormulaInvariantCulture">Инвариантная формула в инвариантной культуре</param>
            /// <returns>Конектсная формула</returns>
            [NotNull]
            public string GetCellFormulaOriginalByInvariantFormulaCalc([NotNull] Cell cell, [NotNull] string invariantFormulaInvariantCulture)
            {
                var formulaEngine = this.Sheet.Workbook.FormulaEngine;
                var culture = this.Sheet.Workbook.Options.Culture;
                this.Sheet.Workbook.Options.Culture = CultureInfo.InvariantCulture;
                var parsedExpression = formulaEngine.Parse(invariantFormulaInvariantCulture);
                this.Sheet.Workbook.Options.Culture = culture;
                parsedExpression.Expression.Visit(new InvariantInvertVisitor(cell));
                return parsedExpression.ToString();
            }

            /// <summary> Визитор, который замещает "текущие" строки на "первую", что б получать одинаковую строку формулы для разных строк </summary>
            /// <remarks>
            /// Так же можно сделать и для растягивания областей
            /// https://docs.devexpress.com/OfficeFileAPI/DevExpress.Spreadsheet.Formulas.ExpressionVisitor
            /// </remarks>
            private class InvariantVisitor : ExpressionVisitor
            {
                public InvariantVisitor([NotNull] Cell cell)
                {
                    this.BaseCell = cell;
                }

                /// <summary> Базовая ячейка для вычисления формулы </summary>
                public Cell BaseCell { get; set; }

                public override void Visit(CellReferenceExpression expression)
                {
                    if (this.BaseCell.RowIndex == expression.CellArea.TopRowIndex
                        && this.BaseCell.RowIndex == expression.CellArea.BottomRowIndex
                    )
                    {
                        expression.CellArea.TopRowIndex = 0;
                        expression.CellArea.BottomRowIndex = 0;
                    }

                    base.Visit(expression);
                }
            }

            /// <summary> Визитор, который замещает "первую" строку на "текущие", что б получать правильную формулу по её инваринату </summary>
            private class InvariantInvertVisitor : ExpressionVisitor
            {
                public InvariantInvertVisitor([NotNull] Cell cell)
                {
                    this.BaseCell = cell;
                }

                /// <summary> Базовая ячейка для вычисления формулы </summary>
                public Cell BaseCell { get; set; }

                public override void Visit(CellReferenceExpression expression)
                {
                    if (expression.CellArea.TopRowIndex == 0
                        && expression.CellArea.BottomRowIndex == 0
                    )
                    {
                        expression.CellArea.TopRowIndex = this.BaseCell.RowIndex;
                        expression.CellArea.BottomRowIndex = this.BaseCell.RowIndex;
                    }

                    base.Visit(expression);
                }
            }

            #endregion Работа с инвариантами формул

            #region Работа с инвариантными координатами

            /// <summary> Кеш наименований колонок области данных </summary>
            private List<Tuple<int, string>> _colSysNames = new List<Tuple<int, string>>(500);

            /// <summary> Кеш кодов ПБС области данных </summary>
            private List<Tuple<int, string>> _rowPbsNames = new List<Tuple<int, string>>(500);

            /// <summary> Сброс кешей структуры  </summary>
            private void ClearStructCache()
            {
                this._colSysNames.Clear();
                this._rowPbsNames.Clear();
            }

            /// <summary> Получить системное имя колонки. Если не область данных или коэффициентов - то null </summary>
            [NotNull]
            private string InvariantGetColumnName([NotNull] Cell cell, InvariantCellCoord.EnumCellCoordRegion area)
            {
                if (area == InvariantCellCoord.EnumCellCoordRegion.Data)
                {
                    var inf = this._colSysNames.SingleOrDefault(x => x.Item1 == cell.LeftColumnIndex);
                    if (inf != null)
                        return inf.Item2;

                    var sysNamesRange = this.GetDefinedRange(EnumDyDocSsRanges.SysColumnNames);
                    var sysColumnName = this.Sheet.Cells[sysNamesRange.Range.TopRowIndex, cell.LeftColumnIndex].Value.ToString();
                    this._colSysNames.Add(Tuple.Create(cell.LeftColumnIndex, sysColumnName));
                    
                    return sysColumnName;
                }

                if (area == InvariantCellCoord.EnumCellCoordRegion.Coeff)
                {
                    var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                    if (coeffRegion != null)
                        return (cell.ColumnIndex - coeffRegion.Range.LeftColumnIndex).ToString();
                    else
                        throw new NotSupportedException("Попытка получения имени колони области коэффициентов при отсутсвии области коэффициентов");
                }

                throw new NotSupportedException("Неизвестная область строк " + area);
            }

            /// <summary> Получить системное имя колонки. Если не область данных или коэффициентов - то null </summary>
            [NotNull]
            private string InvariantGetRowName([NotNull] Cell cell, InvariantCellCoord.EnumCellCoordRegion area)
            {
                if (area == InvariantCellCoord.EnumCellCoordRegion.Data)
                {
                    var inf = this._rowPbsNames.SingleOrDefault(x => x.Item1 == cell.TopRowIndex);
                    if (inf != null)
                        return inf.Item2;

                    var pbsCodeRange = this.GetDefinedRange(EnumDyDocSsRanges.DataPbsCode);
                    var pbsCode = this.Sheet.Cells[cell.TopRowIndex, pbsCodeRange.Range.LeftColumnIndex].Value.ToString();
                    this._rowPbsNames.Add(Tuple.Create(cell.TopRowIndex, pbsCode));
                    return pbsCode;
                }

                if (area == InvariantCellCoord.EnumCellCoordRegion.Coeff)
                {
                    var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                    if (coeffRegion != null)
                        return (cell.RowIndex - coeffRegion.Range.TopRowIndex).ToString();
                    else
                        throw new NotSupportedException("Попытка получения имени строки области коэффициентов при отсутсвии области коэффициентов");
                }

                throw new NotSupportedException("Неизвестная область строк " + area);
            }

            /// <summary> Получить индекс колонки по её имени </summary>
            private int? InvariantGetColumnIndex([NotNull] string columnSysName, InvariantCellCoord.EnumCellCoordRegion area)
            {
                if (area == InvariantCellCoord.EnumCellCoordRegion.Data)
                {
                    var inf = this._colSysNames.SingleOrDefault(x => x.Item2 == columnSysName);
                    if (inf != null)
                        return inf.Item1;

                    var sysNamesRange = this.GetDefinedRange(EnumDyDocSsRanges.SysColumnNames);
                    var sysNameCell = sysNamesRange.Range.SingleOrDefault(x => x.Value.ToString() == columnSysName);
                    if (sysNameCell == null)
                        return null;

                    var tuple = new Tuple<int, string>(sysNameCell.LeftColumnIndex, columnSysName);
                    this._colSysNames.Add(tuple);
                    return tuple.Item1;
                }
                
                if (area == InvariantCellCoord.EnumCellCoordRegion.Coeff)
                {
                    var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                    if (coeffRegion != null)
                        return coeffRegion.Range.LeftColumnIndex + int.Parse(columnSysName);
                    else
                        throw new NotSupportedException("Попытка получения индекса колонки области коэффициентов при отсутсвии области коэффициентов");
                }

                throw new NotSupportedException("Неизвестная область строк " + area);
            }

            /// <summary> Получить индекс строки по её имени </summary>
            private int? InvariantGetRowIndex([NotNull] string rowSysCode, InvariantCellCoord.EnumCellCoordRegion area)
            {
                if (area == InvariantCellCoord.EnumCellCoordRegion.Data)
                {
                    var inf = this._rowPbsNames.SingleOrDefault(x => x.Item2 == rowSysCode);
                    if (inf != null)
                        return inf.Item1;

                    var pbsCodeRange = this.GetDefinedRange(EnumDyDocSsRanges.DataPbsCode);
                    var pbsCodeCell = pbsCodeRange.Range.SingleOrDefault(x => x.Value.ToString() == rowSysCode);
                    if (pbsCodeCell == null)
                        return null;

                    var tuple = new Tuple<int, string>(pbsCodeCell.TopRowIndex, rowSysCode);
                    this._rowPbsNames.Add(tuple);
                    return tuple.Item1;
                }


                if (area == InvariantCellCoord.EnumCellCoordRegion.Coeff)
                {
                    var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                    if (coeffRegion != null)
                        return coeffRegion.Range.LeftColumnIndex + int.Parse(rowSysCode);
                    else
                        throw new NotSupportedException("Попытка получения индекса строки области коэффициентов при отсутсвии области коэффициентов");
                }

                throw new NotSupportedException("Неизвестная область строк " + area);
            }


            /// <summary> Независимые от отображения "координаты" ячейки: <br/>
            /// Для области данных: Код ПБС и системное имя колонки<br/>
            /// Для коээфициентов - ??? 
            /// </summary>
            public struct InvariantCellCoord
            {
                /// <summary> Учитываемые области данных </summary>
                public enum EnumCellCoordRegion
                {
                    /// <summary> Не определено </summary>
                    Undef, 

                    /// <summary> Область данных </summary>
                    Data,
                    
                    /// <summary> Область коэффициентов </summary>
                    Coeff
                }

                public EnumCellCoordRegion CellCoordRegion { get; set; }

                /// <summary> Код ПБС </summary>
                [NotNull]
                public string InvariantRow;

                /// <summary> "Системное имя колонки" </summary>
                [NotNull]
                public string InvariantColumn;

                public InvariantCellCoord(EnumCellCoordRegion region, [NotNull] string invariantRow, [NotNull] string invariantColumn)
                {
                    this.CellCoordRegion = region;
                    this.InvariantRow = invariantRow;
                    this.InvariantColumn = invariantColumn;
                }
            }

            /// <summary>
            /// Получить "Координаты" ячейки области данных.
            /// Код ПБС и системное имя колонки
            /// </summary>
            /// <param name="cell">Ячейка</param>
            /// <returns>"Координаты" или null (если не область данных)</returns>
            [CanBeNull]
            public InvariantCellCoord? GetInvariantCoord([NotNull] Cell cell)
            {
                if (this.DataRange.Range.RegionInclusion(cell))
                {
                    var sysColumnName = this.InvariantGetColumnName(cell, InvariantCellCoord.EnumCellCoordRegion.Data);
                    var pbsCode = this.InvariantGetRowName(cell, InvariantCellCoord.EnumCellCoordRegion.Data);

                    return new InvariantCellCoord(InvariantCellCoord.EnumCellCoordRegion.Data, pbsCode, sysColumnName);
                }

                var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                if (coeffRegion != null && coeffRegion.Range.RegionInclusion(cell))
                {
                    var colId = this.InvariantGetColumnName(cell, InvariantCellCoord.EnumCellCoordRegion.Coeff);
                    var rowId = this.InvariantGetRowName(cell, InvariantCellCoord.EnumCellCoordRegion.Coeff);

                    return new InvariantCellCoord(InvariantCellCoord.EnumCellCoordRegion.Coeff, rowId, colId);
                }

                return null;
            }

            /// <summary> Найти ячейку по "системным координатам" </summary>
            /// <param name="coord">"Координаты"</param>
            /// <returns>Ячейка или null, если ячейка не найдена</returns>
            [CanBeNull]
            public Cell GetCellByInvariantCoord(InvariantCellCoord coord)
            {
                if (coord.CellCoordRegion == InvariantCellCoord.EnumCellCoordRegion.Data)
                {
                    var colIndex = this.InvariantGetColumnIndex(coord.InvariantColumn, InvariantCellCoord.EnumCellCoordRegion.Data);
                    if (colIndex == null)
                        return null;
                    var rowIndex = this.InvariantGetRowIndex(coord.InvariantRow, InvariantCellCoord.EnumCellCoordRegion.Data);
                    if (rowIndex == null)
                        return null;

                    return this.Sheet.Cells[rowIndex.Value, colIndex.Value];
                }
                
                if (coord.CellCoordRegion == InvariantCellCoord.EnumCellCoordRegion.Coeff)
                {
                    var coeffRegion = this.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                    if (coeffRegion != null)
                    {
                        var colIndex = this.InvariantGetColumnIndex(coord.InvariantColumn, InvariantCellCoord.EnumCellCoordRegion.Coeff);
                        if (colIndex == null)
                            return null;

                        var rowIndex = this.InvariantGetRowIndex(coord.InvariantRow, InvariantCellCoord.EnumCellCoordRegion.Coeff);
                        if (rowIndex == null)
                            return null;

                        return this.Sheet.Cells[rowIndex.Value, colIndex.Value];
                    }
                }

                return null;
            }
            #endregion
        }

        public static Color EditableColor = Color.Yellow;

        /// <summary> Проверка, что ячейка редактируема на урвоне шаблона (т.е. без использования Protection.Lock) </summary>
        public bool IsEditableWorkbookCellCriteria(Cell cell)
        {
            return cell.Fill.BackgroundColor.R == EditableColor.R
                   && cell.Fill.BackgroundColor.G == EditableColor.G
                   && cell.Fill.BackgroundColor.B == EditableColor.B;
        }
    }
}