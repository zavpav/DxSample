using System;
using System.Collections.Generic;

using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using JetBrains.Annotations;
using Sphaera.Bp.Services.Core;
using Sphaera.Bp.Services.Log;
using Sphaera.Bp.Services.Math;

namespace Sphaera.Bp.Bl.Excel
{
    /// <summary>
    /// Вспомогательные методы работы с spreadsheet от devexpress
    /// </summary>
    // ReSharper disable InconsistentNaming тут названия типа TLRB и пр. которые мне нравятся
    public static class SpreadSheetHelper
    {
        /// <summary> Установить число с учётом мержа </summary>
        [NotNull]
        public static Range SetValueMerged([NotNull] this Range r, decimal value)
        {
            return r.SetValueMerged((double)value);
        }

        /// <summary> Значение ячейки с учетом мержинга. </summary>
        [NotNull]
        public static Range SetValueMerged([NotNull] this Range r, [CanBeNull] CellValue value)
        {
            if (r.TopRowIndex != r.BottomRowIndex)
                throw new NotSupportedException("Не поддерживается многострочность");

            var firstColumn = r.Worksheet.Columns[r.LeftColumnIndex];
            var prevWdh = firstColumn.Width;

            var width = 0d;
            for (var cl = r.LeftColumnIndex; cl < r.RightColumnIndex; cl++)
                width += r.Worksheet.Columns[cl].Width;

            firstColumn.Width = width;
            var workCell = r.Worksheet.CellLT(r.LeftColumnIndex, r.TopRowIndex);

            workCell.Alignment.WrapText = true;
            workCell.SetValueZ(value);
            r.Merge();
            r.Alignment.WrapText = true;
            r.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            r.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
            firstColumn.Width = prevWdh;

            r.AutofitRow();

            return r;
        }

        /// <summary>
        /// Занчение ячейки с учётом мёржинга
        /// (приехало из тестов, так как мне надоело всё время копировать)
        /// </summary>
        /// <remarks>
        /// В смёрженых ячейках данные пустые, кроме "первой". Ну или какие-то кривые...
        /// </remarks>
        [NotNull]
        public static string GetMergedCellText([NotNull] this Worksheet ws, int parsingRow, int col)
        {
            var cl = ws.Cells[parsingRow, col];
            if (!cl.IsMerged)
                return cl.Value.ToString();

            return string.Join(" ", cl.GetMergedRanges().Select(x => x.Value.ToString()));
        }


        /// <summary> Получить ячейку по индексу колонки и строке </summary>
        [NotNull]
        public static Range CellLT([NotNull] this Worksheet sheet, int columnIndex, int rowIndex)
        {
            return sheet.Range.FromLTRB(columnIndex, rowIndex, columnIndex, rowIndex);
        }

        /// <summary> Получить ячейку по индексу колонки и строке </summary>
        [NotNull]
        public static Range CellLT([NotNull] this Range baseRange, int columnIndex, int rowIndex)
        {
            return baseRange.FromLTRB(columnIndex, rowIndex, columnIndex, rowIndex);
        }


        /// <summary> Обёртка установки decimal (так как в spreadsheet - double) </summary>
        [NotNull, PublicAPI]
        public static Range SetValueZ([NotNull] this Range cell, decimal value)
        {
            return cell.SetValueZ((double)value);
        }

        /// <summary> Установить значение в ячейке с корректной обработкой null </summary>
        [NotNull, PublicAPI]
        public static Range SetValueZ([NotNull] this Range cell, [CanBeNull] CellValue value)
        {
            // ReSharper disable once ConvertIfStatementToNullCoalescingExpression // Мне с if больше нравится
            if (value == null)
                cell.Value = CellValue.Empty;
            else
                cell.Value = value;
            return cell;
        }

        /// <summary> Вспомогательный метод получения подRange по левому-правому углу. </summary>
        [NotNull, PublicAPI]
        public static Range FromLTRB2Absolute([NotNull] this IRangeProvider baseRange, int leftColumn, int topRow, int rightColumn, int bottomRow)
        {
            var range = baseRange.FromLTRB(leftColumn, topRow, rightColumn, bottomRow);
            CodeStyleHelper.Assert(range != null, "range != null");
            return range;
        }


        /// <summary> Вспомогательный метод получения подRange по левому-правому углу. </summary>
        /// <exception cref="IndexOutOfRangeException">Выбираемый диапазон больше базового</exception>
        [NotNull, PublicAPI]
        public static Range FromLTRB([NotNull] this Range baseRange, int leftColumn, int topRow, int rightColumn, int bottomRow)
        {
            var refRange = baseRange.Worksheet.Range.FromLTRB(baseRange.LeftColumnIndex + leftColumn,
                                                              baseRange.TopRowIndex + topRow,
                                                              baseRange.LeftColumnIndex + rightColumn,
                                                              baseRange.TopRowIndex + bottomRow);
            if (refRange.BottomRowIndex > baseRange.BottomRowIndex ||
                refRange.RightColumnIndex > baseRange.RightColumnIndex)
                throw new IndexOutOfRangeException("Выбираемый диапазон больше базового");
            return refRange;
        }

        /// <summary> Отформатировать область как "заголовочную" </summary>
        [NotNull, PublicAPI]
        public static Range FormatCellHeader([NotNull] this Range range, [CanBeNull] Action<Range> addFormat = null)
        {
            range.FormatCellBase(addFormat);
            return range;
        }

        /// <summary> Отформатировать область как "заголовочную" </summary>
        [NotNull, PublicAPI]
        public static Range FormatCellHeader([NotNull] this Range range, [NotNull] string caption, [CanBeNull] Action<Range> addFormat = null)
        {
            range.Value = caption;
            range.Merge();
            range.FormatCellHeader(addFormat);
            return range;
        }

        /// <summary> Рисовать жирным </summary>
        [NotNull, PublicAPI]
        public static Range FormatBold([NotNull] this Range range)
        {
            range.Font.Bold = true;
            return range;
        }

        /// <summary> Установить шрифт на оласть </summary>
        [NotNull, PublicAPI]
        public static Range FormatFontSize([NotNull] this Range range, double size, bool isRelative = false)
        {
            if (isRelative)
                range.Font.Size += size;
            else
                range.Font.Size = size;

            return range;
        }

        /// <summary> Установить шрифт на оласть </summary>
        [NotNull, PublicAPI]
        public static Range FormatCellHeader([NotNull] this Range range, [NotNull] string caption, int columnWidthInCharacters, [CanBeNull] Action<Range> addFormat = null)
        {
            range.FormatCellHeader(caption, addFormat);
            range.ColumnWidthInCharacters = columnWidthInCharacters;
            return range;
        }


        /// <summary> Отформатировать область под целые числа</summary>
        /// <param name="range">Область</param>
        /// <param name="addFormat">Функция дополнительного форматирования</param>
        [NotNull, PublicAPI]
        public static Range FormatCellSm0([NotNull] this Range range, [CanBeNull] Action<Range> addFormat = null)
        {
            return range.FormatCellNum(0, addFormat);
        }

        /// <summary> Отформатировать область под числа с 1 знаком</summary>
        /// <param name="range">Область</param>
        /// <param name="addFormat">Функция дополнительного форматирования</param>
        [NotNull]
        public static Range FormatCellSm1([NotNull] this Range range, [CanBeNull] Action<Range> addFormat = null)
        {
            return range.FormatCellNum(1, addFormat);
        }

        /// <summary> Отформатировать область под числа с 2 знаками</summary>
        /// <param name="range">Область</param>
        /// <param name="addFormat">Функция дополнительного форматирования</param>
        /// <param name="fontSz">Размер шрифта</param>
        [NotNull, PublicAPI]
        public static Range FormatCellSm2([NotNull] this Range range, [CanBeNull] Action<Range> addFormat = null, double fontSz = 8)
        {
            return range.FormatCellNum(2, addFormat, fontSz);
        }

        /// <summary> Отформатировать область под числа </summary>
        /// <param name="range">Область</param>
        /// <param name="decision">Знаков после запятой</param>
        /// <param name="addFormat">Функция дополнительного форматирования</param>
        /// <param name="fontSz">Размер шрифта</param>
        [NotNull]
        public static Range FormatCellNum([NotNull] this Range range, int decision, [CanBeNull] Action<Range> addFormat = null, double fontSz = 8)
        {
            range.FormatCellBase(addFormat, fontSz);
            range.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right;

            var format = "#,##0" + (decision > 0 ? "." + new string('0', decision) : "");
            range.NumberFormat = format;


            return range;
        }

        /// <summary> Базовое форматирование ячеек </summary>
        /// <param name="range">Форматируемый диапазон</param>
        /// <param name="addFormat">Дополниетльное форматирование</param>
        /// <param name="fontSz">Размер шрифта</param>
        /// <returns></returns>
        [NotNull, PublicAPI]
        public static Range FormatCellBase([NotNull] this Range range, [CanBeNull, InstantHandle] Action<Range> addFormat = null, double fontSz = 8)
        {
            range.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            range.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            range.Alignment.WrapText = true;

            range.Font.Size = fontSz;
            range.Font.Name = "Arial";

            if (addFormat != null)
                addFormat(range);

            return range;
        }

        /// <summary> Сделать сумму между двух ячеек </summary>
        /// <param name="formulaCell">Где рисуем формулы</param>
        /// <param name="cell1">Начальный диапазон</param>
        /// <param name="cell2">Конечный диапазон</param>
        /// <returns></returns>
        [NotNull, PublicAPI]
        public static Range SetFormulaSumRange([NotNull] this Range formulaCell, [NotNull] Range cell1, [NotNull] Range cell2)
        {
            formulaCell.Formula = string.Format("=SUM({0}:{1})", cell1.GetReferenceA1(), cell2.GetReferenceA1());
            return formulaCell;
        }

        /// <summary>  Поиск именованной области.  Если не найдно - Exception  </summary>
        /// <exception cref="NoDataFoundException">Не найдена область</exception>
        [NotNull, Pure]
        public static DefinedName GetNamedRange([NotNull] this IWorkbook document, [NotNull] string rangeName)
        {
            try
            {
                return document.DefinedNames.Single(x => x.Name == rangeName);
            }
            catch (InvalidOperationException)
            {
                throw new NoDataFoundException("Не найдена область " + rangeName);
            }
            catch (ArgumentException)
            {
                throw new NoDataFoundException("Не найдена область " + rangeName);
            }
        }

        // ReSharper restore InconsistentNaming


        /// <summary> Сделать автофит ячейки </summary>
        public static void AutofitRow([NotNull] this Range range)
        {
            // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
            if (range is Cell)
            {
                var cl = (Cell) range;
                if (cl.IsMerged)
                {
                    var rngs = cl.GetMergedRanges().ToList();
                    range = range.Worksheet.Range.FromLTRB(rngs.Min(x => x.LeftColumnIndex), rngs.Min(x => x.TopRowIndex), rngs.Max(x => x.RightColumnIndex), rngs.Max(x => x.BottomRowIndex));
                }
            }

            range.Alignment.WrapText = true;
            var fontName = range.Font.Name;
            var fontSz = range.Font.Size;
            var isBold = range.Font.Bold;
            var isItalic = range.Font.Italic;
            var text = range.Value != null ? range.Value.TextValue : string.Empty;
            var cell = range.Worksheet.Cells[range.TopRowIndex, 999];
            var nWidth = cell.ColumnWidth;
            var nHeight = cell.RowHeight;

            var calcRangeWidth = CalcRangeWidth(range.Worksheet, range.LeftColumnIndex, range.RightColumnIndex);
            try
            {
                cell.ColumnWidth = calcRangeWidth;
            }
            catch (Exception ex)
            {
                LoggerS.ErrorLogger("SpreadSheetHelper", ex, "Ошибка установки ширины ячейки {0}:\n{1}\n{2}", calcRangeWidth, ex.Message, text);
                cell.ColumnWidth = 255;
            }

            cell.Alignment.WrapText = true;
            cell.Font.Name = fontName;
            cell.Font.Size = fontSz;
            cell.Font.Bold = isBold;
            cell.Font.Italic = isItalic;
            cell.SetValue(text);
            cell.AutoFitRows();
            cell.SetValue(null);
            cell.Worksheet.ClearFormats(cell);
            cell.ColumnWidth = nWidth;
            var nHeightNew = cell.RowHeight;
            if (!SphaeraMath.IsEqual(nHeightNew, nHeight))
                nHeightNew += (nHeightNew / nHeight) * 4;
            if (nHeight > nHeightNew)
                nHeightNew = nHeight;

            try
            {
                cell.RowHeight = nHeightNew;
            }
            catch (Exception ex)
            {
                cell.RowHeight = 408;
                LoggerS.ErrorLogger("SpreadSheetHelper", ex, "Ошибка установки высоты ячейки:\n{0}\n{1}", ex.Message, text);
                //throw new ArgumentException("Ошибка установки значения:\n" + ex.Message + "\n" + text, ex);
            }
        }

        /// <summary> Вычисление ширины колонки </summary>
        /// <returns></returns>
        private static double CalcRangeWidth([NotNull] Worksheet wSheet, int start, int end)
        {
            //var cell = range.Worksheet.Cells[range.TopRowIndex, range.BottomRowIndex];
            //var cellValue = cell.Value.ToString();
            //var model = (DevExpress.XtraSpreadsheet.Model.DocumentModel)range.Worksheet.Workbook.Model;
            //var modelCell = model.ActiveSheet[0, 0];
            //var measurer = new CellFormatStringMeasurer(modelCell);
            //var nWidth = measurer.MeasureStringWidth(cellValue);

            var width = 0d;
            for (var i = start; i <= end; i++)
                width += wSheet.Columns[i].Width;

            return width;
        }


        /// <summary> Вставить строку именованного пространства </summary>
        [NotNull]
        public static Range InsertRow([NotNull] this DefinedName namedRange)
        {
            var range = namedRange.Range;
            var worksheet = namedRange.Range.Worksheet;
            var rowRange = worksheet.Range[(range.TopRowIndex+1) + ":" + (range.BottomRowIndex+1)];
            worksheet.InsertCells(rowRange, InsertCellsMode.ShiftCellsDown);
            var fromRowRange = worksheet.Range[(namedRange.Range.TopRowIndex + 1) + ":" + (namedRange.Range.BottomRowIndex + 1)];
            rowRange.CopyFrom(fromRowRange, PasteSpecial.All);
            return range;
        }


        /// <summary> Удалить закрывающую строку именованного пространства</summary>
        public static void DeleteLastNamedRow([NotNull] this DefinedName namedRange, Dictionary<string, List<ICellDataSumInfo>> smInfos = null)
        {
            var worksheet = namedRange.Range.Worksheet;

            var fromRow = namedRange.Range.TopRowIndex + 1;
            var toRow = namedRange.Range.BottomRowIndex + 1;
            var rowRange = worksheet.Range[fromRow + ":" + toRow];
            worksheet.DeleteCells(rowRange, DeleteMode.ShiftCellsUp);

            // Перетрясаем дом информацию, если надо
            if (smInfos != null)
            {
                var reCellRefName = new Regex(@"^(?<mN>.*?!\D+)(?<rowIdx>\d+)$", RegexOptions.Compiled);
                foreach (var smInfo in smInfos.ToList())
                {
                    try
                    {
                        var row = reCellRefName.Match(smInfo.Key);
                        var idx = int.Parse(row.Groups["rowIdx"].Value);
                        if (idx >= fromRow && idx <= toRow)
                        {
                            LoggerS.ErrorMessage("Удаляемая строка находится в дополнительной информации: {Ключ} {Откуда} {Куда}", smInfos);
                            continue;
                        }
                        if (idx > fromRow)
                        {
                            idx = idx - (toRow - fromRow + 1);
                            smInfos.Remove(smInfo.Key);
                            smInfos.Add(reCellRefName.Replace(smInfo.Key, @"${mN}" + idx), smInfo.Value);
                        }
                    }
                    catch (Exception e)
                    {
                        LoggerS.ErrorMessage(e, "Фатальная ошибка подмены информации о сумме {Код}", smInfo.Key);
                    }
                }
            }


        }


        /// <summary> Вставить строку именованного пространства </summary>
        /// <param name="namedRange">Область, откуда берём данны</param>
        /// <returns></returns>
        [NotNull]
        public static Range InsertColumns([NotNull] this DefinedName namedRange)
        {
            var range = namedRange.Range;
            var worksheet = namedRange.Range.Worksheet;
            worksheet.InsertCells(namedRange.Range, InsertCellsMode.ShiftCellsRight);
            range.CopyFrom(namedRange.Range, PasteSpecial.All);
            return range;
        }

        /// <summary> Удалить закрывающую строку именованного пространства</summary>
        /// <remarks>
        /// Колонка удается ЦЕЛИКОМ (через DeleteCells не сработало)
        /// </remarks>
        public static void DeleteLastNamedColumn([NotNull] this DefinedName namedRange)
        {
            var worksheet = namedRange.Range.Worksheet;
            // Через DelteCells почему-то не сработало... Удаляется колнка только целиком
            // Обязательно кешируем, так как там объект разрушается
            var rangeRightColumnIndex = namedRange.Range.RightColumnIndex;
            var rangeLeftColumnIndex = namedRange.Range.LeftColumnIndex;
            for (var c = rangeRightColumnIndex; c >= rangeLeftColumnIndex; c--)
                worksheet.Columns.Remove(c);
        }


        /// <summary> Вычисление информации о колонках </summary>
        public class RealColumnIndexData
        {
            /// <summary> Индекс колонки данных </summary>
            public int ColIdx { get; private set; }

            /// <summary> Левая позиция ячейки для колонки </summary>
            public int CellLeft { get; private set; }

            /// <summary> Правая позиция ячейки для колонки </summary>
            public int CellRight { get; private set; }

            /// <summary> Вычисление информации о колонках </summary>
            /// <param name="colIdx">Индекс колонки данных</param>
            /// <param name="cellLeft">Левая позиция ячейки для колонки</param>
            /// <param name="cellRight">Правая позиция ячейки для колонки</param>
            public RealColumnIndexData(int colIdx, int cellLeft, int cellRight)
            {
                this.ColIdx = colIdx;
                this.CellLeft = cellLeft;
                this.CellRight = cellRight;
            }
        }

        /// <summary>
        /// Функция вычисления реальных значений индексов колонок в строке (range).
        /// Не совпадает, если есть merge и т.д.
        /// Идекс берётся внутри области
        /// </summary>
        /// <returns>Возвращает тройки [Колонка;Реальный индекс начала;Реальный иднекс конца]</returns>
        public static List<RealColumnIndexData> CalcRealColumnIndex([NotNull] this Range rowRange)
        {
            var columns = new List<RealColumnIndexData>();

            var cell = rowRange[0, 0];
            var colIdx = 0;
            while (cell != null)
            {
                // Вычисляем ячейку в следующей логической колонке
                var nextPosCol = cell.RightColumnIndex + 1;
                if (cell.IsMerged)
                    nextPosCol = cell.GetMergedRanges().Max(x => x.RightColumnIndex) + 1;
                nextPosCol -= rowRange.LeftColumnIndex;

                if (rowRange.ColumnCount < nextPosCol)
                    cell = null;
                else
                {
                    columns.Add(new RealColumnIndexData(colIdx, cell.LeftColumnIndex - rowRange.LeftColumnIndex, nextPosCol - 1));
                    cell = rowRange[0, nextPosCol];
                    colIdx++;
                }
            }

            return columns;
        }

        /// <summary> Поколоночное заполнение строки </summary>
        /// <param name="rowRange">Заполняемая область строки</param>
        /// <param name="cellDatas">Значение в строке. Должно быть строго равно количеству "логических колонок" (т.е. учитывается megre). </param>
        /// <param name="smInfos">Дополнительная информация о сумме</param>
        public static void FillRow(
                        [NotNull] this Range rowRange, 
                        [NotNull] List<CellData> cellDatas,
                        [CanBeNull] Dictionary<string, List<ICellDataSumInfo>> smInfos = null)
        {
            var columnTuples = rowRange.CalcRealColumnIndex();
            if (columnTuples.Count > cellDatas.Count)
                throw new NotSupportedException("Ненайдено необходимое количество колонок");
            if (columnTuples.Count < cellDatas.Count)
                throw new NotSupportedException("Слишком много данных");

            foreach (var columnTuple in columnTuples)
            {
                var cell = rowRange[0, columnTuple.CellLeft];
                CodeStyleHelper.Assert(cell != null, "cell != null");

                var cellData = cellDatas[columnTuple.ColIdx];
                var cellVal = cellData.Val;
                // ReSharper disable once CanBeReplacedWithTryCastAndCheckForNull
                if (cellVal is string)
                {
                    var txtVal = (string)cellVal;
                    if (txtVal.StartsWith("="))
                        cell.FormulaInvariant = rowRange.ReplaceColumnRefs(cell, txtVal, columnTuples);
                    else
                        cell.Value = txtVal;
                }
                else if (cellVal is decimal)
                    cell.Value = (double)(decimal)cellVal;
                else if (cellVal is double)
                    cell.Value = (double)cellVal;
                else if (cellVal is int)
                    cell.Value = (int)cellVal;
                else if (cellVal is long)
                    cell.Value = (long)cellVal;
                else if (cellVal is CellData.EnumSpecCellData && ((CellData.EnumSpecCellData)cellVal) == CellData.EnumSpecCellData.Skip)
                {// Ничего не делаем. Дальше только следующее значение вычисляем
                }
                else
                    throw new NotSupportedException("Неизвестный тип значения " + cellVal.GetType().FullName);

                if (!string.IsNullOrWhiteSpace(cellData.CellComment))
                {
//                    var sys = Locator.Resolve<IMainContext>();
                    var cellRange = rowRange.FromLTRB(columnTuple.CellLeft, 0, columnTuple.CellRight, 0);
//                    var cmm = cell.Worksheet.Comments.Add(cellRange, sys.Project.RusName());
//                    cmm.Text = cellData.CellComment;
                }
                    

                cell.AutofitRow();

                if (smInfos != null)
                {
                    // Если есть - добиваем информацию по суммам
                    if (cellData.SumInfos != null)
                    {
                        if (cell.IsMerged)
                            foreach (var cl in cell.GetMergedRanges()
                                                .SelectMany(x => x.ExistingCells)
                                                .Distinct())
                            {
                                smInfos.Add(cl.GetReferenceA1(ReferenceElement.IncludeSheetName), cellData.SumInfos);
                            }
                        else
                            smInfos.Add(cell.GetReferenceA1(ReferenceElement.IncludeSheetName), cellData.SumInfos);

                    }
                }

                if (cellData.Format != null)
                    cellData.Format(cell);
            }
        }

        /// <summary> Часто используемый regex поиска имени колонки </summary>
        [NotNull]
        private static readonly Regex ReFindColName = new Regex("^[A-Za-z]+", RegexOptions.Compiled);

        /// <summary> Часто используемый regex поиска ссылок на другие колонки в формулах </summary>
        [NotNull]
        private static readonly Regex ReFindColumnRefs = new Regex(@"<col(?<ref>[+\-]\d+)?>", RegexOptions.Compiled);

        /// <summary> Произвести замену ссылок колонок в формуле в ячейке </summary>
        /// <param name="rowRange">Обрабатываемая строка</param>
        /// <param name="cell">Обрабатываемая ячейка</param>
        /// <param name="txtVal">"Вводимое" значение</param>
        /// <param name="columnTuples">Привязки логических колонок к физичеким индексам. Может быть null. Оптимизация</param>
        private static string ReplaceColumnRefs([NotNull] this Range rowRange, [NotNull] Cell cell, [NotNull] string txtVal, [CanBeNull] List<RealColumnIndexData> columnTuples = null)
        {
            // ReSharper disable once UnusedVariable // для отладки
            var preTxt = txtVal;
            var mchs = ReFindColumnRefs.Matches(txtVal).OfType<Match>().GroupBy(x => x.Value).ToList();
            if (mchs.Count > 0)
            {
                foreach (var mch in mchs.Select(x => x.First())) // Нам всё равно какое вхождение, так как при group ссылочные элементы одинаковые
                {
                    var refCol = 0;
                    if (!string.IsNullOrEmpty(mch.Groups["ref"].Value))
                        refCol = int.Parse(mch.Groups["ref"].Value);

                    var refCell = cell;
                    if (refCol != 0)
                    {
                        if (columnTuples == null)
                            columnTuples = rowRange.CalcRealColumnIndex();

                        int currColumn;
                        try
                        {
                            currColumn = columnTuples.Single(x => x.CellLeft == cell.LeftColumnIndex).ColIdx;
                        }
                        catch (Exception)
                        {
                            throw new NoDataFoundException("Неудачный поиск текущей логической колонки. Возможно ячейка находится внути merge зоны. Поиск идёт по левому индексу." +
                                                           cell.GetReferenceA1());
                        }

                        if (currColumn + refCol < 0 || currColumn + refCol >= columnTuples.Count)
                            throw new NotSupportedException(string.Format("Ссылка на колонку находящуюся вне области строки {0} (всего {1})", currColumn + refCol, columnTuples.Count));

                        var refCellLeftIndex = columnTuples[currColumn + refCol].CellLeft;
                        refCell = rowRange.FromLTRB(refCellLeftIndex, 0, refCellLeftIndex, 0)[0];
                    }


                    var colName = refCell.GetReferenceA1();
                    colName = ReFindColName.Match(colName).Value;
                    txtVal = txtVal.Replace(mch.Value, colName);
                }
            }

            //if (txtVal != preTxt)
            //    this.Logger.Info("Замена формул в ячейке {Ячейка}: {ИсходнаяФормула} -> {НоваяФормула}", cell.GetReferenceA1(ReferenceElement.IncludeSheetName), preTxt, txtVal);
            return txtVal;
        }

        /// <summary> Смержить колонки в строке с учётом того, что часть ячеек уже помёржена</summary>
        /// <param name="rowRange">Область строки</param>
        /// <param name="fromCol">Начальная колонка (индекс с 0)</param>
        /// <param name="toCol">Коненая колонка (индекс с 0)</param>
        public static void MergeRow([NotNull] this Range rowRange, int fromCol, int toCol)
        {
            var columnTuples = rowRange.CalcRealColumnIndex();
            if (fromCol < 0 || fromCol >= columnTuples.Count)
                throw new NotSupportedException(string.Format("Неверно задана левая колонка {0} (Всего колонок {1}", fromCol, columnTuples.Count));
            if (toCol < 0 || toCol >= columnTuples.Count)
                throw new NotSupportedException(string.Format("Неверно задана правая колонка {0} (Всего колонок {1}", toCol, columnTuples.Count));

            // Проверяем что заполнено одинаково (пустые ячейки игнорируются при проверки и дозаполняются "нужным" значением, так как если мёржить ["пусто" с "данные" (в такой последовательности)] в итоге будет "пусто".

            var allCellVals = new List<CellValue>();
            for (var cIdx = columnTuples[fromCol].CellLeft; cIdx <= columnTuples[toCol].CellRight; cIdx++)
                allCellVals.Add(rowRange.FromLTRB(cIdx, 0, cIdx, 0).Value);
            var notEmptyLst = allCellVals.Select(x => x.ToString()).Where(x => !string.IsNullOrEmpty(x)).ToList();
            if (notEmptyLst.Distinct().Count() > 1)
                throw new NotSupportedException("Нельзя объеденить ячейки с разными значениями");
            //if (notEmptyLst.Count != allCellVals.Count)
            //{
            //    var anyVal = allCellVals.FirstOrDefault(x => !string.IsNullOrEmpty(x.ToString()));
            //    if (anyVal != null && allCellVals.Any(x => string.IsNullOrEmpty(x.ToString()))) // Есть непустые, и пустые. Если всё пусто или 
            //    {
            //        for (var cIdx = columnTuples[fromCol].CellLeft; cIdx <= columnTuples[toCol].CellRight; cIdx++)
            //            rowRange.FromLTRB(cIdx, 0, cIdx, 0).Value = anyVal;
            //    }
            //}

            rowRange.FromLTRB(columnTuples[fromCol].CellLeft, 0, columnTuples[toCol].CellRight, 0).Merge();
        }

        /// <summary> Замена "статических" полей по всей книге. Статические переменные в скобках, имена русские, без пробелов и т.д.</summary>
        /// <param name="workbook">Книга</param>
        /// <param name="dic">Словарь данных</param>
        /// <param name="bracket">Скобки, как определять переменные. Формат записи: Скобки, точка, скобки.</param>
        public static void ReplaceStaticFields([NotNull] this IWorkbook workbook,
            [NotNull] Dictionary<string, string> dic,
            [NotNull] string bracket = "{.}")
        {
            foreach (var worksheet in workbook.Worksheets)
                worksheet.ReplaceStaticFields(dic, bracket);
        }

        /// <summary> Замена "статических" полей на листе. Статические переменные в скобках, имена русские, без пробелов и т.д.</summary>
        /// <param name="sheet">Лист</param>
        /// <param name="dic">Словарь данных</param>
        /// <param name="bracket">Скобки, как определять переменные. Формат записи: Скобки, точка, скобки.</param>
        public static void ReplaceStaticFields([NotNull] this Worksheet sheet, 
                    [NotNull] Dictionary<string, string> dic, 
                    [NotNull] string bracket = "{.}")
        {
            var reBracket = Regex.Escape(bracket);
            var reFindVar = new Regex(reBracket.Replace("\\.", "([A-Za-zА-Яа-я0-9:_]+)"), RegexOptions.Compiled);

            foreach (var cell in sheet.GetExistingCells())
            {
                if (!string.IsNullOrWhiteSpace(cell.Formula))
                    continue;
                var cellValue = cell.Value.ToString();
                if (!string.IsNullOrWhiteSpace(cellValue))
                {
                    while (reFindVar.IsMatch(cellValue))
                    {
                        var variableName = reFindVar.Match(cellValue).Groups[1].Value;
                        string variableValue;
                        if (dic.TryGetValue(variableName, out variableValue))
                        {
                            cellValue = cellValue.Replace(bracket.Replace(".", variableName), variableValue);
                            cell.Value = cellValue;
                        }
                        else
                        {
                            cellValue = cellValue.Replace(bracket.Replace(".", variableName), "err" + variableName);

                            LoggerS.ErrorMessage("Ошибка поиска переменных для замены в словаре: {ИмяЛиста} {Переменная}", sheet.Name, variableName);
                        }
                    }
                }
            }
        }

        /// <summary> Защитить книгу по цвету </summary>
        /// <param name="workbook">Книга</param>
        /// <param name="unprotectColor">Цвет редактируемых ячеек</param>
        /// <param name="password">Пароль. По умолчанию sphaera</param>
        [PublicAPI]
        public static void ProtectByColor([NotNull] this IWorkbook workbook, Color unprotectColor, [CanBeNull] string password = null)
        {
            password = password ?? "sphaera";
            foreach (var worksheet in workbook.Worksheets)
            {
                worksheet.Protect(password, WorksheetProtectionPermissions.Default);
                worksheet.GetExistingCells().Where(x => 
                           x.Fill.BackgroundColor.R == unprotectColor.R
                        && x.Fill.BackgroundColor.G == unprotectColor.G
                        && x.Fill.BackgroundColor.B == unprotectColor.B
                    ).ForEach(x => x.Protection.Locked = false);
            }
        }


        /// <summary> Защитить книгу по цвету. РАЗРЕШИТЬ редактирование по цвету, отличному от белого </summary>
        /// <param name="workbook">Книга</param>
        /// <param name="password">Пароль, если не задан, то дефолтовый</param>
        public static void ProtectByAnotherColor(this IWorkbook workbook, [CanBeNull] string password = null)
        {
            var colors = workbook.Worksheets
                        .SelectMany(x => x.GetExistingCells())
                        .Select(x => x.Fill.BackgroundColor)
                        // Не белый или чёрный
                        .Where(x => !((x.R == 0 && x.G == 0 && x.B == 0) || (x.R == 255 && x.G == 255 && x.B == 255)))
                        .Distinct()
                        .ToList();

            if (colors.Count == 0) // Такого быть не должно. Если мы защищаем, то ХОТЬ ЧТО-ТО должно редактироваться.
                throw new NotSupportedException("Не нашли ячеек для защиты (количество цветов заливки = 1)");
            if (colors.Count != 1)
                throw new NotSupportedException("Слишком много цветов в файле. Используйте для защиты метод ProtectByColor с указанием нужного цвета.");

            workbook.ProtectByColor(colors[0], password);
        }

    }
}