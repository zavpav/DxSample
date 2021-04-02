using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using DyDocTestSS.Domain;
using DyDocTestSS.DyTemplates;
using JetBrains.Annotations;
using Sphaera.Bp.Bl.Excel;
using Sphaera.Bp.Services.Log;

namespace DyDocTestSS.Visual
{
    public class DynamicSheetController
    {
        public enum EnumColoringType
        {
            None,
            CellRegionChristmas,
        }

        private EnumColoringType ColoringType { get; set; }

        [NotNull]
        public DyDocSs DocumentInfo { get; set; }

        /// <summary> Заблокировать пересайз колонок и строк (для внутреннего форматирования и т.д.) </summary>
        public bool DisableColumnRowSizing { get; set; }

        public DynamicSheetController([NotNull] DyDocSs dyTemplateInfo)
        {
            this.DocumentInfo = dyTemplateInfo;
            this.ColoringType = EnumColoringType.None;
        }

        /// <summary> Заблокировать ввод данных </summary>
        public void ProtectByColor()
        {
            this.DocumentInfo.Wb.BeginUpdate();
            foreach (var worksheet in this.DocumentInfo.Wb.Worksheets)
            {
                var prms = WorksheetProtectionPermissions.Default | WorksheetProtectionPermissions.FormatColumns | WorksheetProtectionPermissions.FormatRows;
                worksheet.Protect("1", prms);
                worksheet.GetExistingCells()
                    .Where(this.DocumentInfo.IsEditableWorkbookCellCriteria)
                    .ForEach(x => x.Protection.Locked = false);
            }
            this.DocumentInfo.Wb.EndUpdate();
        }


        #region Скрытие регионов

        /// <summary> Отформатировать документ для показа в форме </summary>
        public void FormatForShow()
        {
            this.HideRangeInternal();
            this.HideRangeByPrefix(EnumDyDocSsRanges.Header
                                   | EnumDyDocSsRanges.Footer
                                   | EnumDyDocSsRanges.System);
        }

        /// <summary> Отображение всех скрытых колонок и т.д. по областям </summary>
        public void FormatInternal()
        {
            this.DisableColumnRowSizing = true;
            this.ShowRangeByPrefix("ForDelete");
            this.ShowRangeByPrefix("InternalData");

            this.ShowRangeByPrefix(EnumDyDocSsRanges.RowBind);
            this.ShowRangeByPrefix(EnumDyDocSsRanges.SysColumnNames);

            this.ShowRangeByPrefix(EnumDyDocSsRanges.Header
                                   | EnumDyDocSsRanges.Footer);

            //this.DisableColumnRowSizing = false;
        }

        /// <summary> Отобразить область по имени </summary>
        private void ShowRangeByPrefix(EnumDyDocSsRanges enmPrefixes)
        {
            foreach (var vl in Enum.GetValues(typeof(EnumDyDocSsRanges)).OfType<EnumDyDocSsRanges>())
            {
                if ((enmPrefixes & vl) == vl)
                    this.ShowRangeByPrefix(vl + "_");
            }
        }

        /// <summary> Отобразить область по имени </summary>
        private void ShowRangeByPrefix([NotNull] string prefixName)
        {
            foreach (var docRng in this.DocumentInfo.Wb.DefinedNames)
            {
                if (docRng.Name.StartsWith(prefixName))
                {
                    var templateInfo = this.DocumentInfo.WbTemplateInfo ?? throw new NotSupportedException("Ошибка задания шаблона для восстановления документа");
                    var tmpltRng = templateInfo.DefinedNames.SingleOrDefault(x => x.Name == docRng.Name);
                    if (tmpltRng == null)
                        throw new NotSupportedException("Пытаемся отобразить область, которой нет в шаблоне " + docRng.Name);

                    this.ShowRangeAsTemplate(docRng, tmpltRng);
                }
            }
        }

        /// <summary> Отобразить область по шаблону </summary>
        private void ShowRangeAsTemplate([NotNull] DefinedName docRng, [NotNull] DefinedName tmpltRng)
        {
            var templateRange = tmpltRng.Range;
            var documentRange = docRng.Range;

            if (templateRange.IsRangeFullCol())
            {
                if (templateRange.ColumnCount != documentRange.ColumnCount)
                    throw new NotSupportedException("Ошибка отображения региона " + docRng.Name + " не совпадает количество колонок");

                for (var i = 0; i < templateRange.ColumnCount; i++)
                    documentRange.Worksheet.Columns[i + documentRange.LeftColumnIndex].Width =
                        templateRange.Worksheet.Columns[i + templateRange.LeftColumnIndex].Width;

                return;
            }

            if (templateRange.IsRangeFullRow())
            {
                if (templateRange.RowCount != documentRange.RowCount)
                    throw new NotSupportedException("Ошибка отображения региона " + docRng.Name + " не совпадает количество строк");

                for (var i = 0; i < templateRange.RowCount; i++)
                    documentRange.Worksheet.Rows[i + documentRange.TopRowIndex].Height =
                        templateRange.Worksheet.Rows[i + templateRange.TopRowIndex].Height;

                return;
            }

            LoggerS.WarnMessage("Провальная попытка скрыть область {RegionName}. Область не строки и не колонки", tmpltRng.Name);
        }

        /// <summary> Скрыть области по наименованию области </summary>
        private void HideRangeByPrefix(EnumDyDocSsRanges enmPrefixes)
        {
            foreach (var vl in Enum.GetValues(typeof(EnumDyDocSsRanges)).OfType<EnumDyDocSsRanges>())
            {
                if ((enmPrefixes & vl) == vl)
                    this.HideRangeByPrefix(vl + "_");
            }
        }

        /// <summary> Скрыть области по наименованию области </summary>
        private void HideRangeByPrefix([NotNull] string prefixName)
        {
            foreach (var namedRng in this.DocumentInfo.Wb.DefinedNames)
            {
                if (namedRng.Name.StartsWith(prefixName))
                    this.HideRange(namedRng);
            }
        }

        /// <summary> Скрыть внутренние данные </summary>
        private void HideRangeInternal()
        {
            this.HideRangeByPrefix("ForDelete");
            this.HideRangeByPrefix("InternalData");

            //foreach (var rowBindRange in this.DocumentInfo.DocumentInfo.Wb.DefinedNames
            //    .Where(x => x.Name.StartsWith(DyTemplateInfo.EnumTemplateRanges.RowBind.ToString())))
            //{
            //    for (int i = 0; i < rowBindRange.Range.RowCount; i++)
            //        rowBindRange.Range.Worksheet.Rows[i + rowBindRange.Range.TopRowIndex].Height = 0;
            //}

            this.HideRangeByPrefix(EnumDyDocSsRanges.RowBind);
            this.HideRangeByPrefix(EnumDyDocSsRanges.SysColumnNames);
        }

        /// <summary> Спрятать область </summary>
        private void HideRange([NotNull] DefinedName namedRng)
        {
            var processingRegion = namedRng.Range;
            
            if (processingRegion.IsRangeFullCol())
            {
                for (var i = 0; i < processingRegion.ColumnCount; i++)
                    processingRegion.Worksheet.Columns[i + processingRegion.LeftColumnIndex].Width = 0;
                
                return;
            }
            
            if (processingRegion.IsRangeFullRow())
            {
                for (var i = 0; i < processingRegion.RowCount; i++)
                    processingRegion.Worksheet.Rows[i + processingRegion.TopRowIndex].Height = 0;
                
                return;
            }

            LoggerS.WarnMessage("Провальная попытка скрыть область {RegionName}. Область не строки и не колонки", namedRng.Name);
        }
        #endregion

        #region Условное форматирование бекграунда

        // ReSharper disable InconsistentNaming
        
        private readonly Brush BrushUndef = new HatchBrush(HatchStyle.DiagonalBrick, Color.Aqua, Color.Beige);
        private readonly Brush BrushNoDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Crimson, Color.Gold);
        private readonly Brush BrushNoBindDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Beige, Color.Red);
        private readonly Brush BrushInfoDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Khaki, Color.Aquamarine);
        private readonly Brush BrushSkipDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.GhostWhite, Color.DeepSkyBlue);
        private readonly Brush BrushAksiokDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.OrangeRed, Color.DeepSkyBlue);
        private readonly Brush BrushFinDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Aquamarine, Color.DeepSkyBlue);
        private readonly Brush BrushLastSmDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Gold, Color.DeepSkyBlue);
        private readonly Brush BrushDataRange = new HatchBrush(HatchStyle.DiagonalBrick, Color.Khaki, Color.PaleTurquoise);


        /// <summary> Редактируемые заголовки (пользовательские) </summary>
        private readonly Brush BrushEditableHeader = new HatchBrush(HatchStyle.LightHorizontal, Color.Aquamarine, Color.Azure);

        /// <summary> Редактируемые ячейки </summary>
        private readonly Brush BrushEditableCell = new HatchBrush(HatchStyle.DiagonalBrick, Color.Yellow, Color.Gold);

        /// <summary> Формула </summary>
        private readonly Brush BrushFillInfoFormula = new HatchBrush(HatchStyle.DiagonalBrick, Color.Aquamarine, Color.Gold);

        /// <summary> "Отличающаяся" формула </summary>
        private readonly Brush BrushFillDifferFormula = new HatchBrush(HatchStyle.DiagonalBrick, Color.Coral, Color.AliceBlue);

        /// <summary> Очень подозрительная формула (ссылки на другие строчки области данных) </summary>
        private readonly Brush BrushFillWarnFormula = new HatchBrush(HatchStyle.Cross, Color.DarkRed, Color.Gold);

        /// <summary> Ошибки в формулах редактируемых ячеек #ССЫЛКА и др. </summary>
        private readonly Brush BrushEditableCellErrors = new HatchBrush(HatchStyle.DiagonalCross, Color.Red, Color.Gold);
        
        /// <summary> Ошибки деления на ноль (потом убрать, наверное) </summary>
        private readonly Brush BrushEditableDevidedByZero = new HatchBrush(HatchStyle.ZigZag, Color.Firebrick, Color.Gold);

        /// <summary> Циклические формулы </summary>
        private readonly Brush BrushCrossRefsInvalid = new HatchBrush(HatchStyle.Cross, Color.Yellow, Color.DarkRed);

        // ReSharper restore InconsistentNaming

        public bool CustomDrawBackgroundFormatting([NotNull] Cell cell, [NotNull] Graphics graphics, Rectangle bounds)
        {
            if (!cell.Protection.Locked)
            {
                var sheet = this.DocumentInfo.SystemSheetInfos.SingleOrDefault(x => x.Sheet == cell.Worksheet);
                if (sheet == null)
                    return false;

                #region Разные виды ошибок

                if (!string.IsNullOrEmpty(cell.FormulaInvariant))
                {
                    // Общие ошибки в формулах
                    if (cell.Value.IsError)
                    {
                        graphics.FillRectangle(BrushEditableCellErrors, bounds);
                        return true;
                    }

                    //Циклические ссылки
                    if (this.CalcErrorCrossRefsInvalid(cell))
                    {
                        graphics.FillRectangle(BrushCrossRefsInvalid, bounds);
                        return true;
                    }

                    // Проверка что ссылаемся только на колонки из текущей строки. Если из другой - возмнож это ошибка
                    // Проверка только для области Data
                    if (this.CalcErrorInvalidFormualRow(cell))
                    {
                        graphics.FillRectangle(BrushFillWarnFormula, bounds);
                        return true;
                    }
                }

                #endregion

                // Редактируемые ячейки заголовка
                var columnHeaderRange = sheet.TryGetDefinedRange(EnumDyDocSsRanges.ColumnHeaders);
                if (columnHeaderRange != null &&
                    columnHeaderRange.Range.RegionInclusion(cell))
                {
                    graphics.FillRectangle(BrushEditableHeader, bounds);
                    return true;
                }

                #region Предупреждения

                // Основная формула колонки (существует только для ячеек области data)
                var columnFormulaInvariant = "";
                if (sheet.DataRange.Range.RegionInclusion(cell))
                    columnFormulaInvariant = sheet.GetMainFormulaInvariantForColumn(cell.ColumnIndex);


                if (!string.IsNullOrEmpty(columnFormulaInvariant))
                {
                    // "Отличающаяся формула"
                    var cellFormulaInvariant = sheet.GetCellFormulaInvariant(cell, cell.FormulaInvariant);
                    if (columnFormulaInvariant != cellFormulaInvariant)
                    {
                        graphics.FillRectangle(BrushFillDifferFormula, bounds);
                        return true;
                    }
                }

                #endregion


                #region Просто подсветка редактируемых данных

                    // Деление на ноль (в нашем случае - это не особо-то и ошибка)
                if (!string.IsNullOrEmpty(columnFormulaInvariant))
                {
                    var recalcedVal = this.DocumentInfo.Wb.FormulaEngine.Evaluate(cell.Formula);
                    if (recalcedVal.IsError && recalcedVal.ErrorValue.Type == ErrorType.DivisionByZero)
                    {
                        graphics.FillRectangle(BrushEditableDevidedByZero, bounds);
                        return true;
                    }
                }

                if (!string.IsNullOrEmpty(cell.FormulaInvariant))
                {
                    graphics.FillRectangle(BrushFillInfoFormula, bounds);
                    return true;
                }
                else 
                {
                    graphics.FillRectangle(BrushEditableCell, bounds);
                    return true;
                }

                #endregion

            }

            return false;

            if (cell.Fill.BackgroundColor.R == Color.Yellow.R
                && cell.Fill.BackgroundColor.G == Color.Yellow.G
                && cell.Fill.BackgroundColor.B == Color.Yellow.B)
            {
                if (!string.IsNullOrEmpty(cell.FormulaInvariant))
                {
                    var sheet = this.DocumentInfo.SystemSheetInfos
                        .SingleOrDefault(x => x.Sheet == cell.Worksheet);
                    if (sheet == null)
                        return false;

                    var columnFormulaInvariant = sheet.GetMainFormulaInvariantForColumn(cell.ColumnIndex);
                    if (!string.IsNullOrEmpty(columnFormulaInvariant))
                    {
                        var cellFormulaInvariant = sheet.GetCellFormulaInvariant(cell, cell.FormulaInvariant);
                        if (columnFormulaInvariant != cellFormulaInvariant)
                        {
                            graphics.FillRectangle(BrushFillDifferFormula, bounds);
                            return true;
                        }
                    }

                    graphics.FillRectangle(BrushFillInfoFormula, bounds);
                    return true;



                    //// Проверка что ссылаемся только на колонки из текущей строки. Если из другой - возмнож это ошибка
                    //var formulaEngine = cell.Worksheet.Workbook.FormulaEngine;
                    //var parsedExpression = formulaEngine.Parse(cell.Formula);
                    //if (parsedExpression
                    //        .GetRanges()
                    //        .SelectMany(x => x.ExistingCells)
                    //        .Any(x => x.RowIndex != cell.RowIndex))
                    //{
                    //    graphics.FillRectangle(BrushFillWarnFormula, bounds);
                    //    return true;
                    //}
                    //else
                    //{
                    //    var sheet = this.DocumentInfo.DocumentInfo.SystemSheetInfos
                    //                        .SingleOrDefault(x => x.Sheet == cell.Worksheet);
                    //    if (sheet == null)
                    //        return false;

                    //    var frm = sheet.GetMainFormulaInvariantForColumn(cell.ColumnIndex);

                        
                    //    graphics.FillRectangle(BrushFillInfoFormula, bounds);
                    //    return true;
                    //}
                }
                else
                {
                    graphics.FillRectangle(BrushEditableCell, bounds);
                    return true;
                }
            }

            else if (this.ColoringType == EnumColoringType.CellRegionChristmas)
            {
                var rowBindInfos = this.DocumentInfo.Wb
                    .DefinedNames.Where(x => x.Name.StartsWith(EnumDyDocSsRanges.RowBind.ToString()))
                    .ToList();
                
                // Информация по листу
                var shtInfo = this.DocumentInfo.SystemSheetInfos.SingleOrDefault(x => x.Sheet == cell.Worksheet);
                if (shtInfo == null)
                {
                    graphics.FillRectangle(BrushUndef, bounds);
                    return true;
                }

                Debug.Assert(shtInfo.Sheet != null, "shtInfo.Sheet != null");
                if (shtInfo.DataRange != null && shtInfo.DataRange.Range.RegionInclusion(
                        shtInfo.Sheet.Range.FromLTRB2Absolute(cell.ColumnIndex, cell.RowIndex, cell.ColumnIndex, cell.RowIndex)
                    ))
                {   // Расцветка области данных
                //    var bindDataInfo = shtInfo.GetColumnBind(cell.ColumnIndex);
                //    if (bindDataInfo == null)
                //    {
                //        graphics.FillRectangle(BrushNoBindDataRange, bounds);
                //        return true;
                //    }
                //    else
                //    {
                //        if (bindDataInfo.MainColumnData == DyTemplateColumnBind.EnumColumnDataType.Info)
                //        {
                //            graphics.FillRectangle(BrushInfoDataRange, bounds);
                //            return true;
                //        }
                //        else if (bindDataInfo.MainColumnData == DyTemplateColumnBind.EnumColumnDataType.Skip)
                //        {
                //            graphics.FillRectangle(BrushSkipDataRange, bounds);
                //            return true;
                //        }
                //        else if (bindDataInfo.MainColumnData == DyTemplateColumnBind.EnumColumnDataType.Aksiok)
                //        {
                //            graphics.FillRectangle(BrushAksiokDataRange, bounds);
                //            return true;
                //        }
                //        else if (bindDataInfo.MainColumnData == DyTemplateColumnBind.EnumColumnDataType.CurrentFinance)
                //        {
                //            graphics.FillRectangle(BrushFinDataRange, bounds);
                //            return true;
                //        }
                //        else if (bindDataInfo.MainColumnData == DyTemplateColumnBind.EnumColumnDataType.LastYearSumm)
                //        {
                //            graphics.FillRectangle(BrushLastSmDataRange, bounds);
                //            return true;
                //        }
                //        else
                //        {
                //            graphics.FillRectangle(BrushDataRange, bounds);
                //            return true;
                //        }
                //    }
                }
                else
                {
                    graphics.FillRectangle(BrushNoDataRange, bounds);
                    return true;
                }
            }

            return false;
        }

        #endregion



        /// <summary> Чистка кешей </summary>
        private void ClearCaches()
        {
            this._crossRefsInvalidCache.Clear();
            this._wrongRowsRefsCache.Clear();
        }

        #region Поиск рекурсивных ссылок

        /// <summary> Кеш результатов рекурсивных ссылок </summary>
        private readonly Dictionary<Cell, bool> _crossRefsInvalidCache = new Dictionary<Cell, bool>(10000);

        /// <summary> Проверка циклических ссылок у ячейки </summary>
        private bool CalcErrorCrossRefsInvalid([NotNull] Cell cell)
        {
            bool crossRefsInvalid;
            if (this._crossRefsInvalidCache.TryGetValue(cell, out crossRefsInvalid))
                return crossRefsInvalid;

            if (!string.IsNullOrEmpty(cell.FormulaInvariant))
            {
                var checkedCells = new List<Cell>();
                crossRefsInvalid = this.CalcCrossRefsInvalidRecursive(cell, cell, checkedCells);
            }
            else
            {
                crossRefsInvalid = false; // Для не формул сразу пусто
            }

            this._crossRefsInvalidCache.Add(cell, crossRefsInvalid);

            return crossRefsInvalid;
        }

        /// <summary> Рекурсивная проверка циклических ссылок </summary>
        /// <param name="chkCell">Проверяемая ячейка</param>
        /// <param name="processCell">Обрабатываемая ячейка</param>
        /// <param name="checkedCells">Уже проверенные ячейки</param>
        /// <returns></returns>
        private bool CalcCrossRefsInvalidRecursive([NotNull] Cell chkCell, [NotNull] Cell processCell, [NotNull] List<Cell> checkedCells)
        {
            foreach (var cl in processCell.ParsedExpression.GetRanges().SelectMany(x => x.ExistingCells))
            {
                if (checkedCells.Contains(cl))
                    continue;

                if (chkCell.Equals(cl))
                    return true;

                checkedCells.Add(cl);

                if (!string.IsNullOrWhiteSpace(cl.Formula))
                    return this.CalcCrossRefsInvalidRecursive(chkCell, cl, checkedCells);
            }

            return false;
        }

        #endregion

        #region Поиск ошибочных строк блока Data

        /// <summary> Кеш результатов ошибочных строк</summary>
        private readonly Dictionary<Cell, bool> _wrongRowsRefsCache = new Dictionary<Cell, bool>(10000);

        /// <summary> Поиск ошибок связанных со ссылками на другие строки области данных </summary>
        /// <remarks>
        /// Проверка что ссылаемся только на колонки из текущей строки. Если из другой - возмнож это ошибка
        /// Проверка только для области Data
        /// </remarks>
        private bool CalcErrorInvalidFormualRow([NotNull] Cell cell)
        {
            bool wrongRowRef;
            if (this._wrongRowsRefsCache.TryGetValue(cell, out wrongRowRef))
                return wrongRowRef;

            var sheet = this.DocumentInfo.SystemSheetInfos.SingleOrDefault(x => x.Sheet == cell.Worksheet);
            if (sheet == null)
                return false;


            wrongRowRef = false;
            var dataRangeRange = sheet.DataRange.Range;
            if (dataRangeRange.LeftColumnIndex <= cell.ColumnIndex
                && dataRangeRange.RightColumnIndex >= cell.ColumnIndex
                && dataRangeRange.TopRowIndex <= cell.RowIndex
                && dataRangeRange.BottomRowIndex >= cell.RowIndex)
            {
                var parsedExpression = cell.ParsedExpression;
                if (parsedExpression
                    .GetRanges()
                    .SelectMany(x => x.ExistingCells)
                    .Where(x => x.RowIndex != cell.RowIndex)
                    .Any(x => dataRangeRange.LeftColumnIndex <= x.ColumnIndex
                              && dataRangeRange.RightColumnIndex >= x.ColumnIndex
                              && dataRangeRange.TopRowIndex <= x.RowIndex
                              && dataRangeRange.BottomRowIndex >= x.RowIndex)
                )
                {
                    wrongRowRef = true;
                }
            }
            
            this._wrongRowsRefsCache.Add(cell, wrongRowRef);
            
            return wrongRowRef;
        }

        #endregion

        /// <summary> Получить подсказку по ячейке </summary>
        [CanBeNull]
        public ToolTipControlInfo GetCellToolTip([NotNull] Cell cell)
        {
            if (!string.IsNullOrWhiteSpace(cell.Formula))
            {
                if (this.CalcErrorCrossRefsInvalid(cell))
                    return new ToolTipControlInfo(cell, "Формула содержит циклические ссылки", ToolTipIconType.Error);

                if (this.CalcErrorInvalidFormualRow(cell))
                    return new ToolTipControlInfo(cell, "Формула содержит ссылки на другие строки области данных", ToolTipIconType.Warning);

                var recalcedVal = this.DocumentInfo.Wb.FormulaEngine.Evaluate(cell.Formula);
                if (recalcedVal.IsError && recalcedVal.ErrorValue.Type == ErrorType.DivisionByZero)
                    return new ToolTipControlInfo(cell, "Деление на ноль", ToolTipIconType.Asterisk);

                return new ToolTipControlInfo(cell, cell.Formula, ToolTipIconType.Information);
            }

            return null;
        }



        #region Обновление вводимых данных

        /// <summary> Простое определение, что вводили число </summary>
        // ReSharper disable once InconsistentNaming
        private readonly Regex IsDigitValueRegex = new Regex(@"^-?\d+(,\d+)?$");


        /// <summary> Обновление вводимых данных при копипасте</summary>
        public void UpdateEditCopyPasteData([NotNull] Cell cell)
        {
            this.ClearCaches();

            var precision = cell.GetPrecisionFromNumberFormat();
            if (precision != null && cell.Value.IsNumeric)
                cell.Value = Math.Round(cell.Value.NumericValue, precision.Value);
        }

        ///// <summary> Обновить формулу </summary>
        //public void TryUpdateFormula([NotNull] string editText, [NotNull] Cell cell)
        //{
        //    if (cell.Protection.Locked)
        //        throw new NotSupportedException("Попытка редактирования закрытую ячейку");
            
        //    var sheetInfo = this.DocumentInfo.DocumentInfo.SystemSheetInfos.SingleOrDefault(x => x.Sheet == cell.Worksheet);
        //    if (sheetInfo == null)
        //        throw new NotSupportedException("Ненайден редактируемый лист");

        //    if (sheetInfo.DataRange.Range.LeftColumnIndex > cell.LeftColumnIndex
        //            || sheetInfo.DataRange.Range.RightColumnIndex < cell.RightColumnIndex)
        //        throw new NotSupportedException("Попытка редактирования формулы вне области данных");

            
        //    if (!editText.StartsWith("="))
        //        editText = "=" + editText;

        //    var oldFormula = cell.DisplayText;
        //    if (oldFormula.StartsWith("="))
        //        oldFormula = oldFormula.TrimStart('=');

        //    if (oldFormula == editText || editText == "=" + oldFormula)
        //        return;

        //    cell.Worksheet.Workbook.BeginUpdate();
        //    for (var rIdx = sheetInfo.DataRange.Range.TopRowIndex;
        //                            rIdx <= sheetInfo.DataRange.Range.BottomRowIndex;
        //                            rIdx++)
        //    {
        //        var edCell = sheetInfo.Sheet.Range.FromLTRB2Absolute(cell.ColumnIndex, rIdx, cell.ColumnIndex, rIdx);
                                        
        //        if (edCell.Formula == oldFormula || edCell.Formula == "=" + oldFormula || edCell.Value.ToString() == "")
        //            edCell.Formula = editText;
        //    }

        //    cell.Worksheet.Workbook.EndUpdate();
        //}

        /// <summary> Обновление вводимых данных </summary>
        [NotNull]
        public string UpdateEditData([NotNull] string editorText, [NotNull] Cell cell)
        {
            var sheetInfo = this.DocumentInfo.SystemSheetInfos.SingleOrDefault(x => x.Sheet == cell.Worksheet);
            
            // Если нечего нет - игнорим действия
            if (sheetInfo == null)
                return editorText;

            this.ClearCaches();

            if (!string.IsNullOrEmpty(editorText)
                    && editorText.StartsWith("=")
                    && sheetInfo.DataRange.Range.Contains(cell)
            )
            {
                return this.AfterEditCellAddRoundInFormula(editorText, cell);
            }

            if (this.IsDigitValueRegex.IsMatch(editorText)
                    && sheetInfo.DataRange.Range.Contains(cell)
            )
            {
                // НЕ ЗАБЫТЬ Аналогичная логика при копипасте
                return this.AfterEditCellRoundEditValue(editorText, cell);
            }

            return editorText;
        }
        
        /// <summary> Округление вводимых данных до нужного формата </summary>
        [NotNull]
        private string AfterEditCellRoundEditValue([NotNull] string editorText, [NotNull] Cell cell)
        {
            var formatPrecision = cell.GetPrecisionFromNumberFormat();
            if (formatPrecision != null)
            {
                decimal val;
                if (decimal.TryParse(editorText, out val))
                {
                    return Math.Round(val, formatPrecision.Value).ToString(CultureInfo.CurrentCulture);
                }
            }

            return editorText;
        }

        /// <summary> Обрамление округлением формул до нужной размерности </summary>
        [NotNull]
        private string AfterEditCellAddRoundInFormula([NotNull] string editorText, [NotNull] Cell cell)
        {
            if (!string.IsNullOrEmpty(editorText) && editorText.StartsWith("="))
            {
                if (!editorText.StartsWith("=ROUND(") && !editorText.StartsWith("=ОКРУГЛ("))
                {
                    var formatPrecision = cell.GetPrecisionFromNumberFormat();
                    if (formatPrecision != null)
                    {
                        return "=ROUND(" + editorText.TrimStart('=') + ";" + formatPrecision + ")";
                    }
                }
                else
                {
                    var formatPrecision = cell.GetPrecisionFromNumberFormat();
                    if (formatPrecision != null)
                    {
                        var reFnd = new Regex(@"^=(ROUND\(|ОКРУГЛ\().*?;(\d+)\)\s*$");
                        var mchPrecision = reFnd.Match(editorText);
                        if (mchPrecision.Success)
                        {
                            var strPrecision = mchPrecision.Groups[2].Value;
                            int precision;
                            if (int.TryParse(strPrecision, out precision))
                            {
                                if (precision > formatPrecision)
                                {
                                    return "=ROUND(" + editorText.TrimStart('=') + ";" + formatPrecision + ")";
                                }
                            }
                        }
                        else
                        {
                            return "=ROUND(" + editorText.TrimStart('=') + ";" + formatPrecision + ")";
                        }
                    }
                }
            }

            return editorText;
        }


        #endregion

        /// <summary> Выгрузка данных в Excel </summary>
        /// <remarks>
        /// Убирает все наимнования областей, удаляет ненужные области
        /// </remarks>
        public void ExportExcel(string fileName, bool isSystem)
        {
            this.DisableColumnRowSizing = true;
            
            this.ShowRangeByPrefix(EnumDyDocSsRanges.Header
                                   | EnumDyDocSsRanges.Footer);
            
            var ms = new MemoryStream();
            this.DocumentInfo.Wb.SaveDocument(ms, DocumentFormat.OpenXml);
            ms.Position = 0;
            var wbCopy = new Workbook();
            wbCopy.LoadDocument(ms, DocumentFormat.OpenXml);
            wbCopy.BeginUpdate();

            if (!isSystem)
            {
                foreach (var rangeNamePrefix in new[]
                {
                    "ForDelete" , "InternalData" ,
                    EnumDyDocSsRanges.RowBind + "_",
                    EnumDyDocSsRanges.SysColumnNames + "_",
                })
                {
                    foreach (var namedRange in wbCopy.DefinedNames.Where(x => x.Name == rangeNamePrefix))
                    {
                        if (namedRange.Range.IsRangeFullCol())
                            namedRange.Range.Delete(DeleteMode.EntireColumn);
                        else if (namedRange.Range.IsRangeFullRow())
                            namedRange.Range.Delete(DeleteMode.EntireRow);

                        wbCopy.DefinedNames.Remove(namedRange);
                    }
                }

                foreach (var namedRange in wbCopy.DefinedNames.ToList())
                    wbCopy.DefinedNames.Remove(namedRange);
            }
            else
            {
                foreach (var worksheet in wbCopy.Worksheets)
                    worksheet.Unprotect("1");
            }
            wbCopy.EndUpdate();

            wbCopy.SaveDocument(fileName, DocumentFormat.OpenXml);

            this.DisableColumnRowSizing = false;
        }
    }
}