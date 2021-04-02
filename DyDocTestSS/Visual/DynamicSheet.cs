using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using DevExpress.Utils;
using DevExpress.Utils.Controls;
using DevExpress.Utils.Menu;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraSpreadsheet.Services;
using DyDocTestSS.Bl;
using DyDocTestSS.Domain;
using DyDocTestSS.DyTemplates;
using JetBrains.Annotations;
using Sphaera.Bp.Services.Core;
using Sphaera.Bp.Services.Math;
using Cell = DevExpress.Spreadsheet.Cell;
using Worksheet = DevExpress.Spreadsheet.Worksheet;

namespace DyDocTestSS.Visual
{
    public partial class DynamicSheet : UserControl
    {
        public DynamicSheet()
        {
            InitializeComponent();
            this.ssControl.Options.Behavior.Column.Changed += ColumnOnChanged;

            this.MenuItemAddColumn = new DXMenuItem("Добавить колонку", this.MenuClickAddColumn);
            this.MenuItemRemoveColumn = new DXMenuItem("Удалить колонку", this.MenuClickRemoveColumn);
        }

        private void ColumnOnChanged(object sender, BaseOptionChangedEventArgs e)
        {
            //throw new NotImplementedException();
        }


        /// <summary> Хитрый Форматировщик </summary>
        [CanBeNull] 
        private DynamicSheetController _dynamicSheetController;
        
        /// <summary> Хитрый Форматировщик </summary>
        [NotNull]
        public DynamicSheetController DynamicSheetController { get { return this._dynamicSheetController.NotNull("this._dynamicSheetController != null"); } }

        public void TestLoadSheet()
        {
            var templateManager = new DyDocumentManager();
            var doc = templateManager.GetDocument(
                new DocToPbs {TopFullSprKey = "320.0105.9050090012.121.211.003.30" }, 
                new DocToPbs { TopFullSprKey = "320.0105.9050090012.121.211.003.30" }, 
                this.ssControl.Document);

            var msgSrv = this.ssControl.GetService<IMessageBoxService>();
            this.ssControl.ReplaceService((IMessageBoxService)new MessageBoxService(msgSrv));

//            this.ssControl.AddService(typeof(ICustomCalculationService), new DySsFormulaEngine());

            this._dynamicSheetController = new DynamicSheetController(doc);
            this.DynamicSheetController.FormatForShow();
            this.DynamicSheetController.ProtectByColor();
            
            this.FillInternalCache();
        }

        /// <summary> Кеш колонок, которые мы скрыли </summary>
        // ReSharper disable once InconsistentNaming
        private readonly HashSet<Tuple<string, int>> CacheZeroColumns = new HashSet<Tuple<string, int>>();
        
        /// <summary> Кеш строк, которые мы скрыли </summary>
        // ReSharper disable once InconsistentNaming
        private readonly HashSet<Tuple<string, int>> CacheZeroRows = new HashSet<Tuple<string, int>>();

        /// <summary> Заполнение кеша скрытых колонок </summary>
        private void FillInternalCache()
        {
            this.CacheZeroColumns.Clear();
            this.CacheZeroRows.Clear();

            foreach (var sheet in ssControl.Document.Worksheets)
            {
                for (var c = 0; c < 1000; c++)
                {
                    if (SphaeraMath.IsEqual(sheet.Columns[c].Width, 0))
                        this.CacheZeroColumns.Add(Tuple.Create(sheet.Name, c));
                }
                for (var r = 0; r < 1000; r++)
                {
                    if (SphaeraMath.IsEqual(sheet.Rows[r].Height, 0))
                        this.CacheZeroRows.Add(Tuple.Create(sheet.Name, r));
                }
            }
        }

        private void ssControl_ContentChanged(object sender, EventArgs e)
        {
            if (this._dynamicSheetController == null)
                return;

            // Зануляем ширину скрытых столбцов и строк
            if (this.DynamicSheetController.DisableColumnRowSizing)
                return;

            var activeNm = this.ssControl.ActiveSheet.Name;
            var ws = this.ssControl.ActiveSheet as Worksheet;

            if (ws == null)
                return;

            foreach (var zeroColumn in this.CacheZeroColumns.Where(x => x.Item1 == activeNm))
                ws.Columns[zeroColumn.Item2].Width = 0;

            foreach (var zeroRow in this.CacheZeroRows.Where(x => x.Item1 == activeNm))
                ws.Rows[zeroRow.Item2].Height = 0;
        }

        private void ssControl_DocumentLoaded(object sender, EventArgs e)
        {
        }


        private void ssControl_CellValueChanged(object sender, SpreadsheetCellEventArgs e)
        {
            if (e.Cell.Protection.Locked == false)
            {
                // При копипасте часто слетает backcolor. Восстанавливаем
               // e.Cell.Fill.BackgroundColor = Color.Yellow;
            }
        }

        private void ssControl_CustomDrawCell(object sender, CustomDrawCellEventArgs e)
        {
            
        }

        private void ssControl_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (this.ssControl.GetSelectedRanges().Count != 1)
                throw new NotSupportedException("Что-то пошло не так");
        }

        

        /// <summary> Добавить колонку </summary>
        private DXMenuItem MenuItemAddColumn { get; set; }
        
        /// <summary> Добавить колонку </summary>
        private void MenuClickAddColumn(object sender, EventArgs e)
        {
            var selectedRanges = this.ssControl.GetSelectedRanges();
            if (selectedRanges.Count != 1 || selectedRanges[0].LeftColumnIndex != selectedRanges[0].RightColumnIndex)
                return;
            
            var sht = this.DynamicSheetController.DocumentInfo.SystemSheetInfos
                                .SingleOrDefault(x => x.Sheet == this.ssControl.ActiveSheet);
            if (sht == null)
                return;

            var addParams = new DyDocSs.DyDocSsSheetInfo.AddColumnParam();
            addParams.Caption = "Пользовательская колонка " + DateTime.Now.Ticks;
            addParams.Precision = 2;

            sht.AddColumnBefore(selectedRanges[0].LeftColumnIndex, addParams);
        }

        /// <summary> Удалить колонку </summary>
        private DXMenuItem MenuItemRemoveColumn { get; set; }

        /// <summary> Удалить колонку </summary>
        private void MenuClickRemoveColumn(object sender, EventArgs e)
        {
            var selectedRanges = this.ssControl.GetSelectedRanges();
            if (selectedRanges.Count != 1 || selectedRanges[0].LeftColumnIndex != selectedRanges[0].RightColumnIndex)
                return;

            var sht = this.DynamicSheetController.DocumentInfo.SystemSheetInfos
                .SingleOrDefault(x => x.Sheet == this.ssControl.ActiveSheet);
            if (sht == null)
                return;
            
            sht.RemoveColumn(selectedRanges[0].LeftColumnIndex);
        }

        private void ssControl_PopupMenuShowing(object sender, PopupMenuShowingEventArgs e)
        {
            if (e.MenuType == SpreadsheetMenuType.ColumnHeading)
            {
                foreach (var menuItem in e.Menu.Items.ToList())
                    e.Menu.Items.Remove(menuItem);

                var selectedRanges = this.ssControl.GetSelectedRanges();
                if (selectedRanges.Count != 1 || selectedRanges[0].LeftColumnIndex != selectedRanges[0].RightColumnIndex)
                    return;

                var sheetInfos = this.DynamicSheetController
                    .DocumentInfo.SystemSheetInfos
                    .Where(x => x.Sheet == this.ssControl.ActiveSheet)
                    .ToList();

                if (sheetInfos.Count != 1)
                    throw new NotSupportedException("Ошибка поиска листа");

                var sheetInfo = sheetInfos.Single();
                var firstInsertRange = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.FirstInsertColumn);

                if (firstInsertRange != null &&
                    firstInsertRange.Range.LeftColumnIndex <= selectedRanges[0].LeftColumnIndex &&
                    firstInsertRange.Range.RightColumnIndex >= selectedRanges[0].RightColumnIndex
                )
                {
                    e.Menu.Items.Add(this.MenuItemAddColumn);
                    if (firstInsertRange.Range.RightColumnIndex != selectedRanges[0].RightColumnIndex)
                        e.Menu.Items.Add(this.MenuItemRemoveColumn);
                }

            }
            else if (e.MenuType == SpreadsheetMenuType.RowHeading)
            {
                foreach (var menuItem in e.Menu.Items.ToList())
                    e.Menu.Items.Remove(menuItem);
            }
        }

        private void ssControl_RowsInserting(object sender, RowsChangingEventArgs e)
        {

        }

        private void ssControl_ColumnsInserted(object sender, ColumnsChangedEventArgs e)
        {
            this.FillInternalCache();
        }

        private void ssControl_ColumnsInserting(object sender, ColumnsChangingEventArgs e)
        {

        }

        private void ssControl_ColumnsRemoved(object sender, ColumnsChangedEventArgs e)
        {
            this.FillInternalCache();
        }

        private void ssControl_ColumnsRemoving(object sender, ColumnsChangingEventArgs e)
        {

        }


        private void ssControl_CellEndEdit(object sender, SpreadsheetCellValidatingEventArgs e)
        {
            e.EditorText = this.DynamicSheetController.UpdateEditData(e.EditorText, e.Cell);

            
            var sheetInfo = this.DynamicSheetController.DocumentInfo.SystemSheetInfos
                            .SingleOrDefault(x => x.Sheet == e.Cell.Worksheet)
                                            ?? throw new NotSupportedException("Ошибка получения листа");

            var mainFormulaRange = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.MainFormula);
            if (mainFormulaRange != null &&
                mainFormulaRange.Range.Contains(e.Cell))
            {
                throw new NotSupportedException("Не стали пока делать работу с основной формулой");
                //this.DynamicSheetController.TryUpdateFormula(e.EditorText, e.Cell);
            }

        }

        private void ssControl_ClipboardDataPasted(object sender, ClipboardDataPastedEventArgs e)
        {
            // Переокругляем после вставки
            var sheet = this.ssControl.ActiveSheet as Worksheet;
            if (sheet != null)
            {
                var range = sheet.Range.Parse(e.TargetRange);
                foreach (var cell in range.ExistingCells)
                    this.DynamicSheetController.UpdateEditCopyPasteData(cell);
            }
        }

        private void ssControl_ClipboardDataObtained(object sender, ClipboardDataObtainedEventArgs e)
        {
            // При вставке берём только значения
            e.Flags = PasteSpecial.Values;
        }

        private void ssControl_CopiedRangePasting(object sender, CopiedRangePastingEventArgs e)
        {
            var aa = e.TargetRange;

            var selectedCells = this.ssControl.GetSelectedRanges();
            if (selectedCells.Count == 1
                && selectedCells[0].LeftColumnIndex == e.TargetRange.LeftColumnIndex
                && selectedCells[0].RightColumnIndex == e.TargetRange.RightColumnIndex
                && selectedCells[0].TopRowIndex <= e.TargetRange.TopRowIndex
                && selectedCells[0].BottomRowIndex >= e.TargetRange.BottomRowIndex
            )
            {
                e.PasteSpecialFlags = PasteSpecial.Values | PasteSpecial.Formulas;
            }
            else
            {
                // При вставке берём только значения
                e.PasteSpecialFlags = PasteSpecial.Values;
            }

        }

        private void ssControl_CustomDrawCellBackground(object sender, CustomDrawCellBackgroundEventArgs e)
        {
            if (this._dynamicSheetController == null)
                return;

            e.Handled = this.DynamicSheetController.CustomDrawBackgroundFormatting(e.Cell, e.Graphics, e.Bounds);
        }

        public void ExportExcel([NotNull] string fileName, bool isSystem = false)
        {
            this.DynamicSheetController.ExportExcel(fileName, isSystem);
        }

        private void ssControl_CellBeginEdit(object sender, SpreadsheetCellCancelEventArgs e)
        {

        }

        private void ssControl_CustomCellEdit(object sender, SpreadsheetCustomCellEditEventArgs e)
        {

        }

        private void ssControl_RangeCopied(object sender, RangeCopiedEventArgs e)
        {

        }

        private void ssControl_RangeCopying(object sender, RangeCopyingEventArgs e)
        {

        }

        private void ssControl_ClipboardDataPasting(object sender, EventArgs e)
        {
        }

        private void ssControl_CopiedRangePasted(object sender, CopiedRangePastedEventArgs e)
        {

        }

        private void ssControlToolTip_GetActiveObjectInfo(object sender, DevExpress.Utils.ToolTipControllerGetActiveObjectInfoEventArgs e)
        {
            if (this._dynamicSheetController == null)
                return;
            
            var cell = this.ssControl.GetCellFromPoint(e.ControlMousePosition);
            if (cell != null)
                e.Info = this.DynamicSheetController.GetCellToolTip(cell);
        }
    }
}
