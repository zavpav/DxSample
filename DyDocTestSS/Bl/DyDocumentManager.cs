using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraSpreadsheet.Services;
using DyDocTestSS.Bl.Aksiok;
using DyDocTestSS.Bl.DataGetters;
using DyDocTestSS.Domain;
using DyDocTestSS.DyTemplates;
using JetBrains.Annotations;
using SdBp.Domain.Spr;
using Sphaera.Bp.Bl.Excel;
using Sphaera.Bp.Services.Core;
using Sphaera.Bp.Services.Log;

namespace DyDocTestSS.Bl
{
    public class DyDocumentManager
    {
        private IAksiokStorage AksiokStorage { get; set; }

        private IDataCalculator DataCalculator { get; set; }

        public DyDocumentManager()
        {
            this.Logger = new LoggerProxyImp("aa");
            this.AksiokStorage = new AksiokStorage();
            this.DataCalculator = new DataCalculator();
        }

        [NotNull]
        private ILoggerProxy Logger { get; set; }

        /// <summary> Информация для создания расчетов </summary>
        private class DyTemplateCreateInfo
        {
            public DyTemplateCreateInfo([NotNull] DocToPbs doc)
            {
                this.Doc = doc;
                this.NamedParams = new List<Parameter>();
                this.TemplateName = "";
            }

            /// <summary> Загружена ли книга полностью (она может быть загружен, когда получаем данные из базы) </summary>
            public bool IsWorkbookLoaded { get; set; }
            
            /// <summary> Имя шаблона для загрузки </summary>
            [NotNull]
            public string TemplateName { get; set; }
            
            /// <summary> Обрабатываемый документ </summary>
            [NotNull]
            public DocToPbs Doc { get; set; }

            /// <summary> Предыдущий документ </summary>
            [CanBeNull]
            public DocToPbs PrevDoc { get; set; }
            
            /// <summary> Именованные параметры </summary>
            [NotNull]
            public List<Parameter> NamedParams { get; set; }

            public class Parameter
            {
                public string Name { get; set; }
                public string Data { get; set; }
            }
        }

        /// <summary> Получить листы расчетов </summary>
        /// <param name="docToPbs">Документ, для которого получаем данные</param>
        /// <param name="prevDocPbs">"Предыдущий" докумнет из которого тянем пользовательские данные</param>
        /// <param name="wbData">Книга отображения (что б корректно были привязаны ref, надо получать wb из контрола)</param>
        [NotNull]
        public DyDocSs GetDocument([NotNull] DocToPbs docToPbs, [CanBeNull] DocToPbs prevDocPbs, [CanBeNull] IWorkbook wbData)
        {
            var defineXlsTemplateInfo = this.GetTemplateInfo(docToPbs);
            defineXlsTemplateInfo.PrevDoc = prevDocPbs;
            return this.CreateDyDocByCreateInfo(defineXlsTemplateInfo, wbData);
        }

        /// <summary> Создать документ на основе данных для создания </summary>
        [NotNull]
        private DyDocSs CreateDyDocByCreateInfo([NotNull] DyTemplateCreateInfo createInfo, [CanBeNull] IWorkbook wbData)
        {
            var swParts = new StopwatchParts("DyDocCreate");
            try
            {
                swParts.NewPart("Загрузка " + createInfo.TemplateName);
                var wb = wbData ?? new Workbook();

                if (!createInfo.IsWorkbookLoaded)
                {
                    wb.LoadDocument(createInfo.TemplateName);
                    wb.DocumentSettings.Calculation.PrecisionAsDisplayed = true;
                }
                wb.DocumentSettings.Calculation.Mode = CalculationMode.Manual;

                swParts.NewPart("Проверка структуры шаблона и создание документа");
                var dyDoc = this.CreateDocumentFromTemplate(wb);
                wb.AddService(typeof(ICustomCalculationService), new DySsFormulaEngine(dyDoc));
                
                if (!createInfo.IsWorkbookLoaded)
                {
                    swParts.NewPart("Полное создание структуры");
                    this.FullStructCreate(dyDoc, createInfo);

                    swParts.NewPart("Заполнение данными");
                    this.FillStaticData(dyDoc, createInfo);
                }

                swParts.NewPart("Синхронизация старых пользовательских данных");
                if (createInfo.PrevDoc != null)
                    this.FillDocumentUser(dyDoc, createInfo.PrevDoc);

                swParts.NewPart("Дубликат в шаблон");

                wb.CalculateFullRebuild();
                wb.DocumentSettings.Calculation.Mode = CalculationMode.Automatic;

                // Докопируем что получилось в TemplateInfo
                var ms = new MemoryStream();
                dyDoc.Wb.SaveDocument(ms, DocumentFormat.OpenXml);
                var wbDub = new Workbook();
                ms.Position = 0;
                wbDub.LoadDocument(ms, DocumentFormat.OpenXml);
                dyDoc.WbTemplateInfo = wbDub;

                swParts.StopAndLogging();

                return dyDoc;
            }
            catch (Exception e)
            {
                swParts.StopAndLogging();
                this.Logger.Error(e, "Ошибка создания динамического документа");
                throw;
            }

        }

        #region Поиск и обработка xml шаблонов 

        /// <summary> Получить информацию по поиску и использованию шаблона </summary>
        /// <param name="docToPbs">Почти загруженный документ (возможно, без информации о расчётных листах)</param>
        [NotNull]
        private DyTemplateCreateInfo GetTemplateInfo([NotNull] DocToPbs docToPbs)
        {
            var xDoc = XDocument.Load("DyDocBinds2020.xml");

            var xRootTemplates = xDoc.Element("BpTemplates") ?? throw new NotSupportedException("Ненайден блок BpTemplates");
            var xBinds = xRootTemplates.Element("Binds") ?? throw new NotSupportedException("Ненайден блок Binds");
            var xTemplateBind = this.FindXElementByFilter(xBinds, docToPbs.TopFullSprKey);
            
            var createInfo = new DyTemplateCreateInfo(docToPbs)
            {
                TemplateName = (xTemplateBind.Element("Template") ?? throw new NotSupportedException("Ненайден блок Template для " + xTemplateBind )).Value
            };

            var xParameters = xTemplateBind.Element("ColumnDataBindParameter");
            if (xParameters != null)
            {
                var xBindParam = this.FindXElementByFilter(xParameters, docToPbs.TopFullSprKey);

                foreach (var xParm in xBindParam.Elements("Parameter"))
                {
                    createInfo.NamedParams.Add(
                        new DyTemplateCreateInfo.Parameter
                        {
                            Name = (xParm.Attribute("name") ?? throw new NotSupportedException("Для параметра не определен блок name " + xParm)).Value,
                            Data = (xParm.Element("Data") ?? throw new NotSupportedException("Для параметра не определен блок Data " + xParm)).Value
                        }
                    );
                }
            }

            return createInfo;
        }

        /// <summary> Поиск узла удовлетворяющему фильтру </summary>
        /// <param name="xContainer">Рутовый элемент xml-контейнера</param>
        /// <param name="fsk">Искомый код</param>
        /// <remarks>
        /// Ищет в контейнере элементы Bind и в них Filter. 
        /// </remarks>
        /// <returns>Элемент Bind</returns>
        [NotNull]
        private XElement FindXElementByFilter([NotNull] XElement xContainer, [NotNull] string fsk)
        {
            var prpList = new List<Tuple<ITemplateFilter, XElement>>();
            foreach (var xBind in xContainer.Elements("Bind"))
            {
                var xFilter = xBind.Element("Filter");
                if (xFilter == null)
                {
                    this.Logger.Error("Ненайден блок Filter в Bind " + xBind);
                    throw new NotSupportedException("Ошибка разбора блока Filter. " + xBind);
                }

                var bind = BindFilter.CreateFromXml(xFilter);
                prpList.Add(Tuple.Create((ITemplateFilter)bind, xBind));
            }

            var possibleBinds = prpList.Where(x => x.Item1.IsSuitable(fsk)).ToList();
            if (possibleBinds.Count == 0)
            {
                this.Logger.Error("Ненайден блок Filter в Bind для " + fsk);
                throw new NotSupportedException("Ненайден блок Filter в Bind для " + fsk);
            }

            if (possibleBinds.Count > 1)
            {
                this.Logger.Error("Найдено слишком много блоков Filter в Bind для " + fsk);
                throw new NotSupportedException("Найдено слишком много блоков Filter в Bind для " + fsk);
            }

            return possibleBinds.Single().Item2;
        }

        #endregion

        #region Создание и заполнение шаблона

        /// <summary> Создание обвязки расчётных листов и проверка шаблона на наличие необходимых данных в шаблоне</summary>
        [NotNull]
        private DyDocSs CreateDocumentFromTemplate([NotNull] IWorkbook wb)
        {
            var dyDoc = new DyDocSs(wb);

            var systemNames = wb.DefinedNames.Where(x => x.Name.StartsWith("System_")).ToList();
            if (systemNames.Count == 0)
                throw new NotSupportedException("Ненайдено системный описаний листов");

            this.Logger.Info("-> Системных листов {SysCnt} всего листов {TtlCnt}", systemNames.Count, wb.Worksheets.Count);

            // Заполняем системную информацию по листам
            var anyErr = false;
            foreach (var worksheet in systemNames.Select(x => x.Range.Worksheet))
            {
                var sheetInfo = new DyDocSs.DyDocSsSheetInfo(worksheet);
                dyDoc.SystemSheetInfos.Add(sheetInfo);

                foreach (var definedName in wb.DefinedNames.Where(x => x.Name.EndsWith("_" + sheetInfo.PostfixName)))
                {
                    if (definedName.Range.Worksheet != sheetInfo.Sheet)
                    {
                        anyErr = true;
                        this.Logger.Error("-> Область " + definedName.Name + " не принадлежит листу " + sheetInfo.Sheet.Name);
                    }

                }

                anyErr |= this.CheckRequireRange(sheetInfo, EnumDyDocSsRanges.Data);
                anyErr |= this.CheckRequireRange(sheetInfo, EnumDyDocSsRanges.System);
                anyErr |= this.CheckRequireRange(sheetInfo, EnumDyDocSsRanges.DataPbsCode);
                anyErr |= this.CheckRequireRange(sheetInfo, EnumDyDocSsRanges.DataPbsName);
                anyErr |= this.CheckRequireRange(sheetInfo, EnumDyDocSsRanges.SysColumnNames);

                anyErr |= this.CheckRegionInclusion(sheetInfo, sheetInfo.DataRange, EnumDyDocSsRanges.DataPbsName);
                anyErr |= this.CheckRegionInclusion(sheetInfo, sheetInfo.DataRange, EnumDyDocSsRanges.DataPbsCode);
                anyErr |= this.CheckRegionInclusion(sheetInfo, sheetInfo.DataRange, EnumDyDocSsRanges.FirstInsertColumn);
                anyErr |= this.CheckRegionInclusion(sheetInfo, sheetInfo.DataRange, EnumDyDocSsRanges.TotalYearSum);

                anyErr |= this.CheckRegionTotalYearSum(sheetInfo);

                anyErr |= this.CheckRegionOutcomeAdditionalBinding(sheetInfo, EnumDyDocSsRanges.Residue);
                anyErr |= this.CheckRegionOutcomeAdditionalBinding(sheetInfo, EnumDyDocSsRanges.ForDistrib);
                anyErr |= this.CheckRegionOutcomeAdditionalBinding(sheetInfo, EnumDyDocSsRanges.MainFormula);

                anyErr |= this.CheckSysColumnNames(sheetInfo);

                anyErr |= this.ParseRowBind(sheetInfo);

                anyErr |= this.CheckCoeffRegion(sheetInfo);

                anyErr |= this.FillSysInfo(sheetInfo);
            }

            if (anyErr)
                throw new NotSupportedException("Возникла какая-то ошибка. См. логи.");

            return dyDoc;
        }

        /// <summary> Полное создание структуры документы (добавляем динамические строки и колонки) </summary>
        private void FullStructCreate([NotNull] DyDocSs dyDoc, [NotNull] DyTemplateCreateInfo createInfo)
        {
            var year = 2020;
            var parsedFsk = new ParsedFsk2019(createInfo.Doc.TopFullSprKey);

            var staticFields = new Dictionary<string, string>
            {
                {"ТекущийГод", year.ToString()},
                {"ПредыдущийГод", (year-1).ToString() },
                {"ОчереднойГод", (year+1).ToString() },
                {"1ППГод", (year+1).ToString() },
                {"2ППГод", (year+2).ToString() },
                {"ПолныйКод", createInfo.Doc.TopFullSprKey },
                {"КбкГлаваКод",  parsedFsk.Grbs},
                {"КбкРзПрзКод",  parsedFsk.RzPrz},
                {"КбкРзПрзНаименование",  "Наименование " + parsedFsk.RzPrz},
                {"КбкЦСРКод",  parsedFsk.Csr},
                {"КбкЦСРНаименование",  "Наименование " + parsedFsk.Csr},
                {"КбкВРКод",  parsedFsk.Vr},
                {"КбкВРНаименование",  "Наименование " + parsedFsk.Vr},
                {"КбкКОСГУКод",  parsedFsk.Kosgu},
                {"КбкКОСГУНаименование",  "Наименование " + parsedFsk.Kosgu},
                {"КбкНРКод",  parsedFsk.Nr},
                {"КбкНРНаименование",  "Наименование " + parsedFsk.Nr},
                {"КбкСФРКод",  parsedFsk.Sfr},
                {"КбкСФРНаименование",  "Наименование " + parsedFsk.Sfr},
            };

            dyDoc.Wb.ReplaceStaticFields(staticFields);

            var pbsList = this.GetPbsList(createInfo);

            foreach (var sheetInfo in dyDoc.SystemSheetInfos)
            {
                // Обработка динамических строк
                var pbsCodeRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.DataPbsCode);
                if (sheetInfo.RowType == DyDocSs.DyDocSsSheetInfo.EnumSheetRowType.Dynamic)
                {
                    // Обработка динамических колонок
                    var dataBindRow = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.RowBind);
                    if (dataBindRow != null)
                    {
                        var colIdx = 0;
                        while (sheetInfo.DataRange.Range.RightColumnIndex >= colIdx) //sheetInfo.DataRange.Range плывёт, поэтому только так
                        {
                            var bindingCell = dataBindRow.Range.FromLTRB(sheetInfo.DataRange.Range.LeftColumnIndex + colIdx, 0,
                                    sheetInfo.DataRange.Range.LeftColumnIndex + colIdx, 0);

                            // Тут нас интересуют только данные АКСИОКа, так как могут быть множественные колонки
                            var columnBindInfo = this.GetColumnDataBindInfo(bindingCell, createInfo);

                            if (!columnBindInfo.Columns.Any())
                                throw new NotSupportedException("Не смогли разобрать параметры " + bindingCell.Value);

                            if (columnBindInfo.IsGroup && columnBindInfo.Columns.All(x => x is ColumnDataBindInfoGetterAksiok))
                            {
                                // В параметрах у нас список необходимых параметров. Если там null тогда надо "всё".
                                var aksiokDatas = columnBindInfo.Columns
                                                .OfType<ColumnDataBindInfoGetterAksiok>()
                                                .SelectMany(this.AksiokStorage.GetAksiokData).ToList();
                                if (aksiokDatas.Count == 0)
                                {
                                    sheetInfo.Sheet.DeleteCells(bindingCell, DeleteMode.EntireColumn);
                                }
                                else
                                {
                                    foreach (var aksiokDat in aksiokDatas)
                                    {
                                        var addParams = new DyDocSs.DyDocSsSheetInfo.AddColumnParam();
                                        addParams.Caption = aksiokDat.ParamName;
                                        var tmpBind = aksiokDat.CreateColumnDataBindInfoAksiok();
                                        addParams.ColumnBinding = tmpBind.ToStringBind();
                                        addParams.IsEditable = false;
                                        addParams.SystemName = tmpBind.ToStringBind();
                                        addParams.Precision = 2;
                                        sheetInfo.AddColumnBefore(sheetInfo.DataRange.Range.LeftColumnIndex + colIdx, addParams);
                                        colIdx++;
                                    }
                                    bindingCell = dataBindRow.Range.FromLTRB(sheetInfo.DataRange.Range.LeftColumnIndex + colIdx, 0,
                                            sheetInfo.DataRange.Range.LeftColumnIndex + colIdx, 0);
                                    sheetInfo.Sheet.DeleteCells(bindingCell, DeleteMode.EntireColumn);
                                    colIdx--;
                                }
                            }
                            else if (columnBindInfo.IsGroup)
                                throw new NotSupportedException("Пока множественные колонки возможны только для АКСИОКа");

                            colIdx++;
                        }
                    }

                    var pbsNameRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.DataPbsName);
                    var dataRange = sheetInfo.DataRange;
                    foreach (var pbs in pbsList.Select((x, i) => new {Idx = i, Pbs = x}))
                    {
                        Range rowRange;
                        if (pbs.Idx == 0)
                        {
                            rowRange = dataRange.Range.FromLTRB(0, 0, dataRange.Range.ColumnCount-1, 0);
                        }
                        else
                        {
                            var range = dataRange.Range;
                            rowRange = sheetInfo.Sheet.Range[(range.BottomRowIndex + 1) + ":" + (range.BottomRowIndex + 1)];
                            sheetInfo.Sheet.InsertCells(rowRange, InsertCellsMode.ShiftCellsDown);
                            var fromRowRange = sheetInfo.Sheet.Range[(range.BottomRowIndex + 2) + ":" + (range.BottomRowIndex + 2)];
                            rowRange.CopyFrom(fromRowRange, PasteSpecial.All);
                        }

                        var pbsCell = sheetInfo.Sheet.Range.FromLTRB(pbsCodeRange.Range.LeftColumnIndex, rowRange.TopRowIndex,
                            pbsCodeRange.Range.LeftColumnIndex, rowRange.BottomRowIndex);
                        pbsCell.Value = pbs.Pbs.Item1;

                        var nameCell = sheetInfo.Sheet.Range.FromLTRB(pbsNameRange.Range.LeftColumnIndex, rowRange.TopRowIndex,
                            pbsNameRange.Range.LeftColumnIndex, rowRange.BottomRowIndex);
                        nameCell.Value = pbs.Pbs.Item2;

                        this.FullStructCreateFillDynamicException(sheetInfo, rowRange, pbs.Pbs.Item1, pbsCodeRange, dataBindRow);
                    }
                    
                    // Подчищаем пустую строку
                    var delRange = sheetInfo.Sheet.Range[(dataRange.Range.BottomRowIndex + 1) + ":" + (dataRange.Range.BottomRowIndex + 1)];
                    sheetInfo.Sheet.DeleteCells(delRange, DeleteMode.ShiftCellsUp);
                }

                this.ValidatePbsCodes();

            }

        }


        /// <summary> Заполнение исключений в шаблонах с динамическими строчками </summary>
        /// <param name="sheetInfo"></param>
        /// <param name="rowRange">Заполняемая строка</param>
        /// <param name="pbsCode">Код ПБС</param>
        /// <param name="pbsCodeRange"></param>
        /// <param name="dataBindRow">Область данных с биндингом строк для поиска исключений</param>
        private void FullStructCreateFillDynamicException(
                            [NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo, 
                            [NotNull] Range rowRange,
                            [NotNull] string pbsCode,
                            [NotNull] DefinedName pbsCodeRange, 
                            [CanBeNull] DefinedName dataBindRow)
        {
            if (dataBindRow == null || dataBindRow.Range.RowCount < 2)
                return; // Нет описания биндинга или нет исключений

            Row exceptionBindingRow = null;

            for (var rowBindIdx = 1; rowBindIdx < dataBindRow.Range.RowCount; rowBindIdx++)
            {
                var pbsExceptionCell = sheetInfo.Sheet[dataBindRow.Range.TopRowIndex + rowBindIdx, pbsCodeRange.Range.LeftColumnIndex];
                if (pbsExceptionCell.Value.ToString() == pbsCode)
                {
                    exceptionBindingRow = sheetInfo.Sheet.Rows[dataBindRow.Range.TopRowIndex + rowBindIdx];
                    break;
                }
            }

            if (exceptionBindingRow == null)
                return;

            for (var dataColumnIdx = 0; dataColumnIdx < rowRange.ColumnCount; dataColumnIdx++)
            {
                // Коды ПБС игнорим
                if (dataBindRow.Range.LeftColumnIndex + dataColumnIdx == pbsCodeRange.Range.LeftColumnIndex)
                    continue;

                var exceptionBindingCell = exceptionBindingRow[0, dataBindRow.Range.LeftColumnIndex + dataColumnIdx];
                var exceptionVal = exceptionBindingCell.FormulaInvariant;
                if (string.IsNullOrWhiteSpace(exceptionVal))
                    exceptionVal = exceptionBindingCell.Value.ToString();
                if (string.IsNullOrWhiteSpace(exceptionVal))
                    continue;

                var cell = sheetInfo.Sheet[rowRange.TopRowIndex, dataBindRow.Range.LeftColumnIndex + dataColumnIdx];
                if (exceptionVal.Trim('\'').StartsWith("="))
                {
                    var frm = exceptionVal.Trim('\'');
                    var invFrm = sheetInfo.GetCellFormulaInvariantCalc(exceptionBindingCell, frm);
                    var newFrm = sheetInfo.GetCellFormulaOriginalByInvariantFormulaCalc(cell, invFrm);

                    cell.FormulaInvariant = "=" + newFrm;


                }
                else
                    cell.Value = exceptionVal;

            }
        }

        /// <summary> Поиск биндинга в строке биндинга </summary>
        [NotNull]
        private ColumnDataBind GetBindingFromDataBinding([CanBeNull] DefinedName dataBindRow, 
                    [NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo,
                    int cI,
                    [NotNull] DyTemplateCreateInfo createInfo
            )
        {
            // Если не распарсено, значит лезем в binding
            if (dataBindRow == null)
                throw new NotSupportedException("Кривая область привязки данных RowBind");

            var bindingCell = dataBindRow.Range[0, sheetInfo.DataRange.Range.LeftColumnIndex + cI];
            // TODO тут можно вставить поиск исключений
            return this.GetColumnDataBindInfo(bindingCell, createInfo);
        }

        /// <summary> Собственно заполнение данными </summary>
        private void FillStaticData(DyDocSs dyDoc, DyTemplateCreateInfo defineXlsTemplateInfo)
        {
            dyDoc.Wb.BeginUpdate();
            foreach (var sheetInfo in dyDoc.SystemSheetInfos)
            {
                var pbsCodeRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.DataPbsCode);
                var pbsNameRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.DataPbsName);
                var dataBindRow = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.RowBind);
                var pbsCodes = new Dictionary<int, string>();
                for (var cI = 0; cI < sheetInfo.DataRange.Range.ColumnCount; cI++)
                {
                    for (var rI = 0; rI < sheetInfo.DataRange.Range.RowCount; rI++)
                    {
                        string pbsCode;
                        if (!pbsCodes.TryGetValue(rI, out pbsCode))
                        {
                            pbsCode = pbsCodeRange.Range[rI, 0].Value.ToString();
                            pbsCodes.Add(rI, pbsCode);
                        }

                        var cell = sheetInfo.DataRange.Range[rI, cI];

                        if (pbsCodeRange.Range.Contains(cell))
                            continue;
                        if (pbsNameRange.Range.Contains(cell))
                            continue;

                        ColumnDataBind bindingInfo;
                        var cellVal = cell.Value.ToString();
                        if (cellVal != "0" && !string.IsNullOrWhiteSpace(cellVal) && cell.Value.IsText)
                        {
                            try
                            {
                                bindingInfo = this.GetColumnDataBindInfo(cell, defineXlsTemplateInfo); // Если это "странный текст" - пытаемся его разобрать на получение данных
                            }
                            catch (Exception)
                            {
                                bindingInfo = this.GetBindingFromDataBinding(dataBindRow, sheetInfo, cI, defineXlsTemplateInfo);
                            }
                        }
                        else
                            bindingInfo = this.GetBindingFromDataBinding(dataBindRow, sheetInfo, cI, defineXlsTemplateInfo); // Если это не текст или не "стандартный какой-то", сразу лезем в биндниг

                        if (bindingInfo.Columns.All(x => x  is ColumnDataBindInfoGetterEmpty)) { }
                        else if (bindingInfo.Columns.All(x => x is ColumnDataBindInfoGetterUserEdit))
                        {
                            cell.Fill.BackgroundColor = DyDocSs.EditableColor;
                            cell.Value = 0;
                        }
                        else
                        {
                            var sm = this.DataCalculator.GetDecimal(pbsCode, bindingInfo);
                            cell.Value = sm;
                        }
                    }
                }
            }
            dyDoc.Wb.EndUpdate();
        }

        /// <summary> Копирование пользовательских данных для документа </summary>
        private void FillDocumentUser([NotNull] DyDocSs dyDoc, [NotNull] DocToPbs prevDocPbs)
        {

            return;

            var createinfo = new DyTemplateCreateInfo(prevDocPbs);
            createinfo.IsWorkbookLoaded = true;
            
            var wb = new Workbook();
            wb.LoadDocument("Откуда обновляем.xlsx");


            var prevDyDoc = this.CreateDyDocByCreateInfo(createinfo, wb);

            // Добавляем колонки из старого документа
            foreach (var sheetInfo in dyDoc.SystemSheetInfos)
            {
                var oldSheetInfo = prevDyDoc.SystemSheetInfos.SingleOrDefault(x => x.SubType == sheetInfo.SubType);
                
                if (oldSheetInfo == null)
                    continue; // Похоже, убрали ненужное

                var oldAddedColumnsRange = oldSheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.FirstInsertColumn);
                if (oldAddedColumnsRange == null)
                    continue;
                
                var newAddedColumnsRange = oldSheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.FirstInsertColumn);


                var oldColumnsHeaders = oldSheetInfo.GetDefinedRange(EnumDyDocSsRanges.ColumnHeaders);
                var oldSysNames = oldSheetInfo.GetDefinedRange(EnumDyDocSsRanges.SysColumnNames);
                var oldRowBinds = oldSheetInfo.GetDefinedRange(EnumDyDocSsRanges.RowBind);
                
                var newSysNames = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.SysColumnNames);


                // Смотрим колонки, которые могут быть пользовательскими
                for (var cIdx = oldAddedColumnsRange.Range.ColumnCount-1; cIdx >= 0; cIdx--)
                {
                    var oldColumnIndex = cIdx + oldAddedColumnsRange.Range.LeftColumnIndex;
                    var oldColumn = oldSheetInfo.Sheet.Columns[oldColumnIndex];
                    var oldSysName = oldColumn[oldSysNames.Range.TopRowIndex].Value.ToString();
                    if (newSysNames.Range.ExistingCells.Any(x => x.Value.ToString() == oldSysName))
                        continue; // Если это какая-то системная колонка в области пользовательских колонок - пропускаем

                    if (newAddedColumnsRange == null) // Раньше эту проверку нельзя делать, так как колонок на добавление может и не быть.
                        throw new NotSupportedException("Пропала область ввода пользовательских данных при их наличии");


                    // Собираем заголовок. Пока так.. потом можно подумать.
                    var caption = string.Join("\n", 
                        oldSheetInfo.Sheet.Range.FromLTRB2Absolute(oldColumnIndex,
                                        oldColumnsHeaders.Range.TopRowIndex, oldColumnIndex, oldColumnsHeaders.Range.BottomRowIndex)
                        .Where(x => !string.IsNullOrWhiteSpace(x.Value.ToString()))
                        .Select(x => x.Value.ToString()));

                    var addColumnInfo = new DyDocSs.DyDocSsSheetInfo.AddColumnParam
                    {
                        SystemName = oldSysName,
                        IsEditable = true,
                        Caption = caption,
                        ColumnBinding = oldSheetInfo.Sheet.Cells[oldRowBinds.Range.TopRowIndex, oldColumnIndex].Value.ToString()
                    };
                    var precision = oldSheetInfo.Sheet.Cells[oldSheetInfo.DataRange.Range.TopRowIndex, oldColumnIndex].GetPrecisionFromNumberFormat();
                    if (precision == null)
                        throw new NotSupportedException("Не смогли определить формат из формата для пользовательской колонки");
                    addColumnInfo.Precision = precision.Value;

                    sheetInfo.AddColumnBefore(newAddedColumnsRange.Range.LeftColumnIndex, addColumnInfo);
                    sheetInfo.Sheet.Columns[newAddedColumnsRange.Range.LeftColumnIndex].Width = oldColumn.Width;
                }

            }

            // Восстанавливаем формулы. (Только пользовательские данные) (Одновременно с колонками восстанавливать нельзя) (Все необходимые колонки должны быть в наличии)
            foreach (var oldSheetInfo in prevDyDoc.SystemSheetInfos)
            {
                var newSheetInfo = dyDoc.SystemSheetInfos.SingleOrDefault(x => x.SubType == oldSheetInfo.SubType);
                if (newSheetInfo == null)
                    continue; // Похоже, убрали ненужное

                var visitor = new UpdateFormulaVisitor(oldSheetInfo, newSheetInfo);

                foreach (var oldFormulaCell in oldSheetInfo.Sheet.GetExistingCells()
                        .Where(x => !string.IsNullOrWhiteSpace(x.Formula))
                        .Where(x => !x.Value.IsError)
                        .Where(dyDoc.IsEditableWorkbookCellCriteria)
                )
                {
                    var parsedFormula = oldFormulaCell.ParsedExpression;
                    visitor.InvalidRecoding = false;
                    parsedFormula.Expression.Visit(visitor);

                    Cell newFormulaCell;

                    var pbsSysNameCoord = oldSheetInfo.GetInvariantCoord(oldFormulaCell);
                    if (pbsSysNameCoord != null)
                    {   // Область данных
                        newFormulaCell = newSheetInfo.GetCellByInvariantCoord(pbsSysNameCoord.Value);
                    }
                    else
                    {
                        var coeffRegion = oldSheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
                        if (coeffRegion != null && coeffRegion.Range.ExistingCells.Any(x => x.Equals(oldFormulaCell)))
                        {
                            // Область коэффициентов (?)
                            this.Logger.Error("Пока не сделал обновление формулы области коэффициентов");
                            newFormulaCell = null;
                        }
                        else
                            throw new NotSupportedException("Формула НЕ области данных и НЕ области коэффициентов");
                    }

                    if (newFormulaCell != null)
                    {
                        if (visitor.InvalidRecoding)
                            newFormulaCell.Value = CellValue.ErrorReference;
                        else
                            newFormulaCell.Formula = parsedFormula.ToString();
                    }

                }
            }

        }

        private class UpdateFormulaVisitor : ExpressionVisitor
        {
            [NotNull] 
            private DyDocSs.DyDocSsSheetInfo OldSheetInfo { get; set; }
            
            [NotNull] 
            private DyDocSs.DyDocSsSheetInfo NewSheetInfo { get; set; }

            public UpdateFormulaVisitor([NotNull] DyDocSs.DyDocSsSheetInfo oldSheetInfo, [NotNull] DyDocSs.DyDocSsSheetInfo newSheetInfo)
            {
                this.OldSheetInfo = oldSheetInfo;
                this.NewSheetInfo = newSheetInfo;
                this.InvalidRecoding = false;
            }

            private bool RegionInclusion([CanBeNull] DefinedName rangeMain, [NotNull] CellReferencePosition referencePosition)
            {
                if (rangeMain == null)
                    return false;

                if (referencePosition.Row < rangeMain.Range.TopRowIndex ||
                    referencePosition.Row > rangeMain.Range.BottomRowIndex ||
                    referencePosition.Column < rangeMain.Range.LeftColumnIndex ||
                    referencePosition.Column > rangeMain.Range.RightColumnIndex
                )
                {
                    return false;
                }

                return true;
            }

            /// <summary> Проблемы декодирования. Надо ставить ошибку ссылок </summary>
            public bool InvalidRecoding { get; set; }

            public override void Visit(CellReferenceExpression expression)
            {
                var newPositionTopLeft = this.GetNewCell(expression.CellArea.TopLeft);
                var newPositionBottomRight = this.GetNewCell(expression.CellArea.BottomRight);
                if (newPositionTopLeft == null || newPositionBottomRight == null)
                {
                    this.InvalidRecoding = true;
                }
                else
                {
                    expression.CellArea.TopRowIndex = newPositionTopLeft.TopRowIndex;
                    expression.CellArea.LeftColumnIndex = newPositionTopLeft.LeftColumnIndex;

                    expression.CellArea.BottomRowIndex = newPositionBottomRight.TopRowIndex;
                    expression.CellArea.RightColumnIndex = newPositionBottomRight.LeftColumnIndex;
                }
            }


            [CanBeNull]
            private Cell GetNewCell(CellReferencePosition cellPosition)
            {
                var isCoeff = this.RegionInclusion(this.OldSheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.Coeff), cellPosition);
                var isData = this.RegionInclusion(this.OldSheetInfo.DataRange, cellPosition);

                if (isData || isCoeff)
                {
                    var oldCell = this.OldSheetInfo.Sheet[cellPosition.Row, cellPosition.Column];
                    
                    var invariantCoord = this.OldSheetInfo.GetInvariantCoord(oldCell);
                    if (invariantCoord == null)
                        throw new NotSupportedException("Неопределили координаты ячейки " + oldCell.GetReferenceA1());

                    var newCell = this.NewSheetInfo.GetCellByInvariantCoord(invariantCoord.Value);
                    return newCell;
                }

                // Вообще неизвестно что
                return null;
            }

            public override void Visit(RangeExpression expression)
            {
                base.Visit(expression);
            }

            public override void Visit(RangeUnionExpression expression)
            {
                base.Visit(expression);
            }

            public override void Visit(RangeIntersectionExpression expression)
            {
                base.Visit(expression);
            }


            public override void Visit(CellErrorReferenceExpression expression)
            {
                base.Visit(expression);
            }

            public override void Visit(DefinedNameReferenceExpression expression)
            {
                base.Visit(expression);
            }

            public override void Visit(TableReferenceExpression expression)
            {
                base.Visit(expression);
            }
        }

        #endregion

        #region Получение и обработка данных

        /// <summary> Получить описания получения данных для колонки по текстовому описанию </summary>
        [NotNull]
        private ColumnDataBind GetColumnDataBindInfo([NotNull] Range bindingCell, [NotNull] DyTemplateCreateInfo createInfo)
        {
            var bindText = bindingCell.Value.ToString();
            var bindArr = bindText.Split('\n', '|');
            if (bindArr.Length == 0)
                throw new NotSupportedException("Пустая строка разбора");

            var mainDataString = bindArr[0].Trim().ToUpper(CultureInfo.InvariantCulture);

            // Пока решил не разделять то что не требует явного получения данных
            if (mainDataString == "ПУСТАЯ СТРОКА" || mainDataString == "ПУСТО")
                return new ColumnDataBind(new ColumnDataBindInfoGetterEmpty()); ;

            if (mainDataString == "РУЧНОЙ ВВОД")
                return new ColumnDataBind(new ColumnDataBindInfoGetterUserEdit()); ;

            if (mainDataString == "ИНФОРМАЦИЯ")
                return new ColumnDataBind(new ColumnDataBindInfoGetterEmpty());

            if (mainDataString == "ПОЛЬЗОВАТЕЛЬ")
                return new ColumnDataBind(new ColumnDataBindInfoGetterEmpty());

            if (mainDataString == "ПЛАНИРОВАНИЕ")
            {
                return new ColumnDataBind(new ColumnDataBindInfoGetterEmpty());
            }

            if (mainDataString == "ФИНАНСИРОВАНИЕ")
            {
                return new ColumnDataBind(new ColumnDataBindInfoGetterBr {BrType = ColumnDataBindInfoGetterBr.EnumPrjBrType.Fin, FullSprKey = string.Empty});
            }

            if (mainDataString == "АКСИОК")
            {
                return this.ParseAksiokData(bindArr);
            }

            if (mainDataString == "ПАРАМЕТР")
            {
                if (bindArr.Length < 2)
                    throw new NotSupportedException("Неопределено имя параметра " + bindText);
                var param = createInfo.NamedParams.SingleOrDefault(x => x.Name.ToUpper() == bindArr[1].ToUpper());
                if (param == null)
                {
                    this.Logger.Error("Ошибка поиска параметра " + bindText);
                    throw new NotSupportedException("Ошибка поиска параметра " + bindText);
                }

                bindingCell.Value = param.Data.Trim();
                return this.GetColumnDataBindInfo(bindingCell, createInfo);
            }

            throw new DyLoadTemplateException("Ошибка парсинга данных '" + bindText + "'");
        }

        /// <summary> Разобрать получение данных АКСИОКа </summary>
        [NotNull]
        private ColumnDataBind ParseAksiokData(string[] bindArr)
        {
            var contrainer = new ColumnDataBind();

            for (var attrInfoIdx = 1; attrInfoIdx < bindArr.Length; attrInfoIdx++)
            {
                var aksInfo = bindArr[attrInfoIdx];
                if (string.IsNullOrWhiteSpace(aksInfo))
                    continue;
                    
                if (aksInfo.ToUpper().StartsWith("СВОД"))
                {
                    var paramsStrSplt = aksInfo.Split(';');
                    var akiokData = new ColumnDataBindInfoGetterAksiok();

                    #region Собираем параметры АКСИОКа
                    foreach (var aksiokPart in paramsStrSplt)
                    {
                        if (string.IsNullOrWhiteSpace(aksiokPart))
                            continue;

                        var paramSplit = aksiokPart.Split(':');
                        if (paramSplit.Length != 2)
                            throw new NotSupportedException("Ошибка разбора прааметров АКСИОКа " + aksiokPart);

                        var paramNm = paramSplit[0].ToUpper();
                        switch (paramNm)
                        {
                            case "СВОД":
                            case "SVOD":
                                akiokData.Svod = paramSplit[1];
                                break;
                            case "ПАРАМЕТР":
                            case "PARAM":
                                akiokData.Param = paramSplit[1];
                                break;
                            case "DK":
                            case "ДК":
                                akiokData.Dk = paramSplit[1];
                                break;
                            case "ГОД":
                                akiokData.Year = paramSplit[1];
                                break;
                            case "РЗПРЗ":
                                akiokData.RzPrz = paramSplit[1];
                                break;
                            case "ЦСР":
                                akiokData.Csr = paramSplit[1];
                                break;
                            case "ВР":
                                akiokData.Vr = paramSplit[1];
                                break;
                            case "КОСГУ":
                                akiokData.Kosgu = paramSplit[1];
                                break;
                            case "НР":
                                akiokData.Nr = paramSplit[1];
                                break;
                            case "СФР":
                                akiokData.Sfr = paramSplit[1];
                                break;
                            default:
                                throw new NotSupportedException("Ошибка разбора шаблона. Неопределен параметр АКСИОКа " + aksiokPart + " в " + aksInfo);
                        }
                    }
                    #endregion

                    if (akiokData.Param == null || akiokData.Svod == null)
                        throw new NotSupportedException("Не заданы парметры СВОДа или Параметров " + aksInfo);
                    
                    #region Разбираемся с параметрами. Они могут быть множественными и формулами
                    var reManyParams = new Regex(@"^(\d+)#(\d+)$");
                    var reSingleNum = new Regex(@"^(\d+)$");
                    var reSimpleFormula = new Regex(@"^(\d+)((\+|\-)(\d+))+$");

                    if (reSingleNum.IsMatch(akiokData.Param))
                    {   // Простое получение праметров АКСИОКа
                        contrainer.AddColumn(akiokData);
                    }
                    else if (reManyParams.IsMatch(akiokData.Param))
                    {   // Множественные параметры
                        var mch = reManyParams.Match(akiokData.Param);
                        int fromVal;
                        if (!int.TryParse(mch.Groups[1].Value, out fromVal))
                            throw new NotSupportedException("Неопределен порядок разбора параметров АКСИОКа " + akiokData.Param);
                        int toVal;
                        if (!int.TryParse(mch.Groups[2].Value, out toVal))
                            throw new NotSupportedException("Неопределен порядок разбора параметров АКСИОКа " + akiokData.Param);
                        
                        for (var i = fromVal; i <= toVal; i++)
                        {
                            var p = akiokData.Clone();
                            p.Param = i.ToString();
                            contrainer.AddColumn(p);
                        }
                    }
                    else if (reSimpleFormula.IsMatch(akiokData.Param))
                    {
                        var res = new ColumnDataBindInfoGetterSomeFormula();
                        var str = akiokData.Param.Trim();
                        var currDgt = new StringBuilder();
                        for (var i = 0; i < str.Length; i++)
                        {
                            if (str[i] >= '0' && str[i] <= '9')
                                currDgt.Append(str[i]);
                            else if (str[i] == '+')
                            {
                                var aks = akiokData.Clone();
                                aks.Param = currDgt.ToString();
                                if (string.IsNullOrWhiteSpace(aks.Param))
                                    throw new NotSupportedException("Ошибка разбора формулы " + str);
                                res.PlusOperation(aks);
                                currDgt = new StringBuilder();
                            }
                            else if (str[i] == '-')
                            {
                                var aks = akiokData.Clone();
                                aks.Param = currDgt.ToString();
                                if (string.IsNullOrWhiteSpace(aks.Param))
                                    throw new NotSupportedException("Ошибка разбора формулы " + str);
                                res.MinusOperation(aks);
                                currDgt = new StringBuilder();
                            }
                            else
                                throw new NotSupportedException("Ошибка парсинга простой формулы АКСИОК " + str);
                        }

                        var last = akiokData.Clone();
                        last.Param = currDgt.ToString();
                        res.FinishOperand(last);
                        contrainer.AddColumn(res);
                    }
                    else 
                        throw new NotSupportedException("Не смогли разобраться с параметрами АКСИОКа " + akiokData.Param);


                    #endregion
                }
                else 
                    throw new NotSupportedException("Не разобрался с получением данных АКСИОК");
            }

            return contrainer;
        }

        #endregion

        /// <summary> Проверить корректность заполнения данных листов </summary>
        private void ValidatePbsCodes()
        {
            // TODO пока не придумал... но надо наверное
            // коды по строчкам в разных листах сравнить,
            // сравнить коды из Росписи
        }

        /// <summary> Метод получения списка организаций из БР для кода </summary>
        private List<Tuple<string, string>> GetPbsList([NotNull] DyTemplateCreateInfo createInfo)
        {
            //TODO сделать нормально

            return new List<Tuple<string, string>>
            {
                Tuple.Create("1", "Имя 1"),
                Tuple.Create("2", "Имя 2"),
                Tuple.Create("3", "Имя 3"),
                Tuple.Create("4", "Имя 4"),
                Tuple.Create("5", "Имя 5"),
                Tuple.Create("6", "Имя 6"),
                Tuple.Create("7", "Имя 7"),
            };
        }


        #region Проверка валидности областей шаблона

        /// <summary> Заполнение данных из системной области </summary>
        private bool FillSysInfo([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo)
        {
            var anyErr = false;
            var sys = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.System);
            for (int rIdx = 0; rIdx <= sys.Range.RowCount; rIdx++)
            {
                if (sys.Range[rIdx, 0].Value.ToString().ToUpper() == "ТИП")
                {
                    var vl = sys.Range[rIdx, 1].Value.ToString().ToUpper();
                    switch (vl)
                    {
                        case "РАСЧЕТ":
                            sheetInfo.SubType = DyDocSs.DyDocSsSheetInfo.EnumSheetSubType.Calc;
                            break;
                        case "ЧИСЛЕННОСТЬ":
                            sheetInfo.SubType = DyDocSs.DyDocSsSheetInfo.EnumSheetSubType.Count;
                            break;
                        case "ЗАЯВКИ":
                            sheetInfo.SubType = DyDocSs.DyDocSsSheetInfo.EnumSheetSubType.Summ;
                            break;
                        default:
                            this.Logger.Error("Ошибка определения типа листа " + vl);
                            anyErr = true;
                            break;
                    }
                }
                else if (sys.Range[rIdx, 0].Value.ToString().ToUpper() == "СТРОКИ")
                {
                    var vl = sys.Range[rIdx, 1].Value.ToString().ToUpper();
                    switch (vl)
                    {
                        case "СТАТИЧЕСКИЕ":
                            sheetInfo.RowType = DyDocSs.DyDocSsSheetInfo.EnumSheetRowType.Static;
                            break;
                        case "ДИНАМИЧЕСКИЕ":
                            sheetInfo.RowType = DyDocSs.DyDocSsSheetInfo.EnumSheetRowType.Dynamic;
                            break;
                        default:
                            this.Logger.Error("Ошибка определения типа строк " + vl);
                            anyErr = true;
                            break;
                    }
                }

            }

            return anyErr;
        }

        /// <summary> Проверка региона коэффициентов. </summary>
        /// <returns>Так как идёт хитрая перестройка формул, то запрещаем пересечение области коэффициентов с областью первой пользовательской колонки </returns>
        private bool CheckCoeffRegion([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo)
        {
            var coeffRegion = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.Coeff);
            if (coeffRegion == null)
                return false; // Нет ножек - нет печенек

            var firstInsert = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.FirstInsertColumn);
            if (firstInsert == null)
                return false; // Нет ножек - нет печенек

            if (coeffRegion.Range.RightColumnIndex >= firstInsert.Range.LeftColumnIndex
                && coeffRegion.Range.LeftColumnIndex <= firstInsert.Range.RightColumnIndex)
            {
                this.Logger.Error("Регион коэффициентов не должен пересекаться с регионом первой пользовательской колонки.");
                return true;
            }

            return false;
        }

        /// <summary> Проверка сумм по годам </summary>
        private bool CheckRegionTotalYearSum([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo)
        {
            var yearRange = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.TotalYearSum);
            if (yearRange == null)
                return false; // Нет ножек - нет печенек


            if (sheetInfo.DataRange.Range.LeftColumnIndex > yearRange.Range.LeftColumnIndex ||
                sheetInfo.DataRange.Range.RightColumnIndex < yearRange.Range.RightColumnIndex)
            {
                this.Logger.Error("-> Область TotalYearSum не попадает в Data.");
                return true;
            }

            return false;
        }

        /// <summary> Проверка привязки дополнительных исходящих данных
        /// (к распределению, нераспределенный остаток, формула)
        /// Проверяем, что у нас области этих данных и область итога по годам пересекается и можно работать.
        /// </summary>
        private bool CheckRegionOutcomeAdditionalBinding([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo, EnumDyDocSsRanges enumTemplateRanges)
        {
            var chkRange = sheetInfo.TryGetDefinedRange(enumTemplateRanges);
            if (chkRange == null)
                return false; // Нет ножек - нет печенек

            if (!chkRange.Range.IsRangeFullRow())
            {
                this.Logger.Error("-> Область {Name} должна быть полной строкой.", chkRange.Name);
                return true;
            }

            if (chkRange.Range.RowCount != 1)
            {
                this.Logger.Error("-> Область {Name} должна быть одной строкой.", chkRange.Name);
                return true;
            }

            var yearRange = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.TotalYearSum);
            if (yearRange == null)
            {
                this.Logger.Error("-> Область TotalYearSum отсутствует при заданной области {Name}.", chkRange.Name);
                return true;
            }

            return false;
        }

        /// <summary> Проверка информации по системным наименованиям колонок </summary>
        /// <returns>Есть ли ошибка парсинга</returns>
        private bool CheckSysColumnNames([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo)
        {
            var sheet = sheetInfo.Sheet;

            Debug.Assert(sheet != null, nameof(sheet) + " != null");
            var dataRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.Data);
            var dataSysNamesRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.SysColumnNames);

            if (!dataSysNamesRange.Range.IsRangeFullRow())
            {
                this.Logger.Error("-> Область {Name} не является полной строкой.", dataSysNamesRange.Name);
                return true;
            }

            var sysNames = new List<string>();
            for (var datColNum = 0; datColNum < dataRange.Range.ColumnCount; datColNum++)
            {
                var cl = dataSysNamesRange.Range.CellLT(datColNum, 0);
                if (cl.IsMerged)
                {
                    this.Logger.Error("В области {RegionName} содержаться объедененные ячейки. Я пока так не умею.", dataSysNamesRange.Name);
                    return true;
                }

                sysNames.Add(cl.Value.ToString());
            }

            if (sysNames.Any(string.IsNullOrWhiteSpace))
            {
                this.Logger.Error("В области {RegionName} есть пустые системные имена.", dataSysNamesRange.Name);
                return true;
            }

            if (sysNames.GroupBy(x => x).Any(x => x.Count() > 1) )
            {
                foreach (var dubCol in sysNames.GroupBy(x => x).Where(x => x.Count() > 1))
                    this.Logger.Error("В области {RegionName} есть дубликаты системных имён {SysColName}.", dataSysNamesRange.Name, dubCol.Key);

                return true;
            }


            return false;

        }

        /// <summary> Проверка наличия обязательных областей </summary>
        /// <returns>Есть ли ошибка парсинга</returns>
        private bool CheckRequireRange([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo, EnumDyDocSsRanges enumRangeName)
        {
            var rangeValue = sheetInfo.TryGetDefinedRange(enumRangeName);
            if (rangeValue == null)
            {
                Debug.Assert(sheetInfo.Sheet != null, "sheetInfo.Sheet != null");
                this.Logger.Error("-> Ненайдена область '{RangeName}' для листа {SysNm}.", enumRangeName, sheetInfo.Sheet.Name);
                return true;
            }

            return false;
        }

        /// <summary> Парсинг привязки строк </summary>
        /// <returns>Есть ли ошибка парсинга</returns>
        private bool ParseRowBind([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo)
        {
            var dataRange = sheetInfo.GetDefinedRange(EnumDyDocSsRanges.Data);
            var rowBindRange = sheetInfo.TryGetDefinedRange(EnumDyDocSsRanges.RowBind);
            var anyErr = false;
            if (rowBindRange != null)
            {
                for (var datColNum = 0; datColNum < dataRange.Range.ColumnCount; datColNum++)
                {
                    var cell = rowBindRange.Range.CellLT(datColNum, 0).ExistingCells.Single();
                    try
                    {
                        if (string.IsNullOrWhiteSpace(cell.Value.ToString()))
                        {
                            this.Logger.Error("Ошибка биндинга колонок (пустой) " + cell.GetReferenceA1());
                            anyErr = true;
                            continue;
                        }

                        //var columnBindData = new DyTemplateColumnBind(cell.Value.ToString());
                        //sheetInfo.AddColumnBindInfo(cell.ColumnIndex, columnBindData);

                    }
                    catch (DyLoadTemplateException e)
                    {
                        anyErr = true;
                        this.Logger.Error("Ошибка биндинга колонок (ошибка) " + cell.GetReferenceA1() + " " + e.Message);
                        continue;
                    }
                }
            }

            return anyErr;
        }

        /// <summary> Проверка вхождения одной области в другую </summary>
        /// <param name="sheetInfo">Основной лист</param>
        /// <param name="dataRange">Проверяемая область</param>
        /// <param name="enumRangeName">Имя свойства для получения информации о вхождении</param>
        /// <returns>Есть или нет ошибки</returns>
        private bool CheckRegionInclusion([NotNull] DyDocSs.DyDocSsSheetInfo sheetInfo, [CanBeNull] DefinedName dataRange, EnumDyDocSsRanges enumRangeName)
        {
            if (dataRange == null)
                throw new NotSupportedException("Ошибка задания dataRange==null");
            
            var rangeValue = sheetInfo.TryGetDefinedRange(enumRangeName); 
            if (rangeValue == null)
                return false;

            return !dataRange.RegionInclusion(rangeValue);
        }

        #endregion
    }
}