using System.Collections.Generic;
using System.Linq;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraSpreadsheet.Services;
using DyDocTestSS.Domain;

namespace DyDocTestSS.Bl
{
    public class DySsFormulaEngine : ICustomCalculationService
    {
        public DyDocSs DyDoc { get; }

        public DySsFormulaEngine(DyDocSs dyDoc)
        {
            this.DyDoc = dyDoc;
        }

        public bool OnBeginCalculation()
        {
            return true;
        }

        public bool ShouldMarkupCalculateAlwaysCells()
        {
            return true;
        }

        public void OnEndCalculation()
        {
        }

        public void OnBeginCellCalculation(CellCalculationArgs args)
        {
            var wb = this.DyDoc.Wb;

            var sheetInfo = this.DyDoc.SystemSheetInfos.SingleOrDefault(x => x.Sheet == wb.Sheets[args.SheetId]);
            if (sheetInfo == null)
                return;

            var cl = sheetInfo.Sheet.Cells[args.Row, args.Column];
            cl.Calculate();
            if (cl.Value.ErrorValue != null && cl.Value.ErrorValue.Type == ErrorType.DivisionByZero)
            {
                args.Value = 0;
                args.Handled = true;
            }
        }

        public void OnEndCellCalculation(CellKey cellKey, CellValue startValue, CellValue endValue)
        {
        }

        public bool OnBeginCircularReferencesCalculation()
        {
            return true;
        }

        public void OnEndCircularReferencesCalculation(IList<CellKey> cellKeys)
        {
        }
    }
}