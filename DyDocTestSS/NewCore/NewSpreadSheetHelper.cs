using System.Linq;
using System.Text.RegularExpressions;
using DevExpress.Spreadsheet;
using JetBrains.Annotations;
using Sphaera.Bp.Services.Log;

namespace DyDocTestSS.DyTemplates
{
    /// <summary> Дополнительные методы, которые я хочу перетащить в SpreadSheetHelper </summary>
    public static class NewSpreadSheetHelper
    {
        /// <summary> Основной метод проверяющий вхождение rangeInclusion в rangeMain </summary>
        /// <param name="rangeMain">Основной регион</param>
        /// <param name="rangeInclusion">Проверяемый регион</param>
        /// <returns>Входит или нет</returns>
        public static bool RegionInclusion([NotNull] this DefinedName rangeMain, [NotNull] DefinedName rangeInclusion)
        {
            if (rangeInclusion.Range.TopRowIndex < rangeMain.Range.TopRowIndex ||
                rangeInclusion.Range.BottomRowIndex > rangeMain.Range.BottomRowIndex ||
                rangeInclusion.Range.LeftColumnIndex < rangeMain.Range.LeftColumnIndex ||
                rangeInclusion.Range.RightColumnIndex > rangeMain.Range.RightColumnIndex
            )
            {
                LoggerS.InfoMessage("Область {InclusionName} не входит в {MainName}.", rangeInclusion.Name, rangeMain.Name);
                return false;
            }

            return true;
        }

        /// <summary> Основной метод проверяющий вхождение rangeInclusion в rangeMain </summary>
        /// <param name="rangeMain">Основной регион</param>
        /// <param name="rangeInclusion">Проверяемый регион</param>
        /// <returns>Входит или нет</returns>
        public static bool RegionInclusion([NotNull] this Range rangeMain, [NotNull] Range rangeInclusion)
        {
            if (rangeInclusion.TopRowIndex < rangeMain.TopRowIndex ||
                rangeInclusion.BottomRowIndex > rangeMain.BottomRowIndex ||
                rangeInclusion.LeftColumnIndex < rangeMain.LeftColumnIndex ||
                rangeInclusion.RightColumnIndex > rangeMain.RightColumnIndex
            )
            {
                //LoggerS.InfoMessage("Область не входит в другую.");
                return false;
            }

            return true;
        }


        private static readonly Regex ReIsFullRow = new Regex(@"^\d+:\d+$", RegexOptions.Compiled);
        /// <summary> Является ли область "полной строкой" </summary>
        public static bool IsRangeFullRow([NotNull] this Range range)
        {
            // Не смог добраться до ModelRange в Range. Через GetProperty не стал. Что так костыль, что эдак.
            var regionName = range.GetReferenceA1(ReferenceElement.ColumnAbsolute | ReferenceElement.ColumnAbsolute);
            return ReIsFullRow.IsMatch(regionName);
        }

        private static readonly Regex ReIsFullCol = new Regex(@"^\$[A-Z]+:\$[A-Z]+$", RegexOptions.Compiled);
        
        /// <summary> Является ли область "полной колонкой" </summary>
        public static bool IsRangeFullCol([NotNull] this Range range)
        {
            // Не смог добраться до ModelRange в Range. Через GetProperty не стал. Что так костыль, что эдак.
            var regionName = range.GetReferenceA1(ReferenceElement.ColumnAbsolute | ReferenceElement.ColumnAbsolute);
            return ReIsFullCol.IsMatch(regionName);
        }

        /// <summary> Вернуть размерность из формата </summary>
        /// <returns>Размерность или null, если ничего не понятно.</returns>
        public static int? GetPrecisionFromNumberFormat([NotNull] this Cell cell)
        {
            var format = cell.NumberFormat;
            var rePrecision = new Regex(@"^[#,0]*?(\.([#0]+))?$");
            var mch = rePrecision.Match(format);

            if (mch.Success)
            {
                if (mch.Groups[1].Success)
                {
                    var precisionFormat = mch.Groups[2].Value;
                    return precisionFormat.Count(c => c == '#' || c == '0');
                }
                else
                    return 0; // Нет дробной части
            }

            return null;
        }

    }
}