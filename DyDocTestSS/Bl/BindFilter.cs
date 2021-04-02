using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using JetBrains.Annotations;
using SdBp.Domain.Spr;

namespace DyDocTestSS.Bl
{

    /// <summary> Фильтр привязки </summary>
    public interface ITemplateFilter
    {
        /// <summary> Применим ли фильтр </summary>
        /// <param name="fsk">Полный код</param>
        bool IsSuitable([NotNull] string fsk);
    }

    public class BindFilter : ITemplateFilter
    {
        public BindFilter()
        {
            var xEmpty = new XElement("root");
            this.RzPrz = new FilterPart("RzPrz", xEmpty);
            this.Csr = new FilterPart("Csr", xEmpty);
            this.Vr = new FilterPart("Vr", xEmpty);
            this.Kosgu = new FilterPart("Kosgu", xEmpty);
            this.Nr = new FilterPart("Nr", xEmpty);
            this.Sfr = new FilterPart("Sfr", xEmpty);
        }

        public void FillFromXml([NotNull] XElement xFilter)
        {
            this.RzPrz = new FilterPart("RzPrz", xFilter);
            this.Csr = new FilterPart("Csr", xFilter);
            this.Vr = new FilterPart("Vr", xFilter);
            this.Kosgu = new FilterPart("Kosgu", xFilter);
            this.Nr = new FilterPart("Nr", xFilter);
            this.Sfr = new FilterPart("Sfr", xFilter);
        }

        /// <summary> Создать на основе xml </summary>
        [NotNull]
        public static BindFilter CreateFromXml([NotNull] XElement xFilter)
        {
            var bind = new BindFilter();
            bind.FillFromXml(xFilter);
            return bind;
        }


        public bool IsSuitable(string fsk)
        {
            var parsedFsk = new ParsedFsk2019(fsk);
            // ReSharper disable once ReplaceWithSingleAssignment.True
            var isSuitable = true;
            // ReSharper disable once ConditionIsAlwaysTrueOrFalse
            if (isSuitable && this.RzPrz.IsActive && !this.RzPrz.IsSuitable(parsedFsk.RzPrz))
                isSuitable = false;
            if (isSuitable && this.Csr.IsActive && !this.Csr.IsSuitable(parsedFsk.Csr))
                isSuitable = false;
            if (isSuitable && this.Vr.IsActive && !this.Vr.IsSuitable(parsedFsk.Vr))
                isSuitable = false;
            if (isSuitable && this.Kosgu.IsActive && !this.Kosgu.IsSuitable(parsedFsk.Kosgu))
                isSuitable = false;
            if (isSuitable && this.Nr.IsActive && !this.Nr.IsSuitable(parsedFsk.Nr))
                isSuitable = false;
            if (isSuitable && this.Sfr.IsActive && !this.Sfr.IsSuitable(parsedFsk.Sfr))
                isSuitable = false;

            return isSuitable;
        }


        [NotNull]
        private FilterPart RzPrz { get; set; }

        [NotNull]
        private FilterPart Csr { get; set; }

        [NotNull]
        private FilterPart Vr { get; set; }

        [NotNull]
        private FilterPart Kosgu { get; set; }

        [NotNull]
        private FilterPart Nr { get; set; }

        [NotNull]
        private FilterPart Sfr { get; set; }

        /// <summary> Часть фильтра </summary>
        public class FilterPart
        {
            private interface IFilterPartExec
            {
                bool IsSuitable([NotNull] string fskPart);
            }

            /// <summary> "Обычный" фильтр </summary>
            private class FilterPartList : IFilterPartExec
            {
                public FilterPartList([NotNull] string filterData)
                {
                    this.Parts = filterData.Split('|').ToList();
                }

                private List<string> Parts { get; set; }

                public bool IsSuitable(string fskPart)
                {
                    return this.Parts.Contains(fskPart);
                }

                public override string ToString()
                {
                    return string.Join("|", this.Parts);
                }
            }

            /// <summary> "Обычный" фильтр </summary>
            private class FilterPartRegex : IFilterPartExec
            {
                public FilterPartRegex([NotNull] string filterData)
                {
                    this.Re = new Regex(filterData, RegexOptions.Compiled);
                }

                private Regex Re { get; set; }

                public bool IsSuitable(string fskPart)
                {
                    return this.Re.IsMatch(fskPart);
                }


                public override string ToString()
                {
                    return "#Re(" + this.Re + ")";
                }
            }


            public FilterPart([NotNull] string partName, [NotNull] XElement xFilter)
            {
                this.PartName = partName;
                this.IsActive = false;
                this.IsNot = false;

                var xPart = xFilter.Element(partName);

                if (xPart != null && !string.IsNullOrWhiteSpace(xPart.Value))
                {
                    this.IsActive = true;

                    var xNot = xPart.Attribute("not");
                    if (xNot != null && xNot.Value == "true")
                        this.IsNot = true;

                    var xRegex = xPart.Attribute("regex");
                    if (xRegex != null && xRegex.Value == "true")
                        this.PartExec = new FilterPartRegex(xPart.Value);
                    else
                        this.PartExec = new FilterPartList(xPart.Value);
                }
            }

            private IFilterPartExec PartExec { get; set; }

            /// <summary> Наименование "части" он же xml </summary>
            [NotNull]
            public string PartName { get; private set; }

            /// <summary> Отрицание условия </summary>
            private bool IsNot { get; set; }

            /// <summary> Активный ли фильтр </summary>
            public bool IsActive { get; set; }


            public bool IsSuitable([CanBeNull] string fskPart)
            {
                if (fskPart == null)
                    return true;

                var isSuitable = this.PartExec.IsSuitable(fskPart);
                if (this.IsNot)
                    isSuitable = !isSuitable;

                return isSuitable;
            }

            public override string ToString()
            {
                if (!this.IsActive)
                    return "";

                return string.Format("<{0}>{1}{2}{3}</{0}>", this.PartName, this.IsNot ? "НЕ(" : "",
                    this.PartExec.ToString(), this.IsNot ? ")" : "");
            }
        }

        public override string ToString()
        {
            var filterString = this.RzPrz.ToString()
                               + this.Csr
                               + this.Kosgu
                               + this.Nr
                               + this.Sfr;

            return filterString;
        }

    }
}