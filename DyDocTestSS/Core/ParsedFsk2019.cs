using JetBrains.Annotations;
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable MemberCanBePrivate.Global

namespace SdBp.Domain.Spr
{
    /// <summary> Вспомогательный класс разбора полного кода БР </summary>
    /// <remarks> Может менятся в зависимости от проетка и года </remarks>
    public class ParsedFsk2019
    {
        public ParsedFsk2019([NotNull] string fsk)
        {
            var fskParts = fsk.Split('.');

            if (fskParts.Length > 0)
                this.Grbs = fskParts[0];
            if (fskParts.Length > 1)
                this.RzPrz = fskParts[1];
            if (fskParts.Length > 2)
                this.Csr = fskParts[2];
            if (fskParts.Length > 3)
                this.Vr = fskParts[3];
            if (fskParts.Length > 4)
                this.Kosgu = fskParts[4];
            if (fskParts.Length > 5)
                this.Nr = fskParts[5];
            if (fskParts.Length > 6)
                this.Sfr = fskParts[6];
            if (fskParts.Length > 7)
                this.Pbs = fskParts[7];
        }

        /// <summary> Глава </summary>
        [CanBeNull]
        public string Grbs { get; set; }

        /// <summary> РзПРз </summary>
        [CanBeNull]
        public string RzPrz { get; set; }

        /// <summary> ЦСР </summary>
        [CanBeNull]
        public string Csr { get; set; }

        /// <summary> ВР </summary>
        [CanBeNull]
        public string Vr { get; set; }

        /// <summary> КОСГУ </summary>
        [CanBeNull]
        public string Kosgu { get; set; }

        /// <summary> НР </summary>
        [CanBeNull]
        public string Nr { get; set; }

        /// <summary> СФР </summary>
        [CanBeNull]
        public string Sfr { get; set; }

        /// <summary> ПБС </summary>
        [CanBeNull]
        public string Pbs { get; set; }

        public override string ToString()
        {
            var res = "";

            if (this.Pbs != null)
                res = "." + this.Pbs;

            if (this.Sfr != null)
                res = "." + this.Sfr + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.Nr != null)
                res = "." + this.Nr + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.Kosgu != null)
                res = "." + this.Kosgu + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.Vr != null)
                res = "." + this.Vr + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.Csr != null)
                res = "." + this.Csr + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.RzPrz != null)
                res = "." + this.RzPrz + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (this.Grbs != null)
                res = "." + this.Grbs + res;
            else if (!string.IsNullOrEmpty(res))
                res = "." + "#" + res;

            if (!string.IsNullOrEmpty(res))
                res = res.Substring(1);

            return res;
        }
    }
}