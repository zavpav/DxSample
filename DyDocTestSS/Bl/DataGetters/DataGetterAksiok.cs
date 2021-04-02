using System;
using System.Linq;
using DyDocTestSS.Bl.Aksiok;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl.DataGetters
{
    public class DataGetterAksiok : IDataGetter
    {
        public DataGetterAksiok()
        {
            this.AksiokStorage = new Aksiok.AksiokStorage();
        }

        /// <summary> Работа с АКСИОКом </summary>
        [NotNull]
        private IAksiokStorage AksiokStorage { get; set; }

        public Type ColumnDataBindInfoGetterType  {  get { return typeof(ColumnDataBindInfoGetterAksiok); } }

        public decimal GetValueDecimal(string pbsCode, ColumnDataBindInfoGetter bindGetterInfo)
        {
            var aksiokInfo = bindGetterInfo as ColumnDataBindInfoGetterAksiok;
            if (aksiokInfo == null)
                throw new NotSupportedException("Плохой биндинг для ColumnDataBindInfoGetter " + bindGetterInfo.GetType());

            var aksiokData = this.AksiokStorage.GetAksiokData(aksiokInfo);
            var pbsDatas = aksiokData
                .SelectMany(x => x.AksiokTableData)
                .Where(x => x.Item1 == pbsCode)
                .ToList();

            if (pbsDatas.Count == 0)
                return 0;
            else if (pbsDatas.Count == 1)
                return pbsDatas[0].Item2;

            throw new NotSupportedException("Ошибка поиска данных АКСИОК. Слишком много данных для " + aksiokInfo.ToStringBind());
        }
    }
}