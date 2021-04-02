using System;
using System.Collections.Generic;

namespace DyDocTestSS.Bl.Aksiok
{
    public interface IAksiokStorage
    {
        List<AksiokDataStub> GetAksiokData(ColumnDataBindInfoGetterAksiok aksiokInfo);
    }

    /// <summary> Заглушка данных АКСИОКа. Потом на что-то умное сменить </summary>
    public class AksiokDataStub
    {
        public string Svod { get; set; }
        public string ParamName { get; set; }
        public string Param { get; set; }
        public string Dk { get; set; }
        public string Year { get; set; }
        public List<Tuple<string, decimal>> AksiokTableData { get; set; }

        /// <summary>
        /// Создание информации по текущему куску АКСИОКа
        /// </summary>
        /// <returns></returns>
        public ColumnDataBindInfoGetterAksiok CreateColumnDataBindInfoAksiok()
        {
            return new ColumnDataBindInfoGetterAksiok
            {
                Svod = this.Svod,
                Param = this.Param,
                Dk = this.Dk,
                Year = this.Year
            };
        }
    }


    public class AksiokStorage : IAksiokStorage
    {
        public List<AksiokDataStub> GetAksiokData(ColumnDataBindInfoGetterAksiok aksiokInfo)
        {
            var funcGenTempData = new Action<AksiokDataStub>(x =>
            {
                var tmptmptmpstub = 1m;
                try
                {
                    tmptmptmpstub = decimal.Parse(x.Svod);
                }
                catch
                {
                    tmptmptmpstub = 12345;
                }


                x.AksiokTableData = new List<Tuple<string, decimal>>(100);
                for (int i = 0; i < 100; i++)
                {
                    x.AksiokTableData.Add(Tuple.Create(i.ToString(), i + tmptmptmpstub + decimal.Parse(x.Param)));
                }
            });

            var aksDat = new AksiokDataStub();
            aksDat.Svod = aksiokInfo.Svod;
            aksDat.ParamName = "Параметр №" + aksiokInfo.Param;
            aksDat.Param = aksiokInfo.Param;
            funcGenTempData(aksDat);
            return new List<AksiokDataStub> { aksDat };
            //}
            //else
            //{
            //    if (filter.Param == null)
            //        throw new NotSupportedException("Генерация стабов для ВСЕГО не поддерживается без наличия этого ВСЕГО где-то там.");

            //    var res = new List<AksiokDataStub>();
            //    foreach (var paramVal in filter.Param)
            //    {
            //        var aksDat = new AksiokDataStub();
            //        aksDat.Svod = filter.Svod;
            //        aksDat.ParamName = "Параметр №" + paramVal;
            //        aksDat.Param = paramVal;
            //        funcGenTempData(aksDat);
            //        res.Add(aksDat);
            //    }

            //    return res;
            //}
        }
    }
}