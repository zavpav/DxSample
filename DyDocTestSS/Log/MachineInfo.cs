using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JetBrains.Annotations;

namespace Sphaera.Bp.Services.Log
{
    /// <summary> Информация о текущем комьютере </summary>
    [UsedImplicitly(ImplicitUseTargetFlags.WithMembers)]
    public class MachineInfo
    {
        /// <summary> Ошибки получения данных </summary>
        [NotNull]
        private List<Exception> _errs;

        /// <summary>
        /// Структура описания процессора
        /// </summary>
        [UsedImplicitly(ImplicitUseTargetFlags.Members)]
        public struct CpuInfo
        {
            /// <summary>
            /// Конструктор
            /// </summary>
            /// <param name="name">Наименование процессора</param>
            /// <param name="mHz">Частота</param>
            public CpuInfo([NotNull] string name, int mHz)
                : this()
            {
                this.Name = name;
                this.MHz = mHz;
            }

            /// <summary> Имя процессора </summary>
            public string Name { get; private set; }
            
            /// <summary> Мегагерцы </summary>
            public int MHz { get; set; }
        }

        /// <summary> Конструктор </summary>
        public MachineInfo()
        {
            this._errs = new List<Exception>();
            this.DnsName = "<не определено>";
            this.OsVersion = "<не определено>";
        }

        /// <summary> IP адреса </summary>
        public string[] Addressess { get; set; }


        /// <summary> Версия системы </summary>
        [NotNull]
        public string OsVersion { get; set; }


        /// <summary> DNS имя </summary>
        [NotNull]
        public string DnsName { get; set; }

        /// <summary> Количество ОЗУ </summary>
        public float Ram { get; set; }

        /// <summary> IP адреса </summary>
        // ReSharper disable InconsistentNaming
        public CpuInfo[] CPUs { get; set; }
        // ReSharper restore InconsistentNaming


        /// <summary> Ошибки получения чего-либо из нужного. </summary>
        public IEnumerable<Exception> Errs { get { return this._errs; } }

        /// <summary> Информация по используемым мониторам </summary>
        public string[] ScreensInfo { get; set; }

        /// <summary> Добавление ошибок в лог </summary>
        public void AddErr(Exception exception)
        {
            this._errs.Add(exception);
        }


        /// <summary> Строка форматирования для Serilog </summary>
        public string ToSeriParamString()
        {
            try
            {
                var sb = new StringBuilder(1000);
                sb.Append("ОЗУ:{ОЗУ:N1} Гб");
                sb.Append(" Количество экранов: {Экранов} Разрешение экранов: [");
                sb.Append("{Экран1}");
                for (var i = 1; i < this.ScreensInfo.Length; i++)
                {
                    sb.Append(",");
                    sb.Append("{Экран" + (i + 1) + "}");
                }
                sb.Append("]");
                sb.Append(" Хост:{ИмяХоста} ");
                
                sb.Append(" Процессор(Ядер):{ПроцессорЯдер} [");
                sb.Append("{Процессор1Имя} ({Процессор1Частота})");
                for (var i = 1; i < this.CPUs.Distinct().Count(); i++)
                {
                    sb.Append(",");
                    sb.Append("{Процессор" + (i + 1) + "Имя} ({Процессор" + (i + 1) + "Частота})");
                }
                sb.Append("]");

                sb.Append(" ОС:{ОперационнаяCистема}");
                sb.Append(" IPs:[");
                sb.Append("{IP1}");
                for (var i = 1; i < this.Addressess.Length; i++)
                {
                    sb.Append(",");
                    sb.Append("{IP" + (i + 1) + "}");
                }
                sb.Append("]");
                return sb.ToString();
            }
            catch (Exception)
            {
                return "Ошибка формирования строки информации о хосте";
            }
        }

        /// <summary> Параметры в соответсвии с ToSeriParamString </summary>
        public IEnumerable<object> ToSeriParam()
        {
            var parm = new List<object> {this.Ram, this.ScreensInfo.Length};
            parm.AddRange(this.ScreensInfo);
            parm.Add(this.DnsName);
            
            parm.Add(this.CPUs.Length);
            foreach (var cpu in this.CPUs.Distinct())
            {
                parm.Add(cpu.Name);
                parm.Add(cpu.MHz);
            }

            parm.Add(this.OsVersion);
            parm.AddRange(this.Addressess);

            
            return parm;
        }
    };
}