using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using JetBrains.Annotations;

namespace Sphaera.Bp.Services.Log
{
    /// <summary> Реализация нетепизированного прокси к логам </summary>
    public class LoggerProxyImp : ILoggerProxy
    {
        /// <summary> Имя логгера </summary>
        [NotNull]
        private string LoggerName { get; set; }

        /// <summary> Имя доменного объекта (не обязательно, если есть доменный объект, жалательно использовать типизированный логгер) </summary>
        [CanBeNull]
        private string DomainName { get; set; }

        /// <summary> Конструктор </summary>
        public LoggerProxyImp([NotNull] string loggerName, [NotNull] string domainName)
        {
            this.LoggerName = loggerName;
            this.DomainName = domainName;

            this._defaultStepStopwatch.Start();
            this._defaultTotalStopwatch.Start();
        }

        /// <summary> Конструктор </summary>
        public LoggerProxyImp([NotNull] string loggerName)
        {
            this.LoggerName = loggerName;
            this.DomainName = null;

            this._defaultStepStopwatch.Start();
            this._defaultTotalStopwatch.Start();
        }

        /// <summary> Получить дополнительную информацию о логгере. </summary>
        /// <remarks>Формирует enrich для логов (логгер, домен и т.д. что может) </remarks>
        [NotNull]
        private LoggerS.LogEntryInfo GetEntryInfo(LoggerS.EnumLevel lvl)
        {
            return new LoggerS.LogEntryInfo(lvl, this.LoggerName, "", this.DomainName);
        }

        public void Debug(Exception exception, string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Debug), exception, msg, args);
        }

        public void Debug(string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Debug), null, msg, args);
        }


        public void Info(Exception exception, string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Info), exception, msg, args);
        }
        
        public void Info(string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Info), null, msg, args);
        }

        
        
        public void Warn(Exception exception, string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Warn), exception, msg, args);
        }

        public void Warn(string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Warn), null, msg, args);
        }



        public void Error(Exception exception, string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Error), exception, msg, args);
        }

        public void Error(string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Error), null, msg, args);
        }



        public void Fatal(Exception exception, string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Fatal), exception, msg, args);
        }

        public void Fatal(string msg, params object[] args)
        {
            LoggerS.InternalLogger(GetEntryInfo(LoggerS.EnumLevel.Fatal), null, msg, args);
        }

        #region Статистика
        /// <summary> Кеш сообщений логирования. Нужен для оптимизации подбора элемента arg </summary>
        [NotNull] private readonly Dictionary<string, string> _statisticMsgArg = new Dictionary<string, string>(100);

        /// <summary> Регексп поиска элементов </summary>
        [NotNull] private readonly Regex _reMsgArg = new Regex(@"\{(\d+)([^\}]*?)\}", RegexOptions.Compiled);

        /// <summary> Изменение строки сообщения. Добавляются параметры для Stopwatch и порога </summary>
        [NotNull]
        private string ChangeStatisticMessage([NotNull] string msg)
        {
            string changedMsg;
            if (_statisticMsgArg.TryGetValue(msg, out changedMsg))
                return changedMsg;


            var mchs = _reMsgArg.Matches(msg);
            var maxV = 0;
            if (mchs.Count > 0)
                maxV = mchs.OfType<Match>()
                            .Select(x => x.Groups[1].Value)
                            .Select(int.Parse)
                            .Max(x => x + 1);

            changedMsg = string.Format("Стат. {0} (Время: {{{1}}}. Порог: {{{2}:N3}} сек)", msg, maxV, maxV + 1);
            _statisticMsgArg.Add(msg, changedMsg);

            if (_statisticMsgArg.Count > 1000)
            {
                this.Warn("Ошика использования логгера. Слишком много шаблонов сообщений.");
            }

            return changedMsg;
        }

        /// <summary> Дефолтовый сопвотч для шагов</summary>
        /// <remarks>
        /// Используется, что бы не заводить стопвотчи в доменных объектах. 
        /// Используется в методах StatisticFast*Step.
        /// После каждого вызова - сбрасывается. (типа на следующий степ)
        /// </remarks>
        [NotNull]
        private readonly Stopwatch _defaultStepStopwatch = new Stopwatch();
        
        /// <summary> Счетчик для "общих операций". Очень часто мерием время не только "шага", но и полного выполнения </summary>
        [NotNull]
        private readonly Stopwatch _defaultTotalStopwatch = new Stopwatch();

        public void ResetDefaultStepwatch()
        {
            this._defaultStepStopwatch.Restart();
            this._defaultTotalStopwatch.Restart();
        }

        public TimeSpan TotalTime()
        {
            return this._defaultTotalStopwatch.Elapsed;
        }

        public void StatisticFastStep(double secTime, string msg, params object[] args)
        {
            this._defaultStepStopwatch.Stop();
            this.LogStatisticFastQ(this._defaultStepStopwatch, secTime, msg, args);
            this._defaultStepStopwatch.Restart();
        }

        public void StatisticFast(double secTime, Stopwatch sw, string msg, params object[] args)
        {
            this.LogStatisticFastQ(sw, secTime, msg, args);
        }

        /// <summary> Вывести в лог сообщение, если время выполнения не укладывается в норматив </summary>
        /// <param name="sw">Время выполнения "шага"</param>
        /// <param name="secTime">Норматив выполнения шага</param>
        /// <param name="msg">Сообщение</param>
        /// <param name="args">Аргументы</param>
        private void LogStatisticFastQ([NotNull] Stopwatch sw, double secTime, [NotNull] string msg, [NotNull] params object[] args)
        {
            if (sw.ElapsedMilliseconds > secTime * 1000)
            {
                var changedArgs = args.ToList();
                changedArgs.Add(sw.Elapsed);
                changedArgs.Add(secTime);
                LoggerS.InternalLogger(this.GetEntryInfo(LoggerS.EnumLevel.Info), null, this.ChangeStatisticMessage(msg), changedArgs.ToArray());
            }
        }
        #endregion

    }
}