using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using JetBrains.Annotations;
using Sphaera.Bp.Services.Log;

namespace Sphaera.Bp.Services.Core
{
    /// <summary> Вспомогательный класс счетчиков логирования времени </summary>
    /// <remarks>
    /// Содержит полное время выполнения и его "части". <br/>
    /// При созаднии объекта сразу стартует общий счётчик. <br/>
    /// При добавлении новой части сначала завершается предыдущая часть. (На данный момент вложенные части не допускаются)<br/>
    /// </remarks>
    public class StopwatchParts
    {
        /// <summary> Имя "процесса", который выполняется </summary>
        [NotNull]
        private readonly string _processName;

        /// <summary> Итоговый счетчик </summary>
        [NotNull]
        private readonly Stopwatch _swTotal;

        /// <summary> Счетчик "частей". Содержит "имя" части и собственно счетчик. "Имя" может быть не уникальным. </summary>
        [NotNull, ItemNotNull]
        private readonly List<Tuple<string, Stopwatch>> _swParts;

        /// <summary> Запущена ли последняя часть (нужна для "остановки") </summary>
        private bool _isLastPartRunning;

        public StopwatchParts([NotNull] string processName)
        {
            this._processName = processName;
            this._swTotal = Stopwatch.StartNew();
            this._swParts = new List<Tuple<string, Stopwatch>>(6);
            this._isLastPartRunning = false;
        }

        /// <summary> Запустить подсчёт новой части (автоматом останавливается предыдущая) </summary>
        /// <param name="name">Имя (участвует в логах)</param>
        public void NewPart([NotNull] string name)
        {
            if (this._isLastPartRunning)
                this.StopLastPart();
            this._swParts.Add(new Tuple<string, Stopwatch>(name, Stopwatch.StartNew()));
            this._isLastPartRunning = true;
        }

        /// <summary> Остановить последнюю запущенную часть </summary>
        private void StopLastPart()
        {
            if (!this._isLastPartRunning)
                return;

            var lastSw = this._swParts.Last();
            lastSw.Item2.Stop();

            this._isLastPartRunning = false;
        }

        /// <summary> Остановить подсчёт и записать в логи данные </summary>
        public void StopAndLogging()
        {
            this.StopAndLogging(LoggerS.GetLogProxy("sw"), 0);
        }

        /// <summary> Остановить подсчёт и записать в логи данные, если общее время выполнения превышает "sec" секунд </summary>
        /// <param name="logger">Логгер, куда пишем</param>
        /// <param name="secTime">Время отсечки</param>
        public void StopAndLogging([NotNull] ILoggerProxy logger, double secTime)
        {
            this.StopLastPart();
            this._swTotal.Stop();
            if (this._swTotal.ElapsedMilliseconds > secTime * 1000)
            {
                logger.StatisticFast(secTime, this._swTotal, this._processName);
                foreach (var swPart in this._swParts)
                    logger.Info("{0} - {1}", swPart.Item2.Elapsed, swPart.Item1);
            }
        }

        public override string ToString()
        {
            var prt = "Общее время: {0} - {1}".StringFormat(this._swTotal.Elapsed, this._processName);

            foreach (var swPart in this._swParts)
                prt += "\n {0}  {1} - {2}".StringFormat(swPart.Item2.IsRunning ? "[running]" : "", swPart.Item2.Elapsed, swPart.Item1);

            return prt;
        }
    }
}