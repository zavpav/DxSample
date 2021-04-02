using System;
using System.Diagnostics;
using JetBrains.Annotations;

namespace Sphaera.Bp.Services.Log
{
    /// <summary>
    /// Прокси к логгеру.
    /// Нужен для удобства и сохранеия переменной логгера.
    /// Так же инкапуслирует пару счетчиков времени выполнения "шага".
    /// - StatisticFast     - лог выводится только если переданный параметр stopwatch больше определенного значения
    /// - StatisticFastStep - лог выводится только если внутренний стандартный stopwatch (_defaultStopwatch) больше значения \d+
    ///                               DefaultStepwatch запускается при создании объекта, сбрасывается при вызове _любого_ метода Fast(\d+)Step.
    ///                               Так же его можно сбросить методом ResetDefaultStepwatch
    /// </summary>
    public interface ILoggerProxy
    {
        /// <summary> Вывести отладочную информацию (не попадает в seq и др)</summary>
        void Debug([NotNull] Exception exception, [NotNull] string msg, params object[] args);
        /// <summary> Вывести отладочную информацию (не попадает в seq и др)</summary>
        void Debug([NotNull] string msg, params object[] args);

        /// <summary> Вывести информацию </summary>
        void Info([NotNull] Exception exception, [NotNull] string msg, params object[] args);
        /// <summary> Вывести информацию </summary>
        void Info([NotNull] string msg, params object[] args);

        /// <summary> Вывести предупреждение </summary>
        void Warn([NotNull] Exception exception, [NotNull] string msg, params object[] args);
        /// <summary> Вывести предупреждение </summary>
        void Warn([NotNull] string msg, params object[] args);

        /// <summary> Вывести ошибку </summary>
        void Error([NotNull] Exception exception, [NotNull] string msg, params object[] args);
        /// <summary> Вывести ошибку </summary>
        void Error([NotNull] string msg, params object[] args);

        /// <summary> Вывести фатальную ошибку </summary>
        void Fatal([NotNull] Exception exception, [NotNull] string msg, [CanBeNull] params object[] args);
        /// <summary> Вывести фатальную ошибку </summary>
        void Fatal([NotNull] string msg, params object[] args);

        /// <summary> 
        /// Сбросить таймеры.
        /// Пошаговый. 
        /// Общий 
        /// </summary>
        void ResetDefaultStepwatch();

        /// <summary> Получить время, которое натикалов в Total счетчике </summary>
        TimeSpan TotalTime();

        /// <summary>
        /// Печать статистики по внутреннему таймеру.
        /// Таймер сбрасывается после каждого вызова метода
        /// </summary>
        /// <param name="secTime">Опорное значение в секундах</param>
        /// <param name="msg">Строка форматирования</param>
        /// <param name="args">Аргументы</param>
        void StatisticFastStep(double secTime, [NotNull] string msg, params object[] args);

        /// <summary>
        /// Печать статистики по внутреннему таймеру.
        /// Таймером не управляется.
        /// </summary>
        /// <param name="secTime">Опорное значение в секундах</param>
        /// <param name="sw">Внешний таймер</param>
        /// <param name="msg">Строка форматирования</param>
        /// <param name="args">Аргументы</param>
        void StatisticFast(double secTime, [NotNull] Stopwatch sw, [NotNull] string msg, params object[] args);
    }
}