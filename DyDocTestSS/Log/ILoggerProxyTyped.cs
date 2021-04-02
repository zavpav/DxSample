namespace Sphaera.Bp.Services.Log
{
    /// <summary>
    /// Типизированный логгер.
    /// 
    /// Имеет 100500 методов вывода логов.
    /// Первая группа: 
    /// - StatisticFast     - лог выводится только если переданный параметр stopwatch больше определенного значения
    /// - StatisticFastStep - лог выводится только если внутренний стандартный stopwatch (_defaultStopwatch) больше значения \d+
    ///                               DefaultStepwatch запускается при создании объекта, сбрасывается при вызове _любого_ метода Fast(\d+)Step.
    ///                               Так же его можно сбросить методом ResetDefaultStepwatch
    /// </summary>
    /// <typeparam name="TDomain"></typeparam>

    // ReSharper disable UnusedTypeParameter
    public interface ILoggerProxyTyped<TDomain> : ILoggerProxy
    // ReSharper restore UnusedTypeParameter
    {
    }
}