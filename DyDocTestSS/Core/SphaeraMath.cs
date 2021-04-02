using System;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using JetBrains.Annotations;

namespace Sphaera.Bp.Services.Math
{
    /// <summary>
    /// Вспомогательный класс математики
    /// TODO наверное надо убрать. Нужен был для корректной работы с double
    /// </summary>
    public static class SphaeraMath
    {
        /// <summary> Точность сравнения (цифр после запятой) </summary>
    	public const int Precision = 5;
        
        /// <summary> Точность сравнения чисел (разность между ними) </summary>
        public const decimal Epsilon = 0.000001m; // По хорошему должно вычисляться на основе Precision, но тогда оно не может быть const

        /// <summary> Стандартое округление (с double были проблемы с 10 знаками после запятой) </summary>
        public static decimal DefaultRound(decimal sm)
        {
            return System.Math.Round(sm, Precision, MidpointRounding.AwayFromZero);
        }

        /// <summary> Сравнение 2х сумм до 6 разряда </summary>
        public static bool IsEqual(double sm1, double sm2)
        {
            return System.Math.Abs(System.Math.Round(sm1 - sm2, Precision, MidpointRounding.AwayFromZero)) < (double)Epsilon;
        }

        /// <summary> Сравнение 2х сумм до 6 разряда </summary>
        public static bool IsEqual(decimal sm1, decimal sm2)
        {
            return sm1 == sm2;
        }


        /// <summary> Распарсить объект в число </summary>
        public static decimal ParseDecimalIgnoreSeparator(object value)
        {
            var strVal = value.ToString();
            return strVal.ParseDecimalIgnoreSeparator();
        }

        /// <summary>
        /// Распарсить строку в число.
        /// Если из БД.
        /// В качестве разделителя - ТОЧКА или Запятая!
        /// Разделителей групп - НЕТ!
        /// Парсит по стандартным культурам
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <exception cref="InvalidCastException">Неверный формат числа.</exception>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        public static decimal ParseDecimalIgnoreSeparator([NotNull] this string value)
        {
            // Парсим "наш" формат
            value = value.Trim(' ', '\t');
            if (Regex.IsMatch(value, @"^\s*-?\d+([,.]\d+)?\s*$"))
                return Decimal.Parse(value
                        .Replace(',', '.')
                        .Replace(".", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
            
            // Парсим
            // ReSharper disable EmptyGeneralCatchClause
            try { return Decimal.Parse(value, System.Threading.Thread.CurrentThread.CurrentCulture); } catch { }
            try { return Decimal.Parse(value, System.Globalization.CultureInfo.CurrentCulture); } catch { }
            //try { return Decimal.Parse(value, DbFormatProvider.CultureInfo); } catch { }
            // ReSharper restore EmptyGeneralCatchClause

            throw new InvalidCastException("Неверный формат числа '" + value + "'");

        }

        /// <summary>
        /// То же, что и ParseDecimalIgnoreSeparator, для чисел, имеющих не более двух знаков после запятой (иначе - InvalidCastException)
        /// 
        /// Распарсить строку в число.
        /// Если из БД.
        /// В качестве разделителя - ТОЧКА или Запятая!
        /// Разделителей групп - НЕТ!
        /// Парсит по стандартным культурам
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static decimal ParseDecimalIgnoreSeparator2Digits(this string value)
        {
            // Парсим "наш" формат
            value = value.Trim(' ', '\t');
            if (Regex.IsMatch(value, @"^\s*-?\d+([,.]\d+)?\s*$") && !Regex.IsMatch(value, @"^\s*-?\d+([,.]\d\d?)?\s*$"))
            {
                throw new InvalidCastException("Неверный формат: число имеет больше двух знаков после запятой '" + value + "'. ");
            }
            return value.ParseDecimalIgnoreSeparator();
        }

        /// <summary> Проверка, что является "нулём" (были проблемы с double) </summary>
        public static bool IsZero(decimal sm1)
		{
			return IsEqual(sm1, 0);
		}
    }
}
