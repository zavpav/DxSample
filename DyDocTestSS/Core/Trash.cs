using System;
using System.Collections.Generic;
using DyDocTestSS;
using JetBrains.Annotations;

namespace Sphaera.Bp.Bl.Excel
{
    public class NoDataFoundException : Exception
    {
        public NoDataFoundException(string неНайденаОбласть)
        {
            throw new NotImplementedException();
        }
    }



    public static class JJJ
    {
        /// <summary> Применить операцию для всех элементов последовательности </summary>
        /// <typeparam name="TA">Обрабатываемый тип</typeparam>
        /// <param name="srcLst">Исходный список</param>
        /// <param name="execFunct">Применяемая функция</param>
        public static void ForEach<TA>([NotNull, ItemNotNull, InstantHandle] this IEnumerable<TA> srcLst, [NotNull, InstantHandle] Action<TA> execFunct)
        {
            foreach (var obj in srcLst)
                execFunct(obj);
        }

    }
}

namespace Sphaera.Bp.Services.Core
{
    /// <summary> Вспомогательные методы работы со стркоами </summary>
    public static class StringHelper
    {
        /// <summary> Форматирование стринга </summary>
        [NotNull, StringFormatMethod("formatString")]
        public static string StringFormat([NotNull] this string formatString, [NotNull] params object[] args)
        {
            return string.Format(formatString, args);
        }

        /// <summary> Вычисление расстояния Левенштейна. </summary>
        /// <remarks>Стырено из ru.wikibooks.org</remarks>
        public static int LevenshteinDistance([NotNull] string string1, [NotNull] string string2)
        {
            if (string1 == null) throw new ArgumentNullException("string1");
            if (string2 == null) throw new ArgumentNullException("string2");

            var m = new int[string1.Length + 1, string2.Length + 1];

            for (var i = 0; i <= string1.Length; i++) { m[i, 0] = i; }
            for (var j = 0; j <= string2.Length; j++) { m[0, j] = j; }

            for (var i = 1; i <= string1.Length; i++)
            {
                for (var j = 1; j <= string2.Length; j++)
                {
                    var diff = (string1[i - 1] == string2[j - 1]) ? 0 : 1;

                    m[i, j] = System.Math.Min(System.Math.Min(m[i - 1, j] + 1,
                            m[i, j - 1] + 1),
                        m[i - 1, j - 1] + diff);
                }
            }
            return m[string1.Length, string2.Length];
        }


    }
}
