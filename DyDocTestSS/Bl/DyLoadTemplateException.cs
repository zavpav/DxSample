using System;
using JetBrains.Annotations;

namespace DyDocTestSS.Bl
{
    /// <summary> Эксепшн загрузки шаблона </summary>
    /// <remarks> Используется, когда мы не можем явно сказать, что происходит при загрузке. </remarks>
    public class DyLoadTemplateException : Exception
    {
        public DyLoadTemplateException([NotNull] string message) : base(message)
        {
        }
    }
}