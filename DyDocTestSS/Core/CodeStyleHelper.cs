using System;
using System.Diagnostics;
using JetBrains.Annotations;
// ReSharper disable ArrangeStaticMemberQualifier
// ReSharper disable HeuristicUnreachableCode

namespace Sphaera.Bp.Services.Core
{
    public static class CodeStyleHelper
    {
        [ContractAnnotation("condition:false => halt"), AssertionMethod]
        public static void Assert(bool condition, [CanBeNull] string descr = null)
        {
            if (!condition)
            {
                if (!string.IsNullOrEmpty(descr))
                {
                    // ReSharper disable once AssignNullToNotNullAttribute
                    Debug.Assert(condition, descr);
                    throw new NotSupportedException(descr);
                }
                else
                {
                    Debug.Assert(condition);
                    throw new NotSupportedException();
                }
            }
        }


        [ContractAnnotation("arg:null => halt")]
        public static void ThrowIfNull([CanBeNull] this object arg, [CanBeNull] string descr = null)
        {
            // ReSharper disable once ReturnValueOfPureMethodIsNotUsed
            arg.NotNull(descr);
        }

        /// <exception cref="NullReferenceException">NullRef</exception>
        [ContractAnnotation("arg:null => halt")]
        [NotNull, Pure]
        public static T NotNull<T>([CanBeNull, NoEnumeration] this T arg, [CanBeNull] string descr = null)
            where T : class
        {
            if (arg == null)
            {
                if (!string.IsNullOrWhiteSpace(descr))
                    throw new NullReferenceException(descr);
                else
                    throw new NullReferenceException();
            }

            return arg;
        }

        /// <exception cref="NullReferenceException">NullRef</exception>
        [ContractAnnotation("arg:null => halt")]
        [NotNull, Pure]
        public static T NotNull<T>([CanBeNull, NoEnumeration] this T arg, [NotNull] Func<string> descrFunc)
                    where T : class
        {
            if (arg == null)
            {
                var descr = descrFunc();

                if (!string.IsNullOrWhiteSpace(descr))
                    throw new NullReferenceException(descr);
                else
                    throw new NullReferenceException();
            }

            return arg;
        }

    }
}