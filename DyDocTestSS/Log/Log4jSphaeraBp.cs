using System;

using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using JetBrains.Annotations;
using log4net.Core;
using Sphaera.Bp.Services.Core;

namespace Sphaera.Bp.Services.Log
{
    /// <summary> Обёртка log4net для посылки в Log2Console </summary>
    /// <remarks>
    /// Подменяем информацию о точке логирования (LocationInfo) вместо sink'а serilog'а выводим нормаьлный класс (ищем по стеку).
    /// 
    /// </remarks>
    [UsedImplicitly]
    // ReSharper disable once InconsistentNaming
    public class Log4jSphaeraBp : log4net.Layout.XmlLayoutSchemaLog4j
    {
        /// <summary> Строка поиска информации о точки логирования. Если её нет, наверное прямой вызов log4net </summary>
        private const string Log4NetFind = "<log4j:locationInfo class=\"Serilog.Sinks.Log4Net.Log4NetSink\" method=\"Emit\" file=\"\" line=\"0\" />";

        /// <summary> Отдельная ошибка при длительной опреации. Там перетрясам лог и locationInfo </summary>
        /// <remarks>
        ///Не нашли информацию о серилоге <log4j:event logger="LongOperation" timestamp="1484568977294" level="ERROR" thread="addRr"><log4j:message><![CDATA[Ошибка при длительной операции Последовательность не содержит соответствующий элемент
        ///   в System.Linq.Enumerable.Single[TSource](IEnumerable`1 source, Func`2 predicate)
        ///   в Sphaera.Bp.Fbpf.Bl.Rbs.Rr.Tg.DocRbsTgFactory.CreateTg(DocTgSetting domainObject, ISprPbs pbs, List`1 baseRows) в D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Fbpf.Bl\Rbs\Rr\Tg\DocRbsTgFactory.cs:строка 659
        ///   в Sphaera.Bp.Bl.Rr.Tg.DocTgFactoryBase`2.CreateTgForPbs(DocTgSetting tgs, ISprPbs pbs, Boolean isDetail) в D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Bl\Rr\Tg\DocTgFactoryBase.cs:строка 231
        ///   в Sphaera.Bp.Bl.Rr.Tg.DocTgSettingFactory.CheckSummAllTg(DocTgSetting tgs, LongOperationMessageDelegate updateLongOperation) в D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Bl\Rr\Tg\DocTgSettingFactory.cs:строка 504
        ///   в Sphaera.Bp.Visual.Rr.Tg.DocGrrTgSettingEditFormController.<>c__DisplayClass39_1.<AddRr>b__1() в D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Visual\Rr\Tg\DocTgSettingEditFormController.cs:строка 251
        ///   в Sphaera.Bp.Services.Threads.ThreadPoolInstance.<>c__DisplayClass8_0.<NewThread>b__0() в D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Services\Threads\ThreadPoolInstance.cs:строка 48]]></log4j:message><log4j:properties><log4j:data name="log4net:UserName" value="ZAVJVLOV-VM10\Admin" /><log4j:data name="log4jmachinename" value="Zavjvlov-vm10" /><log4j:data name="log4japp" value="Sphaera.Bp.Fbpf.vshost.exe" /><log4j:data name="log4net:HostName" value="Zavjvlov-vm10" /></log4j:properties>
        ///   <log4j:locationInfo class="Sphaera.Bp.Services.Threads.ThreadPoolInstance+&lt;&gt;c__DisplayClass8_0" method="&lt;NewThread&gt;b__0" file="D:\Projects\Fsf\trunk\CS\Sphaera.Bp\Sphaera.Bp.Services\Threads\ThreadPoolInstance.cs" line="69" />
        ///   </log4j:event>
        /// </remarks>
        // ReSharper disable once UnusedMember.Local 
        private const string Log4NetThreadPool = "<log4j:locationInfo class=\"Sphaera.Bp.Services.Threads.ThreadPoolInstance";

        /// <summary> Поиски в "текстовом описании stack trace" </summary>
        private readonly Regex _findInStringStackTrace = new Regex(@"^.*?(?<class>Sphaera\.Bp\..*?)\.(?<method>[^\.]+\(.*?\)).*?(?<path>.:[^:]+):[^\d]+(?<line>\d+)\s*$", RegexOptions.Compiled);

        /// <summary> Регекс замены LocationInfo (при вылете из threadpool) </summary>
        private readonly Regex _replaceLocationInfo = new Regex(@"<log4j:locationInfo.*?/>", RegexOptions.Compiled);


        /// <summary> Определение, что просто сообщение (INFO) из thareadPool </summary>
        private readonly Regex _threadPoolMessage = new Regex(@"<log4j:event .*? level=""INFO"" ", RegexOptions.Compiled);

        /// <summary> Строка по которой игнорируем страшности </summary>
        private const string IgnoreWarn1 = "<log4j:locationInfo class=\"Test";
        /// <summary> Строка по которой игнорируем страшности </summary>
        private const string IgnoreWarn2 = "<log4j:locationInfo class=\"Sphaera.Bp.Wcf";
        

        public override void Format(TextWriter writer, LoggingEvent loggingEvent)
        {
            if (!this.LocationInfo)
                base.Format(writer, loggingEvent);
            else
            {
                var internalWriter = new StringWriter();
                base.Format(internalWriter, loggingEvent);
                internalWriter.Close();
                var txt = internalWriter.GetStringBuilder().ToString();
                var isFnd = true;
                if (!txt.Contains(Log4NetFind))
                {
                    isFnd = false; 
                    if (!txt.Contains(IgnoreWarn1) && !txt.Contains(IgnoreWarn2))
                        for (int i = 0; i < 100; i++)
                        { // выводим 100500 раз что б заметили в output
                            System.Diagnostics.Debug.Print("Не нашли информацию о серилоге " + txt);
                        }
                }



                if (isFnd && loggingEvent.LocationInformation.StackFrames.Length < 100)
                {
                    var isAttention = false;

                    if (txt.Contains("Sphaera.Bp.Services.Threads.ThreadPoolInstance"))
                    {
                        // Вылетели по ошибке из LongOperation stackFrames порушен. Так что просто находим и подменяем LocationInfo. На остальное забиваем
                        var innerException = loggingEvent.ExceptionObject;
                        if (innerException == null)
                        {
                            // Если "обычное сообщение" - игнорим.
                            if (!_threadPoolMessage.IsMatch(txt))
                            {
                                for (int i = 0; i < 100; i++)
                                { // выводим 100500 раз что б заметили в output
                                    System.Diagnostics.Debug.Print("Вылетили из ThreadPool, но без exception. " + txt);
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                foreach (var stackString in innerException.StackTrace.Split('\n'))
                                {
                                    var mchs = _findInStringStackTrace.Matches(stackString);
                                    if (mchs.Count > 0)
                                    {
                                        var className = mchs[0].Groups["class"].Value;
                                        var methodName = mchs[0].Groups["method"].Value;
                                        var fileName = mchs[0].Groups["path"].Value;
                                        var line = mchs[0].Groups["line"].Value;


                                        txt = _replaceLocationInfo.Replace(txt,
                                            string.Format("<log4j:locationInfo class=\"{0}\" method=\"{1}\" file=\"{2}\" line=\"{3}\" />",
                                                className.Replace("<", "&lt;").Replace(">", "&gt;"),
                                                methodName.Replace("<", "&lt;").Replace(">", "&gt;"),
                                                fileName,
                                                line)
                                            );
                                        break;
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                // Что-то пошло не так

                                for (int i = 0; i < 100; i++)
                                { // выводим 100500 раз что б заметили в output
                                    System.Diagnostics.Debug.Print("Вылетили из ThreadPool, но без exception. " + txt + "\n" + e);
                                }
                            }
                            txt = txt.Replace("</log4j:properties>", string.Format("<log4j:data name=\"Exceptions\" value=\"{0}\"/></log4j:properties>",
                                                                            innerException.StackTrace.Replace("<", "&lt;").Replace(">", "&gt;")));

                        }
                    }
                    else
                    {
                        // Стандартное поведение

                        // Так как сюда попадаем через десяток обёрток - их надо все раскрутить.
                        for (var i = 0; i < loggingEvent.LocationInformation.StackFrames.Length; i++)
                        {
                            var stackFrame = loggingEvent.LocationInformation.StackFrames[i];//isReverseFind ? loggingEvent.LocationInformation.StackFrames.Length - 1 - i : i];
                            var methodInfo = stackFrame.GetMethod();
                            var declarationType = methodInfo.DeclaringType;
                            if (declarationType == null)
                                continue;

                            var typeFullName = declarationType.FullName;
                            CodeStyleHelper.Assert(typeFullName != null, "typeFullName != null");

                            if (isAttention)
                            {
                                if (declarationType.Name == "LoggerProxyTypedImp`1" || declarationType.Name == "LoggerS" || declarationType.Name == "LoggerProxyImp")
                                    continue;

                                if (typeFullName.StartsWith("Sphaera.Bp.") || typeFullName.StartsWith("TestCore.") || typeFullName.StartsWith("Test."))
                                { // Докопались до нужного класса
                                    txt = txt.Replace(Log4NetFind, string.Format("<log4j:locationInfo class=\"{0}\" method=\"{1}\" file=\"{2}\" line=\"{3}\" />",
                                            typeFullName.Replace("<", "&lt;").Replace(">", "&gt;"),
                                            methodInfo.Name.Replace("<", "&lt;").Replace(">", "&gt;"),
                                            stackFrame.GetFileName(),
                                            stackFrame.GetFileLineNumber()
                                        ));
                                    var sb2 = new StringBuilder(500);
                                    for (var iFrm = i; iFrm > 0; iFrm--)
                                    {
                                        var frame = loggingEvent.LocationInformation.StackFrames[iFrm].ToString();
                                        if (frame.Contains("Sphaera") || frame.Contains("Test"))
                                            sb2.Append(frame);
                                    }

                                    txt = txt.Replace("</log4j:properties>", string.Format("<log4j:data name=\"Exceptions\" value=\"{0}\"/></log4j:properties>",
                                                sb2.Replace("<", "&lt;").Replace(">", "&gt;")));
                                }


                            }
                            else if (typeFullName.StartsWith("Sphaera.Bp."))
                            {
                                isAttention = true;
                            }
                        }
                    }
                }
                else if (loggingEvent.LocationInformation.StackFrames.Length > 100)
                {
                    txt = txt.Replace("</log4j:properties>", string.Format("<log4j:data name=\"Exceptions\" value=\"{0}\"/></log4j:properties>",
                                "Глубина стека больше 100 пропускается (текущая: " + loggingEvent.LocationInformation.StackFrames.Length + ")"));
                }
                writer.Write(txt);
            }
        }
    }
}