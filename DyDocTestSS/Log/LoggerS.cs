using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using JetBrains.Annotations;
using Serilog;
using Serilog.Context;
using Serilog.Core;
using Serilog.Events;
using Sphaera.Bp.Services.Core;
// ReSharper disable NotAllowedAnnotation

namespace Sphaera.Bp.Services.Log
{
    /// <summary> Обёртка над логерами. </summary>
    /// <remarks>
    /// Методы называть так:
    /// Если сообщение "просто", без детализации по логгерам (по идее - одна из используемых операций в "левых" классах) - постфикc - [ErrorLeve]Message
    /// Если надо какой-то логгер использовать постфикс [ErrorLeve]Logger. Мне кажется это я уберу через пару месяцев вообще.
    /// Если мы работаем с "доменно-ориентированным" (ListController, EditFormController, Repository, Factory и т.д. что может быть строго привязано к домену) логом: методы называем Log[ErrorLevel]. Плюс метод должен быть generic.
    ///     Так сделано, что б в ListController было написано this.LogDebug(...) а не просто this.Debug() мне показалось так лучше.
    /// 
    /// Если сообщений в классе много-много и он не "доменный", то можно получить "настроенную прокси" для логгера GetLogProxy. Который возвращает ILoggerProxy с нужной кучей методов.
    /// 
    /// Порядок аргументов:
    /// Если есть имя логгера - оно первое.
    /// Если есть эксепшн, то оно либо первое, либо сразу после имени логгера,
    /// Если есть привязка к домену - то оно первое и тогда не должно быть имени логгера.
    /// 
    /// После описания метаинформации сообщения - пишется строка форматирования. _ПОКА_ пишем в стандартной нотации {0}, {1} .... Потом - будем думать.
    /// Дальше - не обязательные аргументы.
    /// 
    /// Полный вариант _без_ привязаки к домену:
    ///         void InfoLogger(string loggerName, Exception ex, string msg, params object[] args)
    /// 
    /// Полный вариант _с_ привязкой к домену:
    ///         void LogWarn[TDomain](this ILogDomain[TDomain] associatedWithDomain, Exception ex, string msg, params object[] args)
    /// </remarks>
    public static class LoggerS
    {

        #region Дебаг
        /// <summary> Вывести отладочную информацию в дефолтовый логгер (не попадает в seq и др)</summary>
        [PublicAPI]
        public static void DebugMessage([NotNull] string msg, [CanBeNull] params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Debug), null, msg, args);
        }
        
        /// <summary> 
        /// Вывести отладочную информацию (не попадает в seq и др)
        /// С возможность задания логгера
        /// </summary>
        [PublicAPI]
        public static void DebugLogger([NotNull] string loggerName, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Debug, loggerName), null, msg, args);
        }
        
        /// <summary> 
        /// Вывести отладочную информацию (не попадает в seq и др)
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogDebug<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Debug), null, msg, args);
        }
        #endregion

        #region Информация
        /// <summary> Вывести информацию в дефолтовый логгер </summary>
        [PublicAPI]
        public static void InfoMessage([NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Info), null, msg, args);
        }

        /// <summary> Вывести информацию в дефолтовый логгер </summary>
        [PublicAPI]
        public static void InfoMessage([NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Info), ex, msg, args);
        }

        /// <summary> 
        /// Вывести информацию 
        /// С возможностью задания логгера
        /// </summary>
        [PublicAPI]
        public static void InfoLogger([NotNull] string loggerName, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Info, loggerName), ex, msg, args);
        }

        /// <summary> 
        /// Вывести информацию 
        /// С возможностью задания логгера
        /// </summary>
        [PublicAPI]
        public static void InfoLogger([NotNull] string loggerName, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Info, loggerName), null, msg, args);
        }

        /// <summary> 
        /// Вывести информацию
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogInfo<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Info), null, msg, args);
        }

        #endregion

        #region Предупреждения
        /// <summary> 
        /// Вывести предупреждение в дефолтовый логгер
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void WarnMessage([NotNull] string msg, [CanBeNull] params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Warn), null, msg, args);
        }

        /// <summary> 
        /// Вывести предупреждение в дефолтовый логгер
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void WarnMessage([NotNull] Exception ex, [NotNull] string msg, [CanBeNull] params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Warn), ex, msg, args);
        }

        /// <summary> 
        /// Вывести предупреждение
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void WarnLogger([NotNull] string loggerName, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Warn, loggerName), null, msg, args);
        }

        /// <summary> 
        /// Вывести предупреждение
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void WarnLogger([NotNull] string loggerName, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Warn, loggerName), ex, msg, args);
        }

        /// <summary> 
        /// Вывести предупреждение
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogWarn<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Warn), null, msg, args);
        }

        /// <summary> 
        /// Вывести предупреждение
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogWarn<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Warn), ex, msg, args);
        }
        #endregion

        #region Ошибки
        /// <summary>  Вывести ошибку в дефолтовый логгер </summary>
        [PublicAPI]
        public static void ErrorMessage([NotNull] string msg, [CanBeNull] params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Error), null, msg, args);
        }

        /// <summary>  Вывести ошибку в дефолтовый логгер </summary>
        [PublicAPI]
        public static void ErrorMessage([NotNull] Exception ex, [NotNull] string msg, [CanBeNull]  params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Error), ex, msg, args);
        }

        /// <summary> 
        /// Вывести ошибку
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void ErrorLogger([NotNull] string loggerName, [NotNull] string msg, [CanBeNull] params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Error, loggerName), null, msg, args);
        }

        /// <summary> 
        /// Вывести ошибку
        /// С возможностью задания имени логгера
        /// </summary>
        [PublicAPI]
        public static void ErrorLogger([NotNull] string loggerName, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Error, loggerName), ex, msg, args);
        }

        /// <summary> 
        /// Вывести ошибку
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogError<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Error), null, msg, args);
        }

        /// <summary> 
        /// Вывести ошибку
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogError<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Error), ex, msg, args);
        }

        #endregion

        #region Фаталы
        /// <summary> Вывести фатальную ошибку в дефолтовый логгер </summary>
        public static void FatalMessage([NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Fatal), ex, msg, args);
        }

        /// <summary> Вывести фатальную ошибку в дефолтовый логгер </summary>
        public static void FatalMessage([NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Fatal), null, msg, args);
        }

        /// <summary> 
        /// Вывести фатальную ошибку
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        public static void LogFatal<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] Exception ex, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Fatal), ex, msg, args);
        }

        /// <summary> 
        /// Вывести фатальную ошибку
        /// Extension для ассоциированный с доменом объектов
        /// </summary>
        [PublicAPI]
        public static void LogFatal<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, [NotNull] string msg, params object[] args)
        {
            InternalLogger(associatedWithDomain.GetLogEntryByLogDomain(EnumLevel.Fatal), null, msg, args);
        }
        #endregion

        #region Спецлогирование выполнения опреации пользователем

        /// <summary> Спецфункция логирования действий пользователя. </summary>
        public static void UserAction([NotNull] string msg, params object[] args)
        {
            InternalLogger(new LogEntryInfo(EnumLevel.Info), null, msg, args);
        }
        #endregion

        #region Вспомогательные действия

        /// <summary> Список логгеров. Пока на log4net </summary>
        [NotNull]
        // ReSharper disable InconsistentNaming
        private static readonly ConcurrentDictionary<string, ILogger> _loggers;
        // ReSharper restore InconsistentNaming

        /// <summary> Сессия </summary>
        [NotNull]
        public static string SessionId { get; private set; }

        /// <summary> Имя текущего пользователя </summary>
        [NotNull]
        private static string _currentLogin = "";

        /// <summary> Текущая организация </summary>
        /// <remarks>Сейчас только код орги. Вообще нужна для более удобной отладки параллельой работы</remarks>
        [NotNull]
        private static string _currentOrganization = "";

        /// <summary> Имя проетка (ФБПФ, АСУД и т.д.) </summary>
        [NotNull]
        private static string _projectName = "";

        /// <summary> Включен ли log4net в конфиге? </summary>
        // ReSharper disable once RedundantDefaultMemberInitializer
        private static bool _isLog4NetEnabled = true;

        /// <summary> Попытаться получить логгер. Если находим - возвращаем, если нет - создаем </summary>
        /// <param name="loggerName">Имя логгера</param>
        /// <returns>Логгер по имени</returns>
        [CanBeNull]
        private static ILogger GetLogger([NotNull] string loggerName)
        {
            return _loggers.AddOrUpdate(loggerName, x => ConfigureSerilogLogger(loggerName), (x, l1) => l1);
        }


        /// <summary> Описание доп полей "имя пользователя" </summary>
        private class UserNameEnricher : ILogEventEnricher
        {
            public void Enrich([NotNull] LogEvent logEvent, [NotNull] ILogEventPropertyFactory propertyFactory)
            {
                logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty("Login", _currentLogin));
            }
        }

        private class OrganizatoinEnricher : ILogEventEnricher
        {
            public void Enrich([NotNull] LogEvent logEvent, [NotNull] ILogEventPropertyFactory propertyFactory)
            {
                logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty("Organization", _currentOrganization));
            }
        }


        /// <summary> Описание доп полей "сессия" </summary>
        private class SessionEnricher : ILogEventEnricher
        {
            public void Enrich([NotNull] LogEvent logEvent, [NotNull] ILogEventPropertyFactory propertyFactory)
            {
                logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty("Session", SessionId));
            }
        }

        /// <summary> Описание доп полей "имя проекта" </summary>
        private class ProjectNameEnricher : ILogEventEnricher
        {
            public void Enrich([NotNull] LogEvent logEvent, [NotNull] ILogEventPropertyFactory propertyFactory)
            {
                logEvent.AddPropertyIfAbsent(propertyFactory.CreateProperty("Project", _projectName));
            }
        }

        /// <summary> Сконфигурировать системму логгеров </summary>
        /// <param name="loggerName">Имя логгера</param>
        /// <returns>Null если нет конфига, или логгер, если конфиг есть.</returns>
        [CanBeNull]
        private static Serilog.ILogger ConfigureSerilogLogger([NotNull] string loggerName)
        {

            {
                // Конфигурим логи и их Enrich
                var serilog = new Serilog
                        .LoggerConfiguration()
                        //.MinimumLevel.Verbose()
                        .Enrich.With(new UserNameEnricher())
                        .Enrich.With(new OrganizatoinEnricher())
                        .Enrich.With(new SessionEnricher())
                        .Enrich.With(new ProjectNameEnricher())
                        .Enrich.FromLogContext();
                try
                {
                    serilog.Enrich.WithProperty("Host", Dns.GetHostName());
                }
                catch (SocketException)
                {
                    Debug.Print("Давим потенциальную ошибку получения имени хоста.");
                }

                serilog = serilog.WriteTo.LiterateConsole(outputTemplate: "{Timestamp:yyyy.MM.dd HH:mm:ss.ff} [{Level}] {Message}{NewLine}{Exception}", restrictedToMinimumLevel: LogEventLevel.Debug);

                serilog = serilog.WriteTo.Log4Net(defaultLoggerName: loggerName, restrictedToMinimumLevel: LogEventLevel.Verbose);
                _isLog4NetEnabled = true;

                // Можно вместо Seq отправлять всё в graylog https://github.com/whir1/serilog-sinks-graylog
                // Система помощнее, но сложнее в настройке.

//                serilog = serilog.WriteTo.Seq(sphaeraConfig.Loggers.Seq.Address, restrictedToMinimumLevel: LogEventLevel.Information);

                return serilog.CreateLogger();
            }
            return null;
        }

        /// <summary>
        /// Получить локальную информацию по логируемому сообщению.
        /// </summary>
        /// <typeparam name="TDomain">Тим доменного объекта. По факту - любой объект</typeparam>
        /// <param name="associatedWithDomain">Ассоциированный с доменом лог-класс (реализует ILogDomain)</param>
        /// <param name="lvl">Уровень ошибки</param>
        /// <returns></returns>
        [NotNull]
        internal static LogEntryInfo GetLogEntryByLogDomain<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain, EnumLevel lvl)
        {
            return new LogEntryInfo(lvl, DefinitionLoggerNameByAssociatedObject(associatedWithDomain), DefinitionSubLoggerNameByAssociatedObject(associatedWithDomain), DefinitionDomainNameByAssociatedObject(associatedWithDomain));
        }

        /// <summary> Информация сообщения (уровень ошибки, логгер и т.д. </summary>
        public class LogEntryInfo
        {
            /// <summary>
            /// Конструктор
            /// </summary>
            /// <param name="lvl">Уровень ошибки</param>
            /// <param name="loggerName">Логгер</param>
            /// <param name="subLogger">Подлоггер</param>
            /// <param name="domainName">Домен</param>
            public LogEntryInfo(EnumLevel lvl, [NotNull] string loggerName = "Default", [NotNull] string subLogger = "", [CanBeNull] string domainName = null)
            {
                this.Lvl = lvl;
                this.LoggerName = loggerName;
                this.SubLogger = subLogger;
                this.DomainName = domainName;
            }

            /// <summary> Уровень ошибки </summary>
            public EnumLevel Lvl { get; private set; }

            /// <summary> Имя логгера (в основном - default)</summary>
            [NotNull]
            public string LoggerName { get; private set; }

            /// <summary> Подлоггер (EditForm, Repository и т.д.) </summary>
            [NotNull]
            public string SubLogger { get; private set; }

            /// <summary> Доменное имя основного объекта </summary>
            [CanBeNull]
            public string DomainName { get; private set; }
        }

        /// <summary> Уровень ошибки </summary>
        public enum EnumLevel
        {
            /// <summary> Отладка </summary>
            Debug,
            
            /// <summary> Информация </summary>
            Info,
            
            /// <summary> Предупреждение </summary>
            Warn,

            /// <summary> Ошибка </summary>
            Error,

            /// <summary> Полный треш </summary>
            Fatal
        }

        /// <summary>
        /// Собственно сам вывод информации в логи
        /// </summary>
        /// <param name="entryInfo">Дополнительная информация по сообщению</param>
        /// <param name="ex">Исключение (может быть null)</param>
        /// <param name="msg">Строка форматирования</param>
        /// <param name="args">Аргументы</param>
        internal static void InternalLogger([NotNull] LogEntryInfo entryInfo, [CanBeNull] Exception ex, [NotNull] string msg, [CanBeNull] params object[] args)
        {
            var log = GetLogger(entryInfo.LoggerName);
            if (log == null)
                return;

            // Можно то что попало в Enrich пихать в properties
            // http://stackoverflow.com/questions/12139486/log4net-how-to-add-a-custom-field-to-my-logging 
            // Тогда в Log2Console внизу может отображаться
            //log4net.LogicalThreadContext.Properties["CustomColumn"] = "Custom value";
            //log.Info("Message");
            //// ...or global properties
            //log4net.GlobalContext.Properties["CustomColumn"] = "Custom value"; 
            var listProperty = new List<Tuple<string, IDisposable>>();
            try
            {
                if (!string.IsNullOrEmpty(entryInfo.DomainName))
                {
                    CodeStyleHelper.Assert(entryInfo.DomainName != null, "entryInfo.DomainName != null");
                    listProperty.Add(new Tuple<string, IDisposable>("Domain", LogContext.PushProperty("Domain", entryInfo.DomainName.Split('.').Last())));
                    if (_isLog4NetEnabled)
                        log4net.LogicalThreadContext.Properties["Domain"] = entryInfo.DomainName.Split('.').Last();
                }

                //if (!string.IsNullOrEmpty(entryInfo.LoggerName))
                //    listProperty.Add(LogContext.PushProperty("Logger", entryInfo.LoggerName));

                if (!string.IsNullOrEmpty(entryInfo.SubLogger))
                {
                    listProperty.Add(new Tuple<string, IDisposable>("SubLogger", LogContext.PushProperty("SubLogger", entryInfo.SubLogger)));
                    if (_isLog4NetEnabled)
                        log4net.LogicalThreadContext.Properties["SubLogger"] = entryInfo.SubLogger;
                }

                if (ex != null)
                {

                    switch (entryInfo.Lvl)
                    {
                        case EnumLevel.Debug:
                            log.Debug(ex, msg, args);
                            break;
                        case EnumLevel.Info:
                            log.Information(ex, msg, args);
                            break;
                        case EnumLevel.Warn:
                            log.Warning(ex, msg, args);
                            break;
                        case EnumLevel.Error:
                            log.Error(ex, msg, args);
                            break;
                        case EnumLevel.Fatal:
                            log.Fatal(ex, msg, args);
                            break;
                        default:
                            // ReSharper disable ExceptionNotDocumented
                            throw new NotSupportedException("Не нашли уровень логов");
                        // ReSharper restore ExceptionNotDocumented
                    }
                }
                else
                {
                    switch (entryInfo.Lvl)
                    {
                        case EnumLevel.Debug:
                            log.Debug(msg, args);
                            break;
                        case EnumLevel.Info:
                            log.Information(msg, args);
                            break;
                        case EnumLevel.Warn:
                            log.Warning(msg, args);
                            break;
                        case EnumLevel.Error:
                            log.Error(msg, args);
                            break;
                        case EnumLevel.Fatal:
                            log.Fatal(msg, args);
                            break;
                        default:
                            // ReSharper disable ExceptionNotDocumented
                            throw new NotSupportedException("Не нашли уровень логов");
                        // ReSharper restore ExceptionNotDocumented
                    }
                }
            }
            finally
            {
                foreach (var property in listProperty)
                {
                    if (property != null)
                    {
                        if (_isLog4NetEnabled)
                            log4net.LogicalThreadContext.Properties.Remove(property.Item1);
                        property.Item2.Dispose();
                    }
                }
            }
            
            
        }

        /// <summary> Пустая инициализация. До логина </summary>
        public static void EmptyInitialize()
        {
            log4net.Config.XmlConfigurator.Configure();
            //UserAction("Запуск приложения");
        }

        /// <summary> Инициализировать сессию </summary>
        /// <param name="login">Логин пользователя</param>
        /// <param name="sessionId">Некий идентификатор сессии для seq</param>
        /// <param name="projectName">Имя проекта для удобства фильтрации</param>
        public static void InitializeSession([NotNull] string login, [NotNull] string sessionId, [NotNull] string projectName)
        {
            // Чистим всё
            _currentLogin = login;
            SessionId = sessionId;
            _projectName = projectName;
            _currentOrganization = "<undef>";
        }

        /// <summary>
        /// Дополнительная инициализация логгера (на момент отрабатывания InitializeSession ещё не вся информация для Enrich в наличии)
        /// </summary>
        /// <param name="orgInfo"></param>
        public static void InitializeSessionSecond([NotNull] string orgInfo)
        {
            _currentOrganization = orgInfo;
        }

        /// <summary> Завершить работу </summary>
        public static void LogOff()
        {
            UserAction("Пользователь закрыл приложение");
        }

        #region Определение логгера

        /// <summary> Стандартный подлоггер </summary>
        private const string SubNameUndefConst = "Undef";

        /// <summary> Ассоциированные с классом логгеры </summary>
        [NotNull]
        private static readonly ConcurrentDictionary<string, string> AssociatedDomainLogger = new ConcurrentDictionary<string, string>();

        /// <summary> Ассоциированные с классом подтипы </summary>
        [NotNull]
        private static readonly ConcurrentDictionary<string, string> AssociatedDomainSubLogger = new ConcurrentDictionary<string, string>();

        /// <summary>
        /// Определяем тип логгера по объекту.
        /// По хорошему там должно быть [Db, Repository, Edit, List, Factory] of TDomain
        /// Если дефолтово не нашли - возвращаем FullName типа
        /// </summary>
        /// <returns></returns>
        /// <exception cref="NotSupportedException">Не смогли определить логгер</exception>
        [NotNull]
        private static string DefinitionLoggerNameByAssociatedObject<TDomain>([NotNull] ILogDomain<TDomain> associatedObject)
        {
            var associateType = associatedObject.GetType();
            var associateTypeFullName = associateType.FullName;
            CodeStyleHelper.Assert(associateTypeFullName != null, "associateTypeFullName != null");

            string associateLoggerName;

            if (AssociatedDomainLogger.TryGetValue(associateTypeFullName, out associateLoggerName))
                return associateLoggerName;

            FillLoggerName<TDomain>(associateType);
            if (AssociatedDomainLogger.TryGetValue(associateTypeFullName, out associateLoggerName))
                return associateLoggerName;

            throw new NotSupportedException("Не смогли определить логгер");
        }


        /// <summary>
        /// Определяем тип подлоггера по объекту (визуальщина, репозиторий и т.д.).
        /// По хорошему там должно быть [Db, Repository, Edit, List, Factory] of TDomain
        /// Для нестандартных случаев может не найтись
        /// </summary>
        /// <returns></returns>
        [NotNull]
        private static string DefinitionSubLoggerNameByAssociatedObject<TDomain>([NotNull] ILogDomain<TDomain> associatedObject)
        {
            var associateType = associatedObject.GetType();
            var associateTypeFullName = associateType.FullName;
            CodeStyleHelper.Assert(associateTypeFullName != null, "associateTypeFullName != null");
            string associateLoggerName;

            if (AssociatedDomainSubLogger.TryGetValue(associateTypeFullName, out associateLoggerName))
                return associateLoggerName;

            FillLoggerName<TDomain>(associateType);
            if (AssociatedDomainSubLogger.TryGetValue(associateTypeFullName, out associateLoggerName))
                return associateLoggerName;

            return "";
        }

        /// <summary>
        /// Заполнение информации по логгерам
        /// </summary>
        /// <typeparam name="TDomain"></typeparam>
        /// <param name="associateType"></param>
        private static void FillLoggerName<TDomain>([NotNull] Type associateType)
        {
            string associateTypeFullName = associateType.FullName;
            var associateLoggerName = associateTypeFullName;
            CodeStyleHelper.Assert(associateTypeFullName != null, "associateTypeFullName != null");
            // Определяем подтип вызова
            var associateTypeArr = associateTypeFullName.Split('.');
            if (associateTypeArr.Length >= 4)
            {
                var mainSubType = SubNameUndefConst;
                if (associateTypeArr[2] == "Visual" || associateTypeArr[3] == "Visual")
                    mainSubType = "Visual";
                else if (associateTypeArr[2] == "Bl" || associateTypeArr[3] == "Bl")
                    mainSubType = "Bl";
                else if (associateTypeArr[2] == "Dal" || associateTypeArr[3] == "Dal")
                    mainSubType = "Dal";
                else if (associateTypeArr[2] == "Domain" || associateTypeArr[3] == "Domain")
                    mainSubType = "Domain";

                if (mainSubType != SubNameUndefConst)
                {
                    string subName;
                    if (!associateType.IsNested)
                    {
                        subName = GetSubNameForAssociated<TDomain>(mainSubType, associateType);
                    }
                    else
                    {
                        var tp = associateType;
                        subName = GetSubNameForAssociated<TDomain>(mainSubType, tp);
                        while (subName == SubNameUndefConst && tp.DeclaringType != null)
                        {
                            tp = tp.DeclaringType;
                            subName = GetSubNameForAssociated<TDomain>(mainSubType, tp);
                        }

                        if (subName != SubNameUndefConst)
                            subName += "." + associateType.Name;
                    }

                    if (subName == SubNameUndefConst)
                        subName = associateTypeFullName;

                    var domainType = typeof (TDomain);

                    associateLoggerName = domainType.FullName + "." + subName;
                    AssociatedDomainSubLogger.TryAdd(associateTypeFullName, subName);
                }
            }

            AssociatedDomainLogger.TryAdd(associateTypeFullName, associateLoggerName);
        }

        /// <summary> Regex для поиска интерфейсов являющихся фабриками </summary>
        [NotNull]
        private static readonly Regex ReFactory = new Regex("^I.*?Factory$", RegexOptions.Compiled);

        /// <summary> Статический конструктор </summary>
        static LoggerS()
        {
            _loggers = new ConcurrentDictionary<string, ILogger>();
            SessionId = "";
        }

        /// <summary>
        /// Попытка определить какой частью приложения является ассоциированный тип
        /// </summary>
        /// <param name="mainSubType"> Основная часть приложения (определяется по namespace) </param>
        /// <param name="associateType"> Ассоциированный тип </param>
        /// <returns></returns>
        [NotNull]
        // ReSharper disable UnusedTypeParameter
        private static string GetSubNameForAssociated<TDomain>([NotNull] string mainSubType, [NotNull] Type associateType)
        // ReSharper restore UnusedTypeParameter
        {
            // ReSharper disable ExceptionNotDocumented
            var interfaces = associateType.GetInterfaces();
            // ReSharper restore ExceptionNotDocumented
            string subName;
            if (interfaces.Any(x => x.Name == "IRepository`1" || x.Name == "IBrStructRepository"))
            {
                subName = "Repository";
            }
            else if (interfaces.Any(x => x.Name == "IEditFormController" || x.Name == "IGridEditFormController"))
            {
                subName = "EditForm";
            }
            else if (interfaces.Any(x => x.Name == "IListController"))
            {
                subName = "List";
            }
            else if (interfaces.Any(x => x.Name == "IBrDocFactory" || x.Name == "IBrDocFactory2`1"))
            {
                subName = "Factory";
            }
            else if (interfaces.Any(x => x.Name == "IDbRepository`1" || x.Name == "IDbBrStruct"))
            {
                subName = "Dal";
            }
            else if (interfaces.Any(x => x.Name == "IBrEditor"))
            {
                subName = "BrEditor";
            }
            else
            {
                if (mainSubType == "Bl")
                {
                    if (interfaces.Any(x => ReFactory.IsMatch(x.Name))) // Для чистых файбрик по имени
                        subName = "Factory";
                    else
                        subName = SubNameUndefConst;
                }
                else
                    subName = SubNameUndefConst;
            }
            return subName;
        }
        #endregion

        /// <summary> Получить имя типа доменного объекта </summary>
        /// <typeparam name="TDomain">Тип доменного объекта</typeparam>
        /// <param name="associatedObject">Ассоциированный класс</param>
        /// <returns>Имя типа обеъкта</returns>
        /// <remarks>
        /// Подразумевалось, что я буду давать "умные" имена, но пока я остановился на FullName
        /// </remarks>
        [NotNull]
        // ReSharper disable UnusedParameter.Local удалить параметр нельзя. Мне он для PublicApi нужен.
        private static string DefinitionDomainNameByAssociatedObject<TDomain>([NotNull] ILogDomain<TDomain> associatedObject)
        // ReSharper restore UnusedParameter.Local
        {
            // ReSharper disable once AssignNullToNotNullAttribute
            return typeof(TDomain).FullName;
        }
        #endregion

        /// <summary> Получить прокси на логгер. Нужен, что б по 100500 раз не писать loggerName в вызовах сообщений </summary>
        /// <param name="loggerName"></param>
        /// <returns></returns>
        [NotNull]
        public static ILoggerProxy GetLogProxy([NotNull] string loggerName)
        {
            return new LoggerProxyImp(loggerName);
        }

        /// <summary> Получить прокси на логгер. Нужен, что б по 100500 раз не писать loggerName в вызовах сообщений </summary>
        /// <param name="associatedObject"></param>
        /// <returns></returns>
        [NotNull]
        public static ILoggerProxy GetLogProxy<TDomain>([NotNull] this ILogDomain<TDomain> associatedObject)
        {
            return new LoggerProxyImp(DefinitionLoggerNameByAssociatedObject(associatedObject), DefinitionDomainNameByAssociatedObject(associatedObject));
        }

        /// <summary> Получить доменный логгер </summary>
        /// <typeparam name="TDomain">Тим доменного объекта</typeparam>
        /// <param name="associatedWithDomain">Ассоциированный с доменным объектом класс (реализующий ILogDomain)</param>
        /// <returns>Имплементация доменного логгера</returns>
        [NotNull]
        public static ILoggerProxyTyped<TDomain> NewDomainLogger<TDomain>([NotNull] this ILogDomain<TDomain> associatedWithDomain)
        {
            return new LoggerProxyTypedImp<TDomain>(associatedWithDomain);
        }
    }
}