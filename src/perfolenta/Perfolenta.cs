using System;
using System.IO;
using System.Reflection;
using ScriptEngine.Machine.Contexts;
using ScriptEngine.Machine;
using ScriptEngine.Machine.Values;
using ScriptEngine.HostedScript.Library;
using System.Collections.Generic;
using System.Collections;
using Microsoft.Win32;

namespace perfolenta
{
    /// <summary>
    /// Запускает скрипты на языке Перфолента.Net из скриптов на OneScript
    /// </summary>
    [ContextClass("Перфолента", "Perfolenta")]
    public class Perfolenta : AutoContext<Perfolenta>
    {
        delegate object ExecuteMethodDelegate(string text, params object[] mas);
        delegate Assembly CompileMethodDelegate(string text);

        delegate object StartFuncDelegate(object parameter);
        delegate void StartProcDelegate(object parameter);

        private string CompilerPath;
        private Assembly CompilerAssm;

        private Assembly ScriptAssm;

        private CompileMethodDelegate CompileScriptMethod;
        private CompileMethodDelegate CompileScriptFromFileMethod;

        private ExecuteMethodDelegate ExecuteScriptMethod;
        private ExecuteMethodDelegate ExecuteScriptFromFileMethod;

        private StartFuncDelegate StartFuncMethod;
        private StartProcDelegate StartProcMethod;

        private bool IsCompiled;

        /// <summary>
        /// Создаёт экземпляр объекта класса <see cref="Perfolenta"/>.
        /// </summary>
        public Perfolenta()
        {
            //
        }

        /// <summary>
        /// Создаёт экземпляр объекта класса <see cref="Perfolenta"/>.
        /// </summary>
        public Perfolenta(string ScriptText)
        {
            Text = ScriptText;
        }



        // ========= КОНСТРУКТОРЫ ===============

        /// <summary>
        /// Конструктор по-умолчанию. Создаёт экземпляр объекта класса Перфолента.
        /// </summary>
        /// <returns>Экземпляр объекта класса Перфолента.</returns>
        [ScriptConstructor]
        public static IRuntimeContextInstance Constructor()
        {
            return new Perfolenta();
        }

        /// <summary>
        /// Создаёт экземпляр объекта класса Перфолента с установкой текста скрипта.
        /// </summary>
        /// <param name="ScriptText">Текст скрипта на языке Перфолента. Строка.</param>
        /// <returns>Экземпляр объекта класса Перфолента с установленным текстом скрипта.</returns>
        [ScriptConstructor]
        public static IRuntimeContextInstance Constructor(IValue ScriptText)
        {
            return new Perfolenta(ScriptText.AsString());
        }




        // ========= СВОЙСТВА ===============

        /// <summary>
        /// Текст скрипта на языке Перфолента. Должен быть установлен до вызова метода Компилировать. После компиляции, если текст большой, то для экономии памяти можно присвоить свойству Текст пустую строку. Не используется методом КомпилироватьИзФайла.
        /// </summary>
        /// <value>
        /// тип Строка - Текст скрипта на языке Перфолента.
        /// </value>
        [ContextProperty("Текст", "Text")]
        public string Text
        {
            get;
            set;
        }

        /// <summary>
        /// Если в скрипте для элемента Программа, в котором находится метод Старт, указано пространство имён, то следует задать соответсвующее пространство имён в этом свойстве.
        /// </summary>
        /// <value>
        /// тип Строка - пространство имён элемента Программа, в котором находится метод Старт.
        /// </value>
        [ContextProperty("ПространствоИмен", "NameSpace")]
        public string NameSpace
        {
            get;
            set;
        }


        // ========= МЕТОДЫ ===============


        /// <summary>
        /// Запускает на выполнение метод Старт модуля Программа откомпилированного скрипта.
        /// </summary>
        [ContextMethod("Выполнить", "Execute")]
        public IValue Execute(IValue ParamArray)
        {
            if (!IsCompiled)
                throw new System.Exception("Текст скрипта не скомпилирован.");

            if (!(StartProcMethod is null))
            {
                StartProcMethod(ConvertParam(ParamArray));
                return ValueFactory.CreateNullValue();
            }
            else {
                return ConvertResult(StartFuncMethod(ConvertParam(ParamArray)));
            };

        }

        /// <summary>
        /// Компилирует скрипт находящийся в свойстве Текст.
        /// </summary>
        /// <param name="ProgramModuleName">тип IValue - имя модуля Программа, если оно задано. Не обязательный.</param>
        [ContextMethod("Компилировать", "Compile")]
        public void Compile(IValue ProgramModuleName = null)
        {
            IsCompiled = false;

            if (Text is null)
                throw new System.Exception("Не задан текст скрипта.");

            if (CompileScriptMethod is null)
            {

                if (CompilerPath is null)
                    CompilerPath = FindCompilerPath();

                if (CompilerAssm is null)
                    CompilerAssm = System.Reflection.Assembly.LoadFrom(CompilerPath);

                Type tp = CompilerAssm.GetType("Промкод.Перфолента.Компилятор");
                System.Reflection.MethodInfo met = tp.GetMethod("КомпилироватьСкрипт");
                CompileScriptMethod = (CompileMethodDelegate)Delegate.CreateDelegate(typeof(CompileMethodDelegate), null, met);

            };

            ScriptAssm = CompileScriptMethod(Text);

            CreateStartMethodDelegates(ProgramModuleName?.AsString());

            IsCompiled = true;

        }

        /// <summary>
        /// Компилирует текст скрипта из указанного файла. Файл должен содержать текст в кодировке UTF-8. Этот метод не использует свойство Текст.
        /// </summary>
        /// <param name="ScriptFilePath">тип IValue - путь к файлу скрипта, который будет загружен и компилирован.</param>
        /// <param name="ProgramModuleName">тип IValue - имя модуля Программа, если оно задано. Не обязательный.</param>
        [ContextMethod("КомпилироватьИзФайла", "CompileFromFile")]
        public void CompileFromFile(IValue ScriptFilePath, IValue ProgramModuleName = null)
        {
            IsCompiled = false;

            if (CompileScriptFromFileMethod is null)
            {

                if (CompilerPath is null)
                    CompilerPath = FindCompilerPath();

                if (CompilerAssm is null)
                    CompilerAssm = System.Reflection.Assembly.LoadFrom(CompilerPath);

                Type tp = CompilerAssm.GetType("Промкод.Перфолента.Компилятор");
                System.Reflection.MethodInfo met = tp.GetMethod("КомпилироватьСкриптИзФайла");
                CompileScriptFromFileMethod = (CompileMethodDelegate)Delegate.CreateDelegate(typeof(CompileMethodDelegate), null, met);

            };

            ScriptAssm = CompileScriptFromFileMethod(ScriptFilePath.AsString());

            CreateStartMethodDelegates(ProgramModuleName.AsString());

            IsCompiled = true;
        }

        /// <summary>
        /// Подключает скомпилированную сборку аналогично методу глобального контекста ПодключитьВнешнююКомпоненту.
        /// </summary>
        [ContextMethod("ПодключитьКакВнешнююКомпоненту", "AttachAsAddIn")]
        public void AttachAsAddIn()
        {
            if (ScriptAssm is null)
                throw new System.Exception("Текст скрипта не скомпилирован.");

            var EngineInstance = GlobalsManager.GetGlobalContext<SystemGlobalContext>().EngineInstance;
            EngineInstance.AttachExternalAssembly(ScriptAssm, EngineInstance.Environment);
        }

        /// <summary>
        /// Компилирует и выполняет текст программы на языке Перфолента. 
        /// В тексте программы должна присутствовать процедура или функция Старт. 
        /// В случае, если метод Старт это процедура, возвращается значение Неопределено. 
        /// Этот метод компилирует указанный скрипт при каждом запуске, поэтому его не рекомендуется использовать в циклах.
        /// </summary>
        /// <param name="ScriptText">тип IValue - текст скрипта, который будет компилирован и выполнен.</param>
        /// <param name="ParamArray">тип IValue - параметр процедуры или функции Старт. Если необходимо передать больше параметров, передайте Массив.</param>
        /// <returns>
        /// тип Объект - результат выполнения метода Старт.
        /// </returns>
        [ContextMethod("КомпилироватьИВыполнить", "CompileAndExecute")]
        public IValue CompileAndExecute(IValue ScriptText, IValue ParamArray)
        {

            if (ExecuteScriptMethod is null)
            {

                if (CompilerPath is null)
                    CompilerPath = FindCompilerPath();

                if (CompilerAssm is null)
                    CompilerAssm = System.Reflection.Assembly.LoadFrom(CompilerPath);

                Type tp = CompilerAssm.GetType("Промкод.Перфолента.Компилятор");
                System.Reflection.MethodInfo met = tp.GetMethod("ВыполнитьСкрипт");
                ExecuteScriptMethod = (ExecuteMethodDelegate)Delegate.CreateDelegate(typeof(ExecuteMethodDelegate), null, met);

            };

            return ConvertResult(ExecuteScriptMethod(ScriptText.AsString(), ConvertParam(ParamArray)));

        }

        /// <summary>
        /// Компилирует и выполняет текст программы на языке Перфолента из указанного файла. 
        /// В тексте программы должна присутствовать процедура или функция Старт. 
        /// В случае, если метод Старт это процедура, возвращается значение Неопределено. 
        /// Этот метод компилирует указанный скрипт при каждом запуске, поэтому его не рекомендуется использовать в циклах. 
        /// Файл должен содержать текст в кодировке UTF-8.
        /// </summary>
        /// <param name="ScriptFilePath">тип IValue - путь к файлу скрипта, который будет компилирован и выполнен.</param>
        /// <param name="ParamArray">тип IValue - параметр процедуры или функции Старт. Если необходимо передать больше параметров, передайте Массив.</param>
        /// <returns>
        /// тип Объект - результат выполнения метода Старт.
        /// </returns>
        [ContextMethod("КомпилироватьИВыполнитьИзФайла", "CompileAndExecuteFromFile")]
        public IValue CompileAndExecuteFromFile(IValue ScriptFilePath, IValue ParamArray)
        {

            if (ExecuteScriptFromFileMethod is null)
            {

                if (CompilerPath is null)
                    CompilerPath = FindCompilerPath();

                if (CompilerAssm is null)
                    CompilerAssm = System.Reflection.Assembly.LoadFrom(CompilerPath);

                Type tp = CompilerAssm.GetType("Промкод.Перфолента.Компилятор");
                System.Reflection.MethodInfo met = tp.GetMethod("ВыполнитьСкриптИзФайла");
                ExecuteScriptFromFileMethod = (ExecuteMethodDelegate)Delegate.CreateDelegate(typeof(ExecuteMethodDelegate), null, met);

            };

            return ConvertResult(ExecuteScriptFromFileMethod(ScriptFilePath.AsString(), ConvertParam(ParamArray)));

        }

        private void CreateStartMethodDelegates(string ProgramModuleName)
        {
            var isEmpty = String.IsNullOrWhiteSpace(ProgramModuleName);

            if (isEmpty)
                ProgramModuleName = "Программа";

            var ns = NameSpace;
            if (String.IsNullOrWhiteSpace(ns))
            {
                ns = ProgramModuleName;
            }
            else
            {
                ns = ns + "."+ ProgramModuleName;
            };

            StartProcMethod = null;
            StartFuncMethod = null;

            Type tp = ScriptAssm.GetType(ns);
            if (tp is null) {
                if (!isEmpty) throw new System.Exception("Модуль "+ProgramModuleName+" не найден.");
            }
            else
            {

                System.Reflection.MethodInfo met = tp.GetMethod("Старт", BindingFlags.IgnoreCase | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static );

                if (met is null)
                    throw new System.Exception("Метод Старт в модуле " + ProgramModuleName + " не найден.");

                if (met.ReturnType == typeof(void))
                {
                    StartProcMethod = (StartProcDelegate)Delegate.CreateDelegate(typeof(StartProcDelegate), null, met);
                }
                else
                {
                    StartFuncMethod = (StartFuncDelegate)Delegate.CreateDelegate(typeof(StartFuncDelegate), null, met);
                };
            };
        }


        private object ConvertParam(IValue result)
        {
            if (result is null)
            {
                return null;
            }
            else if (result is StringValue)
            {
                return result.AsString();
            }
            else if (result is NumberValue)
            {
                return result.AsNumber();
            }
            else if (result is BooleanValue)
            {
                return result.AsBoolean();
            }
            else if (result is DateValue)
            {
                return result.AsDate();
            }
            else
            {
                return result.AsObject();
            };
        }

        private IValue ConvertResult(object result)
        {
            if (result is null)
            {
                return ValueFactory.CreateNullValue();
            }
            else if (result is string)
            {
                return ValueFactory.Create((string)result);
            }
            else if (result is Decimal)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is float)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is double)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is long)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is ulong)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is uint)
            {
                return ValueFactory.Create((Decimal)result);
            }
            else if (result is int)
            {
                return ValueFactory.Create((int)result);
            }
            else if (result is byte)
            {
                return ValueFactory.Create((int)result);
            }
            else if (result is bool)
            {
                return ValueFactory.Create((bool)result);
            }
            else if (result is DateTime)
            {
                return ValueFactory.Create((DateTime)result);
            }
            else if (result is IRuntimeContextInstance)
            {
                return ValueFactory.Create((IRuntimeContextInstance)result);
            }
            else
            {
                return COMWrapperContext.Create(result);
            };
        }

        private string FindCompilerPath() {

            //найдем путь к компилятору
            var pflexe = "\\pflc.exe";

            //сначала ищем в каталоге стартового скрипта
            var CodeSource = GlobalsManager.GetGlobalContext<SystemGlobalContext>().CodeSource;
            var scrctx = new ScriptInformationContext(CodeSource);
            var path = scrctx.Path + pflexe;
            if (File.Exists(path))
            {
                return path;
            };

            //теперь в каталоге этой библиотеки
            var StartUpPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            StartUpPath = new Uri(StartUpPath).LocalPath;
            path = StartUpPath + pflexe;
            if (File.Exists(path)) {
                return path;
            };

            string programFiles = Environment.ExpandEnvironmentVariables("%ProgramW6432%");
            //string programFilesX86 = Environment.ExpandEnvironmentVariables("%ProgramFiles(x86)%");
            
            //теперь по стандартному пути инсталляции перфоленты
            StartUpPath = programFiles + "\\Promcod\\Perfolenta";
            path = StartUpPath + pflexe;
            if (File.Exists(path))
            {
                return path;
            };

            EnvironmentVariableTarget[] evts = new EnvironmentVariableTarget[] {EnvironmentVariableTarget.Process, EnvironmentVariableTarget.User, EnvironmentVariableTarget.Machine};

            //теперь в переменной окружения PERFOLENTA_HOME
            string varName = "PERFOLENTA_HOME";
            foreach (var evt in evts)
            {
                StartUpPath = Environment.GetEnvironmentVariable(varName, evt);
                if (!(StartUpPath is null))
                {
                    path = StartUpPath + pflexe;
                    if (File.Exists(path))
                    {
                        return path;
                    };
                };
            };

            //теперь в переменной окружения Path
            varName = "Path";
            foreach (var evt in evts)
            {
                StartUpPath = Environment.GetEnvironmentVariable(varName, evt);
                if (!(StartUpPath is null))
                {
                    foreach (var substr in StartUpPath.Split(';'))
                    {
                        path = substr + pflexe;
                        if (File.Exists(path))
                        {
                            return path;
                        };
                    };
                };
            };


            //теперь проверим в реестре, куда инсталлятор должен был записать путь
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Promcod\\Perfolenta.Net"))
                {
                    if (key != null)
                    {
                        Object o = key.GetValue("Install Directory");
                        if (o != null)
                        {
                            StartUpPath = o as String;
                            path = StartUpPath + pflexe;
                            if (File.Exists(path))
                            {
                                return path;
                            };
                        }
                    }
                }
            }
            catch  
            {
                //...
            };


            throw new System.Exception("Не удалось найти путь к компилятору pflc.exe");
        }

    }
}

