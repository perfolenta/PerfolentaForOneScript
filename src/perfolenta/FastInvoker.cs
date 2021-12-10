using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Reflection.Emit;


namespace perfolenta
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="target">тип Object - The target.</param>
    /// <param name="paramters">тип Object[] - The paramters.</param>
    /// <returns></returns>
    public delegate object FastInvokeHandler(object
                    target, object[] paramters);


    /// <summary>
    /// 
    /// </summary>
    public class FastInvoke
    {
        FastInvokeHandler MyDelegate;
        
        /// <summary>
        /// My method information
        /// </summary>
        public MethodInfo MyMethodInfo;

        /// <summary>
        /// My parameters
        /// </summary>
        public ParameterInfo[] MyParameters;

        Object HostObject;

        /// <summary>
        /// The number of arguments
        /// </summary>
        public int NumberOfArguments;

        /// <summary>
        /// Создаёт экземпляр объекта класса <see cref="FastInvoke"/>.
        /// </summary>
        /// <param name="MyObject">тип Object - My object.</param>
        /// <param name="MyName">тип String - My name.</param>
        public FastInvoke(Object MyObject, String MyName)
        {
            HostObject = MyObject;
            Type t2 = MyObject.GetType();
            MethodInfo m2 = t2.GetMethod(MyName);
            MyDelegate = GetMethodInvoker(m2);
            NumberOfArguments = m2.GetParameters().Length;
            MyMethodInfo = m2;
            MyParameters = m2.GetParameters();
        }

        /// <summary>
        /// Создаёт экземпляр объекта класса <see cref="FastInvoke"/>.
        /// </summary>
        /// <param name="MyObjectType">тип Type - Type of object.</param>
        /// <param name="MyName">тип String - My name.</param>
        public FastInvoke(Type MyObjectType, String MyName)
        {
            HostObject = null;
            MethodInfo m2 = MyObjectType.GetMethod(MyName, BindingFlags.IgnoreCase | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static);
            MyDelegate = GetMethodInvoker(m2);
            NumberOfArguments = m2.GetParameters().Length;
            MyMethodInfo = m2;
            MyParameters = m2.GetParameters();
        }

        /// <summary>
        /// Executes the delegate.
        /// </summary>
        /// <param name="FunctionParameters">тип Object[] - The function parameters.</param>
        /// <returns></returns>
        public object ExecuteDelegate(object[] FunctionParameters)
        {
            try
            {
                return (MyDelegate(HostObject, FunctionParameters));
            }
            catch (Exception e)
            {
                Object o = new Object();
                o = e.Message;
                return (o);

            }

        }

        private FastInvokeHandler GetMethodInvoker(MethodInfo methodInfo)
        {
            DynamicMethod dynamicMethod = new DynamicMethod(string.Empty,
                          typeof(object), new Type[] { typeof(object),
                          typeof(object[]) },
                          methodInfo.DeclaringType.Module);
            ILGenerator il = dynamicMethod.GetILGenerator();

            ParameterInfo[] ps = methodInfo.GetParameters();

            Type[] paramTypes = new Type[ps.Length];
            for (int i = 0; i < paramTypes.Length; i++)
            {
                if (ps[i].ParameterType.IsByRef)
                    paramTypes[i] = ps[i].ParameterType.GetElementType();
                else
                    paramTypes[i] = ps[i].ParameterType;
            }
            LocalBuilder[] locals = new LocalBuilder[paramTypes.Length];

            //создаем переменные
            for (int i = 0; i < paramTypes.Length; i++)
            {
                locals[i] = il.DeclareLocal(paramTypes[i], true);
            }

            for (int i = 0; i < paramTypes.Length; i++)
            {
                il.Emit(OpCodes.Ldarg_1);
                EmitFastInt(il, i);
                il.Emit(OpCodes.Ldelem_Ref);
                EmitCastToReference(il, paramTypes[i]);
                il.Emit(OpCodes.Stloc, locals[i]);
            }

            if (!methodInfo.IsStatic)
            {
                il.Emit(OpCodes.Ldarg_0);
            }

            for (int i = 0; i < paramTypes.Length; i++)
            {
                if (ps[i].ParameterType.IsByRef)
                    il.Emit(OpCodes.Ldloca_S, locals[i]);
                else
                    il.Emit(OpCodes.Ldloc, locals[i]);
            }

            //вызываем метод
            if (methodInfo.IsStatic)
                il.EmitCall(OpCodes.Call, methodInfo, null);
            else
                il.EmitCall(OpCodes.Callvirt, methodInfo, null);

            //уст. возвращаемое значение
            if (methodInfo.ReturnType == typeof(void))
                il.Emit(OpCodes.Ldnull);
            else
                EmitBoxIfNeeded(il, methodInfo.ReturnType);

            //значения передаваемые по ссылке загружаем обратно в массив
            for (int i = 0; i < paramTypes.Length; i++)
            {
                if (ps[i].ParameterType.IsByRef)
                {
                    il.Emit(OpCodes.Ldarg_1);
                    EmitFastInt(il, i);
                    il.Emit(OpCodes.Ldloc, locals[i]);
                    if (locals[i].LocalType.IsValueType)
                        il.Emit(OpCodes.Box, locals[i].LocalType);
                    il.Emit(OpCodes.Stelem_Ref);
                }
            }

            //генерируем команду возврата
            il.Emit(OpCodes.Ret);

            //возвращаем готовый динамический метод
            FastInvokeHandler invoder = (FastInvokeHandler)
               dynamicMethod.CreateDelegate(typeof(FastInvokeHandler));
            return invoder;
        }

        private static void EmitCastToReference(ILGenerator il,
                                                System.Type type)
        {
            if (type.IsValueType)
            {
                il.Emit(OpCodes.Unbox_Any, type);
            }
            else
            {
                il.Emit(OpCodes.Castclass, type);
            }
        }

        private static void EmitBoxIfNeeded(ILGenerator il, System.Type type)
        {
            if (type.IsValueType)
            {
                il.Emit(OpCodes.Box, type);
            }
        }

        private static void EmitFastInt(ILGenerator il, int value)
        {
            switch (value)
            {
                case -1:
                    il.Emit(OpCodes.Ldc_I4_M1);
                    return;
                case 0:
                    il.Emit(OpCodes.Ldc_I4_0);
                    return;
                case 1:
                    il.Emit(OpCodes.Ldc_I4_1);
                    return;
                case 2:
                    il.Emit(OpCodes.Ldc_I4_2);
                    return;
                case 3:
                    il.Emit(OpCodes.Ldc_I4_3);
                    return;
                case 4:
                    il.Emit(OpCodes.Ldc_I4_4);
                    return;
                case 5:
                    il.Emit(OpCodes.Ldc_I4_5);
                    return;
                case 6:
                    il.Emit(OpCodes.Ldc_I4_6);
                    return;
                case 7:
                    il.Emit(OpCodes.Ldc_I4_7);
                    return;
                case 8:
                    il.Emit(OpCodes.Ldc_I4_8);
                    return;
            }

            if (value > -129 && value < 128)
            {
                il.Emit(OpCodes.Ldc_I4_S, (SByte)value);
            }
            else
            {
                il.Emit(OpCodes.Ldc_I4, value);
            }
        }

        /// <summary>
        /// Types the convert.
        /// </summary>
        /// <param name="source">тип Object - The source.</param>
        /// <param name="DestType">Type of the dest.</param>
        /// <returns></returns>
        public object TypeConvert(object source, Type DestType)
        {

            object NewObject = System.Convert.ChangeType(source, DestType);

            return (NewObject);
        }

    }
}

