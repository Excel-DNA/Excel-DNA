//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq.Expressions;
using System.Linq;
using ExcelDna.Integration;

namespace ExcelDna.Loader
{
    static class XlDirectMarshal
    {
        // Not cleaning up - we don't expect this to use a lot of memory, and think the pool of Excel calculation threads is stable
        readonly static ThreadLocal<XlMarshalContext> MarshalContext = new ThreadLocal<XlMarshalContext>(() => new XlMarshalContext());
        public static XlMarshalContext GetMarshalContext() => MarshalContext.Value;

        // This method is only called via AutoFree (for an instance ?)
        // We assume it runs on the same thread as the function call that returned the free-bit value
        public static void FreeMemory()
        {
            GetMarshalContext().FreeMemory();
        }

        // NOTE: This is called in parallel, from a ThreadPool thread
        public static void SetDelegateAndFunctionPointer(XlMethodInfo methodInfo)
        {
            //var xlDelegate = GetNativeDelegate(methodInfo);
            var xlDelegate = GetLazyDelegate(methodInfo);
            methodInfo.DelegateHandle = GCHandle.Alloc(xlDelegate);
            if (NativeAOT.IsActive)
                methodInfo.FunctionPointer = MarshalNativeAOT.GetFunctionPointerForDelegate(xlDelegate, methodInfo);
            else
                methodInfo.FunctionPointer = Marshal.GetFunctionPointerForDelegate(xlDelegate);
        }

        // NOTE: This is called in parallel, from a ThreadPool thread
        static Delegate GetNativeDelegate(XlMethodInfo methodInfo)
        {
            // We convert
            //
            // double MyFancyFunc(object input1, string input2) { ... }
            //
            // To
            //
            // IntPtr MyFancyFuncWrapped(IntPtr<XlOper12> xlinput0, IntPtr<XlString12 xlinput1)
            // {
            //
            //    XlMarshalContext ctx = GetContextForThisThread();
            //    try
            //    {
            //  
            //      input0 = Convert1(xlinput0);
            //      input1 = Convert2(xlinput1);
            //
            //      result = MyFancyFunc(input0, input1);
            //      xlresult = ctx.ConvertRet(ctx, result),
            //      return xlresult;
            //    }
            //    catch(Exception ex)
            //    {
            //       resultx = HandleEx(ex);
            //       xlresultx = ctx.ConvertRet(resultex);
            //    }
            //
            //    return resultx
            //
            //
            // }

            // Create the new parameters and return value for the wrapper
            // TODO/DM: Consolidate to a single select
            var outerParams = methodInfo.Parameters.Select(p => Expression.Parameter(typeof(IntPtr), p.Name)).ToArray();
            var innerParamExprs = new Expression[outerParams.Length];

            for (int i = 0; i < methodInfo.Parameters.Length; i++)
            {
                var pi = methodInfo.Parameters[i];
                if (pi.XlMarshalConvert == null)
                {
                    innerParamExprs[i] = outerParams[i];
                }
                else
                {
                    innerParamExprs[i] = Expression.Call(pi.XlMarshalConvert, outerParams[i]);
                }
            }

            Expression innerCall;
            if (methodInfo.MethodInfo != null) // Method and optional Target
            {
                innerCall = Expression.Call(methodInfo.Target == null ? null : Expression.Constant(methodInfo.Target), methodInfo.MethodInfo, innerParamExprs);
            }
            else // LambdaExpression
            {
                innerCall = Expression.Invoke(methodInfo.LambdaExpression, innerParamExprs);
            }

            // variable to hold XlMarshalContext
            var ctx = Expression.Variable(typeof(XlMarshalContext), "xlMarshalContext");
            var assignCtx = Expression.Assign(ctx, Expression.Call(typeof(XlDirectMarshal), nameof(XlDirectMarshal.GetMarshalContext), null));
            if (methodInfo.HasReturnType)
            {
                var result = Expression.Variable(typeof(IntPtr), "returnValue");
                var resultExpr = result; // Overwrite with conversion if applicable
                if (methodInfo.ReturnType.XlMarshalConvert != null)
                {
                    innerCall = Expression.Call(ctx, methodInfo.ReturnType.XlMarshalConvert, innerCall);
                }
            }

            // Prepare the ex(ception) local variable
            var ex = Expression.Variable(typeof(Exception), "ex");
            Expression catchExpression;
            if (methodInfo.IsExcelAsyncFunction)
            {
                // HandleUnhandledException is called by ExcelAsyncHandle.SetException
                var asyncHandle = Expression.TypeAs(innerParamExprs.Last(), typeof(ExcelAsyncHandleNative));
                var setExceptionMethod = typeof(ExcelAsyncHandleNative).GetMethod(nameof(ExcelAsyncHandleNative.SetException));
                // Need to get instance from parameter list
                catchExpression = Expression.Block(
                                        Expression.Call(asyncHandle, setExceptionMethod, ex), // Ignore bool return !?
                                        Expression.Empty());
            }
            else
            {
                var handlerMethod = typeof(ExcelIntegration).GetMethod(nameof(ExcelIntegration.HandleUnhandledException), System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);
                var exHandler = Expression.Call(handlerMethod, ex);
                if (methodInfo.HasReturnType)
                {
                    if (methodInfo.ReturnType.XlType == XlTypes.Xloper)
                    {
                        // We return whatever the result is from the unhandled exception handler
                        catchExpression = Expression.Call(ctx, XlMarshalConversions.ObjectReturn, exHandler);
                    }
                    else
                    {
                        // We return #NUM!, which is better than crashing
                        catchExpression = Expression.Block(exHandler, Expression.Constant(IntPtr.Zero));
                    }
                }
                else
                {
                    catchExpression = Expression.Block(exHandler, Expression.Empty());
                }
            }

            Type delegateType;
            Expression body;
            if (methodInfo.HasReturnType)
            {
                delegateType = XlDirectMarshalTypes.XlFuncs[methodInfo.Parameters.Length];
                body = Expression.Block(
                    typeof(IntPtr),
                    new ParameterExpression[] { ctx },
                    assignCtx,
                    Expression.TryCatch(
                        innerCall,
                        Expression.Catch(ex, catchExpression)));
            }
            else
            {
                delegateType = XlDirectMarshalTypes.XlActs[methodInfo.Parameters.Length];
                if (methodInfo.IsExcelAsyncFunction)
                {
                    body = Expression.Block(
                        Expression.TryCatch(
                            innerCall,
                            Expression.Catch(ex, catchExpression)));
                }
                else
                {
                    body = Expression.TryCatch(
                            innerCall,
                            Expression.Catch(ex, catchExpression));
                }
            }

            return Expression.Lambda(delegateType, body, methodInfo.Name, outerParams).Compile();
        }

        static Delegate GetLazyDelegate(XlMethodInfo methodInfo)
        {
            var lazyLambda = new XlDirectMarshalLazy(() => GetNativeDelegate(methodInfo));

            // now we need to return the right method from lazyLambda, to be sure it can be assigned to the intended delegate type
            if (methodInfo.HasReturnType)
            {
                var delegateType = XlDirectMarshalTypes.XlFuncs[methodInfo.Parameters.Length];
                var method = typeof(XlDirectMarshalLazy).GetMethod($"Func{methodInfo.Parameters.Length}");
                return Delegate.CreateDelegate(delegateType, lazyLambda, method);
            }
            else
            {
                var delegateType = XlDirectMarshalTypes.XlActs[methodInfo.Parameters.Length];
                var method = typeof(XlDirectMarshalLazy).GetMethod($"Act{methodInfo.Parameters.Length}");
                return Delegate.CreateDelegate(delegateType, lazyLambda, method);
            }
        }
    }
}
