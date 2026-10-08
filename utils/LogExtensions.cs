using log4net;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace it.sealink.DocumentProcessing.utils
{
    internal static class LogExtensions
    {
        private static ILog defaultLogger = LogManager.GetLogger("LogExtensions");
        static public ILog GetClassLogger()
        {
            try
            {
                Type? theType = GetPreviousType(2);
                if (theType is not null)
                    return LogManager.GetLogger(theType);
            }
            catch (Exception ex)
            {
                defaultLogger.Error($"GetClassLogger failed: {ex.ToString()}");
            }
            return defaultLogger;
        }

        static public void MethodEnter(this ILog logger)
        {
            try
            {
                if (logger is not null && logger.IsDebugEnabled)
                {
                    MethodBase? theMethod = GetPreviousMethod(2);
                    if (theMethod is not null)
                        logger.DebugFormat("-->> START {0} >>>>>>>>>>>", theMethod.Name);
                }
            }
            catch (Exception ex)
            {
                logger.Error("MethodEnter", ex);
            }
        }

        static public void MethodLeave(this ILog logger)
        {
            try
            {
                if (logger is not null && logger.IsDebugEnabled)
                {
                    MethodBase? theMethod = GetPreviousMethod(2);
                    if (theMethod is not null)
                        logger.DebugFormat("<<-- END {0} <<<<<<<<<<<<<", theMethod.Name);
                }
            }
            catch (Exception ex)
            {
                logger.Error("MethodLeave", ex);
            }
        }

        #region internal (private, internal, protected) section ---------------

        private static MethodBase? GetPreviousMethod(int count)
        {
            try
            {
                return new StackFrame(count).GetMethod();
            }
            catch (System.Exception) { }

            return null;
        }

        private static Type? GetPreviousType(int count)
        {
            try
            {
                return (new StackFrame(count))?.GetMethod()?.DeclaringType;
            }
            catch (System.Exception) { }

            return null;
        }

        #endregion internal (private, internal, protected) section ------------
    }
}
