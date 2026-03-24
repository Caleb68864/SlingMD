using System;
using System.Runtime.InteropServices;

namespace SlingMD.Outlook.Helpers
{
    /// <summary>
    /// Wraps COM interop calls with structured exception handling and logging.
    /// </summary>
    public static class SafeComAction
    {
        /// <summary>
        /// Executes a COM call that returns a value, returning defaultValue on any exception.
        /// </summary>
        public static T Execute<T>(Func<T> action, string context, T defaultValue)
        {
            try
            {
                return action();
            }
            catch (COMException ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
                return defaultValue;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
                return defaultValue;
            }
        }

        /// <summary>
        /// Executes a COM void call, swallowing and logging any exception.
        /// </summary>
        public static void Execute(Action action, string context)
        {
            try
            {
                action();
            }
            catch (COMException ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Executes a COM call returning a COM object. On failure, releases the object if non-null before returning defaultValue.
        /// </summary>
        public static T ExecuteAndRelease<T>(Func<T> action, string context, T defaultValue) where T : class
        {
            T result = null;
            try
            {
                result = action();
                return result;
            }
            catch (COMException ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
                if (result != null)
                {
                    Marshal.ReleaseComObject(result);
                }
                return defaultValue;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"{context}: {ex.Message}", ex);
                if (result != null)
                {
                    Marshal.ReleaseComObject(result);
                }
                return defaultValue;
            }
        }
    }
}
