using System;
using System.Runtime.InteropServices;
using System.Threading;

using EasyHook;

namespace Cogito.VisualBasic6.VB6C.EasyHook
{

    /// <summary>
    /// Instantiated within the VB6.exe process.
    /// </summary>
    public class RemoteEntryPoint : IEntryPoint
    {

        public const string USER32_DLL = "user32.dll";
        public const string ADVAPI32_DLL = "advapi32.dll";
        public const string OLE32_DLL = "ole32.dll";

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Ansi)]
        delegate bool DMessageBeep(uint beepType);

        [DllImport(USER32_DLL, SetLastError = true)]
        static extern bool MessageBeep(uint beepType);

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Auto)]
        delegate int DCoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        [DllImport(OLE32_DLL, CharSet = CharSet.Auto, SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern int CoInitializeEx([In, Optional] IntPtr pvReserved, [In] uint dwCoInit);

        readonly RemoteExecutor executor;

        /// <summary>
        /// Initializes a new instance.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="channelName"></param>
        public RemoteEntryPoint(RemoteHooking.IContext context, string channelName)
        {
            executor = RemoteHooking.IpcConnectClient<RemoteExecutor>(channelName);
            executor.Ping();
        }

        /// <summary>
        /// Invoked inside of the VB6 executable.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="channelName"></param>
        public void Run(RemoteHooking.IContext context, string channelName)
        {
            var messageBeep = LocalHook.Create(
                LocalHook.GetProcAddress(USER32_DLL, nameof(MessageBeep)),
                new DMessageBeep(MessageBeepHook),
                this);
            messageBeep.ThreadACL.SetExclusiveACL(new[] { 0 });

            var coInitializeEx = LocalHook.Create(
                LocalHook.GetProcAddress(OLE32_DLL, nameof(CoInitializeEx)),
                new DCoInitializeEx(CoInitializeExHook),
                this);
            coInitializeEx.ThreadACL.SetExclusiveACL(new[] { 0 });

            RemoteHooking.WakeUpProcess();

            try
            {
                // wait for process exit
                while (true)
                {
                    Thread.Sleep(500);
                    executor.Ping();
                }
            }
            catch
            {
                // server failed to respond
            }

            messageBeep.Dispose();
            LocalHook.Release();
        }

        int CoInitializeExHook(IntPtr pvReserved, uint dwCoInit)
        {
            var hr = CoInitializeEx(pvReserved, dwCoInit);
            ActivateComCtx();
            return hr;
        }

        /// <summary>
        /// Activates the appropriate COM context on the default thread.
        /// </summary>
        void ActivateComCtx()
        {
            executor.WriteStdErr("Activating new COM ctx...\n");
        }

        /// <summary>
        /// Silences the beeps.
        /// </summary>
        /// <param name="beepType"></param>
        /// <returns></returns>
        bool MessageBeepHook(uint beepType)
        {
            return true;
        }

    }

}
