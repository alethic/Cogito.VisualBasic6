using System;
using System.ComponentModel;
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

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Auto)]
        delegate void DCoUninitialize();

        [DllImport(OLE32_DLL, CharSet = CharSet.Auto, SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern void CoUninitialize();

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        delegate void DExitProcess(uint uExitCode);

        [DllImport(KERNEL32_DLL, CallingConvention = CallingConvention.StdCall)]
        static extern int ExitProcess(uint uExitCode);

        readonly RemoteExecutor executor;
        readonly IntPtr hActCtx;

        [ThreadStatic]
        static IntPtr lpCookie;

        /// <summary>
        /// Initializes a new instance.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="channelName"></param>
        public RemoteEntryPoint(RemoteHooking.IContext context, string channelName, string manifest)
        {
            executor = RemoteHooking.IpcConnectClient<RemoteExecutor>(channelName);
            executor.Ping();

            if (manifest != null)
            {
                var actCtx = new ActCtx.ACTCTX();
                actCtx.cbSize = Marshal.SizeOf(typeof(ActCtx.ACTCTX));
                actCtx.dwFlags = ActCtx.ACTCTX_FLAG_SET_PROCESS_DEFAULT;
                actCtx.lpSource = manifest;

                hActCtx = ActCtx.CreateActCtx(ref actCtx);
                if (hActCtx == new IntPtr(-1))
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to create activation context.");
            }
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

            var coUninitialize = LocalHook.Create(
                LocalHook.GetProcAddress(OLE32_DLL, nameof(CoUninitialize)),
                new DCoUninitialize(CoUninitializeHook),
                this);
            coUninitialize.ThreadACL.SetExclusiveACL(new[] { 0 });

            var exitProcess = LocalHook.Create(
                LocalHook.GetProcAddress(KERNEL32_DLL, nameof(ExitProcess)),
                new DExitProcess(ExitProcessHook),
                this);
            exitProcess.ThreadACL.SetExclusiveACL(new[] { 0 });

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
            coInitializeEx.Dispose();
            coUninitialize.Dispose();
            exitProcess.Dispose();
            LocalHook.Release();
        }

        /// <summary>
        /// Invoked when COM is initialized on a thread.
        /// </summary>
        /// <param name="pvReserved"></param>
        /// <param name="dwCoInit"></param>
        /// <returns></returns>
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
            if (hActCtx != IntPtr.Zero)
            {
                executor.WriteStdOut("Activating new COM ctx...\n");

                if (ActCtx.ActivateActCtx(hActCtx, out lpCookie) != true)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to activate activation context.");
            }
        }

        /// <summary>
        /// Invoked when COM is uninitialized on a thread.
        /// </summary>
        /// <returns></returns>
        void CoUninitializeHook()
        {
            DeactivateComCtx();
            CoUninitialize();
        }

        /// <summary>
        /// Deactivates the COM context.
        /// </summary>
        void DeactivateComCtx()
        {
            if (lpCookie != IntPtr.Zero)
            {
                executor.WriteStdOut("Deactivating COM ctx...\n");

                if (ActCtx.DeactivateActCtx(0, lpCookie) != true)
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "Failed to deactivate activation context.");

                lpCookie = IntPtr.Zero;
            }
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
