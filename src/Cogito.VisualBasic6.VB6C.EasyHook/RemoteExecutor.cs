using System;

namespace Cogito.VisualBasic6.VB6C.EasyHook
{

    /// <summary>
    /// Class to be made available from the remote entry point.
    /// </summary>
    public abstract class RemoteExecutor : MarshalByRefObject
    {

        /// <summary>
        /// Pings the remote application.
        /// </summary>
        public abstract void Ping();

        /// <summary>
        /// Writes some data to standard out.
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteStdOut(string value);

        /// <summary>
        /// Writes some data to standard error.
        /// </summary>
        /// <param name="value"></param>
        public abstract void WriteStdErr(string value);

    }

}
