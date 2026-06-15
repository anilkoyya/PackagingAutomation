using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace LaunchAsUser
{
    class Program
    {
        const int LOGON_WITH_PROFILE = 0x00000001;
        const int CREATE_UNICODE_ENVIRONMENT = 0x00000400;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct STARTUPINFO
        {
            public int cb;
            public string lpReserved;
            public string lpDesktop;
            public string lpTitle;
            public int dwX;
            public int dwY;
            public int dwXSize;
            public int dwYSize;
            public int dwXCountChars;
            public int dwYCountChars;
            public int dwFillAttribute;
            public int dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PROCESS_INFORMATION
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public int dwProcessId;
            public int dwThreadId;
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool CreateProcessWithLogonW(
            string lpUsername,
            string lpDomain,
            string lpPassword,
            int dwLogonFlags,
            string lpApplicationName,
            string lpCommandLine,
            int dwCreationFlags,
            IntPtr lpEnvironment,
            string lpCurrentDirectory,
            ref STARTUPINFO lpStartupInfo,
            out PROCESS_INFORMATION lpProcessInfo);

        [DllImport("kernel32.dll")]
        static extern uint WaitForSingleObject(
            IntPtr hHandle,
            uint dwMilliseconds);

        [DllImport("kernel32.dll")]
        static extern bool CloseHandle(IntPtr hObject);

        static int Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine(
                    "Usage: LaunchAsUser.exe <domain> <user> <password> <command>");
                return 1;
            }

            string domain = args[0];
            string user = args[1];
            string password = args[2];
            string command = args[3];

            var si = new STARTUPINFO();
            si.cb = Marshal.SizeOf(si);

            PROCESS_INFORMATION pi;

            bool result = CreateProcessWithLogonW(
                user,
                domain,
                password,
                LOGON_WITH_PROFILE,
                null,
                command,
                CREATE_UNICODE_ENVIRONMENT,
                IntPtr.Zero,
                Environment.CurrentDirectory,
                ref si,
                out pi);

            if (!result)
            {
                throw new Win32Exception(
                    Marshal.GetLastWin32Error());
            }

            Console.WriteLine($"Started PID: {pi.dwProcessId}");

            WaitForSingleObject(pi.hProcess, 0xFFFFFFFF);

            CloseHandle(pi.hThread);
            CloseHandle(pi.hProcess);

            return 0;
        }
    }
}
