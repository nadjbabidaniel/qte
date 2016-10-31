
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using System.Security.Principal;
using System.Diagnostics;

// Forward declarations
using LUID = System.Int64;
using HANDLE = System.IntPtr;

namespace QuickTimeEntry
{

    class ProcessUserInfo
    {
        public const int TOKEN_QUERY = 0X00000008;

        const int ERROR_NO_MORE_ITEMS = 259;

        enum TOKEN_INFORMATION_CLASS
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId
        }

        [StructLayout(LayoutKind.Sequential)]
        struct TOKEN_USER
        {
            public _SID_AND_ATTRIBUTES User;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct _SID_AND_ATTRIBUTES
        {
            public IntPtr Sid;
            public int Attributes;
        }
        [DllImport("advapi32")]
        static extern bool OpenProcessToken(
        HANDLE ProcessHandle, // handle to process
        int DesiredAccess, // desired access to process
        ref IntPtr TokenHandle // handle to open access token
        );

        [DllImport("kernel32")]
        static extern HANDLE GetCurrentProcess();

        [DllImport("advapi32", CharSet = CharSet.Auto)]
        static extern bool GetTokenInformation(
        HANDLE hToken,
        TOKEN_INFORMATION_CLASS tokenInfoClass,
        IntPtr TokenInformation,
        int tokeInfoLength,
        ref int reqLength);

        [DllImport("kernel32")]
        static extern bool CloseHandle(HANDLE handle);

        [DllImport("advapi32", CharSet = CharSet.Auto)]
        static extern bool LookupAccountSid
        (
        [In, MarshalAs(UnmanagedType.LPTStr)] string lpSystemName, // name of local or remote computer
        IntPtr pSid, // security identifier
        StringBuilder Account, // account name buffer
        ref int cbName, // size of account name buffer
        StringBuilder DomainName, // domain name
        ref int cbDomainName, // size of domain name buffer
        ref int peUse // SID type
            // ref _SID_NAME_USE peUse // SID type
        );

        [DllImport("advapi32", CharSet = CharSet.Auto)]
        static extern bool ConvertSidToStringSid(
        IntPtr pSID,
        [In, Out, MarshalAs(UnmanagedType.LPTStr)] ref string pStringSid);

        /// <summary>
        /// Collect User Info
        /// </summary>
        /// <param name="pToken">Process Handle</param>
        public static bool DumpUserInfo(HANDLE pToken, out IntPtr SID)
        {
            int Access = TOKEN_QUERY;
            HANDLE procToken = IntPtr.Zero;
            bool ret = false;
            SID = IntPtr.Zero;
            try
            {
                if (OpenProcessToken(pToken, Access, ref procToken))
                {
                    ret = ProcessTokenToSid(procToken, out SID);
                    CloseHandle(procToken);
                }
                return ret;
            }
            catch (Exception err)
            {
                return false;
            }
        }

        private static bool ProcessTokenToSid(HANDLE token, out IntPtr SID)
        {
            TOKEN_USER tokUser;
            const int bufLength = 256;            
            IntPtr tu = Marshal.AllocHGlobal(bufLength);
            bool ret = false;
            SID = IntPtr.Zero;
            try
            {
                int cb = bufLength;
                ret = GetTokenInformation(token,
                        TOKEN_INFORMATION_CLASS.TokenUser, tu, cb, ref cb);
                if (ret)
                {
                    tokUser = (TOKEN_USER)Marshal.PtrToStructure(tu, typeof(TOKEN_USER));
                    SID = tokUser.User.Sid;
                }
                return ret;
            }
            catch (Exception err)
            {
                return false;
            }
            finally
            {
                Marshal.FreeHGlobal(tu);
            }
        }

        public static string ExGetProcessInfoByPID
            (int PID, out string SID)//, out string OwnerSID)
        {
            IntPtr _SID = IntPtr.Zero;
            SID = String.Empty;
            try
            {
                Process process = Process.GetProcessById(PID);
                if (DumpUserInfo(process.Handle, out _SID))
                {
                    ConvertSidToStringSid(_SID, ref SID);
                }
                return process.ProcessName;
            }
            catch
            {
                return "Unknown";
            }
        }

        /*
        public static void Main()
        {
            string processName = Process.GetCurrentProcess().ProcessName;
            Process[] myProcesses = Process.GetProcessesByName(processName);
            if (myProcesses.Length == 0)
                Console.WriteLine("Could not find notepad processes on remote machine");
            foreach (Process myProcess in myProcesses)
            {
                Console.Write("Process Name : " + myProcess.ProcessName + " Process ID : "
                + myProcess.Id + " MachineName : " + myProcess.MachineName + "\n");
                DumpUserInfo(myProcess.Handle);
            }
        }

        static void DumpUserInfo(HANDLE pToken)
        {
            int Access = TOKEN_QUERY;
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("\nToken dump performed on {0}\n\n", DateTime.Now);
            HANDLE procToken = IntPtr.Zero;
            if (OpenProcessToken(pToken, Access, ref procToken))
            {
                sb.Append("Process Token:\n");
                sb.Append(PerformDump(procToken));
                CloseHandle(procToken);
            }
            Console.WriteLine(sb.ToString());
        }
        static StringBuilder PerformDump(HANDLE token)
        {
            StringBuilder sb = new StringBuilder();
            TOKEN_USER tokUser;
            const int bufLength = 256;
            IntPtr tu = Marshal.AllocHGlobal(bufLength);
            int cb = bufLength;
            GetTokenInformation(token, TOKEN_INFORMATION_CLASS.TokenUser, tu, cb, ref cb);
            tokUser = (TOKEN_USER)Marshal.PtrToStructure(tu, typeof(TOKEN_USER));
            sb.Append(DumpAccountSid(tokUser.User.Sid));
            Marshal.FreeHGlobal(tu);
            return sb;
        }

        static string DumpAccountSid(IntPtr SID)
        {
            int cchAccount = 0;
            int cchDomain = 0;
            int snu = 0;
            StringBuilder sb = new StringBuilder();

            // Caller allocated buffer
            StringBuilder Account = null;
            StringBuilder Domain = null;
            bool ret = LookupAccountSid(null, SID, Account, ref cchAccount, Domain, ref cchDomain, ref snu);
            if (ret == true)
                if (Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                    return "Error";
            try
            {
                Account = new StringBuilder(cchAccount);
                Domain = new StringBuilder(cchDomain);
                ret = LookupAccountSid(null, SID, Account, ref cchAccount, Domain, ref cchDomain, ref snu);
                if (ret)
                {
                    sb.Append(Domain);
                    sb.Append(@"\\");
                    sb.Append(Account);
                }
                else
                    Console.WriteLine("logon account (no name) ");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
            }
            string SidString = null;
            ConvertSidToStringSid(SID, ref SidString);
            sb.Append("\nSID: ");
            sb.Append(SidString);
            return sb.ToString();
        }
         * 
         * */

    }

}