using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Web;

namespace CreateUser.App_Codes
{
    static class supportClass
    {
        //static Hashtable ListeFolderPouPaEffacer;
        #region temp folder

        /// <summary>
        /// Get the Application Guid
        /// </summary>
        public static Guid AppGuid
        {
            get
            {
                Assembly asm = Assembly.GetEntryAssembly();
                object[] attr = (asm.GetCustomAttributes(typeof(GuidAttribute), true));
                return new Guid((attr[0] as GuidAttribute).Value);
            }
        }
        /// <summary>
        /// Get the current assembly Guid.
        /// <remarks>
        /// Note that the Assembly Guid is not necessarily the same as the
        /// Application Guid - if this code is in a DLL, the Assembly Guid
        /// will be the Guid for the DLL, not the active EXE file.
        /// </remarks>
        /// </summary>
        public static Guid AssemblyGuid
        {
            get
            {
                Assembly asm = Assembly.GetExecutingAssembly();
                object[] attr = (asm.GetCustomAttributes(typeof(GuidAttribute), true));
                return new Guid((attr[0] as GuidAttribute).Value);
            }
        }
        /// <summary>
        /// Get the current user data folder
        /// </summary>
        public static string UserDataFolder
        {
            get
            {
                Guid appGuid = AppGuid;
                string folderBase = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string dir = string.Format(@"{0}\{1}\", folderBase, appGuid.ToString("B").ToUpper());
                return CheckDir(dir);
            }
        }
        /// <summary>
        /// Get the current user roaming data folder
        /// </summary>
        public static string UserRoamingDataFolder
        {
            get
            {
                Guid appGuid = AppGuid;
                string folderBase = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string dir = string.Format(@"{0}\{1}\", folderBase, appGuid.ToString("B").ToUpper());
                return CheckDir(dir);
            }
        }
        /// <summary>
        /// Get all users data folder
        /// </summary>
        public static string AllUsersDataFolder
        {
            get
            {
                Guid appGuid = AppGuid;
                string folderBase = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
                string dir = string.Format(@"{0}\{1}\", folderBase, appGuid.ToString("B").ToUpper());
                return CheckDir(dir);
            }
        }
        /// <summary>
        /// Check the specified folder, and create if it doesn't exist.
        /// </summary>
        /// <param name="dir"></param>
        /// <returns></returns>
        private static string CheckDir(string dir)
        {
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }
        #endregion
        #region temp impersonation
        public static String sUserName = "administrator", sDomain = "uespoiretudiant", sPass = "allglorytoGod123!!";
        public const int LOGON32_LOGON_INTERACTIVE = 2;
        public const int LOGON32_PROVIDER_DEFAULT = 0;

        static WindowsImpersonationContext impersonationContext;

        [DllImport("advapi32.dll")]

        public static extern int LogonUserA(String lpszUserName,
            String lpszDomain,
            String lpszPassword,
            int dwLogonType,
            int dwLogonProvider,
            ref IntPtr phToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]

        public static extern int DuplicateToken(IntPtr hToken,
            int impersonationLevel,
            ref IntPtr hNewToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]

        public static extern bool RevertToSelf();

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]

        public static extern bool CloseHandle(IntPtr handle);
        #endregion
        #region connString
        public static string Messages = String.Empty;
        // Sa se pou UEspoir
        static string sSqlConnString = @"Data Source=tcp:CCPAP1;Initial Catalog=UEspoirDB;Persist Security Info=True;User ID=amen;Password=amen123!!";

        // Sa se pou CCPS: Gen youn nan nou  ki te gen tach mete l nan GUI wi
        // static string sSqlConnString = @"Data Source=tcp:CCPAP1;Initial Catalog=CCPAP_Web_Edu;Persist Security Info=True;User ID=amen;Password=amen123!!";

        #endregion
    }
}