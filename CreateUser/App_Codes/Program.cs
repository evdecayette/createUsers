using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Web;

namespace CreateUser.App_Codes
{
    public class Program
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
        protected static string Messages = String.Empty;
        // Sa se pou UEspoir
        static string sSqlConnString = @"Data Source=tcp:CCPAP1;Initial Catalog=UEspoirDB;Persist Security Info=True;User ID=amen;Password=amen123!!";

        // Sa se pou CCPS: Gen youn nan nou  ki te gen tach mete l nan GUI wi
        // static string sSqlConnString = @"Data Source=tcp:CCPAP1;Initial Catalog=CCPAP_Web_Edu;Persist Security Info=True;User ID=amen;Password=amen123!!";

        #endregion


        static String PASSWORD_TEXT = "jesus123";
        static int iIndex = 0;
        static void Main(string[] args)
        {
            if (impersonateValidUser(sUserName, sDomain, sPass))
            {
                //UEspoir: Change Initial Catalog in sSqlconnString - Initial Catalog=UEspoirDB
                // Sa se pou UEspoir
                CreateUsersFromDataBase(string.Format("SELECT PersonneID, Prenom, Nom,  UserNameAttribue FROM [UEspoirDB].[dbo].[Personnes] " +
                " WHERE actif = 1 AND ISNULL(UserNameAttribue,'') = '' ORDER BY Nom, Prenom"));    //Tous les étudiants du Niveau 1, après inscription   !!!!!!!!!!!!!!!!!!!!!!!

                // Sa se pou CCPS: Initial Catalog=CCPAP_Web_Edu
                //CreateUsersFromDataBase(string.Format("SELECT P.Prenom, P.Nom,P.PersonneID, C.ClasseID, S.SessionID FROM Classes C, Sessions S, EtudiantsCourants E, Personnes P " +
                //"WHERE S.SessionID = E.SessionID AND C.ClasseID = S.ClasseID AND P.PersonneID = E.PersonneID AND C.Categorie = '{0}' " +
                //"AND C.NiveauClasse = 1 AND S.actif = 1 AND ISNULL(P.UserNameAttribue,'') = '' ORDER BY P.Nom, P.Prenom", "Informatique"));

                #region PAS_UTILISER
                //CreateUsersFromFile("D:\\IT\\users.txt");

                //EffacerInternetTempFolders();
                //EffacerExeNanFolderDownloads();
                //EffacerUsersInactifs();
                //EffacerUserFolderIfUserDoesNotExist("c:\\users\\");

                //DisableStudentsUsers();
                //EffacerInternetTempFoldersUserSa("falexandre");         
                //TestingEmail();

                //string path = Path.GetTempFileName();
                //ListWindowsAccounts();
                // MeteEtudiantsNanGroupStudents();
                //ReCreateUsersFromDataBaseAfterRestore();
                //EffacerDossiers();
                //MeteEtudiantsNanGroupStudentsKipalaDeja();
                #endregion 
            }
            Console.WriteLine("Pressez une clé pour terminer le programme ...");
            Console.ReadKey();
            return;
        }

        private static bool impersonateValidUser(String userName, String domain, String password)
        {
            WindowsIdentity tempWindowsIdentity;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;

            if (RevertToSelf())
            {
                if (LogonUserA(userName, domain, password, LOGON32_LOGON_INTERACTIVE,
                    LOGON32_PROVIDER_DEFAULT, ref token) != 0)
                {
                    if (DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                    {
                        tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                        impersonationContext = tempWindowsIdentity.Impersonate();

                        if (impersonationContext != null)
                        {
                            CloseHandle(token);

                            CloseHandle(tokenDuplicate);
                            return true;
                        }
                    }
                }
            }

            if (token != IntPtr.Zero)
                CloseHandle(token);
            if (tokenDuplicate != IntPtr.Zero)
                CloseHandle(tokenDuplicate);
            return false;
        }
        private void undoImpersonation()
        {
            impersonationContext.Undo();
        }
        public static void TestingEmail()
        {
            String sReturnAddressEmail = "uespoir@calvarypap.org";
            String sHost = "mail.calvarypap.org";
            String sPassWord = "";
            int iPort = 26;

            MailAddress mailfrom = new MailAddress(sReturnAddressEmail);
            MailAddress mailto = new MailAddress("spoteau@yahoo.com");
            MailMessage newmsg = new MailMessage(mailfrom, mailto);

            newmsg.Subject = "Sujet OK";
            newmsg.Body = "Text message OK tout...";
            newmsg.IsBodyHtml = true;


            SmtpClient smtp = new SmtpClient(sHost, iPort);
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new NetworkCredential(sReturnAddressEmail, sPassWord);
            smtp.EnableSsl = true;

            smtp.Send(newmsg);
        }
        public static void MeteEtudiantsNanGroupStudentsKipalaDeja()
        {
            string sUserName = "", sFullName = "";

            FileInfo fiOutput = new FileInfo("c:\\usernames\\users_.txt");
            string sFullPathNameOutput = fiOutput.FullName;
            StreamWriter srResultats = new StreamWriter(sFullPathNameOutput, true);
            try
            {
                // Connect to Active directory with principlecontext
                PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
                GroupPrincipal groupStudents = GroupPrincipal.FindByIdentity(ctx, "Students");

                foreach (UserPrincipal user in groupStudents.Members)
                {
                    // DateTime dtLastLogon = 
                    // sUserName = user.Name;
                    //sFullName = user.DisplayName;
                    // srResultats.WriteLine(sUserName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : {0} ({1})", ex.Message, sUserName);
            }
            srResultats.Close();
        }
        public static void DisableStudentsUsers()
        {
            string sUserName = "";

            FileInfo fiOutput = new FileInfo("c:\\usernames\\users_disabled.txt");
            string sFullPathNameOutput = fiOutput.FullName;
            StreamWriter srResultats = new StreamWriter(sFullPathNameOutput, true);
            try
            {
                // Connect to Active directory with principlecontext
                PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
                GroupPrincipal groupStudents = GroupPrincipal.FindByIdentity(ctx, "Students");

                foreach (UserPrincipal user in groupStudents.Members)
                {
                    sUserName = user.Name;
                    user.Enabled = false;
                    user.Save();
                    srResultats.WriteLine(sUserName);
                    // EffacerFolderEtudiantSa(sUserName);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : {0} ({1})", ex.Message, sUserName);
            }
            srResultats.Close();
        }
        static void ReCreateUsersFromDataBaseAfterRestore()
        {
            //string sSql = string.Format("SELECT P.Prenom, P.Nom,P.PersonneID, IsNull(P.UserNameAttribue, '') as UserNameAttribue, C.ClasseID, S.SessionID FROM Classes C, Sessions S, EtudiantsCourants E, Personnes P " +
            //      "WHERE S.SessionID = E.SessionID AND C.ClasseID = S.ClasseID AND P.PersonneID = E.PersonneID AND C.Categorie = '{0}' " +
            //      "AND C.NiveauClasse = 1 AND S.actif = 1 ORDER BY P.Nom, P.Prenom", "Informatique");
            //       string sSql = string.Format("SELECT * FROM Personnes WHERE PersonneID NOT IN (SELECT P.PersonneID FROM Classes C, Sessions S, EtudiantsCourants E, Personnes P " +
            //      "WHERE S.SessionID = E.SessionID AND C.ClasseID = S.ClasseID AND P.PersonneID = E.PersonneID AND C.Categorie = '{0}' " +
            //      "AND C.NiveauClasse = 1 AND S.actif = 1 ) ORDER BY Nom, Prenom", "Informatique");
            string sSql = "SELECT * FROM Personnes WHERE ISNULL(UserNameAttribue,'')='' ORDER BY Nom, Prenom";

            SqlDataReader dt = null;
            SqlConnection sqlConn = null;
            SqlCommand SqlCmd;
            string sPrenom, sNom, sUserName, sPersonneID, sUerNameAttribue;

            try
            {
                sqlConn = new System.Data.SqlClient.SqlConnection(sSqlConnString);
                sqlConn.Open();

                SqlCmd = new SqlCommand(sSql, sqlConn);
                dt = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

                sqlConn = null;

                if (dt != null)
                {
                    dt.Read();
                    do
                    {
                        sPrenom = dt["Prenom"].ToString();
                        sNom = dt["Nom"].ToString();
                        sPersonneID = dt["PersonneID"].ToString();
                        sUerNameAttribue = dt["UserNameAttribue"].ToString();
                        if (sUerNameAttribue != string.Empty)
                            sUserName = sUerNameAttribue;
                        else
                        {
                            sUserName = sPrenom.Substring(0, 1) + sNom;
                            sUserName = sUserName.Trim();
                            sUserName = sUserName.Replace(" ", "");
                            sUserName = sUserName.Replace("\t", "");
                            sUserName = sUserName.Replace(".", "");
                            sUserName = sUserName.ToLower();
                            sUserName = sUserName.Replace("'", "");
                            sUserName = sUserName.Replace("é", "e");
                        }

                        iIndex = 0;
                        sUserName = CreateUsername(sUserName.ToLower(), sPrenom, sNom, iIndex);

                        if (sUserName != String.Empty)
                        {
                            // Update record in Database

                            IssueCommand(string.Format("UPDATE Personnes SET UserNameAttribue = '{0}' WHERE PersonneID = {1}", sUserName, sPersonneID));
                        }
                        else
                        {
                            //ERROR
                        }
                    } while (dt.Read());
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
            }

        }
        //static void CreateUsersFromDataBase()
        //{
        //    string sSql = string.Format("SELECT P.Prenom, P.Nom,P.PersonneID, C.ClasseID, S.SessionID FROM Classes C, Sessions S, EtudiantsCourants E, Personnes P " +
        //        "WHERE S.SessionID = E.SessionID AND C.ClasseID = S.ClasseID AND P.PersonneID = E.PersonneID AND C.Categorie = '{0}' " +
        //        "AND C.NiveauClasse = 1 AND S.actif = 1 AND ISNULL(P.UserNameAttribue,'') = '' ORDER BY P.Nom, P.Prenom", "Informatique");

        //    SqlDataReader dt = null;
        //    SqlConnection sqlConn = null;
        //    SqlCommand SqlCmd;
        //    string sPrenom, sNom, sUserName, sPersonneID;

        //    try
        //    {
        //        sqlConn = new System.Data.SqlClient.SqlConnection(sSqlConnString);
        //        sqlConn.Open();

        //        SqlCmd = new SqlCommand(sSql, sqlConn);
        //        dt = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

        //        sqlConn = null;

        //        if (dt != null)
        //        {
        //            dt.Read();
        //            do
        //            {
        //                sPrenom = dt["Prenom"].ToString();
        //                sNom = dt["Nom"].ToString();
        //                sPersonneID = dt["PersonneID"].ToString();

        //                sUserName = sPrenom.Substring(0, 1) + sNom;
        //                sUserName = sUserName.Trim();
        //                sUserName = sUserName.Replace(" ", "");
        //                sUserName = sUserName.Replace("'", "");
        //                sUserName = sUserName.Replace("\t", "");
        //                sUserName = sUserName.Replace("-", "");
        //                sUserName = sUserName.Replace("é", "e");
        //                sUserName = sUserName.Replace("è", "e");
        //                sUserName = sUserName.Replace(",", "");
        //                sUserName = sUserName.Trim();
        //                sUserName = sUserName.ToLower();
        //                iIndex = 0;

        //                sUserName = CreateUsername(sUserName.ToLower(), sPrenom, sNom, iIndex);
        //                if (sUserName != String.Empty)
        //                {
        //                    // Update record in Database
        //                    IssueCommand(string.Format("UPDATE Personnes SET UserNameAttribue = '{0}' WHERE PersonneID = {1}", sUserName, sPersonneID));
        //                }
        //                else
        //                {
        //                    //ERROR
        //                }
        //            } while (dt.Read());
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        string error = ex.Message;
        //    }

        //}
        static void CreateUsersFromDataBase(string sSql)
        {
            SqlDataReader dt = null;
            SqlConnection sqlConn = null;
            SqlCommand SqlCmd;
            string sPrenom, sNom, sUserName, sPersonneID;

            try
            {
                sqlConn = new System.Data.SqlClient.SqlConnection(sSqlConnString);
                sqlConn.Open();

                SqlCmd = new SqlCommand(sSql, sqlConn);
                dt = SqlCmd.ExecuteReader(CommandBehavior.CloseConnection);

                sqlConn = null;

                if (dt != null)
                {
                    dt.Read();
                    do
                    {
                        sPrenom = dt["Prenom"].ToString();
                        sNom = dt["Nom"].ToString();
                        sPersonneID = dt["PersonneID"].ToString();

                        sUserName = sPrenom.Substring(0, 1) + sNom;
                        sUserName = sUserName.Trim();
                        sUserName = sUserName.Replace(" ", "");
                        sUserName = sUserName.Replace("'", "");
                        sUserName = sUserName.Replace("\t", "");
                        sUserName = sUserName.Replace("-", "");

                        sUserName = sUserName.Replace("ù", "u");
                        sUserName = sUserName.Replace("û", "u");
                        sUserName = sUserName.Replace("ü", "u");

                        sUserName = sUserName.Replace("ò", "o");
                        sUserName = sUserName.Replace("ö", "o");
                        sUserName = sUserName.Replace("ô", "o");

                        sUserName = sUserName.Replace("ï", "i");

                        sUserName = sUserName.Replace("ç", "c");

                        sUserName = sUserName.Replace("é", "e");
                        sUserName = sUserName.Replace("è", "e");
                        sUserName = sUserName.Replace("ê", "e");
                        sUserName = sUserName.Replace("ë", "e");

                        sUserName = sUserName.Replace("ä", "a");
                        sUserName = sUserName.Replace("â", "a");
                        sUserName = sUserName.Replace("à", "a");
                        sUserName = sUserName.Replace(",", "");
                        sUserName = sUserName.Trim();
                        sUserName = sUserName.ToLower();
                        iIndex = 0;

                        sUserName = CreateUsername(sUserName.ToLower(), sPrenom, sNom, iIndex);
                        if (sUserName != String.Empty)
                        {
                            SqlParameter paramUserName = new SqlParameter("@UserNameAttribue", SqlDbType.Text);
                            paramUserName.Value = sUserName;
                            SqlParameter paramPersonneID = new SqlParameter("@PersonneID", SqlDbType.NVarChar);
                            paramPersonneID.Value = sPersonneID;
                            // Update record in Database
                            //IssueCommandWithParams("UPDATE Personnes SET UserNameAttribue = '@UserNameAttribue' WHERE PersonneID = '@PersonneID'", paramPersonneID, paramUserName);
                            IssueCommand(string.Format("UPDATE Personnes SET UserNameAttribue = '{0}' WHERE PersonneID = '{1}'", sUserName, sPersonneID));
                        }
                        else
                        {
                            //ERROR
                        }
                    } while (dt.Read());
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
            }
        }

        /// <summary>
        /// Helps to issue query command
        /// </summary>
        /// <returns></returns>
        public static bool IssueCommandWithParams(string csSQL, params SqlParameter[] paramList)
        {
            SqlConnection sqlConn = null;
            SqlCommand SqlCmd = null;
            bool bRetVal = false;

            try
            {
                sqlConn = new System.Data.SqlClient.SqlConnection(sSqlConnString);
                sqlConn.Open();

                SqlCmd = new SqlCommand(csSQL, sqlConn);
                // Add parameters
                for (int iIndex = 0; iIndex < paramList.Length; iIndex++)
                {
                    SqlCmd.Parameters.Add(paramList[iIndex]);
                }

                SqlCmd.ExecuteNonQuery();
                bRetVal = true;
                SqlCmd.Parameters.Clear();
                SqlCmd = null;
                sqlConn.Close();
                sqlConn = null;
            }
            catch (Exception ex)
            {
                if (SqlCmd != null)
                    SqlCmd.Parameters.Clear();
                //if (sqlConn.State == ConnectionState.Open)
                //    sqlConn.Close();

                Debug.WriteLine(ex.Message);
            }

            return bRetVal;
        }


        public static bool IssueCommand(string csSQL)
        {
            SqlConnection sqlConn = null;
            SqlCommand SqlCmd;
            bool bRetVal = false;

            try
            {
                sqlConn = new System.Data.SqlClient.SqlConnection(sSqlConnString);
                sqlConn.Open();

                SqlCmd = new SqlCommand(csSQL, sqlConn);
                SqlCmd.ExecuteNonQuery();
                SqlCmd = null;
                sqlConn.Close();
                sqlConn = null;
                bRetVal = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }

            return bRetVal;
        }
        static void EffacerDossiers()
        {
            // EffacerDocumentsMounNanGroupSa("Niveau 1");
            // EffacerDocumentsMounNanGroupSa("Niveau 2");
            // EffacerDocumentsMounNanGroupSa("Niveau 3");
            // EffacerDocumentsMounNanGroupSa("Students");
            Console.WriteLine("Pressez une clé pour terminer!!!");
            Console.ReadKey();
            return;

            //Random rnd = new Random();
            //try
            //{
            //    //string username = SystemInformation.UserName;
            //    DirectoryInfo diFolder = new DirectoryInfo("c:\\usernames");
            //    if (diFolder.Exists)
            //    {
            //        FileInfo fiOutput = new FileInfo("c:\\usernames\\Resultats.txt");
            //        string sFullPathNameOutput = fiOutput.FullName;
            //        StreamWriter srResultats = new StreamWriter(sFullPathNameOutput, true);

            //        FileInfo fiAllFiles = new FileInfo("c:\\usernames\\Juin2013.txt");
            //        string sFullPathName = fiAllFiles.FullName;
            //        try
            //        {
            //            using (StreamReader sr = new StreamReader(sFullPathName))
            //            {
            //                while (!sr.EndOfStream)
            //                {
            //                    String line = sr.ReadLine();
            //                    string[] user = line.Split(' ');
            //                    iIndex = 0;
            //                    String sVraiUserName = user[0].Substring(0, 1).ToLower() + user[1].ToLower();
            //                    sVraiUserName = sVraiUserName.Trim();
            //                    sVraiUserName = sVraiUserName.Replace(" ", "");
            //                    sVraiUserName = sVraiUserName.Replace("\t", "");
            //                    sVraiUserName = CreateUsername(sVraiUserName, user[0], user[1], iIndex);
            //                    if (sVraiUserName != String.Empty)
            //                    {
            //                        srResultats.WriteLine(string.Format("{0},{1},{2}", user[1], user[0], sVraiUserName));
            //                    }
            //                    else
            //                    {
            //                        srResultats.WriteLine(string.Format("{0},{1},{2}", user[1], user[0], "ERROR: !!!!!!!"));
            //                    }
            //                }
            //            }
            //        }
            //        catch (Exception e)
            //        {
            //            srResultats.Close();
            //            Console.WriteLine(e.Message);
            //        }

            //        srResultats.Close();
            //        Console.ReadKey();
            //    }
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine("The file could not be read:");
            //    Console.WriteLine(e.Message);
            //}           
        }
        static void CreateUsersFromFile(string sFileWithUserNames)
        {
            string sFileToProcess = sFileWithUserNames;
            if (sFileWithUserNames.Trim() == "")
            {
                Debug.WriteLine("Fichier ....");
                return;
            }

            try
            {
                FileInfo fiOutput = new FileInfo("D:\\IT\\ResultatsUsers.txt");
                string sFullPathNameOutput = fiOutput.FullName;
                StreamWriter srResultats = new StreamWriter(sFullPathNameOutput, true);

                FileInfo fiAllFiles = new FileInfo(sFileToProcess);
                string sFullPathName = fiAllFiles.FullName;
                try
                {
                    using (StreamReader sr = new StreamReader(sFullPathName))
                    {
                        while (!sr.EndOfStream)
                        {
                            String line = sr.ReadLine();
                            string[] user = line.Split(',');
                            iIndex = 0;
                            String sUserName = user[0].ToLower();
                            sUserName = sUserName.Trim();
                            //sUserName = sUserName.Replace(" ", "");
                            //sUserName = sUserName.Replace("\t", "");
                            sUserName = CreateUsername(sUserName, user[1], iIndex);
                            if (sUserName != String.Empty)
                            {
                                srResultats.WriteLine(string.Format("{0},{1},{2}", user[1], user[0], sUserName));
                            }
                            else
                            {
                                srResultats.WriteLine(string.Format("{0},{1},{2}", user[1], user[0], "ERROR: !!!!!!!"));
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    srResultats.Close();
                    Console.WriteLine(e.Message);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.ReadKey();

        }
        //static public void EffacerUserFolderIfUserDoesNotExist(String sFolder)
        //{
        //    // Connect to Active directory with principlecontext
        //    PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
        //    try
        //    {
        //        DirectoryInfo foldersToDelete = new DirectoryInfo(sFolder);
        //        if (foldersToDelete.Exists)
        //        {
        //            DirectoryInfo[] subFolders = foldersToDelete.GetDirectories();
        //            for (int i = 0; i < subFolders.Length; i++)
        //            {
        //                //Create an instance of UserPriciple
        //                UserPrincipal u = new UserPrincipal(ctx);
        //                // Create an in-memory user object to use as the query example.
        //                u = UserPrincipal.FindByIdentity(ctx, subFolders[i].Name);
        //                if (u == null)  // username does not exist, create it.
        //                {
        //                    EffacerNan1Folder(subFolders[i].FullName);
        //                }                                                
        //            }
        //            FileInfo[] filesToDelete = foldersToDelete.GetFiles();
        //            for (int i = 0; i < subFolders.Length; i++)
        //            {

        //                filesToDelete[i].Delete();
        //                Console.WriteLine(filesToDelete[i].FullName + " --  Deleted ..");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }
        //}
        static public void EffacerUsersInactifs()
        {
            int iCounter = 0;
            try
            {
                string sUserName;
                string sFullName;

                // Connect to Active directory with principlecontext
                PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, "Students");

                foreach (UserPrincipal user in group.Members)
                {
                    sUserName = user.Name;
                    sFullName = user.DisplayName;
                    Console.WriteLine("{0} - {1} ({2})", ++iCounter, sUserName, sFullName);

                    try
                    {
                        if (user.LastLogon == null)
                        {
                            // user pa janm log on
                            user.Delete();
                            Console.WriteLine("User Deleted ... NEVER Logged On.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error : {0}", ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : {0}", ex.Message);
            }
        }

        //static public void EffacerUsersInactifs()
        //{
        //    FileInfo fiOutput = new FileInfo("c:\\usernames\\UsersDeleted.txt");
        //    string sFullPathNameOutput = fiOutput.FullName;
        //    StreamWriter srResultats = new StreamWriter(sFullPathNameOutput, true);
        //    int iCounter = 0;

        //    Hashtable htExceptions = new Hashtable();
        //    htExceptions.Add("administrator", "administrator"); htExceptions.Add("spoteau", "spoteau");
        //    htExceptions.Add("jextra", "jextra"); htExceptions.Add("jborgella", "jborgella");
        //    htExceptions.Add("mnelson", "mnelson"); htExceptions.Add("tsauveur", "tsauveur");
        //    htExceptions.Add("jfenel", "jfenel"); htExceptions.Add("haneas", "haneas");
        //    htExceptions.Add("rmomlouis", "rmomlouis"); htExceptions.Add("eitoral", "eitoral");
        //    htExceptions.Add("rnicolas", "rnicolas"); htExceptions.Add("fsidney", "fsidney");
        //    htExceptions.Add("panthoby", "panthoby"); htExceptions.Add("falexis", "falexis");
        //    htExceptions.Add("vpierre", "vpierre"); htExceptions.Add("nlejeune", "nlejeune");
        //    htExceptions.Add("bdorvilus", "bdorvilus");

        //    try
        //    {
        //        string sUserName;
        //        string sFullName;

        //        // Connect to Active directory with principlecontext
        //        PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
        //        GroupPrincipal groupStudents = GroupPrincipal.FindByIdentity(ctx, "Students");

        //        foreach (UserPrincipal user in groupStudents.Members)
        //        {
        //            sUserName = user.Name;
        //            sFullName = user.DisplayName;
        //            Console.WriteLine("{0} - {1} ({2})", ++iCounter, sUserName, sFullName);

        //            if (htExceptions.Contains(sUserName))
        //                continue;
        //            try
        //            {
        //                if (user.LastLogon != null)
        //                {
        //                    int iYear = user.LastLogon.Value.Year;
        //                    int iMonth = user.LastLogon.Value.Month;
        //                    if (iYear <= 2017)
        //                    {
        //                        srResultats.WriteLine(sUserName + " - " + sFullName);
        //                        user.Delete();
        //                        Console.WriteLine("User Deleted ....");
        //                        //EffacerFolderEtudiantSa(sUserName);
        //                        Console.WriteLine("Folders Deleted ....");
        //                    }
        //                }
        //                else
        //                {
        //                    // user pa janm log on
        //                    srResultats.WriteLine(sUserName + " - " + sFullName);
        //                    user.Delete();
        //                    Console.WriteLine("User Deleted ... NEVER Logged On.");
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Console.WriteLine("Error : {0}", ex.Message);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error : {0}", ex.Message);
        //        srResultats.Close();
        //    }

        //    srResultats.Close();
        //}

        //static public void EffacerTEMP()
        //{
        //    FileInfo fiAllFiles = new FileInfo(@"C:\UserNames\UsersDeleted.txt");
        //    string sFullPathName = fiAllFiles.FullName;
        //    string sUserName;
        //    try
        //    {
        //        using (StreamReader sr = new StreamReader(sFullPathName))
        //        {
        //            while (!sr.EndOfStream)
        //            {
        //                String line = sr.ReadLine();
        //                sUserName = line.Substring(0, line.LastIndexOf("-"));
        //                EffacerFolderEtudiantSa(sUserName.Trim());
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }
        //}

        //static public void EffacerInternetTempFolders()
        //{
        //    try
        //    {
        //        string sUserName;

        //        // Connect to Active directory with principalcontext
        //        PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
        //        GroupPrincipal groupStudents = GroupPrincipal.FindByIdentity(ctx, "Students");

        //        foreach (UserPrincipal user in groupStudents.Members)
        //        {
        //            sUserName = user.Name;
        //            Console.WriteLine("Username : {0}\n)", sUserName);
        //            EffacerNan1Folder("c:\\users\\" + sUserName + "\\AppData\\Local");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error : {0})", ex.Message);
        //    }
        //}

        //static public void EffacerInternetTempFoldersUserSa(string sUserName)
        //{
        //    try
        //    {
        //        //string sUserName;

        //        // Connect to Active directory with principalcontext
        //        //PrincipalContext ctx = new PrincipalContext(ContextType.Machine);
        //        //GroupPrincipal groupStudents = GroupPrincipal.FindByIdentity(ctx, "students");

        //        //foreach (UserPrincipal user in groupStudents.Members)
        //        //{
        //            //sUserName = user.Name;
        //            Console.WriteLine("Username : {0}\n)", sUserName);
        //            EffacerNan1Folder("c:\\users\\" + sUserName + "\\AppData\\Local");
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Error : {0})", ex.Message);
        //    }
        //}

        //static public void EffacerNan1Folder(string sFolder)
        //{
        //    try
        //    {
        //        DirectoryInfo diFolder = new DirectoryInfo(sFolder);
        //        if (diFolder.Exists)
        //        {
        //            // Effacer les files
        //            FileInfo[] fiFiles = diFolder.GetFiles();
        //            for (int i = 0; i < fiFiles.Length; i++)
        //            {
        //                fiFiles[i].Delete();
        //                Console.WriteLine(fiFiles[i].FullName + " -- Deleted ..");
        //            }

        //            // Effacer les subfolders
        //            DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //            if (subFolders.Length > 0)
        //            {
        //                for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //                {
        //                    EffacerNan1Folder(sFolder + "\\" + subFolders[ifd].Name);
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }
        //}

        //static public void Effacer1File(string sFile)
        //{
        //    try
        //    {
        //        FileInfo fiFile = new FileInfo(sFile);
        //        if (fiFile.Exists)
        //        {
        //            fiFile.Delete();
        //            Console.WriteLine(fiFile.FullName + " --  Deleted ..");
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }
        //}

        //static public void EffacerFolderEtudiantSa(string sEtudiant)
        //{
        //    try
        //    {
        //        DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + sEtudiant);
        //        if (diFolder.Exists)
        //        {
        //            DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //            for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //            {
        //                // Deletes all files
        //                DirectoryInfo folderToDelete = new DirectoryInfo("c:\\users\\" + sEtudiant + "\\" + subFolders[ifd].Name);
        //                if (folderToDelete.Exists)
        //                {
        //                    FileInfo[] filesToDelete = folderToDelete.GetFiles();
        //                    for (int i = 0; i < filesToDelete.Length; i++)
        //                    {
        //                        try
        //                        {
        //                            filesToDelete[i].Delete();
        //                            Console.WriteLine(filesToDelete[i].FullName + " --  Deleted ..");
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Console.WriteLine("\t" + ex.Message);
        //                        }
        //                    }
        //                }
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }
        //}

        //static public void EffacerExeNanFolderDownloads()
        //{
        //    DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\");
        //    if (diFolder.Exists)
        //    {
        //        DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //        for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //        {
        //            // Deletes all files
        //            DirectoryInfo Downloads = new DirectoryInfo("c:\\users\\" + subFolders[ifd].Name + "\\Downloads");
        //            if (Downloads.Exists)
        //            {
        //                FileInfo[] filesInDownloads = Downloads.GetFiles();
        //                for (int i = 0; i < filesInDownloads.Length; i++)
        //                {
        //                    if (filesInDownloads[i].Extension.ToLower() == ".exe" || filesInDownloads[i].Extension.ToLower() == ".bat")
        //                    {
        //                        if (filesInDownloads[i].Name.ToLower() != "cmd.exe")
        //                        {
        //                            filesInDownloads[i].Delete();
        //                            Console.WriteLine(filesInDownloads[i].FullName + " --  Deleted ..");
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //    }
        //}

        //static public void EffacerDocumentsMounNanGroupSa(string sGroup)
        //{
        //    Hashtable hUoP = new Hashtable();
        //    FileInfo fiAllFiles = new FileInfo(@"C:\Admin Helpers\Dossiers Uopeople 1a\UoPUsernames.txt");
        //    string sFullPathName = fiAllFiles.FullName;
        //    try
        //    {
        //        using (StreamReader sr = new StreamReader(sFullPathName))
        //        {
        //            while (!sr.EndOfStream)
        //            {
        //                String line = sr.ReadLine();
        //                hUoP.Add(line, line);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Debug.WriteLine(ex.Message.ToString());
        //    }

        //    PrincipalContext context = new PrincipalContext(ContextType.Machine);

        //    GroupPrincipal insGroupPrincipal = new GroupPrincipal(context);
        //    insGroupPrincipal.Name = "*";

        //    int i = 0;

        //    PrincipalSearcher insPrincipalSearcher = new PrincipalSearcher();
        //    insPrincipalSearcher.QueryFilter = insGroupPrincipal;
        //    PrincipalSearchResult<Principal> results = insPrincipalSearcher.FindAll();
        //    foreach (Principal p in results)
        //    {
        //        Debug.WriteLine(p.Name);
        //        if (p.Name.ToLower() == sGroup.ToLower())
        //        {
        //            GroupPrincipal iGroupPrincipal = (GroupPrincipal)p;
        //            List<Principal> insListPrincipal = new List<Principal>();
        //            foreach (Principal yonEdutiant in iGroupPrincipal.Members)
        //            {
        //                if (!hUoP.ContainsValue(yonEdutiant))
        //                {
        //                    Console.WriteLine(Convert.ToString(++i) + " " + yonEdutiant);
        //                    //DeleteDocFiles(yonEdutiant.Name, "Documents");
        //                    //DeleteDocFiles(yonEdutiant.Name, "Desktop");
        //                    //DeleteDocFiles(yonEdutiant.Name, "");

        //                    DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + yonEdutiant.Name + "\\Documents");
        //                    if (diFolder.Exists)
        //                    {

        //                        DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //                        for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //                        {
        //                            DeleteDocFiles(yonEdutiant.Name, subFolders[ifd].Name);
        //                        }
        //                    }
        //                    diFolder = new DirectoryInfo("c:\\users\\" + yonEdutiant.Name + "\\Desktop");
        //                    if (diFolder.Exists)
        //                    {
        //                        DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //                        for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //                        {
        //                            DeleteDocFiles(yonEdutiant.Name, subFolders[ifd].Name);
        //                        }
        //                    }
        //                    diFolder = new DirectoryInfo("c:\\users\\" + yonEdutiant.Name);
        //                    if (diFolder.Exists)
        //                    {
        //                        DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //                        for (int ifd = 0; ifd < subFolders.Length; ifd++)
        //                        {
        //                            DeleteDocFilesExt(yonEdutiant.Name, subFolders[ifd].Name);
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

        //public static void DeleteDocFiles(string sUser, string sFolder)
        //{
        //    try
        //    {
        //        //string username = SystemInformation.UserName;
        //        DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + sUser + "\\" + sFolder);

        //        if (diFolder.Exists)
        //        {
        //            // Deletes all files
        //            FileInfo[] fiAllFiles = diFolder.GetFiles();
        //            for (int i = 0; i < fiAllFiles.Length; i++)
        //            {
        //                if (fiAllFiles[i].Extension.ToLower() == ".doc" || fiAllFiles[i].Extension.ToLower() == ".docx" || fiAllFiles[i].Extension.ToLower() == ".xls" || fiAllFiles[i].Extension.ToLower() == ".xlsx")
        //                {
        //                    fiAllFiles[i].Delete();
        //                    Console.WriteLine(fiAllFiles[i].FullName + " --  Deleted ..");
        //                }
        //            }

        //            // Deletes all sub-folders
        //            DirectoryInfo[] fiAllSubFolders = diFolder.GetDirectories();
        //            for (int i = 0; i < fiAllSubFolders.Length; i++)
        //            {
        //                DeleteDocFiles(sUser, sFolder + "\\" + fiAllSubFolders[i].Name);
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("The file could not be read : ");
        //        Console.WriteLine("\t\t" + e.Message);
        //    }
        //}

        //public static void DeleteDocFilesExt(string sUser, string sFolder)
        //{
        //    try
        //    {
        //        //string username = SystemInformation.UserName;
        //        DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + sUser + "\\" + sFolder);
        //        //DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + sUser + "\\Desktop");
        //        //DirectoryInfo diFolder = new DirectoryInfo("c:\\users\\" + sUser);
        //        if (diFolder.Name == "AppData" ||
        //            diFolder.Name == "Contacts" ||
        //            diFolder.Name == "Favorites" ||
        //            diFolder.Name == "Links")
        //        if (diFolder.Exists)
        //        {
        //            // Deletes all files
        //            FileInfo[] fiAllFiles = diFolder.GetFiles();
        //            for (int i = 0; i < fiAllFiles.Length; i++)
        //            {
        //                if (fiAllFiles[i].Extension.ToLower() == ".doc" || fiAllFiles[i].Extension.ToLower() == ".docx" || fiAllFiles[i].Extension.ToLower() == ".xls" || fiAllFiles[i].Extension.ToLower() == ".xlsx")
        //                {
        //                    fiAllFiles[i].Delete();
        //                    Console.WriteLine(fiAllFiles[i].FullName + " --  Deleted ..");
        //                }
        //            }
        //            // Deletes all files in subfolders
        //            //DirectoryInfo[] subFolders = diFolder.GetDirectories();
        //            //for (int i = 0; i < subFolders.Length; i++)
        //            //{
        //            //    DeleteDocFiles(sUser, subFolders[i].Name);
        //            //}
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine("The file could not be read : ");
        //        Console.WriteLine("\t\t" + e.Message);
        //    }
        //}

        public static String CreateUsername(String sBaseUsername, String sFirstname, String sLastname, int iAddtoUser)
        {
            bool bSuccess = false;
            string sUsername = sBaseUsername;
            if (iAddtoUser > 0)
                sUsername = sBaseUsername + iAddtoUser.ToString();

            // Connect to Active directory with principlecontext
            PrincipalContext ctx = new PrincipalContext(ContextType.Machine);

            //Create an instance of UserPriciple
            UserPrincipal u = new UserPrincipal(ctx);

            // Create an in-memory user object to use as the query example.
            u = UserPrincipal.FindByIdentity(ctx, sUsername);
            //try again
            if (u == null)  // username does not exist, create it.
            {

                if (CreateLocalWindowsAccount(sUsername, PASSWORD_TEXT, sFirstname + " " + sLastname, "", true, false))
                {
                    bSuccess = true;
                    Console.WriteLine(sFirstname + " - " + sLastname + " - " + sUsername);
                }
            }
            else
            {
                // username exists, add index.
                iIndex++;
                return CreateUsername(sBaseUsername, sFirstname, sLastname, iIndex);
            }

            if (bSuccess)
                return sUsername.ToLower();
            else
                return "";
        }
        public static String CreateUsername(String sBaseUsername, String sFullname, int iAddtoUser)
        {
            bool bSuccess = false;
            string sUsername = sBaseUsername;
            if (iAddtoUser > 0)
                sUsername = sBaseUsername + iAddtoUser.ToString();

            // Connect to Active directory with principlecontext
            PrincipalContext ctx = new PrincipalContext(ContextType.Machine);

            //Create an instance of UserPriciple
            UserPrincipal u = new UserPrincipal(ctx);

            // Create an in-memory user object to use as the query example.
            u = UserPrincipal.FindByIdentity(ctx, sUsername);
            //try again
            if (u == null)  // username does not exist, create it.
            {

                if (CreateLocalWindowsAccount(sUsername, PASSWORD_TEXT, sFullname, "", true, false))
                {
                    bSuccess = true;
                    Console.WriteLine(sFullname + " - " + sUsername);
                }
            }
            //else
            //{
            //    // username exists, add index.
            //    iIndex++;
            //    return CreateUsername(sBaseUsername, sFullname, iIndex);
            //}

            if (bSuccess)
                return sUsername.ToLower();
            else
                return "";
        }
        public static bool CreateLocalWindowsAccount(string username, string password, string displayName, string description, bool canChangePwd, bool pwdExpires)
        {
            try
            {
                // Connect to Active directory with principlecontext
                PrincipalContext context = new PrincipalContext(ContextType.Machine);

                UserPrincipal user = new UserPrincipal(context, username, password, true);
                user.Save();
                //user.SetPassword(password);
                user.DisplayName = displayName.Trim();
                //user.Name = username;
                user.Name = displayName.Trim();
                user.Description = description.Trim();
                user.UserCannotChangePassword = false;
                user.PasswordNeverExpires = pwdExpires;
                user.Save();
                user.ExpirePasswordNow(); // force user to change password next time he/she logs on.

                //now add user to "Users" group so it displays in Control Panel
                GroupPrincipal group = GroupPrincipal.FindByIdentity(context, "Users");
                //GroupPrincipal group1 = GroupPrincipal.FindByIdentity(context, "Niveau 1");
                GroupPrincipal group3 = GroupPrincipal.FindByIdentity(context, "Students");
                GroupPrincipal group2 = GroupPrincipal.FindByIdentity(context, "Remote Desktop Users");

                group.Members.Add(user);
                group.Save();
                //group1.Members.Add(user);
                //group1.Save();
                group2.Members.Add(user);
                group2.Save();
                group3.Members.Add(user);
                group3.Save();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error creating account: {0}", ex.Message);
                return false;
            }
        }
        public static bool ListWindowsAccounts()
        {
            try
            {
                // Connect to Active directory with principlecontext
                PrincipalContext context = new PrincipalContext(ContextType.Machine);

                //string strPath = "WinNT://users";
                //DirectoryEntry entry = null;
                //entry = new DirectoryEntry(strPath);
                //DirectorySearcher mySearcher = new DirectorySearcher(entry);
                //mySearcher.Filter = "(ObjectCategory=user)";

                string sUserName, sFullName;
                int iCounter = 0;
                //foreach (SearchResult result in mySearcher.FindAll())
                //GroupPrincipal group = GroupPrincipal.FindByIdentity(context, "Users");
                GroupPrincipal group = GroupPrincipal.FindByIdentity(context, "Niveau 1");
                foreach (Principal user in group.Members)
                {
                    iCounter++;
                    //strName = result.GetDirectoryEntry().Name;
                    sUserName = user.Name;
                    //strName = user.GivenName;
                    sFullName = user.DisplayName;
                    string[] names = sFullName.Split(' ');
                    string sLastName = names[names.Length - 1];
                    string sFirstName = names[0];
                    if (names.Length > 2)
                        sFirstName = sFirstName + " " + names[1];
                    if (names.Length > 3)
                        sFirstName = sFirstName + " " + names[2];
                    IssueCommand(string.Format("UPDATE Personnes SET UserNameAttribue = '{0}' WHERE Nom = '{1}' AND Prenom = '{2}' AND ISNULL(UserNameAttribue,'')=''", sUserName, sLastName, sFirstName));
                    IssueCommand(string.Format("UPDATE Personnes SET UserNameAttribue = '{0}' WHERE Nom = '{1}' AND Prenom = '{2}' AND ISNULL(UserNameAttribue,'')=''", sUserName, sFirstName, sLastName));
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error creating account: {0}", ex.Message);
                return false;
            }
        }
    }
}