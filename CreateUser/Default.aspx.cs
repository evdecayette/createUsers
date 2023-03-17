using CreateUser.App_Codes;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Security.Principal;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Diagnostics;
using Microsoft.AspNetCore.Authentication;
using System.DirectoryServices.AccountManagement;
using System.Drawing;

namespace CreateUser
{
    public partial class Default : System.Web.UI.Page
    {
        static string sSqlConnString = @"Data Source=DESKTOP-GBQB9HI\MSSQLSERVER1;Initial Catalog=UEspoirDB;Integrated Security=True";
        public static String sUserName = "edecayette", sDomain = "DESKTOP-GBQB9HI\\MSSQLSERVER1", sPass = "2892";
        static String PASSWORD_TEXT = "jesus123";
        static int iIndex = 0;


        protected void Page_Load(object sender, EventArgs e)
        {
            selectPreviewUsers();
            //displayCreatedUsers();
        }

        protected void selectPreviewUsers()
        {
          String sql =  string.Format("SELECT PersonneID, (Prenom + (' ')+ Nom) as NomComplet,  UserNameAttribue FROM [UEspoirDB].[dbo].[Personnes] " +
                " WHERE actif = 1 AND ISNULL(UserNameAttribue,'') = '' ORDER BY Nom, Prenom");

            SqlDataAdapter da = new SqlDataAdapter(sql, sSqlConnString);
            DataSet ds = new DataSet();
            da.Fill(ds, "NomComplet");
            ListView1.DataSource = ds.Tables["NomComplet"];
            ListView1.DataBind();
        }

        public String displayCreatedUsers(string sFistName, string sLastName, string sUserName)
        {
            return "<p>" + sFistName + " - " + sLastName + " - " + sUserName + "</p>";
        }

        protected void btnCreatClick(object sender, EventArgs e)
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
            //Console.WriteLine("Pressez une clé pour terminer le programme ...");
            //Console.ReadKey();
            lblMessage.Text = "Pressez une clé pour terminer le programme ...";
            
            return;

        }

        private static bool impersonateValidUser(String userName, String domain, String password)
        {
            WindowsIdentity tempWindowsIdentity;
            IntPtr token = IntPtr.Zero;
            IntPtr tokenDuplicate = IntPtr.Zero;
            WindowsImpersonationContext impersonationContext;

            if (supportClass.RevertToSelf())
            {
                if (supportClass.LogonUserA(userName, domain, password, supportClass.LOGON32_LOGON_INTERACTIVE,
                    supportClass.LOGON32_PROVIDER_DEFAULT, ref token) != 0)
                {
                    if (supportClass.DuplicateToken(token, 2, ref tokenDuplicate) != 0)
                    {
                        tempWindowsIdentity = new WindowsIdentity(tokenDuplicate);
                        impersonationContext = tempWindowsIdentity.Impersonate();

                        if (impersonationContext != null)
                        {
                            supportClass.CloseHandle(token);

                            supportClass.CloseHandle(tokenDuplicate);
                            return true;
                        }
                    }
                }
            }

            if (token != IntPtr.Zero)
                supportClass.CloseHandle(token);
            if (tokenDuplicate != IntPtr.Zero)
                supportClass.CloseHandle(tokenDuplicate);
            return false;
        }

        public void CreateUsersFromDataBase(string sSql)
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

        public String CreateUsername(String sBaseUsername, String sFirstname, String sLastname, int iAddtoUser)
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
                    displayUsers.Text += displayCreatedUsers(sFirstname, sLastname, sUsername);
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