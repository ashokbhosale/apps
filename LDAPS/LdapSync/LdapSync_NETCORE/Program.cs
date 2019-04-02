using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;

namespace LdapSync_NETCORE
{
    public class Users
    {

        public string UserName { get; set; }
        public string DisplayName { get; set; }

    }
    public class Program
    {
        public static List<Users> GetADUsers(bool SSLEnabled, string ldapPath, string username, string pwd)
        {
            try
            {
                List<Users> lstADUsers = new List<Users>();

                DirectoryEntry searchRoot = null;
                if (SSLEnabled)
                {
                    searchRoot = new DirectoryEntry(ldapPath + ":636", username, pwd);
                    searchRoot.AuthenticationType = AuthenticationTypes.SecureSocketsLayer;
                }
                else
                {
                    searchRoot = new DirectoryEntry(ldapPath + ":389", username, pwd);
                }


                DirectorySearcher search = new DirectorySearcher(searchRoot);
                search.Filter = "(&(objectClass=user)(objectCategory=person))";
                search.PropertiesToLoad.Add("samaccountname");
                search.PropertiesToLoad.Add("displayname");//first name
                SearchResult result;
                SearchResultCollection resultCol = search.FindAll();

                if (resultCol != null)
                {
                    Console.WriteLine("=================================");
                    Console.WriteLine("Displaying first {0} users ", (resultCol.Count >= 10 ? 10 : resultCol.Count));
                    Console.WriteLine("=================================");

                    for (int counter = 0; counter < (resultCol.Count >= 10 ? 10 : resultCol.Count); counter++)
                    {
                        string UserNameEmailString = string.Empty;
                        result = resultCol[counter];
                        if (result.Properties.Contains("samaccountname"))
                        {
                            try
                            {
                                Users objSurveyUsers = new Users();
                                objSurveyUsers.UserName = (String)result.Properties["samaccountname"][0];
                                objSurveyUsers.DisplayName = (String)result.Properties["displayname"][0];
                                lstADUsers.Add(objSurveyUsers);
                            }
                            catch { }
                        }
                    }

                }
                return lstADUsers;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error:" + ex.Message);
                Console.ReadLine();
                return null;
            }
        }


        public static  void Main(string[] args)
        {
            Console.WriteLine("Reading data from Command Line..");
            if (args.Count() != 3)
            {
                throw new Exception(" Please follow command line input format   LdapSync.exe <LDAPPATH:eg.LDAP://BLRMGMTAD.abc.com> <username> <password>");
            }
            string ldapPath = args[0];
            string user = args[1];
            string pwd = args[2];
            Console.WriteLine("Retriving all LDAP Users over 389 Port");
            List<Users> list = GetADUsers(false, ldapPath, user, pwd);
            list.ForEach(aduser =>
            {

                Console.WriteLine("User :" + aduser.UserName.PadRight(25) + " Display Name:" + aduser.DisplayName.PadRight(20));
            });
            Console.WriteLine("=================================");
            Console.WriteLine("=================================");

            Console.WriteLine("           ");
            Console.WriteLine("           ");

            list = null;
            Console.WriteLine("Retriving all LDAP Users over Secure Socket Layer 636 Port");
            list = GetADUsers(true, ldapPath, user, pwd);
            list.ForEach(user1 =>
            {

                Console.WriteLine("User :" + user1.UserName.PadRight(25) + " Display Name:" + user1.DisplayName.PadRight(20));
            });
            Console.WriteLine("=================================");
            Console.WriteLine("=================================");
            Console.WriteLine("Please enter to exit !!!");
            Console.ReadKey();
        }
    }
}
