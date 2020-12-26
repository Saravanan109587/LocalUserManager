using System;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Management;

namespace CreateUser
{
    class Program
    {
        public static string Name;
        public static string Pass;


        static void Main(string[] args)
        {


            Console.WriteLine("Windows Account Creator");
            Console.WriteLine("Enter User Name");
            Name = Console.ReadLine();

            Console.WriteLine("Enter User Password");
            Pass = Console.ReadLine();
            Name = "NewUser";
            Pass = "This1sVeryL0ngPa55w0rd!Changed";
            createUser(Name, Pass);



            AssignServiceUser("WpsReportProvider", Name, Pass);
            //createUser(Name, Pass);

        }

        public static bool DoesUserExist(string userName)
        {
            using (var domainContext = new PrincipalContext(ContextType.Machine, Environment.MachineName))
            {
                using (var foundUser = UserPrincipal.FindByIdentity(domainContext, IdentityType.SamAccountName, userName))
                {
                    return foundUser != null;
                }
            }
        }
        public static void createUser(string Name, string Pass)
        {


            try
            {
                DirectoryEntry AD = new DirectoryEntry("WinNT://" +
                                                       Environment.MachineName + ",computer");
                var test = DoesUserExist(Name);
                if (test)
                {
                    Console.WriteLine("Account Already Exists");
                    DirectoryEntry Exist = AD.Children.Find(Name);
                    Exist.Invoke("SetPassword", Pass);
                }

                else
                {
                    DirectoryEntry NewUser = AD.Children.Add(Name, "user");

                    NewUser.Invoke("SetPassword", Pass);
                    NewUser.Invoke("Put", "Description", "Test User from .NET");
                    NewUser.CommitChanges();
                    DirectoryEntry grp;

                    grp = AD.Children.Find("WPSServices", "group");
                    if (grp != null) { grp.Invoke("Add", NewUser.Path); }
                }

                Console.WriteLine("Account Created Successfully");

                Console.WriteLine("Press Enter to continue....");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();

            }
        }

        private static bool AssignServiceUser(string sName, string userName, string password)
        {
            try
            {
                Console.WriteLine("Environment.MachineName" + Environment.MachineName);

                string queryStr = @"select * from Win32_Service where Name='" + sName + @"'";

                ObjectQuery oQuery = new ObjectQuery(queryStr);

                //Execute the query  
                ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oQuery);

                //Get the results
                ManagementObjectCollection oReturnCollection = oSearcher.Get();

                foreach (ManagementObject oReturn in oReturnCollection)
                {
                    string serviceName = oReturn.GetPropertyValue("Name") as string;

                    string fullServiceName = "Win32_Service.Name='";
                    fullServiceName += serviceName;
                    fullServiceName += "'";
                    object[] accountParams = new object[11];
                    accountParams[6] = Environment.MachineName + @"\" + userName;
                    accountParams[7] = password;
                    ManagementObject mo = new ManagementObject(fullServiceName);
                    var state = mo.GetPropertyValue("State").ToString().ToLower();
                    mo.InvokeMethod("StopService", new object[] { null });
                    uint returnCode = (uint)mo.InvokeMethod("Change", accountParams);
                    mo.InvokeMethod("StartService", new object[] { null });

                    mo.Dispose();
                }
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            return true;
        }


    }
}
