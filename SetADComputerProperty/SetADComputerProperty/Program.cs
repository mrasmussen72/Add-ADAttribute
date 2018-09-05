using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace Set_ADComputer_Department
{
    class Program
    {
        static int Main(string[] args)
        {
            //if(args.Length < 3)
            //{
            //    Console.WriteLine("Invalid number of arguments sent, argument count = " + args.Length.ToString() + "\r\n");
            //    return;
            //}
            int ReturnCode = 0;
            string ComputerNameStr = "ComputerName";
            string AttributeNameStr = "AttributeName";
            string AttributeValueStr = "AttributeValue";
            string LDAPPathStr = "LDAPPath";
            string LogFilePathStr = "LogFilePath";


            int i = 0;
            Hashtable table = new Hashtable();
            try
            {
                foreach (string s in args)
                {
                    switch (s.ToLower())
                    {
                        case "-computername":
                            table.Add(ComputerNameStr, args[i + 1]);
                            break;
                        case "-attributename":
                            table.Add(AttributeNameStr, args[i + 1]);
                            break;
                        case "-attributevalue":
                            table.Add(AttributeValueStr, args[i + 1]);
                            break;
                        case "-ldappath":
                            table.Add(LDAPPathStr, args[i + 1]);
                            break;
                        case "-logfilepath":
                            table.Add(LogFilePathStr, args[i + 1]);
                            break;
                    }
                    i++;
                }

                WriteLog("Starting function\r\n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                WriteLog("Computer Name " + table[ComputerNameStr].ToString() + "\r\n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                WriteLog("AD Attribute " + table[AttributeNameStr].ToString() + "\r\n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                WriteLog("LDAP Path " + table[LDAPPathStr].ToString() + "\r\n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                WriteLog("Log file path " + table[LogFilePathStr].ToString() + "\r\n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());

                DirectoryEntry myLdapConnection = createDirectoryEntry(table[LDAPPathStr].ToString());
                DirectorySearcher search = new DirectorySearcher(myLdapConnection);
                string temp = table[ComputerNameStr].ToString();
                search.Filter = "(cn=" + temp + ")";
                search.PropertiesToLoad.Add(AttributeNameStr);

                SearchResult result = search.FindOne();

                if (result != null)
                {

                    DirectoryEntry entryToUpdate = result.GetDirectoryEntry();

                    entryToUpdate.Properties[table[AttributeNameStr].ToString()].Value = table[AttributeValueStr].ToString();
                    entryToUpdate.CommitChanges();

                }

                else Console.WriteLine("Computer not found!");
            }
            catch (Exception ex)
            {
                try
                {
                    ReturnCode = 1;
                    WriteLog("ERROR MESSAGE = " + ex.Message + "/r/n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                    if (ex.InnerException != null)
                    {
                        WriteLog("Inner exception error message = " + ex.InnerException + "/r/n", table[LogFilePathStr].ToString(), table[ComputerNameStr].ToString());
                    }
                }
                catch (Exception exs)
                {
                    ReturnCode = 2;
                }
            }
            return ReturnCode;
        }

        static DirectoryEntry createDirectoryEntry(string ldapPath)
        {
            // create and return new LDAP connection with desired settings  

            DirectoryEntry ldapConnection = new DirectoryEntry(ldapPath);
            //ldapConnection.Path = ldapPath;
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;
            return ldapConnection;
        }

        static void WriteLog(string log, string filePath, string computername)
        {
            File.AppendAllText(filePath + @"\" + computername + "-DeptStamp.log", log);
        }
    }
}
