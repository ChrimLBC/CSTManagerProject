using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CSTMgrProject;
using System.Collections;
using System.IO;
using System.DirectoryServices;


namespace UnitTestProject
{
    public class Tests
    {
        
            
       

        public static CrPanel initalizeVariables()
        {
            CrPanel crp = new CrPanel();
            crp.DC.Text = "project.com";
            crp.DC.SelectedIndex = 1;//project.com
            crp.Groups.SelectedIndex = 1;//AMT
            crp.UserOU.SelectedIndex = 1;//CST
            crp.NumUsers.Text = "3";
            crp.FirstUser.Text = "100";
            
            crp.Password.Text = "password";
            crp.passwordLength.Text = "7";

            
            crp.HomeDrive.SelectedIndex = 0;
            crp.HomeDir.Text = @"\\Prjdc\vol1\";
            
            crp.ProfilePath.Text = @"\\Prjdc\vol1\";
            
            
            return crp;
            
        }

        public static DelPanel initDel()
        {
            DelPanel dp = new DelPanel();
            dp.DC.Text = "project.com";
            dp.DC.SelectedIndex = 1;//Project.com
            dp.UserOU.SelectedIndex = 1;//CST
            dp.ServerPath.Text = @"\\prjdc\";
            dp.GetShares();
            dp.Share.SelectedIndex = 5;//vol1

            return dp;
        }

        public static void testDeleteVariables()
        {
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr();
            target.setDelVars(dp);

            Console.WriteLine("Testing Delete Panel Variables:");
            Console.WriteLine();

            Console.WriteLine("Domain Controller:");
            if (target.DelDC == dp.DC.Items[dp.DC.SelectedIndex].ToString())
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed");
            }
            Console.WriteLine();
            Console.WriteLine("User OU:");
            if (target.DelUserOU == dp.UserOU.Items[dp.UserOU.SelectedIndex].ToString())
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed");
            }

            Console.WriteLine();
            Console.WriteLine("Server Path:");
            if (target.ServerPath == dp.ServerPath.Text)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed");
            }

            Console.WriteLine();
            Console.WriteLine("Share:");
            if (target.Share == dp.Share.Items[dp.Share.SelectedIndex].ToString())
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed");
            }
        }

        /// <summary>
        ///A test for setting variables and getting proper values back
        ///</summary>
        
        
        public static void setVariablesTests()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);

            Console.WriteLine("Testing Create Panel Variables:");

            Console.WriteLine("User Prefix");
            if (target.UserPrefix == crp.UserPrefix.Text)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("User Group");
            
            if (target.UserGroup == crp.GroupPath)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("User OU");
            
            if (target.UserOU == crp.OUPath)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("Domain Controller");

            if (target.DC == crp.DC.Text)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("Number of Users");

            if (target.NumUsers == int.Parse(crp.NumUsers.Text))
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("First User");

            if (target.FirstUser == int.Parse(crp.FirstUser.Text))
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("Home Drive");

            if (target.HomeDrive == crp.HomeDrive.Items[crp.HomeDrive.SelectedIndex].ToString())
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("Home Directory");

            if (target.HomeDir == crp.HomeDir.Text + crp.HomeDirExt.Text.Substring(0, crp.HomeDirExt.Text.Length - 3))
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            Console.WriteLine();
            Console.WriteLine("Profile Path");

            if (target.ProfilePath == crp.ProfilePath.Text + crp.ProfileExt.Text.Substring(0, crp.ProfileExt.Text.Length - 3))
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }

            
            

            
            
        }

        public static void TestPassword()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);
 
            Console.WriteLine("Password Length");
            try
            {
                bool isGood = target.PasswdLength == int.Parse(crp.passwordLength.Text);
                if (isGood)
                {
                    Console.WriteLine("Test Passed!");
                }
                else
                {
                    Console.WriteLine("Test Failed!");
                }

            }

            
            catch (FormatException fe)
            {
                //There was nothing in the passwordlength textbox
                Console.WriteLine("Test Passed!");
            }
        }

        public static void TestDomainToLdapPath()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);

            Console.WriteLine("Testing converting domain name to LDAP path");
            string LDAP = PanelMgr.FriendlyDomainToLdapDomain(target.DC);

            if (LDAP == "DC=project,DC=com")
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
            
        }

        public static void TestGetDomains()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);
            ArrayList domains = new ArrayList();
            ArrayList getDomains = ActiveDirUtilities.getDomains();
            Console.WriteLine("Testing to get all the domains");
            bool isCorrect = true;
            
            for(int i = 1; i < crp.DC.Items.Count; i++)
            {
                domains.Add(crp.DC.Items[i]);
                if (domains[i - 1].ToString() != getDomains[i - 1].ToString())
                {
                    isCorrect = false;
                }
            }

            if (isCorrect)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
        }

        public static void TestEnumerateOU()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);

            ArrayList ous = new ArrayList();
            Array getOUs =  ActiveDirUtilities.EnumerateOU("ou=STU,ou=USR," +
                    PanelMgr.FriendlyDomainToLdapDomain(crp.DC.Items[crp.DC.SelectedIndex].ToString()), "organizationalUnit").ToArray();
            Console.WriteLine("Testing to enumerate OUs within the STU OU");
            bool isCorrect = true;
            ArrayList userOUs = new ArrayList();
            foreach (object o in getOUs)
            {
                DirectoryEntry de = (DirectoryEntry)o;

                userOUs.Add(de.Name.Substring(3));
            }
            for (int i = 1; i < crp.UserOU.Items.Count; i++)
            {
                ous.Add(crp.UserOU.Items[i]);
                if (ous[i - 1].ToString() != userOUs[i - 1].ToString())
                {
                    isCorrect = false;
                }
            }

            if (isCorrect)
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }
        }

        public static void TestErrorChecking()
        {
            CrPanel crp = initalizeVariables();
            DelPanel dp = initDel();
            PanelMgr target = new PanelMgr(); // TODO: Initialize to an appropriate value
            target.setVariables(crp);

            ErrorChecking ec = new ErrorChecking();
            Console.WriteLine("Testing the getMsg() functionality in the ErrorChecking class");
            ec.AddMsg("Testing error message");
            //Must add \n to the end of the line sicne ec.AddMsg() will concat \n to the string
            if (ec.getMsg().Equals("Testing error message\n"))
            {
                Console.WriteLine("Test Passed!");
            }
            else
            {
                Console.WriteLine("Test Failed!");
            }

        }

        public static void TestEnumShares()
        {
            DelPanel dp = initDel();
            CrPanel crp = initalizeVariables();
            PanelMgr target = new PanelMgr();
            target.setVariables(crp);

        }

        
        static void Main(string[] args)
        {
            
            setVariablesTests();
            Console.WriteLine();
            TestPassword();
            Console.WriteLine();
            TestDomainToLdapPath();
            Console.WriteLine();
            TestGetDomains();
            Console.WriteLine();
            TestEnumerateOU();
            Console.WriteLine();
            TestErrorChecking();
            Console.WriteLine();
            testDeleteVariables();
            Console.Read();
        }
    }
}
