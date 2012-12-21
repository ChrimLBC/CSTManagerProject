using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Collections;
using System.Windows.Forms;
namespace CSTMgrProject
{
    /// <summary>
    /// Active Directory utilities class
    /// </summary>
    public class ActiveDirUtilities
    {
        /// <summary>
        /// Get all the domain names from the current forest
        /// </summary>
        /// <returns>ArrayList of domain names in the current forest</returns>
        public static ArrayList getDomains()
        {
            //Define array list for storing the domains
            ArrayList al = new ArrayList();
            //Get the current forest
            Forest forest = Forest.GetCurrentForest();

            //Get the domains from the forest
            DomainCollection dl = forest.Domains;
            //Foreach domain in the forest
            foreach (Domain domain in dl)
            {
                //Add the name of the domain to the array list
                al.Add(domain.Name);
            }
            //Return the array list of domain names

            return al;
        }

        /// <summary>
        /// Find all OUs beneath a given OU
        /// </summary>
        /// <param name="OuDn">Distinguished name of the parent OU</param>
        /// <returns>Children OUs from the given OU distinguished name</returns>
        public static ArrayList EnumerateOU(string OuDn)
        {
            //Define our return type of array list
            ArrayList alObjects = new ArrayList();
            try
            {
                //Create a DirectoryEntry object with the passed in OU distinguished name
                DirectoryEntry directoryObject = new DirectoryEntry("LDAP://" + OuDn);
                //Foreach child in the OU given
                foreach (DirectoryEntry child in directoryObject.Children)
                {
                    //Add the child to our array list to return
                    alObjects.Add(child);
                }
            }
            //If an error occurred
            catch (DirectoryServicesCOMException e)
            {
                //Display the error message in a message box.
                MessageBox.Show("An Error Occurred: " + e.Message.ToString());
            }
            //Return the array list of children of the given ou
            return alObjects;
        }

        /// <summary>
        /// Find all OUs beneath a given OU
        /// </summary>
        /// <param name="OuDn">Distinguished name of the parent OU</param>
        /// <param name="type">Type of objects to return</param>
        /// <returns>Children OUs from the given OU distinguished name</returns>
        public static ArrayList EnumerateOU(string OuDn, string type)
        {
            //Define our return type of array list
            ArrayList alObjects = new ArrayList();
            try
            {
                //Create a DirectoryEntry object with the passed in OU distinguished name
                DirectoryEntry directoryObject = new DirectoryEntry("LDAP://" + OuDn);
                DirectorySearcher ouSearch = new DirectorySearcher(directoryObject.Path);
                ouSearch.Filter = "(objectClass=" + type + ")";
                ouSearch.SearchRoot = directoryObject;
                ouSearch.SearchScope = SearchScope.OneLevel;
                SearchResultCollection allOUS = ouSearch.FindAll();
                //Foreach child in the OU given
                foreach (SearchResult child in allOUS)
                {
                    DirectoryEntry de = child.GetDirectoryEntry();
                    //Add the child to our array list to return
                    alObjects.Add(de);
                }
            }
            //If an error occurred
            catch (DirectoryServicesCOMException e)
            {
                //Display the error message in a message box.
                MessageBox.Show("An Error Occurred: " + e.Message.ToString());
            }
            //Return the array list of children of the given ou
            return alObjects;
        }

        #region Not Implemented Yet
        protected void getUserObjects()
        {
            //Implement in the modify users release
        }

        //protected getGroups()
        //{
        //    ArrayList al = new ArrayList();
        //    string defaultNamingContext;
        //    //TODO 0 - Acquire and display the available OU's
        //    DirectoryEntry rootDSE = new DirectoryEntry("LDAP://RootDSE");
        //    defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
        //    DirectoryEntry entryToQuery = new DirectoryEntry("LDAP://" + defaultNamingContext);
        //    //MessageBox.Show(entryToQuery.Path.ToString());

        //    DirectorySearcher ouSearch = new DirectorySearcher(entryToQuery.Path);
        //    ouSearch.Filter = "(objectCategory=organizationalUnit)";
        //    ouSearch.SearchScope = SearchScope.Subtree;
        //    ouSearch.PropertiesToLoad.Add("name");

        //    SearchResultCollection allOUS = ouSearch.FindAll();

        //    allOUS[0].GetDirectoryEntry().
        //}
        
        //public static ArrayList getSpecificOu(string OuDn)
        //{
        //    ArrayList al = new ArrayList();
        //    string defaultNamingContext = ("LDAP://" + OuDn);
        //    DirectoryEntry entryToQuery = new DirectoryEntry(defaultNamingContext);

        //    DirectorySearcher ouSearch = new DirectorySearcher(entryToQuery);
        //    ouSearch.Filter = "(objectCategory=organizationalUnit)";
        //    ouSearch.SearchScope = SearchScope.Base;
        //    ouSearch.PropertiesToLoad.Add("name");

        //    SearchResult allOUS = ouSearch.FindOne();

        //    foreach (String propkey in allOUS.Properties.PropertyNames)
        //    {
        //        foreach (Object property in allOUS.Properties[propkey])
        //        {
        //            al.Add(property.ToString());
        //        }

        //    }
        //    return al;
        //}
        #endregion


        public static void RadButtonHelper(string ouPath, ComboBox cboDC, ComboBox cboUserOU, ArrayList aUserOUPaths)
        {
            //Clear the combo box UserOu
            cboUserOU.Items.Clear();
            //Readd the default item
            cboUserOU.Items.Add("---Choose One---");
            string domainPath = PanelMgr.FriendlyDomainToLdapDomain(cboDC.Items[cboDC.SelectedIndex].ToString());

            aUserOUPaths.Clear();
            //Get a list of all OUs under the specified OU
            Array alOU = ActiveDirUtilities.EnumerateOU(ouPath + domainPath
            , "organizationalUnit").ToArray();
            //Foreach object in our array of OUs
            foreach (Object de in alOU)
            {
                //Cast to a directory entry
                DirectoryEntry ou = (DirectoryEntry)de;
                //Add the homeDirPath to our UserOU paths array
                aUserOUPaths.Add(ou.Path);
                //Add the name to the combo box UserOu
                cboUserOU.Items.Add(ou.Name.Substring(3));
            }
        }
    }
}
