using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DiskQuotaTypeLibrary;



namespace CSTMgrProject
{
    public class QuotaControl
    {
        private DiskQuotaControl _diskQuotaControl;
        
        //This path has to be in this format or 
        //else is going to give an error of invalid path
        private string FILESHAREVOLUME = @"\\PRJDC\vol1\";
        /// <summary>
        /// Volume in which to search/manage quotas
        /// </summary>
        public string FILESHAREVOLUME1
        {
            get { return FILESHAREVOLUME; }
            set { FILESHAREVOLUME = value; }
        }
        private const int MBTOBYTES = 1048576;


        /// <summary>
        /// DiskQuotaControl property
        /// </summary>
        public DiskQuotaControl DiskQuotaControl
        {
            
            get
            {
                if (this._diskQuotaControl == null)
                {
                    this._diskQuotaControl = new DiskQuotaControl();
                    //Initializes the control to the specified path
                    this._diskQuotaControl.Initialize(FILESHAREVOLUME, true);
                }
                return this._diskQuotaControl;
            }
        }

        public QuotaControl()
        {
        }

        /// <summary>
        /// Removes the user form the Quota Entries List
        /// </summary>
        /// <PARAM name="userName"></PARAM>
        public void Remove(string userName)
        {
            //Here we just use the user, and invoke 
            //the method DeleteUser from the control
            this.DiskQuotaControl.DeleteUser(this.GetUser(userName));
        }

        /// <summary>
        /// A private function to return the user object
        /// </summary>
        /// <PARAM name="userName">The user name</PARAM>
        /// <returns>A DIDiskQuotaUser Object 
        /// of the specified user</returns>
        public DIDiskQuotaUser GetUser(string userName)
        {
            //Invokes the method to find a user in a quota entry list
            return this.DiskQuotaControl.FindUser(userName);
        }

        /// <summary>
        /// Gets the quota of the user
        /// </summary>
        /// <PARAM name="userName">The user name</PARAM>
        /// <returns>A formatted string of the quota 
        /// limit of the user</returns>
        public string GetQuota(string userName)
        {
            //here we return the text of the quota limit
            //0.0 bytes, 0.0 Kb, 0.0 Mb etc
            return this.GetUser(userName).QuotaLimitText;
        }

        /// <summary>
        /// Gets the quota currently used by the user
        /// </summary>
        /// <PARAM name="userName">The user name </PARAM>
        /// <returns>A formatted string of the quota 
        /// used by the user</returns>
        public string GetQuotaUsed(string userName)
        {
            return this.GetUser(userName).QuotaUsedText;
        }


    }

}
