using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Collections;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Threading;



namespace CSTMgrProject
{

    public class DelPanel : Panel
    {
        #region Enumerate Shares Functionality

        /// <summary>
        /// This Section is for getting shares based on server path.
        /// Handles getting shares from the network.
        /// </summary>

        #region External Calls
        [DllImport("Netapi32.dll", SetLastError = true)]
        static extern int NetApiBufferFree(IntPtr Buffer);
        [DllImport("Netapi32.dll", CharSet = CharSet.Unicode)]
        private static extern int NetShareEnum(
             StringBuilder ServerName,
             int level,
             ref IntPtr bufPtr,
             uint prefmaxlen,
             ref int entriesread,
             ref int totalentries,
             ref int resume_handle
             );
        #endregion
        #region External Structures
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHARE_INFO_1
        {
            public string shi1_netname;
            public uint shi1_type;
            public string shi1_remark;
            public SHARE_INFO_1(string sharename, uint sharetype, string remark)
            {
                this.shi1_netname = sharename;
                this.shi1_type = sharetype;
                this.shi1_remark = remark;
            }
            public override string ToString()
            {
                return shi1_netname;
            }
        }
        #endregion
        const uint MAX_PREFERRED_LENGTH = 0xFFFFFFFF;
        private Label lblServer;
        private TextBox txtServer;
        private Button btnDeleteFiles;
        private Button btnDelQuota;
        const int NERR_Success = 0;
        private enum NetError : uint
        {
            NERR_Success = 0,
            NERR_BASE = 2100,
            NERR_UnknownDevDir = (NERR_BASE + 16),
            NERR_DuplicateShare = (NERR_BASE + 18),
            NERR_BufTooSmall = (NERR_BASE + 23),
        }
        private enum SHARE_TYPE : uint
        {
            STYPE_DISKTREE = 0,
            STYPE_PRINTQ = 1,
            STYPE_DEVICE = 2,
            STYPE_IPC = 3,
            STYPE_SPECIAL = 0x80000000,
        }

        public SHARE_INFO_1[] EnumNetShares(string Server)
        {
            List<SHARE_INFO_1> ShareInfos = new List<SHARE_INFO_1>();
            int entriesread = 0;
            int totalentries = 0;
            int resume_handle = 0;
            int nStructSize = Marshal.SizeOf(typeof(SHARE_INFO_1));
            IntPtr bufPtr = IntPtr.Zero;
            StringBuilder server = new StringBuilder(Server);
            int ret = NetShareEnum(server, 1, ref bufPtr, MAX_PREFERRED_LENGTH, ref entriesread, ref totalentries, ref resume_handle);
            if (ret == NERR_Success)
            {
                IntPtr currentPtr = bufPtr;
                for (int i = 0; i < entriesread; i++)
                {
                    SHARE_INFO_1 shi1 = (SHARE_INFO_1)Marshal.PtrToStructure(currentPtr, typeof(SHARE_INFO_1));
                    ShareInfos.Add(shi1);
                    currentPtr = new IntPtr(currentPtr.ToInt32() + nStructSize);
                }
                NetApiBufferFree(bufPtr);
                return ShareInfos.ToArray();
            }
            else
            {
                ShareInfos.Add(new SHARE_INFO_1("ERROR=" + ret.ToString(), 10, string.Empty));
                return ShareInfos.ToArray();
            }
        }
        #endregion

        #region Constructor
        public DelPanel()
        {
            InitializeComponent();
        }
        #endregion

        #region Attributes
        /// <summary>
        /// Variable for handling error checking
        /// </summary>
        private ErrorChecking ec = new ErrorChecking();
        /// <summary>
        /// List of shares
        /// </summary>
        private SHARE_INFO_1[] shares;
        /// <summary>
        /// Panel manager to call certain helper methods
        /// </summary>
        private PanelMgr mgr = new PanelMgr();
        /// <summary>
        /// ArrayList of user LDAP paths
        /// </summary>
        private ArrayList aUserLDAPPaths = new ArrayList();
        /// <summary>
        /// ArrayList of OU LDAP paths
        /// </summary>
        private ArrayList aUserOUPaths = new ArrayList();
        /// <summary>
        /// Radio button for STU OU
        /// </summary>
        private RadioButton radStuOU;
        /// <summary>
        /// Radio button for CST OU
        /// </summary>
        private RadioButton radCstOU;
        /// <summary>
        /// Radio button for SVC OU
        /// </summary>
        private RadioButton radSvcOU;
        /// <summary>
        /// Label for the USER OU
        /// </summary>
        private Label lblUserOU;
        /// <summary>
        /// Combo box for selecting the USER OU
        /// </summary>
        private ComboBox cboUserOU;
        /// <summary>
        /// Groupbox to place all components in
        /// </summary>
        private GroupBox grpContainer;
        /// <summary>
        /// Label for the domain controller
        /// </summary>
        private Label lblDC;
        /// <summary>
        /// Combo box with teh domain controller
        /// </summary>
        private ComboBox cboDC;
        /// <summary>
        /// Checkbox to select all users
        /// </summary>
        private CheckBox chkSelectAll;
        /// <summary>
        /// Button to execute the delete users code
        /// </summary>
        private Button btnDeleteUser;
        /// <summary>
        /// Listbox with all the users in it
        /// </summary>
        private ListBox lstUsers;
        /// <summary>
        /// Button to reset all fields
        /// </summary>
        private Button btnReset;
        /// <summary>
        /// Combo box with share names
        /// </summary>
        private ComboBox cboShares;
        /// <summary>
        /// Label for the shares
        /// </summary>
        private Label lblShares;
        /// <summary>
        /// Label for selecting the OU
        /// </summary>
        private Label lblSetOU;

        #endregion

        #region Properties
        public ComboBox UserOU
        {
            get { return cboUserOU; }
        }
        public ComboBox DC
        {
            get { return cboDC; }
        }
        public TextBox ServerPath
        {
            get { return this.txtServer; }
        }
        public ComboBox Share
        {
            get { return this.cboShares; }
        }
        
        #endregion

        #region Initialize Components
        private void InitializeComponent()
        {
            this.lblSetOU = new System.Windows.Forms.Label();
            this.radStuOU = new System.Windows.Forms.RadioButton();
            this.radCstOU = new System.Windows.Forms.RadioButton();
            this.radSvcOU = new System.Windows.Forms.RadioButton();
            this.lblUserOU = new System.Windows.Forms.Label();
            this.cboUserOU = new System.Windows.Forms.ComboBox();
            this.grpContainer = new System.Windows.Forms.GroupBox();
            this.btnDeleteUser = new System.Windows.Forms.Button();
            this.cboDC = new System.Windows.Forms.ComboBox();
            this.cboShares = new System.Windows.Forms.ComboBox();
            this.lblShares = new System.Windows.Forms.Label();
            this.lblDC = new System.Windows.Forms.Label();
            this.chkSelectAll = new System.Windows.Forms.CheckBox();
            this.lstUsers = new System.Windows.Forms.ListBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.txtServer = new System.Windows.Forms.TextBox();
            this.btnDeleteFiles = new System.Windows.Forms.Button();
            this.lblServer = new System.Windows.Forms.Label();
            this.btnDelQuota = new System.Windows.Forms.Button();
            this.grpContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblSetOU
            // 
            this.lblSetOU.AutoSize = true;
            this.lblSetOU.Location = new System.Drawing.Point(30, 65);
            this.lblSetOU.Name = "lblSetOU";
            this.lblSetOU.Size = new System.Drawing.Size(45, 13);
            this.lblSetOU.TabIndex = 0;
            this.lblSetOU.Text = "Set OU:";
            // 
            // radStuOU
            // 
            this.radStuOU.AutoSize = true;
            this.radStuOU.Checked = true;
            this.radStuOU.Location = new System.Drawing.Point(30, 85);
            this.radStuOU.Name = "radStuOU";
            this.radStuOU.Size = new System.Drawing.Size(47, 17);
            this.radStuOU.TabIndex = 0;
            this.radStuOU.TabStop = true;
            this.radStuOU.Text = "STU";
            this.radStuOU.UseVisualStyleBackColor = true;
            this.radStuOU.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // radCstOU
            // 
            this.radCstOU.AutoSize = true;
            this.radCstOU.Location = new System.Drawing.Point(30, 105);
            this.radCstOU.Name = "radCstOU";
            this.radCstOU.Size = new System.Drawing.Size(46, 17);
            this.radCstOU.TabIndex = 0;
            this.radCstOU.TabStop = true;
            this.radCstOU.Text = "CST";
            this.radCstOU.UseVisualStyleBackColor = true;
            this.radCstOU.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // radSvcOU
            // 
            this.radSvcOU.AutoSize = true;
            this.radSvcOU.Location = new System.Drawing.Point(30, 125);
            this.radSvcOU.Name = "radSvcOU";
            this.radSvcOU.Size = new System.Drawing.Size(46, 17);
            this.radSvcOU.TabIndex = 0;
            this.radSvcOU.TabStop = true;
            this.radSvcOU.Text = "SVC";
            this.radSvcOU.UseVisualStyleBackColor = true;
            this.radSvcOU.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // lblUserOU
            // 
            this.lblUserOU.AutoSize = true;
            this.lblUserOU.Location = new System.Drawing.Point(30, 150);
            this.lblUserOU.Name = "lblUserOU";
            this.lblUserOU.Size = new System.Drawing.Size(51, 13);
            this.lblUserOU.TabIndex = 0;
            this.lblUserOU.Text = "User OU:";
            // 
            // cboUserOU
            // 
            this.cboUserOU.FormattingEnabled = true;
            this.cboUserOU.Items.AddRange(new object[] {
            "---Choose One---"});
            this.cboUserOU.Location = new System.Drawing.Point(30, 170);
            this.cboUserOU.Name = "cboUserOU";
            this.cboUserOU.Size = new System.Drawing.Size(121, 21);
            this.cboUserOU.TabIndex = 0;
            this.cboUserOU.SelectedIndexChanged += new System.EventHandler(this.cboUserOU_SelectedIndexChanged);
            // 
            // grpContainer
            // 
            this.grpContainer.Controls.Add(this.btnDelQuota);
            this.grpContainer.Controls.Add(this.btnDeleteUser);
            this.grpContainer.Controls.Add(this.cboDC);
            this.grpContainer.Controls.Add(this.cboUserOU);
            this.grpContainer.Controls.Add(this.cboShares);
            this.grpContainer.Controls.Add(this.lblShares);
            this.grpContainer.Controls.Add(this.lblDC);
            this.grpContainer.Controls.Add(this.chkSelectAll);
            this.grpContainer.Controls.Add(this.lblSetOU);
            this.grpContainer.Controls.Add(this.lblUserOU);
            this.grpContainer.Controls.Add(this.lstUsers);
            this.grpContainer.Controls.Add(this.radCstOU);
            this.grpContainer.Controls.Add(this.radSvcOU);
            this.grpContainer.Controls.Add(this.radStuOU);
            this.grpContainer.Controls.Add(this.btnReset);
            this.grpContainer.Controls.Add(this.txtServer);
            this.grpContainer.Controls.Add(this.btnDeleteFiles);
            this.grpContainer.Controls.Add(this.lblServer);
            this.grpContainer.Location = new System.Drawing.Point(4, 12);
            this.grpContainer.Name = "grpContainer";
            this.grpContainer.Size = new System.Drawing.Size(612, 294);
            this.grpContainer.TabIndex = 0;
            this.grpContainer.TabStop = false;
            // 
            // btnDeleteUser
            // 
            this.btnDeleteUser.Location = new System.Drawing.Point(180, 165);
            this.btnDeleteUser.Name = "btnDeleteUser";
            this.btnDeleteUser.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteUser.TabIndex = 0;
            this.btnDeleteUser.Text = "Delete User";
            this.btnDeleteUser.UseVisualStyleBackColor = true;
            this.btnDeleteUser.Click += new System.EventHandler(this.btnDeleteUser_Click);
            // 
            // cboDC
            // 
            this.cboDC.FormattingEnabled = true;
            this.cboDC.Items.AddRange(new object[] {
            "---Choose One---"});
            this.cboDC.Items.AddRange(ActiveDirUtilities.getDomains().ToArray());
            this.cboDC.Location = new System.Drawing.Point(30, 40);
            this.cboDC.Name = "cboDC";
            this.cboDC.Size = new System.Drawing.Size(121, 21);
            this.cboDC.TabIndex = 0;
            this.cboDC.SelectedIndexChanged += new System.EventHandler(this.cboDC_SelectedIndexChanged);
            // 
            // cboShares
            // 
            this.cboShares.FormattingEnabled = true;
            this.cboShares.Location = new System.Drawing.Point(400, 100);
            this.cboShares.Name = "cboShares";
            this.cboShares.Size = new System.Drawing.Size(121, 21);
            this.cboShares.TabIndex = 0;
            this.cboShares.Click += new System.EventHandler(this.cboShares_Click);
            // 
            // lblShares
            // 
            this.lblShares.AutoSize = true;
            this.lblShares.Location = new System.Drawing.Point(400, 80);
            this.lblShares.Name = "lblShares";
            this.lblShares.Size = new System.Drawing.Size(59, 13);
            this.lblShares.TabIndex = 0;
            this.lblShares.Text = "Disk Share";
            // 
            // lblDC
            // 
            this.lblDC.AutoSize = true;
            this.lblDC.Location = new System.Drawing.Point(30, 20);
            this.lblDC.Name = "lblDC";
            this.lblDC.Size = new System.Drawing.Size(93, 13);
            this.lblDC.TabIndex = 0;
            this.lblDC.Text = "Domain Controller:";
            // 
            // chkSelectAll
            // 
            this.chkSelectAll.AutoSize = true;
            this.chkSelectAll.Location = new System.Drawing.Point(180, 145);
            this.chkSelectAll.Name = "chkSelectAll";
            this.chkSelectAll.Size = new System.Drawing.Size(70, 17);
            this.chkSelectAll.TabIndex = 0;
            this.chkSelectAll.Text = "Select All";
            this.chkSelectAll.UseVisualStyleBackColor = true;
            this.chkSelectAll.CheckedChanged += new System.EventHandler(this.chkSelectAll_CheckedChanged);
            // 
            // lstUsers
            // 
            this.lstUsers.FormattingEnabled = true;
            this.lstUsers.Location = new System.Drawing.Point(180, 20);
            this.lstUsers.Name = "lstUsers";
            this.lstUsers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstUsers.Size = new System.Drawing.Size(200, 130);
            this.lstUsers.TabIndex = 0;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(280, 165);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 0;
            this.btnReset.Text = "Reset Fields";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // txtServer
            // 
            this.txtServer.Location = new System.Drawing.Point(400, 50);
            this.txtServer.Name = "txtServer";
            this.txtServer.Size = new System.Drawing.Size(100, 20);
            this.txtServer.TabIndex = 0;
            // 
            // btnDeleteFiles
            // 
            this.btnDeleteFiles.Location = new System.Drawing.Point(400, 130);
            this.btnDeleteFiles.Name = "btnDeleteFiles";
            this.btnDeleteFiles.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteFiles.TabIndex = 0;
            this.btnDeleteFiles.Text = "Delete Files";
            this.btnDeleteFiles.UseVisualStyleBackColor = true;
            this.btnDeleteFiles.Click += new System.EventHandler(this.btnDeleteFiles_Click);
            // 
            // lblServer
            // 
            this.lblServer.AutoSize = true;
            this.lblServer.Location = new System.Drawing.Point(400, 30);
            this.lblServer.Name = "lblServer";
            this.lblServer.Size = new System.Drawing.Size(169, 13);
            this.lblServer.TabIndex = 0;
            this.lblServer.Text = "Server Path  | e.g. \\\\ServerName\\";
            // 
            // btnDelQuota
            // 
            this.btnDelQuota.Location = new System.Drawing.Point(400, 160);
            this.btnDelQuota.Name = "btnDelQuota";
            this.btnDelQuota.Size = new System.Drawing.Size(100, 23);
            this.btnDelQuota.TabIndex = 0;
            this.btnDelQuota.Text = "Delete Quota";
            this.btnDelQuota.UseVisualStyleBackColor = true;
            this.btnDelQuota.Click += new System.EventHandler(this.btnDelQuota_Click);
            // 
            // DelPanel
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(127)))), ((int)(((byte)(36)))));
            this.Controls.Add(this.grpContainer);
            this.Size = new System.Drawing.Size(620, 290);
            this.grpContainer.ResumeLayout(false);
            this.grpContainer.PerformLayout();
            this.ResumeLayout(false);

        }

        
        #endregion

        #region Error Checking
        /// <summary>
        /// This method will check all fields for valid entries
        /// and return true if all values entered correctly.
        /// </summary>
        /// <returns></returns>
        private bool isCompleted()
        {
            bool retVal = true;
            this.ec.resetMsg();
            if (this.cboDC.SelectedIndex <= 0)
            {
                this.ec.AddMsg("Please select a domain controller");
                retVal = false;
            }

            if (this.cboUserOU.SelectedIndex <= 0)
            {
                this.ec.AddMsg("Please select an OU");
                retVal = false;
            }

            if (this.lstUsers.SelectedIndex < 0)
            {
                this.ec.AddMsg("Please select a user from the list");
                retVal = false;
            }

            if (this.txtServer.Text == "" || this.cboShares.SelectedIndex < 0)
            {
                DialogResult result = MessageBox.Show("You are about to delete users without deleting quotas, Continue?", "Delete Users", MessageBoxButtons.OKCancel);

                if (result == DialogResult.Cancel)
                {
                    retVal = false;
                }
                else
                {
                    retVal = true;
                }
            }

            return retVal;
        }
        #endregion

        #region Events
        /// <summary>
        /// When the Domain Controller combo box index is changed.
        /// This code will fill in our cboUserOU
        /// combo box with the correct OUs based on the DC
        /// that was chosen from the combo box
        /// </summary>
        /// <param name="sender">Object that was changed</param>
        /// <param name="e">Event that occurred</param>
        private void cboDC_SelectedIndexChanged(object sender, EventArgs e)
        {
            //If a valid DC was chosen
            if (cboDC.SelectedIndex > 0)
            {
                ActiveDirUtilities.RadButtonHelper("ou=STU,ou=USR,",cboDC,cboUserOU,aUserOUPaths);
                //deleteFilesOwnedBy("");
                Process scriptProc = new Process();
                scriptProc.StartInfo.FileName = @"cscript"; 
                scriptProc.StartInfo.Arguments = @"//B //Nologo D:\cst\CSTMgrProject\CSTMgrProject\script.vb";
                scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up
                scriptProc.Start();
                scriptProc.WaitForExit();
                scriptProc.Close();
            }
        }

        /// <summary>
        /// When a radio button is selected repopulate the user OU combo box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rad_CheckedChanged(object sender, EventArgs e)
        {
            OUChanged();
            //check if a Domain controller has been selected
            if (cboDC.SelectedIndex > 0)
            {
                sender = (RadioButton)sender;
                
                string ouPath = "";
                //depending on which button is selected set ouPath
                if (sender.Equals(radStuOU))
                {
                    ouPath ="ou=STU,ou=USR,";
                }
                else if (sender.Equals(radSvcOU))
                {
                    ouPath = "ou=SVC,ou=STU,ou=USR,";
                }
                else if (sender.Equals(radCstOU))
                {
                    ouPath = "ou=CST,ou=STU,ou=USR,";
                }
                ActiveDirUtilities.RadButtonHelper(ouPath, cboDC, cboUserOU, aUserOUPaths);
            }
            else if (cboDC.SelectedIndex < 0 && ((RadioButton)sender).Checked)
            {
                //inform user they haven't selected a Domain controller yet
                MessageBox.Show("Please select a domain controller first!");
            }
        }

        /// <summary>
        /// Method called when the user OU combo box selected index is changed.
        /// Handles filling out the listbox with all users under selected OU
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboUserOU_SelectedIndexChanged(object sender, EventArgs e)
        {
            OUChanged();
            if (cboUserOU.SelectedIndex > 0)
            {
                Array aLDAP = ActiveDirUtilities.EnumerateOU
                    (aUserOUPaths[cboUserOU.SelectedIndex - 1].ToString().Substring(7), "user").ToArray();
                ArrayList aNames = new ArrayList();
                foreach (object o in aLDAP)
                {
                    DirectoryEntry de = (DirectoryEntry)o;
                    aUserLDAPPaths.Add(de.Path);
                    aNames.Add(de.Name.Substring(3));
                }
                lstUsers.Items.AddRange(aNames.ToArray());
            }
        }
        /// <summary>
        /// Method called when the Select All checkbox is checked or unchecked.
        /// Handles selecting or deselecting all items from the list of users
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkSelectAll.Checked)
            {
                for (int i = 0, nLen = lstUsers.Items.Count; i < nLen; i++)
                {
                    lstUsers.SetSelected(i, true);
                }
            }
            else
            {
                for (int i = 0, nLen = lstUsers.Items.Count; i < nLen; i++)
                {
                    lstUsers.SetSelected(i, false);
                }
            }
        }

        /// <summary>
        /// When the user clicks the Delete User button it will go through and 
        /// delete all the selected users
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteUser_Click(object sender, EventArgs e)
        {
            if (isCompleted())
            {
                DialogResult result = MessageBox.Show("Are you sure you want to delete the selected users?", 
                        "Delete Users", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    

                    

                    int index = 0;
                    ArrayList alDelUsers = new ArrayList();
                    ArrayList userToDelete = new ArrayList();
                    //Create array list of profilepaths of each user we are about to delete
                    ArrayList alProfilePath = new ArrayList();
                    

                        foreach (string cn in lstUsers.SelectedItems)
                        {
                            index = lstUsers.Items.IndexOf(cn);
                            string finish = aUserLDAPPaths[index].ToString();
                            DirectoryEntry de = new DirectoryEntry(finish);
                            alDelUsers.Add(de);
                            userToDelete.Add(cn);

                            alProfilePath.Add(de.Properties["homeDirectory"].Value.ToString());  
                            alProfilePath.Add(de.Properties["profilePath"].Value.ToString());

                        }
                        
                        // for each path delete the directory and everything inside
                        foreach (string profilepath in alProfilePath)
                        {
                            //MessageBox.Show(profilepath);
                            DeleteDirectory(profilepath);
                        }
                        int i = 0;
                        ProgressWindow pgWindow = new ProgressWindow("Deleting Users...", userToDelete.Count);
                        pgWindow.Show();
                        pgWindow.SetPercentDone(i);
                        // Remove users from list in our Project
                        foreach (string cn in userToDelete)
                        {
                            pgWindow.SetPercentDone(i);
                            i++;
                            index = lstUsers.Items.IndexOf(cn);
                            lstUsers.Items.Remove(cn);

                            aUserLDAPPaths.RemoveAt(index);
                            
                        }
                        pgWindow.Close();
                        //Delete users in active directory
                        mgr.deleteUsers(alDelUsers);
                        lstUsers.Update();
                    
                }
            }
            else
            {
                MessageBox.Show(ec.getMsg());
            }
        }

        /// <summary>
        /// Method performed when clicking the Fill Shares button.
        /// This method will fill the cboShares combobox with all
        /// shares based on server path entered.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void cboShares_Click(object sender, EventArgs e)
        {
            GetShares();
        }

        /// <summary>
        /// Reset all fields in the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnReset_Click(object sender, EventArgs e)
        {
            this.cboDC.SelectedIndex = 0;
            this.cboUserOU.Items.Clear();
            this.cboUserOU.Text = "";
            this.cboShares.Text = "";
            this.lstUsers.Items.Clear();
            this.txtServer.Text = "";
            this.cboShares.Items.Clear();
            this.chkSelectAll.Checked = false;
        }
        
        /// <summary>
        /// Using the share path picked by the user and the users selected,
        /// delete all files in that share owned by the user selected.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDeleteFiles_Click(object sender, EventArgs e)
        {
            if (ServerPath.Text == "" || cboShares.SelectedIndex < 0)
            {
                MessageBox.Show("Please enter a server path and select a share");
            }
            else
            {
                if (lstUsers.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Please select a user to delete their files");
                }
                else
                {
                    
                    ProgressWindow pwBar = new ProgressWindow("Deleting Files");
                    pwBar.Show();
                    pwBar.Refresh();
                    Thread thProgress = new Thread(new ThreadStart(pwBar.Draw));
                    thProgress.Start();
                    foreach (string cn in lstUsers.SelectedItems)
                    {
                        try
                        {
                            RunScript(cn);
                        }
                        catch (Exception exp)
                        {
                            MessageBox.Show("There was an error running the delete files script.\n " +
                                "Seems like there was no files in the selected share");
                        }
                    }
                    thProgress.Abort();
                    pwBar.Close();
                }
            }
        }

        public void deleteQuota()
        {
            if (isCompleted())
            {
                string txtServerPath = this.txtServer.Text;
                if (txtServerPath.Length > 0)
                {
                    if (txtServerPath.ElementAt(txtServerPath.Length - 1).ToString() != "\\")
                    {
                        txtServerPath += "\\";
                    }
                }

                QuotaControl q = new QuotaControl();
                DialogResult cont = DialogResult.Abort;
                if (this.txtServer.Text != "")
                {
                    q.FILESHAREVOLUME1 = txtServerPath + this.cboShares.Items
                        [this.cboShares.SelectedIndex].ToString() + "\\";
                }

                try
                {
                    q.GetQuota(lstUsers.SelectedItems[0].ToString());
                    cont = DialogResult.Yes;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("The user selected doesn't have a quota on this share");
                }

                if (cont == DialogResult.Yes)
                {
                    int i = 0;
                    ProgressWindow pgWindow = new ProgressWindow("Deleting Users...", lstUsers.SelectedItems.Count);
                    pgWindow.Show();
                    pgWindow.SetPercentDone(i);
                    foreach (string cn in lstUsers.SelectedItems)
                    {

                        pgWindow.SetPercentDone(i);
                        i++;

                        if (txtServer.Text != "")
                        {
                            try
                            {
                                string domain = this.cboDC.Items[cboDC.SelectedIndex].ToString();
                                domain = domain.Substring(0, domain.IndexOf("."));
                                q.Remove(domain + cn);
                            }
                            catch (Exception ex)
                            {
                                //User already knows that the quota won't be deleted,
                                //no need to display error message here.
                            }
                        }
                    }



                    pgWindow.Close();
                    //Delete users in active directory


                }
            }
            else
            {
                MessageBox.Show(ec.getMsg());
            }
        }
        private void btnDelQuota_Click(object sender, EventArgs e)
        {
            deleteQuota();
        }
        /// <summary>
        /// When the user selects a new OU or User OU the list box will be cleared
        /// and the Select All check box will be unchecked
        /// </summary>
        public void OUChanged()
        {
            lstUsers.Items.Clear();
            aUserLDAPPaths.Clear();
            chkSelectAll.CheckState = CheckState.Unchecked;
        }
        #endregion

        #region File Shares Helper Methods
        /// <summary>
        /// Method to recursively turn each file and folder to normal permissions and then delete them
        /// </summary>
        /// <param name="deleteDir">Direcotry Path to delete</param>
        private void DeleteDirectory(string deleteDir)
        {
            if (Directory.Exists(deleteDir))
            {
                string[] files = Directory.GetFiles(deleteDir);
                string[] dirs = Directory.GetDirectories(deleteDir);

                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    File.Delete(file);
                }

                foreach (string dir in dirs)
                {
                    DeleteDirectory(dir);
                }

                Directory.Delete(deleteDir, false);
                //Directory.Delete(profilepath, true);  //recursion is already being done
            }
        }

        /// <summary>
        /// Find all shares within a certain server
        /// </summary>
        public void GetShares()
        {
            string servName = this.txtServer.Text.Trim();
            this.cboShares.Items.Clear();
            try
            {
                this.shares = EnumNetShares(servName);
                foreach (SHARE_INFO_1 s in this.shares)
                {
                    if (s.ToString() == "ERROR=53" || s.ToString() == "ERROR=123")
                    {
                        MessageBox.Show("There was a problem in retrieving the shares from " +
                            servName + ". Please verify that the server name is entered properly.");
                    }
                    else
                    {
                        this.cboShares.Items.Add(s.shi1_netname);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("There was a problem in retrieving the shares from " +
                    servName + ". Please verify that the server name is entered correctly.");
            }

        }

        /// <summary>
        /// Runs powershell script to delete users files in the specified share
        /// </summary>
        /// <param name="userName"></param>
        /// <returns></returns>
        private string RunScript(string userName)
        {
            // create Powershell runspace
            string scriptText = @"$testpath = Test-Path $path
                                    if ($testpath -eq 'True')
                                    {
                                        $files = Get-ChildItem $path -Recurse; 
                                        foreach ($file in $files)
                                        {	
                                           $owner = Get-Acl $file.FullName;	
                                            if ($owner.Owner -eq $user)
                                            {
	                                            Remove-Item $file.FullName -Recurse -Force;
                                            }	
                                        }
                                    }";
            //Get user string as DOMAIN\username
            string user = this.cboDC.Items[cboDC.SelectedIndex].ToString();
            user = user.Substring(0, user.IndexOf(".")).ToUpper() + "\\" + userName;
            //Get path of share
            string serverPath = this.txtServer.Text.Trim();
            if (!serverPath.ElementAt(serverPath.Length - 1).Equals("\\"))
            {
                serverPath += "\\";
            }
            string path = serverPath + this.cboShares.Items[cboShares.SelectedIndex].ToString();
            Runspace runspace = RunspaceFactory.CreateRunspace();


            // open it

            runspace.Open();
            runspace.SessionStateProxy.SetVariable("user", user.ToString());
            runspace.SessionStateProxy.SetVariable("path", path.ToString());
            // create a pipeline and feed it the script text

            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript(scriptText);

            // add an extra command to transform the script
            // output objects into nicely formatted strings

            // remove this line to get the actual objects
            // that the script returns. For example, the script

            // "Get-Process" returns a collection
            // of System.Diagnostics.Process instances.

            pipeline.Commands.Add("Out-String");

            // execute the script

            Collection<PSObject> results = pipeline.Invoke();

            // close the runspace

            runspace.Close();

            // convert the script result into a single string

            StringBuilder stringBuilder = new StringBuilder();
            foreach (PSObject obj in results)
            {
                stringBuilder.AppendLine(obj.ToString());
            }

            return stringBuilder.ToString();
        }

        
        #endregion

        
        
    }


}
