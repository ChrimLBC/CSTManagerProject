using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.Collections;
using System.IO;


namespace CSTMgrProject
{
    
    /// <summary>
    /// Panel class for the create users tab.
    /// </summary>
    public class CrPanel : Panel
    {
        #region  Attributes
        /// <summary>
        /// ArrayList of group LDAP paths
        /// </summary>
        private ArrayList aGroupPaths = new ArrayList();
        /// <summary>
        /// ArrayList of OU LDAP paths
        /// </summary>
        private ArrayList aUserOUPaths = new ArrayList();
        /// <summary>
        /// ErrorChecking variable (used as error message)
        /// </summary>
        private ErrorChecking ec = new ErrorChecking();
        /// <summary>
        /// Label set OU
        /// </summary>
        private Label lblSetOU;


        /// <summary>
        /// User OU label
        /// </summary>
        private Label lblUserOU;
        /// <summary>
        /// GroupBox for the selectables
        /// </summary>
        private GroupBox grpSelectables;
        /// <summary>
        /// ComboBox of user OUs
        /// </summary>
        private ComboBox cboUserOU;
        /// <summary>
        /// Label for the groups combo box
        /// </summary>
        private Label lblGroups;
        /// <summary>
        /// Combo box for the groups
        /// </summary>
        private ComboBox cboGroups;
        /// <summary>
        /// Label for domain controller
        /// </summary>
        private Label lblDC;
        /// <summary>
        /// Combo box for domain controllers
        /// </summary>
        private ComboBox cboDC;
        /// <summary>
        /// Button to create users
        /// </summary>
        private Button btnCreateUsers;
        /// <summary>
        /// Textbox for the OU
        /// </summary>
        private TextBox txtOU;
        /// <summary>
        /// First user label
        /// </summary>
        private Label lblFirstUser;
        /// <summary>
        /// Label for the number of users
        /// </summary>
        private Label lblNumUsers;
        /// <summary>
        /// Textbox with the first user number
        /// </summary>
        private TextBox txtFirstUser;
        /// <summary>
        /// Textbox with teh number of users
        /// </summary>
        private TextBox txtNumUsers;
        /// <summary>
        /// Label for the home drive
        /// </summary>
        private Label lblHomeDrive;
        /// <summary>
        /// Label for the home directory
        /// </summary>
        private Label lblHomeDir;

        /// <summary>
        /// Label for teh password
        /// </summary>
        private Label lblPassword;
        /// <summary>
        /// Textbox for the home directory
        /// </summary>
        private TextBox txtHomeDir;
        /// <summary>
        /// Textbox with teh password
        /// </summary>
        private TextBox txtPassword;
        /// <summary>
        /// GroupBox with all the fill-in fields
        /// </summary>
        private GroupBox grpFillIn;
        /// <summary>
        /// Label for the OU
        /// </summary>
        private Label lblOUString;
        /// <summary>
        /// Label for the user prefix
        /// </summary>
        private Label lblUserPrefix;
        /// <summary>
        /// Textbox for the user prefix
        /// </summary>
        private TextBox txtUserPrefix;
        /// <summary>
        /// Label for the home directory extension
        /// </summary>
        private Label lblHomeDirExt;
        /// <summary>
        /// Label for the OU extension
        /// </summary>
        private Label lblOUExt;
        /// <summary>
        /// Button to create extension users
        /// </summary>
        private Button btnCreateExtUsers;
        /// <summary>
        /// Radio button for CST ou
        /// </summary>
        private RadioButton radCstOu;
        /// <summary>
        /// Radio button for STU ou
        /// </summary>
        private RadioButton radSvcOu;
        /// <summary>
        /// Radio button for SVC ou
        /// </summary>
        private RadioButton radStuOu;
        /// <summary>
        /// Combo box with the home drive letter
        /// </summary>
        private ComboBox cboHomeDrive;
        /// <summary>
        /// Label for password length
        /// </summary>
        private Label lblPasswordLength;
        /// <summary>
        /// Textbox for password length
        /// </summary>
        private TextBox txtPasswordLength;
        /// <summary>
        /// textbox for profile homeDirPath
        /// </summary>
        private TextBox txtProfilePath;
        /// <summary>
        /// Label for profile homeDirPath
        /// </summary>
        private Label lblProfilePath;
        private Label lblProfileExt;
        /// <summary>
        /// Button to reset fields
        /// </summary>
        private Button btnResetFields;
        #endregion

        #region Constructor

        public CrPanel()
        {
            InitializeComponent();
        }
        #endregion

        #region Initialize Panel
        /// <summary>
        /// Initialize the panel
        /// </summary>
        private void InitializeComponent()
        {
            this.lblSetOU = new System.Windows.Forms.Label();
            this.lblUserOU = new System.Windows.Forms.Label();
            this.grpSelectables = new System.Windows.Forms.GroupBox();
            this.radCstOu = new System.Windows.Forms.RadioButton();
            this.radSvcOu = new System.Windows.Forms.RadioButton();
            this.radStuOu = new System.Windows.Forms.RadioButton();
            this.cboUserOU = new System.Windows.Forms.ComboBox();
            this.lblGroups = new System.Windows.Forms.Label();
            this.cboGroups = new System.Windows.Forms.ComboBox();
            this.lblDC = new System.Windows.Forms.Label();
            this.cboDC = new System.Windows.Forms.ComboBox();
            this.btnCreateUsers = new System.Windows.Forms.Button();
            this.txtOU = new System.Windows.Forms.TextBox();
            this.lblFirstUser = new System.Windows.Forms.Label();
            this.lblNumUsers = new System.Windows.Forms.Label();
            this.txtFirstUser = new System.Windows.Forms.TextBox();
            this.txtNumUsers = new System.Windows.Forms.TextBox();
            this.lblHomeDrive = new System.Windows.Forms.Label();
            this.lblHomeDir = new System.Windows.Forms.Label();
            this.lblPassword = new System.Windows.Forms.Label();
            this.txtHomeDir = new System.Windows.Forms.TextBox();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.btnResetFields = new System.Windows.Forms.Button();
            this.grpFillIn = new System.Windows.Forms.GroupBox();
            this.cboHomeDrive = new System.Windows.Forms.ComboBox();
            this.lblOUString = new System.Windows.Forms.Label();
            this.lblUserPrefix = new System.Windows.Forms.Label();
            this.txtUserPrefix = new System.Windows.Forms.TextBox();
            this.lblHomeDirExt = new System.Windows.Forms.Label();
            this.lblOUExt = new System.Windows.Forms.Label();
            this.btnCreateExtUsers = new System.Windows.Forms.Button();
            this.lblPasswordLength = new System.Windows.Forms.Label();
            this.txtPasswordLength = new System.Windows.Forms.TextBox();
            this.lblProfilePath = new System.Windows.Forms.Label();
            this.txtProfilePath = new System.Windows.Forms.TextBox();
            this.lblProfileExt = new System.Windows.Forms.Label();
            this.grpSelectables.SuspendLayout();
            this.grpFillIn.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblSetOU
            // 
            this.lblSetOU.Location = new System.Drawing.Point(30, 60);
            this.lblSetOU.Name = "lblSetOU";
            this.lblSetOU.Size = new System.Drawing.Size(45, 13);
            this.lblSetOU.TabIndex = 0;
            this.lblSetOU.Text = "Set OU:";
            // 
            // lblUserOU
            // 
            this.lblUserOU.AutoSize = true;
            this.lblUserOU.Location = new System.Drawing.Point(30, 145);
            this.lblUserOU.Name = "lblUserOU";
            this.lblUserOU.Size = new System.Drawing.Size(51, 13);
            this.lblUserOU.TabIndex = 0;
            this.lblUserOU.Text = "User OU:";
            // 
            // grpSelectables
            // 
            this.grpSelectables.Controls.Add(this.lblSetOU);
            this.grpSelectables.Controls.Add(this.radCstOu);
            this.grpSelectables.Controls.Add(this.radSvcOu);
            this.grpSelectables.Controls.Add(this.radStuOu);
            this.grpSelectables.Controls.Add(this.lblUserOU);
            this.grpSelectables.Controls.Add(this.cboUserOU);
            this.grpSelectables.Controls.Add(this.lblGroups);
            this.grpSelectables.Controls.Add(this.cboGroups);
            this.grpSelectables.Controls.Add(this.lblDC);
            this.grpSelectables.Controls.Add(this.cboDC);
            this.grpSelectables.Controls.Add(this.btnCreateUsers);
            this.grpSelectables.Location = new System.Drawing.Point(10, 10);
            this.grpSelectables.Name = "grpSelectables";
            this.grpSelectables.Size = new System.Drawing.Size(200, 280);
            this.grpSelectables.TabIndex = 0;
            this.grpSelectables.TabStop = false;
            this.grpSelectables.Text = "Select From The Following:";
            // 
            // radCstOu
            // 
            this.radCstOu.AutoSize = true;
            this.radCstOu.Location = new System.Drawing.Point(30, 100); 
            this.radCstOu.Name = "radCstOu";
            this.radCstOu.Size = new System.Drawing.Size(46, 17);
            this.radCstOu.TabIndex = 2;
            this.radCstOu.TabStop = true;
            this.radCstOu.Text = "CST";
            this.radCstOu.UseVisualStyleBackColor = true;
            this.radCstOu.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // radSvcOu
            // 
            this.radSvcOu.AutoSize = true;
            this.radSvcOu.Location = new System.Drawing.Point(30, 120); 
            this.radSvcOu.Name = "radSvcOu";
            this.radSvcOu.Size = new System.Drawing.Size(46, 17);
            this.radSvcOu.TabIndex = 3;
            this.radSvcOu.TabStop = true;
            this.radSvcOu.Text = "SVC";
            this.radSvcOu.UseVisualStyleBackColor = true;
            this.radSvcOu.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // radStuOu
            // 
            this.radStuOu.AutoSize = true;
            this.radStuOu.Checked = true;
            this.radStuOu.Location = new System.Drawing.Point(30, 80);
            this.radStuOu.Name = "radStuOu";
            this.radStuOu.Size = new System.Drawing.Size(47, 17);
            this.radStuOu.TabIndex = 1;
            this.radStuOu.TabStop = true;
            this.radStuOu.Text = "STU";
            this.radStuOu.UseVisualStyleBackColor = true;
            this.radStuOu.CheckedChanged += new System.EventHandler(this.rad_CheckedChanged);
            // 
            // cboUserOU
            // 
            this.cboUserOU.FormattingEnabled = true;
            //this.cboUserOU.Items.AddRange(new object[] {
            //"---Choose One---"});
            this.cboUserOU.Location = new System.Drawing.Point(30, 160);
            this.cboUserOU.Name = "cboUserOU";
            this.cboUserOU.Size = new System.Drawing.Size(121, 21);
            this.cboUserOU.TabIndex = 4;
            this.cboUserOU.SelectedIndexChanged += new System.EventHandler(this.cboUserOU_SelectedIndexChanged);
            // 
            // lblGroups
            // 
            this.lblGroups.AutoSize = true;
            this.lblGroups.Location = new System.Drawing.Point(30, 185);
            this.lblGroups.Name = "lblGroups";
            this.lblGroups.Size = new System.Drawing.Size(115, 13);
            this.lblGroups.TabIndex = 0;
            this.lblGroups.Text = "Groups in Domain/OU:";
            // 
            // cboGroups
            // 
            this.cboGroups.FormattingEnabled = true;
            //this.cboGroups.Items.AddRange(new object[] {
            //"---Choose One---"});
            this.cboGroups.Location = new System.Drawing.Point(30, 200);
            this.cboGroups.Name = "cboGroups";
            this.cboGroups.Size = new System.Drawing.Size(121, 21);
            this.cboGroups.TabIndex = 5;
            // 
            // lblDC
            // 
            this.lblDC.AutoSize = true;
            this.lblDC.Location = new System.Drawing.Point(30, 15);
            this.lblDC.Name = "lblDC";
            this.lblDC.Size = new System.Drawing.Size(93, 13);
            this.lblDC.TabIndex = 0;
            this.lblDC.Text = "Domain Controller:";
            // 
            // cboDC
            // 
            this.cboDC.FormattingEnabled = true;
            this.cboDC.Items.AddRange(new object[] {
            "---Choose One---"});
            this.cboDC.Items.AddRange(ActiveDirUtilities.getDomains().ToArray());
            this.cboDC.Location = new System.Drawing.Point(30, 30);
            this.cboDC.Name = "cboDC";
            this.cboDC.Size = new System.Drawing.Size(121, 21);
            this.cboDC.TabIndex = 0;
            this.cboDC.SelectedIndex = 0;
            this.cboDC.SelectedIndexChanged += new System.EventHandler(this.cboDC_SelectedIndexChanged);
            // 
            // btnCreateUsers
            // 
            this.btnCreateUsers.Location = new System.Drawing.Point(50, 250);
            this.btnCreateUsers.Name = "btnCreateUsers";
            this.btnCreateUsers.Size = new System.Drawing.Size(75, 23);
            this.btnCreateUsers.TabIndex = 14;
            this.btnCreateUsers.Text = "Create Users";
            this.btnCreateUsers.UseVisualStyleBackColor = true;
            this.btnCreateUsers.Click += new System.EventHandler(this.btnCreateUsers_Click);
            this.btnCreateUsers.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
            // 
            // txtOU
            // 
            this.txtOU.Location = new System.Drawing.Point(100, 55);
            this.txtOU.Name = "txtOU";
            this.txtOU.Size = new System.Drawing.Size(260, 20);
            this.txtOU.ReadOnly = true;
            this.txtOU.TabIndex = 0;
            // 
            // lblFirstUser
            // 
            this.lblFirstUser.AutoSize = true;
            this.lblFirstUser.Location = new System.Drawing.Point(30, 85);
            this.lblFirstUser.Name = "lblFirstUser";
            this.lblFirstUser.Size = new System.Drawing.Size(51, 13);
            this.lblFirstUser.TabIndex = 0;
            this.lblFirstUser.Text = "First User:";
            // 
            // lblNumUsers
            // 
            this.lblNumUsers.AutoSize = true;
            this.lblNumUsers.Location = new System.Drawing.Point(100, 85);
            this.lblNumUsers.Name = "lblNumUsers";
            this.lblNumUsers.Size = new System.Drawing.Size(66, 13);
            this.lblNumUsers.TabIndex = 0;
            this.lblNumUsers.Text = "No. of Users:";
            // 
            // txtFirstUser
            // 
            this.txtFirstUser.Location = new System.Drawing.Point(30, 100);
            this.txtFirstUser.Name = "txtFirstUser";
            this.txtFirstUser.Size = new System.Drawing.Size(51, 20);
            this.txtFirstUser.TabIndex = 7;
            // 
            // txtNumUsers
            // 
            this.txtNumUsers.Location = new System.Drawing.Point(100, 100);
            this.txtNumUsers.Name = "txtNumUsers";
            this.txtNumUsers.Size = new System.Drawing.Size(66, 20);
            this.txtNumUsers.TabIndex = 8;
            // 
            // lblHomeDrive
            // 
            this.lblHomeDrive.AutoSize = true;
            this.lblHomeDrive.Location = new System.Drawing.Point(25, 125);
            this.lblHomeDrive.Name = "lblHomeDrive";
            this.lblHomeDrive.Size = new System.Drawing.Size(63, 13);
            this.lblHomeDrive.TabIndex = 0;
            this.lblHomeDrive.Text = "Home Drive:";
            // 
            // lblHomeDir
            // 
            this.lblHomeDir.AutoSize = true;
            this.lblHomeDir.Location = new System.Drawing.Point(100, 125);
            this.lblHomeDir.Name = "lblHomeDir";
            this.lblHomeDir.Size = new System.Drawing.Size(80, 13);
            this.lblHomeDir.TabIndex = 0;
            this.lblHomeDir.Text = "Home Directory:";
            // 
            // lblPassword
            // 
            this.lblPassword.AutoSize = true;
            this.lblPassword.Location = new System.Drawing.Point(30, 195);
            this.lblPassword.Name = "lblPassword";
            this.lblPassword.Size = new System.Drawing.Size(170, 13);
            this.lblPassword.TabIndex = 0;
            this.lblPassword.Text = "Password (leave blank for random)";
            // 
            // txtHomeDir
            // 
            this.txtHomeDir.Location = new System.Drawing.Point(100, 140);
            this.txtHomeDir.Name = "txtHomeDir";
            this.txtHomeDir.Size = new System.Drawing.Size(200, 20);
            this.txtHomeDir.TabIndex = 10;
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(60, 210);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.Size = new System.Drawing.Size(100, 20);
            this.txtPassword.TabIndex = 12;
            // 
            // btnResetFields
            // 
            this.btnResetFields.Location = new System.Drawing.Point(50, 250);
            this.btnResetFields.Name = "btnResetFields";
            this.btnResetFields.Size = new System.Drawing.Size(75, 23);
            this.btnResetFields.TabIndex = 14;
            this.btnResetFields.Text = "Reset All Fields";
            this.btnResetFields.UseVisualStyleBackColor = true;
            this.btnResetFields.Click += new System.EventHandler(this.btnResetFields_Click);
            this.btnResetFields.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
            // 
            // grpFillIn
            // 
            this.grpFillIn.Controls.Add(this.txtOU);
            this.grpFillIn.Controls.Add(this.lblFirstUser);
            this.grpFillIn.Controls.Add(this.txtFirstUser);
            this.grpFillIn.Controls.Add(this.lblNumUsers);
            this.grpFillIn.Controls.Add(this.txtNumUsers);
            this.grpFillIn.Controls.Add(this.lblHomeDrive);
            this.grpFillIn.Controls.Add(this.cboHomeDrive);
            this.grpFillIn.Controls.Add(this.lblHomeDir);
            this.grpFillIn.Controls.Add(this.txtHomeDir);
            this.grpFillIn.Controls.Add(this.lblPassword);
            this.grpFillIn.Controls.Add(this.txtPassword);
            this.grpFillIn.Controls.Add(this.btnResetFields);
            this.grpFillIn.Controls.Add(this.lblOUString);
            this.grpFillIn.Controls.Add(this.lblUserPrefix);
            this.grpFillIn.Controls.Add(this.txtUserPrefix);
            this.grpFillIn.Controls.Add(this.lblHomeDirExt);
            this.grpFillIn.Controls.Add(this.lblOUExt);
            this.grpFillIn.Controls.Add(this.btnCreateExtUsers);
            this.grpFillIn.Controls.Add(this.lblPasswordLength);
            this.grpFillIn.Controls.Add(this.txtPasswordLength);
            this.grpFillIn.Controls.Add(this.lblProfilePath);
            this.grpFillIn.Controls.Add(this.txtProfilePath);
            this.grpFillIn.Controls.Add(this.lblProfileExt);
            this.grpFillIn.Location = new System.Drawing.Point(210, 10);
            this.grpFillIn.Name = "grpFillIn";
            this.grpFillIn.Size = new System.Drawing.Size(400, 280);
            this.grpFillIn.TabIndex = 0;
            this.grpFillIn.TabStop = false;
            this.grpFillIn.Text = "Fill in the blanks";
            // 
            // cboHomeDrive
            // 
            this.cboHomeDrive.FormattingEnabled = true;
            this.cboHomeDrive.Location = new System.Drawing.Point(35, 140);
            this.cboHomeDrive.Name = "cboHomeDrive";
            this.cboHomeDrive.Size = new System.Drawing.Size(40, 21);
            this.cboHomeDrive.TabIndex = 9;
            string drives = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            for (int i = 0, nLen = drives.Length; i < nLen; i++)
            {
                this.cboHomeDrive.Items.Add(drives.ElementAt(i) + ":");
            }
            // 
            // lblOUString
            // 
            this.lblOUString.AutoSize = true;
            this.lblOUString.Location = new System.Drawing.Point(45, 55);
            this.lblOUString.Name = "lblOUString";
            this.lblOUString.Size = new System.Drawing.Size(26, 13);
            this.lblOUString.TabIndex = 0;
            this.lblOUString.Text = "OU:";
            // 
            // lblUserPrefix
            // 
            this.lblUserPrefix.AutoSize = true;
            this.lblUserPrefix.Location = new System.Drawing.Point(30, 30);
            this.lblUserPrefix.Name = "lblUserPrefix";
            this.lblUserPrefix.Size = new System.Drawing.Size(58, 13);
            this.lblUserPrefix.TabIndex = 0;
            this.lblUserPrefix.Text = "User Prefix:";
            // 
            // txtUserPrefix
            // 
            this.txtUserPrefix.Location = new System.Drawing.Point(100, 30);
            this.txtUserPrefix.Name = "txtUserPrefix";
            this.txtUserPrefix.Size = new System.Drawing.Size(100, 20);
            this.txtUserPrefix.TabIndex = 6;
            this.txtUserPrefix.TextChanged += new System.EventHandler(this.txtUserPrefix_TextChanged);
            // 
            // lblHomeDirExt
            // 
            this.lblHomeDirExt.AutoSize = true;
            this.lblHomeDirExt.Location = new System.Drawing.Point(300, 140);
            this.lblHomeDirExt.Name = "lblHomeDirExt";
            this.lblHomeDirExt.Size = new System.Drawing.Size(0, 13);
            this.lblHomeDirExt.TabIndex = 0;
            // 
            // lblOUExt
            // 
            this.lblOUExt.AutoSize = true;
            this.lblOUExt.Location = new System.Drawing.Point(200, 55);
            this.lblOUExt.Name = "lblOUExt";
            this.lblOUExt.Size = new System.Drawing.Size(0, 13);
            this.lblOUExt.TabIndex = 0;
            // 
            // btnCreateExtUsers
            // 
            this.btnCreateExtUsers.Location = new System.Drawing.Point(150, 250);
            this.btnCreateExtUsers.Name = "btnCreateExtUsers";
            this.btnCreateExtUsers.Size = new System.Drawing.Size(75, 23);
            this.btnCreateExtUsers.TabIndex = 15;
            this.btnCreateExtUsers.Text = "Create Ext Users";
            this.btnCreateExtUsers.UseVisualStyleBackColor = true;
            this.btnCreateExtUsers.Click += new System.EventHandler(this.btnCreateExtUsers_Click);
            this.btnCreateExtUsers.BackColor = System.Drawing.ColorTranslator.FromHtml("#FFFF00");
            // 
            // lblPasswordLength
            // 
            this.lblPasswordLength.AutoSize = true;
            this.lblPasswordLength.Location = new System.Drawing.Point(230, 195);
            this.lblPasswordLength.Name = "lblPasswordLength";
            this.lblPasswordLength.Size = new System.Drawing.Size(89, 13);
            this.lblPasswordLength.TabIndex = 0;
            this.lblPasswordLength.Text = "Password Length";
            // 
            // txtPasswordLength
            // 
            this.txtPasswordLength.Location = new System.Drawing.Point(250, 210);
            this.txtPasswordLength.Name = "txtPasswordLength";
            this.txtPasswordLength.Size = new System.Drawing.Size(30, 20);
            this.txtPasswordLength.TabIndex = 13;
            // 
            // lblProfilePath
            // 
            this.lblProfilePath.AutoSize = true;
            this.lblProfilePath.Location = new System.Drawing.Point(30, 170);
            this.lblProfilePath.Name = "lblProfilePath";
            this.lblProfilePath.Size = new System.Drawing.Size(64, 13);
            this.lblProfilePath.TabIndex = 0;
            this.lblProfilePath.Text = "Profile Path:";
            // 
            // txtProfilePath
            // 
            this.txtProfilePath.Location = new System.Drawing.Point(100, 170);
            this.txtProfilePath.Name = "txtProfilePath";
            this.txtProfilePath.Size = new System.Drawing.Size(200, 20);
            this.txtProfilePath.TabIndex = 11;
            // 
            // lblProfileExt
            // 
            this.lblProfileExt.AutoSize = true;
            this.lblProfileExt.Location = new System.Drawing.Point(300, 170);
            this.lblProfileExt.Name = "lblProfileExt";
            this.lblProfileExt.Size = new System.Drawing.Size(0, 13);
            this.lblProfileExt.TabIndex = 0;
            // 
            // CrPanel
            // 
            this.Controls.Add(this.grpSelectables);
            this.Controls.Add(this.grpFillIn);
            this.grpSelectables.ResumeLayout(false);
            this.grpSelectables.PerformLayout();
            this.grpFillIn.ResumeLayout(false);
            this.grpFillIn.PerformLayout();
            this.ResumeLayout(false);

        }
        #endregion

        #region Error Checking
        /// <summary>
        /// Method for error checking, validating fields
        /// </summary>
        /// <returns>boolean if fields filled in properly.</returns>
        //This method is not finished yet
        private bool isCompleted()
        {
            //Reset the messageXX
            ec.resetMsg();
            //Initialize our return value
            bool filledOut = true;
            //Check each field in our GUI for invalid
            //information entries
            if (cboUserOU.SelectedIndex <= 0)
            {
                filledOut = false;
                ec.AddMsg("Please select an OU");
            }

            if (cboGroups.SelectedIndex <= 0)
            {
                filledOut = false;
                ec.AddMsg("Please select a group");
            }

            if (cboDC.SelectedIndex <= 0)
            {
                filledOut = false;
                ec.AddMsg("Please select a domain controller");
            }

            if (txtUserPrefix.Text == "")
            {
                filledOut = false;
                ec.AddMsg("User prefix wasn't specified");
            }

            if (txtOU.Text == "")
            {
                filledOut = false;
                ec.AddMsg("OU wasn't specified");
            }

            if (txtFirstUser.Text == "")
            {
                filledOut = false;
                ec.AddMsg("First user was not specified");
            }
            else if (txtFirstUser.Text != "")
            {
                try
                {
                    int num = int.Parse(txtFirstUser.Text);
                    if (num < 0)
                    {
                        filledOut = false;
                        ec.AddMsg("First user must be positive");
                    }
                }
                catch (FormatException fe)
                {
                    filledOut = false;
                    ec.AddMsg("First user must be an integer");
                }
            }

            if (txtNumUsers.Text == "")
            {
                filledOut = false;
                ec.AddMsg("Number of users was not specified");
            }
            else if (txtNumUsers.Text != "")
            {
                try
                {
                    int num = int.Parse(txtNumUsers.Text);
                    if (num <= 0)
                    {
                        filledOut = false;
                        ec.AddMsg("Number of users must be positive");
                    }
                }
                catch (FormatException fe)
                {
                    filledOut = false;
                    ec.AddMsg("Number of users must be an integer");
                }
            }

            if (txtHomeDir.Text == "")
            {
                filledOut = false;
                ec.AddMsg("Please specify the home directory");
            }

            if (cboHomeDrive.SelectedIndex < 0)
            {
                filledOut = false;
                ec.AddMsg("Please specify a home drive");
            }

            try
            {
                if (this.txtPasswordLength.Text.Length >= 5)
                {
                    int.Parse(this.txtPasswordLength.Text);
                }
            }
            catch (Exception e)
            {
                filledOut = false;
                ec.AddMsg("Password length must be an integer value");
            }
            string message = "Password field contains an invalid character: ";
            if (this.txtPassword.Text != "")
            {

                for (int i = 0; i < Password.Text.Length; i++)
                {
                    byte ascii = (byte)Password.Text.ElementAt(i);
                    if (!((ascii >= 97 && ascii <= 122) || (ascii >= 64 && ascii <= 90)
                        || (ascii >= 48 && ascii <= 57) || (ascii >= 35 && ascii <= 38)
                        || (ascii == 33) || (ascii == 42) || (ascii == 45)))
                    {

                        message += Password.Text.ElementAt(i);
                        filledOut = false;
                    }
                }
                ec.AddMsg(message);
            }
            if (txtPassword.Text == "")
            {
                try
                {
                    int length = int.Parse(txtPasswordLength.Text.Trim());
                    if (length > 50)
                    {
                        ec.AddMsg("Length of the password must be less than 50");
                        filledOut = false;
                    }
                }
                catch (Exception e)
                {
                    ec.AddMsg("Password length must be an integer value");
                    filledOut = false;
                }
            }

            string homeDirPath = "";
            string profilePath = "";
            try
            {
                //check to see if user entered an extra back slash
                if (txtHomeDir.Text[txtHomeDir.Text.Length - 1] == '\\')
                {
                    txtHomeDir.Text = txtHomeDir.Text.Substring(0, txtHomeDir.Text.Length - 1);
                }
                if (txtProfilePath.Text[txtProfilePath.Text.Length - 1] == '\\')
                {
                    txtProfilePath.Text = txtProfilePath.Text.Substring(0, txtProfilePath.Text.Length - 1);
                }
                homeDirPath = txtHomeDir.Text + lblHomeDirExt.Text.Substring(0, cboUserOU.Text.Length+2);
                profilePath = txtProfilePath.Text + lblProfileExt.Text.Substring(0, cboUserOU.Text.Length+2);
            }
            catch (Exception e)
            {
                //This catch block is to prevent the user from encountering an exception when
                //trying to create a user without a user prefix specified in the txtUserPrefix field.
            }

            if (!Directory.Exists(homeDirPath))
            {
                filledOut = false;
                ec.AddMsg("The home directory specified doesn't exist!\nInvalid directory: "
                    + homeDirPath);
            }

            if (!Directory.Exists(profilePath))
            {
                filledOut = false;
                ec.AddMsg("The profile path specified doesn't exist!\nInvalid directory: "
                    + profilePath);
            }


            return filledOut;
        }
        #endregion

        #region Properties
        /// <summary>
        /// Property for the button that creates extension accounts
        /// </summary>
        public Button CreateExtButton
        {
            get { return this.btnCreateExtUsers; }
        }
        /// <summary>
        /// Property for the button for creating users
        /// </summary>
        public Button CreateButton
        {
            get { return this.btnCreateUsers; }
        }
        /// <summary>
        /// Profile path textbox property
        /// </summary>
        public TextBox ProfilePath
        {
            get { return this.txtProfilePath; }
        }
        /// <summary>
        /// Password length if specified
        /// </summary>
        public TextBox passwordLength
        {
            get { return this.txtPasswordLength; }
            
        }
        /// <summary>
        /// Combobox property for User OU
        /// </summary>
        public ComboBox UserOU
        {
            get { return this.cboUserOU; }
            
        }
        /// <summary>
        /// Property for the LDAP Path for the group
        /// </summary>
        public string GroupPath
        {
            get { return this.aGroupPaths[this.cboGroups.SelectedIndex - 1].ToString(); }
        }
        /// <summary>
        /// Property for the LDAP Path for the User OU
        /// </summary>
        public string OUPath
        {
            get { return this.aUserOUPaths[this.cboUserOU.SelectedIndex - 1].ToString(); }
        }
        /// <summary>
        /// Combobox property for the Groups
        /// </summary>
        public ComboBox Groups
        {
            get { return this.cboGroups; }
        }
        /// <summary>
        /// Combobox property for the Domain Controller
        /// </summary>
        public ComboBox DC
        {
            get { return this.cboDC; }
        }
        /// <summary>
        /// Textbox proeprty for the User prefix
        /// </summary>
        public TextBox UserPrefix
        {
            get { return this.txtUserPrefix; }
        }
        /// <summary>
        /// Textbox property that contains the User OU
        /// </summary>
        public TextBox OU
        {
            get { return this.txtOU; }
        }
        /// <summary>
        /// Textbox property that contains the First User
        /// </summary>
        public TextBox FirstUser
        {
            get { return this.txtFirstUser; }
        }
        /// <summary>
        /// Textbox property that contains the Number of Users
        /// </summary>
        public TextBox NumUsers
        {
            get { return this.txtNumUsers; }
        }
        /// <summary>
        /// Textbox property that contains the Home Drive letter
        /// </summary>
        public ComboBox HomeDrive
        {
            get { return this.cboHomeDrive; }
        }
        /// <summary>
        /// Textbox property for the Home directory homeDirPath
        /// </summary>
        public TextBox HomeDir
        {
            get { return this.txtHomeDir; }
        }
        /// <summary>
        /// Textbox property for the Password
        /// </summary>
        public TextBox Password
        {
            get { return this.txtPassword; }
        }
        /// <summary>
        /// Property for the extension of the Home directory homeDirPath
        /// </summary>
        public Label HomeDirExt
        {
            get { return this.lblHomeDirExt; }
        }
        /// <summary>
        /// Property for the profile extension label
        /// </summary>
        public Label ProfileExt
        {
            get { return this.lblProfileExt; }
        }

        #endregion

        #region Events
        /// <summary>
        /// Code to be performed when the create users button is clicked
        /// </summary>
        /// <param name="sender">Object that sent the request</param>
        /// <param name="e">Event that occurred</param>
        private void btnCreateUsers_Click(object sender, EventArgs e)
        {
            //Instantiate a panel manager
            PanelMgr mgr = new PanelMgr();
            //Call the createUsers method on the panel manager
            //passing in this instance of CrPanel
            if (isCompleted())
            {
                mgr.createUsers(this);
                CreateButton.Enabled = true;
            }
            else
            {
                MessageBox.Show(ec.getMsg());
            }
        }


        /// <summary>
        /// When the Domain Controller combo box index is changed.
        /// This code will fill in our cboGroups and cboUserOU
        /// combo boxes with the correct OUs based on the DC
        /// that was chosen from the combo box
        /// </summary>
        /// <param name="sender">Object that was changed</param>
        /// <param name="e">Event that occurred</param>
        private void cboDC_SelectedIndexChanged(object sender, EventArgs e)
        {
            //If a valid DC was chosen
            if (cboDC.SelectedIndex > 0)
            {
                //Clear the userOU paths and groupPaths
                aUserOUPaths.Clear();
                aGroupPaths.Clear();
                /*
                 * Clear the combo boxes for groups and userOUs and
                 * re-add our default value "---Choose One---"
                 */
                this.cboGroups.Items.Clear();
                this.cboGroups.Items.Add("---Choose One---");
                DefaultCboValues();


                //Get an array of OUs under the STU homeDirPath(default).
                Array alOU = ActiveDirUtilities.EnumerateOU("ou=STU,ou=USR," +
                    //Gets the proper domain LDAP homeDirPath as well
                    PanelMgr.FriendlyDomainToLdapDomain(cboDC.Items[cboDC.SelectedIndex].ToString()), "organizationalUnit").ToArray();
                //Foreach OU that was found
                foreach (Object de in alOU)
                {
                    //Convert to directoryentry object
                    DirectoryEntry ou = (DirectoryEntry)de;
                    //Add the OU homeDirPath to our array of User OU paths
                    aUserOUPaths.Add(ou.Path);
                    //Add only the name of the OU to the combo box of User OUs
                    this.cboUserOU.Items.Add(ou.Name.Substring(3));
                }
                //Get all groups from the domain selected
                Array alGRPS = ActiveDirUtilities.EnumerateOU("ou=GRP,ou=USR," +
                    PanelMgr.FriendlyDomainToLdapDomain(cboDC.Items[cboDC.SelectedIndex].ToString()), "group").ToArray();
                //Foreach group found
                foreach (Object de in alGRPS)
                {
                    //Convert to directoryentry object
                    DirectoryEntry ou = (DirectoryEntry)de;
                    //Add the group LDAP homeDirPath to our array of group Paths
                    aGroupPaths.Add(ou.Path);
                    //Add only the name of the group to the combo box of groups
                    this.cboGroups.Items.Add(ou.Name.Substring(3));
                }
            }
            else if (cboDC.SelectedIndex <= 0)
            {
                //Make User OU and Group OU combo boxes blank
                this.cboUserOU.Items.Clear();
                this.cboUserOU.Text = "";
                this.cboGroups.Items.Clear();
                this.cboGroups.Text = "";
                //Display Message
                MessageBox.Show("Please select a domain controller");
            }
        }

        /// <summary>
        /// Code to perform when combo box UserOU index is changed.
        /// This code will auto fill in the txtOU textbox with
        /// the correct DN for the UserOU selected. This also
        /// fills in part of the home directory homeDirPath for you.
        /// Also sets the user prefix textbox.
        /// </summary>
        /// <param name="sender">Object that was changed</param>
        /// <param name="e">Event that occurred on this object</param>
        private void cboUserOU_SelectedIndexChanged(object sender, EventArgs e)
        {
            //If a valid selection was chosen
            if (cboUserOU.SelectedIndex > 0)
            {
                //Set the user prefix to the name of the user OU
                txtUserPrefix.Text = this.cboUserOU.Items[this.cboUserOU.SelectedIndex].ToString();
                //Set txtOU to the correct DN homeDirPath for the selected OU
                txtOU.Text = this.aUserOUPaths[this.cboUserOU.SelectedIndex - 1].ToString().Substring(7);
                //Set the label for the home directory.
                lblHomeDirExt.Text = "\\" + txtUserPrefix.Text + "\\" + txtUserPrefix.Text + "xxx";
                //Set the label for the profile path extension
                lblProfileExt.Text = "\\" + txtUserPrefix.Text + "\\" + txtUserPrefix.Text + "xxx";
            }
            //If it wasn't a valid selection
            else
            {
                //Reset the user prefix and OU, and lblHomeDirExt, and profileExt label
                txtUserPrefix.Text = "";
                txtOU.Text = "";
                lblHomeDirExt.Text = "";
                lblProfileExt.Text = "";
            }
        }

        /// <summary>
        /// Will set the combo boxes to default to let user know they must reselect the combo boxes
        /// </summary>
        private void DefaultCboValues()
        {
            // reset list in cboUserOU so they can't select something that isn't in said OU
            this.cboUserOU.Items.Clear();
            this.cboUserOU.Items.Add("---Choose One---");
            //populate combo box with the index is "---Choose One---"
            this.cboUserOU.SelectedIndex = 0;

            //populate cboGroups with the index is "---Choose One---"
            this.cboGroups.SelectedIndex = 0;
            
        }

        /// <summary>
        /// Event that occurs when the Create Ext Accounts button is clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCreateExtUsers_Click(object sender, EventArgs e)
        {
            //Call the create extension users method
            if (isCompleted())
            {
                PanelMgr mgr = new PanelMgr();
                mgr.createExtUsers(this);
                CreateButton.Enabled = true;
            }
            else
            {
                MessageBox.Show(ec.getMsg());
            }
        }

        /// <summary>
        /// When a radio button is selected repopulate the user OU combo box
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rad_CheckedChanged(object sender, EventArgs e)
        {

            //check if a Domain controller has been selected
            if (cboDC.SelectedIndex > 0)
            {
                /*
                 * Clear the combo boxes for groups and userOUs and
                 * re-add our default value "---Choose One---"
                 */
                DefaultCboValues();

                sender = (RadioButton)sender;
                string ouPath = "";
                //depending on which button is selected set ouPath
                if (sender.Equals(radStuOu))
                {
                    ouPath = "ou=STU,ou=USR,";
                }
                else if (sender.Equals(radSvcOu))
                {
                    ouPath = "ou=SVC,ou=STU,ou=USR,";
                }
                else if (sender.Equals(radCstOu))
                {
                    ouPath = "ou=CST,ou=STU,ou=USR,";
                }
                ActiveDirUtilities.RadButtonHelper(ouPath, cboDC, cboUserOU, aUserOUPaths);
            }
            else if (cboDC.SelectedIndex < 1 && ((RadioButton)sender).Checked)
            {
                //inform user they haven't selected a Domain controller yet
                MessageBox.Show("Please select a domain controller first!");
            }

        }

        /// <summary>
        /// Reset all fields on the panel.
        /// </summary>
        /// <param name="sender">Object that send the request</param>
        /// <param name="e">Event that occurred</param>
        private void btnResetFields_Click(object sender, EventArgs e)
        {
            //Don't clear items from DC but reset the text
            this.cboDC.Text = "";
            //Clear all items and reset text for the combo boxes
            this.cboUserOU.Items.Clear();
            this.cboHomeDrive.Text = "";
            this.cboGroups.Items.Clear();
            this.cboGroups.Text = "";
            this.cboUserOU.Text = "";
            //Set all textboxes text to nothing
            this.txtFirstUser.Text = "";
            this.txtHomeDir.Text = "";
            this.txtNumUsers.Text = "";
            this.txtOU.Text = "";
            this.lblHomeDirExt.Text = "";
            this.txtUserPrefix.Text = "";
            this.txtPassword.Text = "";
            this.txtPasswordLength.Text = "";
            this.txtProfilePath.Text = "";
            
        }

        private void txtUserPrefix_TextChanged(object sender, EventArgs e)
        {
            if (cboUserOU.SelectedIndex > 0)
            {
                this.lblHomeDirExt.Text = "\\" + this.cboUserOU.Items[this.cboUserOU.SelectedIndex].ToString() + @"\" +
                    this.txtUserPrefix.Text + "xxx";
                this.lblProfileExt.Text = "\\" + this.cboUserOU.Items[this.cboUserOU.SelectedIndex].ToString() + @"\" +
                    this.txtUserPrefix.Text + "xxx";
            }
        }
        #endregion

        #region Generated stuff
        public PanelMgr PanelMgr
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public ErrorChecking ErrorChecking
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }

        public ActiveDirUtilities ActiveDirUtilities
        {
            get
            {
                throw new System.NotImplementedException();
            }
            set
            {
            }
        }
        #endregion
    }
}
