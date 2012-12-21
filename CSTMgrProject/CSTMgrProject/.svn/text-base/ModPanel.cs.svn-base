using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.DirectoryServices;
using System.IO;
using System.Management;
using System.Drawing;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using System.Reflection;
using System.Collections;
using System.ComponentModel;
using System.Security.Permissions;
using System.Net.NetworkInformation;
using System.Threading;
//using System.Activities.Presentation.PropertyEditing;


namespace CSTMgrProject
{

    public class ModPanel : Panel
    {

        #region Constructor
        public ModPanel()
        {
            InitializeComponent();
        }
        #endregion

        #region Attributes
        /// <summary>
        /// Dictionary variable to store key/value pairs
        /// specifically for holding user propertie names/values
        /// </summary>
        private Dictionary<string, string> dict;
        /// <summary>
        /// Panel manager to call certain helper methods
        /// </summary>
        PanelMgr mgr = new PanelMgr();
        /// <summary>
        /// ArrayList of user LDAP paths
        /// </summary>
        private ArrayList aUserLDAPPaths = new ArrayList();
        /// <summary>
        /// ArrayList of OU LDAP paths
        /// </summary>
        private ArrayList aUserOUPaths = new ArrayList();
        /// <summary>
        /// Create array list of users who are selected for modifying their properties
        /// </summary>
        private ArrayList alModifyUsers = new ArrayList();
        /// <summary>
        /// Create array list of writable properties to be put in properties column of dgvUserProperties
        /// </summary>
        private ArrayList alWritableProperties = new ArrayList();
        /// <summary>
        /// Create array list of current property values to be put in properties column of dgvUserProperties
        /// </summary>
        private ArrayList alCurrentPropertyValues = new ArrayList();

        #region Duplicate of Delete User
        /// <summary>
        /// Label for the domain controller
        /// </summary>
        private Label lblDC;
        /// <summary>
        /// Combo box with teh domain controller
        /// </summary>
        private ComboBox cboDC;
        /// <summary>
        /// Label for selecting the OU
        /// </summary>
        private Label lblSetOU;
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
        #endregion

        /// <summary>
        /// Groupbox to place all components in
        /// </summary>
        private GroupBox grpContainer;
        /// <summary>
        /// Label for the User List Box
        /// </summary>
        private Label lblUsers;
        /// <summary>
        /// Label for the User Properties
        /// </summary>
        private Label lblUserProperties;
        /// <summary>
        /// Listbox with all the users in it
        /// </summary>
        private ListBox lstUsers;
        /// <summary>
        /// Data Grid View with all user properties and values
        /// </summary>
        private modDataGridView dgvUserProperties;
        /// <summary>
        /// Checkbox to select all users
        /// </summary>
        private CheckBox chkSelectAll;
        /// <summary>
        /// Button to reset password of selected user(s)
        /// </summary>
        private Button btnResetPassword;
        /// <summary>
        /// Button to save changes of user properties
        /// </summary>
        private Button btnSave;
        /// <summary>
        /// Button to discard changes of user properties
        /// </summary>
        private Button btnDiscard;

        #endregion

        #region Initialize Panel

        /// <summary>
        /// Initialize the ModPanel
        /// </summary>
        private void InitializeComponent()
        {
            this.cboDC = new System.Windows.Forms.ComboBox();
            this.lblDC = new System.Windows.Forms.Label();
            this.lblSetOU = new System.Windows.Forms.Label();
            this.radStuOU = new System.Windows.Forms.RadioButton();
            this.radCstOU = new System.Windows.Forms.RadioButton();
            this.radSvcOU = new System.Windows.Forms.RadioButton();
            this.lblUserOU = new System.Windows.Forms.Label();
            this.cboUserOU = new System.Windows.Forms.ComboBox();
            this.lblUsers = new System.Windows.Forms.Label();
            this.lblUserProperties = new System.Windows.Forms.Label();
            this.lstUsers = new System.Windows.Forms.ListBox();
            this.dgvUserProperties = new CSTMgrProject.modDataGridView();
            this.chkSelectAll = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnResetPassword = new System.Windows.Forms.Button();
            this.btnDiscard = new System.Windows.Forms.Button();
            this.grpContainer = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgvUserProperties)).BeginInit();
            this.grpContainer.SuspendLayout();
            this.SuspendLayout();
            // 
            // cboDC
            // 
            this.cboDC.FormattingEnabled = true;
            this.cboDC.Items.AddRange(new object[] { "---Choose One---" });
            this.cboDC.Items.AddRange(ActiveDirUtilities.getDomains().ToArray());
            //this.cboDC.Items.AddRange(ActiveDirUtilities.getDomains().ToArray());
            this.cboDC.Location = new System.Drawing.Point(30, 40);
            this.cboDC.Name = "cboDC";
            this.cboDC.Size = new System.Drawing.Size(121, 21);
            this.cboDC.TabIndex = 0;
            this.cboDC.SelectedIndexChanged += new System.EventHandler(this.cboDC_SelectedIndexChanged);
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
            // lblUsers
            // 
            this.lblUsers.AutoSize = true;
            this.lblUsers.Location = new System.Drawing.Point(170, 20);
            this.lblUsers.Name = "lblUsers";
            this.lblUsers.Size = new System.Drawing.Size(34, 13);
            this.lblUsers.TabIndex = 0;
            this.lblUsers.Text = "Users";
            // 
            // lblUserProperties
            // 
            this.lblUserProperties.AutoSize = true;
            this.lblUserProperties.Location = new System.Drawing.Point(325, 20);
            this.lblUserProperties.Name = "lblUserProperties";
            this.lblUserProperties.Size = new System.Drawing.Size(79, 13);
            this.lblUserProperties.TabIndex = 0;
            this.lblUserProperties.Text = "User Properties";
            // 
            // lstUsers
            // 
            this.lstUsers.FormattingEnabled = true;
            this.lstUsers.Location = new System.Drawing.Point(170, 40);
            this.lstUsers.Name = "lstUsers";
            this.lstUsers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstUsers.Size = new System.Drawing.Size(139, 95);
            this.lstUsers.TabIndex = 0;
            this.lstUsers.SelectedIndexChanged += new System.EventHandler(this.lstUsers_SelectedIndexChanged);
            // 
            // dgvUserProperties
            // 
            this.dgvUserProperties.AllowUserToAddRows = false;
            this.dgvUserProperties.AllowUserToDeleteRows = false;
            this.dgvUserProperties.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvUserProperties.Location = new System.Drawing.Point(325, 40);
            this.dgvUserProperties.MultiSelect = false;
            this.dgvUserProperties.Name = "dgvUserProperties";
            this.dgvUserProperties.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvUserProperties.Size = new System.Drawing.Size(275, 187);
            this.dgvUserProperties.TabIndex = 0;
            this.dgvUserProperties.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dgvUserProperties_CellMouseDoubleClick);
            this.dgvUserProperties.SelectionChanged += new System.EventHandler(this.dgvUserProperties_SelectionChanged);
            // 
            // chkSelectAll
            // 
            this.chkSelectAll.AutoSize = true;
            this.chkSelectAll.Location = new System.Drawing.Point(170, 145);
            this.chkSelectAll.Name = "chkSelectAll";
            this.chkSelectAll.Size = new System.Drawing.Size(70, 17);
            this.chkSelectAll.TabIndex = 0;
            this.chkSelectAll.Text = "Select All";
            this.chkSelectAll.UseVisualStyleBackColor = true;
            this.chkSelectAll.CheckedChanged += new System.EventHandler(this.chkSelectAll_CheckedChanged);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(325, 230);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(168, 23);
            this.btnSave.TabIndex = 0;
            this.btnSave.Text = "Save Changes";
            this.btnSave.UseVisualStyleBackColor = true;
            // 
            // btnResetPassword
            // 
            this.btnResetPassword.Location = new System.Drawing.Point(170, 169);
            this.btnResetPassword.Name = "btnResetPassword";
            this.btnResetPassword.Size = new System.Drawing.Size(139, 24);
            this.btnResetPassword.TabIndex = 0;
            this.btnResetPassword.Text = "Reset Password(s)";
            this.btnResetPassword.UseVisualStyleBackColor = true;
            this.btnResetPassword.Click += new System.EventHandler(this.btnResetPassword_Click);
            // 
            // btnDiscard
            // 
            this.btnDiscard.Location = new System.Drawing.Point(325, 253);
            this.btnDiscard.Name = "btnDiscard";
            this.btnDiscard.Size = new System.Drawing.Size(168, 23);
            this.btnDiscard.TabIndex = 0;
            this.btnDiscard.Text = "Discard Changes";
            this.btnDiscard.UseVisualStyleBackColor = true;
            // 
            // grpContainer
            // 
            this.grpContainer.Controls.Add(this.lblDC);
            this.grpContainer.Controls.Add(this.cboDC);
            this.grpContainer.Controls.Add(this.lblSetOU);
            this.grpContainer.Controls.Add(this.radCstOU);
            this.grpContainer.Controls.Add(this.radSvcOU);
            this.grpContainer.Controls.Add(this.radStuOU);
            this.grpContainer.Controls.Add(this.lblUserOU);
            this.grpContainer.Controls.Add(this.cboUserOU);
            this.grpContainer.Controls.Add(this.lblUsers);
            this.grpContainer.Controls.Add(this.lblUserProperties);
            this.grpContainer.Controls.Add(this.lstUsers);
            this.grpContainer.Controls.Add(this.dgvUserProperties);
            this.grpContainer.Controls.Add(this.chkSelectAll);
            this.grpContainer.Controls.Add(this.btnResetPassword);
            this.grpContainer.Controls.Add(this.btnSave);
            this.grpContainer.Controls.Add(this.btnDiscard);
            this.grpContainer.Location = new System.Drawing.Point(4, 12);
            this.grpContainer.Name = "grpContainer";
            this.grpContainer.Size = new System.Drawing.Size(612, 294);
            this.grpContainer.TabIndex = 0;
            this.grpContainer.TabStop = false;
            // 
            // ModPanel
            // 
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(127)))), ((int)(((byte)(36)))));
            this.Controls.Add(this.grpContainer);
            this.Size = new System.Drawing.Size(620, 290);
            ((System.ComponentModel.ISupportInitialize)(this.dgvUserProperties)).EndInit();
            this.grpContainer.ResumeLayout(false);
            this.grpContainer.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion


        #region Methods

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
                ActiveDirUtilities.RadButtonHelper("ou=STU,ou=USR,", cboDC, cboUserOU, aUserOUPaths);

                //Find the writable usr properties
                createWriteableAttribList();

                this.cboUserOU.SelectedIndex = 0;
            }
            else if (cboDC.SelectedIndex <= 0)
            {
                //Make User OU blank and list box empty
                this.cboUserOU.SelectedIndex = 0;
                this.cboUserOU.Items.Clear();
                this.cboUserOU.Text = "";
                lstUsers.Items.Clear();
                populateDataGridView();
                //Display Message
                MessageBox.Show("Please select a domain controller");
            }
        }
        /// <summary>
        ///Function: go through an populates an arrayList containing all writable user attributes.
        /// </summary>
        private void createWriteableAttribList()
        {

            //First, get schemea of properties, both single and multivalued
            //Get current server
            string ADServer = Environment.GetEnvironmentVariable("LOGONSERVER");
            //Remove the "//"
            ADServer = ADServer.Substring(2);
            //Get the selected domain
            string domainName = cboDC.Text.ToString();
            //Concatenate to make full server name
            string fullServerName = ADServer + "." + domainName;
            //MessageBox.Show("The directory context is: " + fullServerName);
            //Specify server stuff                                                                ex. "prjdc.project.com"
            DirectoryContext adamContext = new DirectoryContext(DirectoryContextType.DirectoryServer, fullServerName);
            //MessageBox.Show(adamContext.Name);

            //Get schema for that servers active directory
            ActiveDirectorySchema schema = ActiveDirectorySchema.GetSchema(adamContext);

            /*
            // Test for getting all classes
            foreach (ActiveDirectorySchemaClass schemaClass in schema.FindAllClasses())
            {
                MessageBox.Show(schemaClass.GetAllProperties().ToString());
            }
            */

            //Get the user schema class
            ActiveDirectorySchemaClass schemaClass = schema.FindClass("user");

            //Now that we have the correct class, GET ALL PROPERTIES in that class (these properties themselves are readonly because we don't want to alter them)
            ReadOnlyActiveDirectorySchemaPropertyCollection adPropertyCollection = schemaClass.GetAllProperties(); //There are 342





            //http://stackoverflow.com/questions/4931982/how-can-i-check-if-a-user-has-write-rights-in-active-directory-using-c

            //Find the current logged on user that will be modfying user properties
            string userName = Environment.UserName;
            //Get the users domain
            string userDomainName = Environment.UserDomainName;
            //MessageBox.Show("The domain name which the user is in: " + userDomainName);


            DirectoryEntry de = new DirectoryEntry();
            de.Path = "LDAP://" + userDomainName;
            de.AuthenticationType = AuthenticationTypes.Secure;

            DirectorySearcher deSearch = new DirectorySearcher();

            deSearch.SearchRoot = de;
            deSearch.Filter = "(&(objectClass=user) (cn=" + userName + "))";

            SearchResult result = deSearch.FindOne();

            DirectoryEntry deUser = new DirectoryEntry(result.Path);

            //Refresh Cache to get the values, I guess?
            deUser.RefreshCache(new string[] { "allowedAttributesEffective" });

            // Proof this is the user
            //MessageBox.Show("About to show the username: " + deUser.Properties["cn"].Value.ToString());
            //MessageBox.Show("About to show the users distinguished name: " + deUser.Properties["distinguishedName"].Value.ToString());

            //Now get the property["allowedAttributesEffective"] of the user

            //Test on a property that did not have to refresh cache for, just to ensure enumeration was working
            //string propertyName = deUser.Properties["otherHomePhone"].PropertyName;

            string propertyName = deUser.Properties["allowedAttributesEffective"].PropertyName;

            //MessageBox.Show("About to loop through the multi-valued property: " + deUser.Properties["allowedAttributesEffective"].PropertyName);
            //MessageBox.Show("About to loop through the multi-valued property: " + propertyName);
            //MessageBox.Show("Number of elements: " + deUser.Properties[propertyName].Count);




            alWritableProperties.Clear();
            //Go through all the elements of that multi-valued property
            IEnumerator propEnumerator = deUser.Properties[propertyName].GetEnumerator(); //Getting the Enumerator

            propEnumerator.Reset(); //Position at the Beginning


            //Loading Bar for writable properties
            int i = 0;
            ProgressWindow pwBar = new ProgressWindow("Finding Writable User Properties", deUser.Properties[propertyName].Count);
            pwBar.Show();
            pwBar.SetPercentDone(i);

            while (propEnumerator.MoveNext()) //while there is a next one to move to
            {
                //MessageBox.Show("" + propEnumerator.Current.ToString());
                try
                {
                    ActiveDirectorySchemaProperty propertyToTest = new ActiveDirectorySchemaProperty(adamContext, propEnumerator.Current.ToString());
                    //See if property is writable
                    //is property single valued
                    if (adPropertyCollection[adPropertyCollection.IndexOf(propertyToTest)].IsSingleValued == true)
                    {
                        //Single valued comparison
                        deUser.Properties[propEnumerator.Current.ToString()].Value = deUser.Properties[propEnumerator.Current.ToString()].Value;
                    }
                    else
                    {
                        //Multi-valued comparison (Not implemented)

                        //http://stackoverflow.com/questions/5067363/active-directory-unable-to-add-multiple-email-addresses-in-multi-valued-propert

                        //MessageBox.Show("Dealing with multivalued property: " + propEnumerator.Current.ToString());

                        //deUser.Properties[propEnumerator.Current.ToString()].Clear();
                        //deUser.Properties[propEnumerator.Current.ToString()].Count;
                        //MessageBox.Show("Number of elements: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                        //deUser.Properties[propEnumerator.Current.ToString()].Value = deUser.Properties[propEnumerator.Current.ToString()].Value;
                        /*
                        if (propEnumerator.Current.ToString().Equals("departmentNumber"))
                        {
                            MessageBox.Show("Number of elements before: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                            deUser.Properties[propEnumerator.Current.ToString()].Add("9");
                            MessageBox.Show("Number of elements after: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                            //MessageBox.Show("Number of elements before: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                            //object tempObject = deUser.Properties[propEnumerator.Current.ToString()].Value;
                            //deUser.Properties[propEnumerator.Current.ToString()].Clear();
                            //MessageBox.Show("Number of elements now cleared: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                            //deUser.Properties[propEnumerator.Current.ToString()].Value = tempObject;
                            //MessageBox.Show("Number of elements after: " + deUser.Properties[propEnumerator.Current.ToString()].Count.ToString());
                        }
                        */
                    }

                    deUser.CommitChanges();

                    //Add to array since it is writable
                    alWritableProperties.Add(propEnumerator.Current.ToString());
                }
                catch (Exception e)
                {
                    //MessageBox.Show(propEnumerator.Current.ToString() + " can only be viewed" + e.Message);
                }
                pwBar.SetPercentDone(i++);
            }
            deUser.Close();
            //MessageBox.Show("Number of actual writable properties: " + alWritableProperties.Count);

            //End loading bar
            pwBar.Close();


            // OLD STUFF BELOW

            // http://msdn.microsoft.com/en-us/library/ms180940(v=vs.80).aspx 
            //ActiveDirectorySchemaClass
            // http://stackoverflow.com/questions/3290730/how-can-i-read-the-active-directory-schema-programmatically 

            // http://msdn.microsoft.com/en-us/library/bb267453.aspx#sdsadintro_topic3_manageschema 

            /*
            //Get current server
            string ADServer = Environment.GetEnvironmentVariable("LOGONSERVER");
            //Remove the "//"
            ADServer = ADServer.Substring(2);
            //Get the selected domain
            string domainName = cboDC.Text.ToString();
            //Concatenate to make full server name
            string fullServerName = ADServer + "." + domainName;
            MessageBox.Show("The directory context is: " + fullServerName);
            //Specify server stuff                                                                ex. "prjdc.project.com"
            DirectoryContext adamContext = new DirectoryContext(DirectoryContextType.DirectoryServer, fullServerName);
            MessageBox.Show(adamContext.Name);
            
            //Get schema for that servers active directory
            ActiveDirectorySchema schema = ActiveDirectorySchema.GetSchema(adamContext);
            */
            /*
            // Test for getting all classes
            foreach (ActiveDirectorySchemaClass schemaClass in schema.FindAllClasses())
            {
                MessageBox.Show(schemaClass.GetAllProperties().ToString());
            }
            */
            /*
            //Get the user schema class
            ActiveDirectorySchemaClass schemaClass = schema.FindClass("user");
            
            //Now that we have the correct class, GET ALL PROPERTIES in that class (these properties themselves are readonly because we don't want to alter them)
            ReadOnlyActiveDirectorySchemaPropertyCollection adPropertyCollection = schemaClass.GetAllProperties(); //There are 342
            //ActiveDirectorySchemaPropertyCollection adPropertyCollection = schemaClass.OptionalProperties; //There are 335
            //ActiveDirectorySchemaPropertyCollection adPropertyCollection = schemaClass.MandatoryProperties;  //There are 7
            */




            /*
            foreach (ActiveDirectorySchemaProperty schemaProperty in adPropertyCollection)
            {

                // http://msdn.microsoft.com/en-us/library/system.reflection.propertyattributes.aspx
                // Final test with "systemOnly" attribute, if this doesn't work, it won't ever...

                // Get the PropertyAttributes enumeration of the property.
                // Get the type.
                TypeAttributes schemaPropertyAttributes = schemaProperty.GetType().Attributes;
                // Get the property attributes.
                //PropertyInfo schemaPropertyInfo = schemaPropertyType.GetProperty("systemOnly");
                //PropertyAttributes propAttributes = schemaPropertyInfo.Attributes;
                //Display the property attributes value.
                MessageBox.Show("Property: " + schemaProperty.CommonName.ToString() + " has attributes: " + s;
                
            }
            */






            //Have a fake read-only property
            //ActiveDirectorySchemaProperty fakeProperty = new ActiveDirectorySchemaProperty(adamContext, "cn");
            //AttributeCollection fakePropertyAttributes = TypeDescriptor.GetAttributes(fakeProperty);
            //DirectoryServicesPermissionAttribute a = new DirectoryServicesPermissionAttribute(

            /*
            MessageBox.Show("Does fake property contain read-write attribute: " + fakePropertyAttributes.Contains(ReadOnlyAttribute.No).ToString());

            if (a)
            {   
                MessageBox.Show("READ ONLY");
            }
            else
            {
                //Can't be, when try to go fakeProperty.Name = "MEGADEATH" it says can't because prop is read-only
                MessageBox.Show("READ AND WRITE");
            }
            */

            //http://technet.microsoft.com/en-us/library/cc773309(WS.10).aspx

            // CURRENT PROBLEM:
            // Can get all properties, but cannot seem to separate writable from read-only
            // have heard of attributeSchema but no luck
            // now thinking of using systemOnly flag, but no idea how to check for that http://msdn.microsoft.com/en-us/library/aa772300.aspx
            /*
            // Test value of flags using bitwise AND.
            bool test = (meetingDays & Days2.Thursday) == Days2.Thursday; // true
            Console.WriteLine("Thursday {0} a meeting day.", test == true ? "is" : "is not");
            // Output: Thursday is a meeting day.
             * 
             */


            /*
            Type type = typeof(ActiveDirectorySchemaProperty);
            object[] ac = type.GetCustomAttributes(true);
            




            //Test for what's in collection
            MessageBox.Show("Now Testing what is in ReadOnlyActiveDirectorySchemaPropertyCollection");
            int actualNumber = 0;
            foreach (ActiveDirectorySchemaProperty adProperty in adPropertyCollection)
            {
                actualNumber++;

                //MessageBox.Show("Property: " + adProperty.Name + " // Common Name: " + adProperty.CommonName);
                // http://msdn.microsoft.com/en-us/library/system.componentmodel.attributecollection.aspx //
                //Get attributes of that property (ex. is a string, is read only, is writable, etc)
                AttributeCollection attributes = TypeDescriptor.GetAttributes(adProperty);
                
                //List of systemOnly of the property
                MessageBox.Show("Now showing attributes in current property");

                //attributes.Contains(Attribute.GetCustomAttribute(AssemblyFlagsAttribute"systemOnly",typeof(FlagsAttribute)));
                AssemblyName a = new AssemblyName();
                //https://connid.googlecode.com/svn-history/r169/bundles/ad/trunk/src/main/java/org/connid/ad/schema/ADSchemaBuilder.java
                
                //AssemblyNameFlags aName = new AssemblyNameFlags();
                //AssemblyFlagsAttribute afa = new AssemblyFlagsAttribute(aName);
                                
                //See if the attributes collection isn't writable
                //if (attributes.Contains(ReadOnlyAttribute.No) == false)
                //if(attributes.Contains(Attribute.GetCustomAttribute(typeof(ActiveDirectorySchemaProperty),"systemOnly")))



                // More freaking testing
                // http://stackoverflow.com/questions/2051065/check-if-property-has-attribute //
                //http://msdn.microsoft.com/en-us/library/cc220922(v=prot.10).aspx
                
                //Go through all attributes and see if systemOnly is false, if it is then add the property to the array
                
                //Go through all attributes and see if systemOnly is false, if it is then add the property to the array
                foreach (Attribute currentattribute in attributes)
                {
                    MessageBox.Show(currentattribute.TypeId.ToString());
                }

                
                /*
                if ()
                {
                    //Cannot read and write
                }
                else
                {
                    //Our property is read/write!
                    //Add the name of the property to our alWritableProperties array list
                    alWritableProperties.Add(adProperty.Name.ToString());
                }
                */

            /*
            }
            MessageBox.Show("Now Seeing what has been added to the writable properties list");
            MessageBox.Show("Number of Properties: " + actualNumber.ToString() + "\nNumber of Writable Properties: " + alWritableProperties.Count);
            */


            /*
            #region Properties of the schema
            /*
            // This will get the properties of the schema class
            PropertyInfo[] aPropertyInfo = schemaClass.GetType().GetProperties();

            //For each property
            foreach (PropertyInfo property in aPropertyInfo)
            {

                MessageBox.Show("Property: " + property.Name);
                /*
                if (property.PropertyType.Assembly == schemaClass.GetType().Assembly)
                {
                    //Show just property
                    MessageBox.Show("Property: " + property.Name);
                }
                else
                {
                    MessageBox.Show("Property: " + property.Name + " Value: " + propValue);
                }
                */
            /*
            }
            */
            /*
            #endregion
            */
            /* http://msdn.microsoft.com/en-us/library/windows/desktop/ms677167(v=vs.85).aspx */
            //IDEA: We get do foreach madatory, and a foreach optional property, put all into huge property array
            //      Then for each property, we do the whole type thing
            //http://stackoverflow.com/questions/6196413/how-to-recursively-print-the-values-of-an-objects-properties-using-reflection

            /*
            //foreach (ActiveDirectorySchemaProperty schemaProperty in schemaClass.MandatoryProperties)
            //ActiveDirectorySchemaPropertyCollection
            //ActiveDirectorySchemaPropertyCollection[] properties = schemaClass.GetType().GetProperties();
            foreach(ActiveDirectorySchemaProperty schemaProperty in schemaClass.MandatoryProperties)
            {
                PropertyInfo[] properties = schemaProperty.GetType().GetProperties();

                //findAttrValue
                
                
                //See what we have in the properties MessageBox.Show(properties.GetEnumerator().ToString());
                
                
                /*
                http://msdn.microsoft.com/en-us/library/system.componentmodel.defaultvalueattribute(v=vs.71).aspx
                 */
            //[C#] 
            // Gets the attributes for the property.
            //AttributeCollection attributes = TypeDescriptor.GetProperties(this)["MyProperty"].Attributes;
            //AttributeCollection attributes = TypeDescriptor.GetProperties(this)[schemaProperty.Name].Attributes;
            //AttributeCollection attributes = [schemaProperty.Name].Attributes;

            //Prints the default value by retrieving the DefaultValueAttribute
            //from the AttributeCollection. 
            //DefaultValueAttribute myAttribute = (DefaultValueAttribute)attributes[typeof(DefaultValueAttribute)];
            //MessageBox.Show("The default value is: " + myAttribute.Value.ToString());
            /*
            // Checks to see whether the value of the ReadOnlyAttribute is Yes.
            if (attributes[typeof(ReadOnlyAttribute)].Equals(ReadOnlyAttribute.Yes))
            {
                // Insert code here.
                MessageBox.Show("The Property " + schemaProperty.Name + " is read-only");
            }
            */


            //AttributeCollection attributes = TypeDescriptor.GetProperties(schemaProperty)[schemaProperty.Name].Attributes;
            //Attribute a = attributes[schemaProperty].


            /*
                foreach (PropertyInfo property in properties)
                {
                    MessageBox.Show("Property: " + property.Name);
                }

            }
            */



            /*
            
            //Find all mandatory properties for the schemaClass
            foreach (ActiveDirectorySchemaProperty schemaProperty in schemaClass.MandatoryProperties)
            {
                MessageBox.Show("Property: " + schemaProperty.ToString());
                MessageBox.Show("Name(what we write to): " + schemaProperty.Name + ", Common Name(Display Name to show on datagridview): " + schemaProperty.CommonName);
                
                
                //Determine if it is a writable property



                //To get the CanWrite property, first get the class Type.
                //Type propertyType = Type.GetType(schemaProperty.Name);
                Type propertyType = schemaProperty.GetType();
                //Type propertyType = schemaProperty.Name.GetType();
                //Type propertyType = Type.GetType(schemaProperty.Name);
                //From the Type, get the PropertyInfo. From the PropertyInfo, get the CanWrite value.

                MessageBox.Show("Made it past Type: " + propertyType.ToString());


                PropertyInfo[] properties = propertyType.GetProperties();

                foreach (PropertyInfo property in properties)
                {
                    object propValue = property.GetValue(schemaProperty, null);
                    if (property.PropertyType.Assembly == propertyType.Assembly)
                    {
                        MessageBox.Show("Property: " + property.Name);
                    }
                    else
                    {
                        MessageBox.Show("Property: " + property.Name + " Value: " + propValue);
                    }
                }



                /*
                PropertyInfo propInfo = propertyType.GetProperty(schemaProperty.ToString());

                PropertyAttributes pAttributes = propInfo.Attributes;

                MessageBox.Show("Attributes: " + pAttributes.ToString());
                */
            /*
            //MessageBox.Show("Made it past Info! " + propInfo.CanWrite);

            if (propInfo.CanWrite == true)
            {
                MessageBox.Show("We CAN write to this mofo!");
            }
            else
            {
                MessageBox.Show("We CANNOT write to this mofo!");
            }
            */

            //MessageBox.Show("Can we write to this mofo?  " + propInfo.CanWrite.ToString());


            //Old
            //PropertyInfo[] propInfo = propertyType.GetProperties(BindingFlags.Public | BindingFlags.Instance);



            /* http://msdn.microsoft.com/en-us/library/system.reflection.propertyinfo.canwrite.aspx */

            //using reflection
            /* http://www.codersource.net/microsoft-net/c-basics-tutorials/reflection-in-c.aspx */
            /*
            for (int i = 0; i < propInfo.Length ;i++)
            {
                MessageBox.Show(propInfo[0].ToString());
            }
            */
            /*
            if(propertyType)
            {
                //Since this is a writable property, add it to the array
                    
            }
            */
            /*
            }
            */

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

                //Clear out list box and ensure the combo box's value is now "---Choose One---"
                lstUsers.Items.Clear();
                this.cboUserOU.SelectedIndex = 0;

                sender = (RadioButton)sender;

                string ouPath = "";
                //depending on which button is selected set ouPath
                if (sender.Equals(radStuOU))
                {
                    ouPath = "ou=STU,ou=USR,";
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
            else if (cboDC.SelectedIndex < 1 && ((RadioButton)sender).Checked)
            {
                //Clear cboUser box text
                this.cboUserOU.Items.Clear();
                this.cboUserOU.Text = "";

                //inform user they haven't selected a Domain controller yet
                MessageBox.Show("Please select a domain controller first!");
            }
            populateDataGridView();
        }

        /// <summary>
        /// Method called when the user OU combo box selected index is changed.
        /// Handles filling out the listbox with all users under selected OU
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cboUserOU_SelectedIndexChanged(object sender, EventArgs e)
        {
            lstUsers.Items.Clear();
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
        /// Method called when the User list box has a user selected or unselected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void lstUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Clear modify users array list since user list will re-populate it
            alModifyUsers.Clear();
            //Index of selected user in the user list box to get
            int index = 0;
            //For each selected user, add to modify user array list (alModifyUsers)
            foreach (string cn in lstUsers.SelectedItems)
            {
                index = lstUsers.Items.IndexOf(cn);
                string finish = aUserLDAPPaths[index].ToString();

                DirectoryEntry de = new DirectoryEntry(finish);
                alModifyUsers.Add(de);
            }

            //Just making sure we get the exact users we want to modify
            //MessageBox.Show("Users Selected: " + alModifyUsers.Count.ToString());
            /** CURRENTLY THIS METHOD RUNS "NUMBER OF USERS SELECTED" TIMES WHEN CHECKBOX IS SELECTED **/

            // Populate dgvUserProperties with list of writable user properties
            populateDataGridView();

        }

        /// <summary>
        /// Populates dgvUserProperties with list of writable user properties
        /// 
        ///     Property | Value
        ///     Name     | CST
        ///     Country  | Canada
        /// </summary>
        private void populateDataGridView()
        {
            /* More info on stylizing columns
             * http://msdn.microsoft.com/en-us/library/system.windows.forms.datagridview.defaultcellstyle.aspx
             */

            dgvUserProperties.Columns.Clear();
            dict = new Dictionary<string, string>();
            //Populate datagridview if we have 1 or more users selected from the list box
            if (lstUsers.SelectedItems.Count >= 1)
            {

                //Add Column "Property"
                dgvUserProperties.Columns.Add("Property", "Property");
                dgvUserProperties.Columns["Property"].ReadOnly = true;
                dgvUserProperties.Columns["Property"].DefaultCellStyle.BackColor = Color.LightGray;
                //Add Column "Value"
                dgvUserProperties.Columns.Add("Value", "Value");

                //Clear rows just in case someone started editing something now that we have rules in place
                dgvUserProperties.Rows.Clear();

                /* Testing datagridview */
                /*
                
                //Get [i] user out of list
                DirectoryEntry de = (DirectoryEntry)alModifyUsers[0];

                //add values
                dgvUserProperties.Rows.Add("Name", de.Properties["name"].Value);
                dgvUserProperties.Rows.Add("First Name", de.Properties["givenName"].Value);
                dgvUserProperties.Rows.Add("Last Name", de.Properties["sn"].Value);
                
                */

                //TODO: User comparison

                alCurrentPropertyValues.Clear();
                //Loop through each user in the array list
                foreach (DirectoryEntry currentUser in alModifyUsers)
                {
                    //MessageBox.Show(currentUser.Properties["cn"].Value.ToString());
                    foreach (PropertyValueCollection prop in currentUser.Properties)
                    {

                        //dict.Add(propertyName, prop.Value.ToString());
                        dgvUserProperties.Rows.Add(prop.PropertyName, prop.Value);

                    }
                    //Loop through the properties of the currentUser based on the alWritableProperties(strings)
                    foreach (string propertyName in alWritableProperties)
                    {
                        //NOTE: if we're comparing to a custom passed in list, we need error checking here 
                        //Add the property to alCurrentUserProperties

                        //dict.Add(propertyName, currentUser.Properties[propertyName].Value.ToString());
                        //dgvUserProperties.Rows.Add(propertyName, currentUser.Properties[propertyName].Value.ToString());
                    }

                    //Now compare the alCurrentUserPropertyValues to alCurrentPropertyValues

                    //if alCurrentPropertyValues already has the property, then compare values
                    //if the values are different, change value in alCurrentPropertyValues to "[Multiple Values]"
                    //else
                    //add the property from alCurrentUserPropertyValues to alCurrentPropertyValues


                }

                /* End of testing datagridview */

                dgvUserProperties.Focus();
                //dgvUserProperties.CurrentCell = dgvUserProperties[dgvUserProperties.Columns["Value"].Index, dgvUserProperties.Rows.GetFirstRow(DataGridViewElementStates.None)];
            }
            else
            {
                dgvUserProperties.DataSource = null;
            }
        }

        /// <summary>
        /// Overwrites the double clicking of a cell in a row to automatically start editing the "value" cell in that row
        /// </summary>
        private void dgvUserProperties_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dgvUserProperties.CurrentCell = dgvUserProperties[dgvUserProperties.Columns["Value"].Index, dgvUserProperties.CurrentCell.RowIndex];
            dgvUserProperties.BeginEdit(true);
        }
        /// <summary>
        /// Sets the current cell in a row to the "value" so if a person goes to a different row and starts typing, it'll start editing the value
        /// </summary>
        /// 
        private void dgvUserProperties_SelectionChanged(object sender, EventArgs e)
        {
            dgvUserProperties.CurrentCell = dgvUserProperties[dgvUserProperties.Columns["Value"].Index, dgvUserProperties.CurrentCell.RowIndex];
        }

        /// <summary>
        /// Method called when the Select All checkbox is checked or unchecked.
        /// Handles selecting or deselecting all items from the list of users
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkSelectAll_CheckedChanged(object sender, EventArgs e)
        {

            //Temporarly remove event of comparing users and creating list of values unless we are down to the last one
            this.lstUsers.SelectedIndexChanged -= new System.EventHandler(this.lstUsers_SelectedIndexChanged);

            if (this.chkSelectAll.Checked)
            {
                for (int i = 0, nLen = lstUsers.Items.Count; i < nLen; i++)
                {
                    //if last one to select, enable the event
                    if (i == (nLen - 1))
                    {
                        this.lstUsers.SelectedIndexChanged += new System.EventHandler(this.lstUsers_SelectedIndexChanged);
                    }

                    lstUsers.SetSelected(i, true);
                }
            }
            else
            {
                for (int i = 0, nLen = lstUsers.Items.Count; i < nLen; i++)
                {
                    lstUsers.SetSelected(i, false);
                }
                //Re-add event
                this.lstUsers.SelectedIndexChanged += new System.EventHandler(this.lstUsers_SelectedIndexChanged);
                //Clear datagridview of values

            }
        }

        #endregion

        /// <summary>
        /// Resets the password for each of the selected users
        /// The user will be required to change the password at the next logon process
        /// </summary>
        private void btnResetPassword_Click(object sender, EventArgs e)
        {

        }


    }
}
