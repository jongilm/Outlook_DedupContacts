using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

using OutLook = Microsoft.Office.Interop.Outlook;

namespace OutlookContactsAnalyser
{
    public partial class frmNewContact : Form
    {
        Microsoft.Office.Interop.Outlook.MAPIFolder m_CustomFolder = null;

        public frmNewContact()
        {
            InitializeComponent();
        }

        /// <summary>
        /// SAVING THE NEW CONTACT TO OUTLOOK.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (rdoCustom.Checked)
            {
                // IF THE CUSTOM FOLDER OPTION IS SELECTED PROMPTING FOR THE NAME.
                if (string.IsNullOrEmpty(txFolder.Text.Trim().ToString()))
                {
                    MessageBox.Show("Please Give Your CustomFolder Name");
                    return;
                }
                // IF WE NEED A CUSTOM FOLDER FOR CREATING CONTACTS .HERE IT CHECKS FOR FOLDER OR OTHER 
                // WISE IT CREATES A NEW CUSTOMER FOLDER WITH GIVEN NAME .

                if (!CheckCustomFolderExists(txFolder.Text.Trim().ToString()))
                {
                    CreateCustomFolder(txFolder.Text.Trim().ToString());
                }
                if (chkUpdateExistingContact.Checked)
                {
                    //orig OutLook._Application outlookApplication = new OutLook.Application();
                    /*MyContact contact = new MyContact();*/
                    //orig cntact.CustomProperty = txtProp1.Text.Trim().ToString();

                    // CREATING CONTACT ITEM OBJECT AND FINDING THE CONTACT ITEM 
                    String KeyPetName = txtProp1.Text.Trim().ToString();
                    OutLook.ContactItem newContact = (OutLook.ContactItem)FindContactItemWithPetName(KeyPetName, /*contact, */m_CustomFolder);

                    // THE VALUES WE CAN GET FROM WEB SERVICES OR DATA BASE OR CLASS. WE HAVE TO ASSIGN THE VALUES 
                    // TO OUTLOOK CONTACT ITEM OBJECT .
                    if (newContact != null)
                    {
                        newContact.FirstName                = txtFirstName.Text.Trim().ToString();
                        newContact.LastName                 = txtLastName.Text.Trim().ToString();
                        newContact.Email1Address            = txtEmail.Text.Trim().ToString();
                        newContact.Business2TelephoneNumber = txtPhone.Text.Trim().ToString();
                        newContact.BusinessAddress          = txtAddress.Text.Trim().ToString();
                        if (chkAddCustomProperty.Checked)
                        {
                            // HERE WE CAN CREATE OUR OWN CUSTOM PROPERTY TO IDENTIFY OUR APPLICATION. 
                            if(string.IsNullOrEmpty(txtProp1.Text.Trim().ToString()))
                            {
                                MessageBox.Show("Please add value to your Custom Property");
                                return;
                            }
                            newContact.UserProperties.Add("myPetName", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
                            newContact.UserProperties["myPetName"].Value = txtProp1.Text.Trim().ToString();
                        }
                        newContact.Save();
                        this.Close();
                    }
                    else // FindContactItem was not successful
                    {
                        // IF THE CONTACT DOES NOT EXIST WITH SAME CUSTOM PROPERTY CREATES THE CONTACT.
                        newContact = (OutLook.ContactItem)m_CustomFolder.Items.Add(OutLook.OlItemType.olContactItem);
                        newContact.FirstName                = txtFirstName.Text.Trim().ToString();
                        newContact.LastName                 = txtLastName.Text.Trim().ToString();
                        newContact.Email1Address            = txtEmail.Text.Trim().ToString();
                        newContact.Business2TelephoneNumber = txtPhone.Text.Trim().ToString();
                        newContact.BusinessAddress          = txtAddress.Text.Trim().ToString();
                        if (chkAddCustomProperty.Checked)
                        {
                            // HERE WE CAN CREATE OUR OWN CUSTOM PROPERTY TO IDENTIFY OUR APPLICATION. 
                            if (string.IsNullOrEmpty(txtProp1.Text.Trim().ToString()))
                            {
                                MessageBox.Show("Please add value to your Custom Property");
                                return;
                            }
                            newContact.UserProperties.Add("myPetName", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
                            newContact.UserProperties["myPetName"].Value = txtProp1.Text.Trim().ToString();
                        }
                        newContact.Save();
                        this.Close();
                    }
                }
                else // chkUpdateExistingContact not Checked
                {
                    //orig OutLook._Application outlookApplication = new OutLook.Application();
                    OutLook.ContactItem newContact = (OutLook.ContactItem)m_CustomFolder.Items.Add(OutLook.OlItemType.olContactItem);
                    newContact.FirstName                = txtFirstName.Text.Trim().ToString();
                    newContact.LastName                 = txtLastName.Text.Trim().ToString();
                    newContact.Email1Address            = txtEmail.Text.Trim().ToString();
                    newContact.Business2TelephoneNumber = txtPhone.Text.Trim().ToString();
                    newContact.BusinessAddress          = txtAddress.Text.Trim().ToString();
                    if (chkAddCustomProperty.Checked)
                    {
                        // HERE WE CAN CREATE OUR OWN CUSTOM PROPERTY TO IDENTIFY OUR APPLICATION. 
                        if (string.IsNullOrEmpty(txtProp1.Text.Trim().ToString()))
                        {
                            MessageBox.Show("Please add value to your Custom Property");
                            return;
                        }
                        newContact.UserProperties.Add("myPetName", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
                        newContact.UserProperties["myPetName"].Value = txtProp1.Text.Trim().ToString();
                    }
                    newContact.Save();
                    this.Close();
                }
            }
            else // rdoCustom is not checked - Save to default folder
            {
                // CREATES THE OUTLOOK CONTACT IN DEFAULT CONTACTS FOLDER.
                OutLook._Application outlookApplication = new OutLook.Application();
                OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
                OutLook.ContactItem newContact = (OutLook.ContactItem)contactsFolder.Items.Add(OutLook.OlItemType.olContactItem);
                // THE VALUES WE CAN GET FROM WEB SERVICES OR DATA BASE OR CLASS.
                // WE HAVE TO ASSIGN THE VALUES TO OUTLOOK CONTACT ITEM OBJECT .
                newContact.FirstName                = txtFirstName.Text.Trim().ToString();
                newContact.LastName                 = txtLastName.Text.Trim().ToString();
                newContact.Email1Address            = txtEmail.Text.Trim().ToString();
                newContact.Business2TelephoneNumber = txtPhone.Text.Trim().ToString();
                newContact.BusinessAddress          = txtAddress.Text.Trim().ToString();
                newContact.Save();
                this.Close();
            }

        }

        /// <summary>
        /// ENABLING AND DISABLING THE CUSTOM FOLDER AND PROPERY OPTIONS.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void rdoCustom_CheckedChanged(object sender, EventArgs e)
        {
            if (rdoCustom.Checked)
            {
                txFolder.Enabled = true;
                chkAddCustomProperty.Enabled = true;
                chkUpdateExistingContact.Enabled = true;
                txtProp1.Enabled = true;
            }
            else
            {
                txFolder.Enabled = false;
                chkAddCustomProperty.Enabled = false;
                chkUpdateExistingContact.Enabled = false;
                txtProp1.Enabled = false;
            }
        }

        /// <summary>
        /// FIND THE CONTACT WITH SAME USER PROPERTY VALUE.
        /// </summary>
        /// <param name="contact"></param>
        /// <param name="folder"></param>
        /// <returns></returns>
        private OutLook._ContactItem FindContactItemWithPetName(String KeyPetName, /*MyContact contact, */OutLook.MAPIFolder folder)
        {
            object missing = System.Reflection.Missing.Value;
            foreach (OutLook._ContactItem OutlookContact in folder.Items)
            {
              OutLook.UserProperty userProperty = OutlookContact.UserProperties.Find("myPetName", missing);
                if (userProperty != null)
                {
                    if (userProperty.Value.Equals(KeyPetName))
                      return OutlookContact;
                }
            }
            return null;
        }
       
        /// <summary>
        /// Check myconacts(custom folder) exists or not
        /// </summary>
        /// <returns></returns>
        public bool CheckCustomFolderExists(String argFolder)
        {
            OutLook._Application outlookApplication = new OutLook.Application();
            OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
            // VERIFYING THE MYCONTACTS (CUSTOM) SUB FOLDER IN CONTACTS FOLDER IN OUTLOOK.
            foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
            {
                if (subFolder.Name == argFolder)
                {
                    m_CustomFolder = subFolder;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// CREATES THE CUSTOM FOLDER IN OUTLOOK CONTACTS FOLDER
        /// </summary>
        public void CreateCustomFolder(String argFolder)
        {
            OutLook._Application outlookApplication = new OutLook.Application();
            OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
            // VERIFYING THE MYCONTACTS(CUSTOM) FOLDER IN OUTLOOK .
            m_CustomFolder = null;
            foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
            {
                if (subFolder.Name == argFolder) 
                {
                    m_CustomFolder = subFolder;
                    break;
                }
            }
            // IF MYCONTACTS FOLDER DOES NOT EXIST CREATE A NEW FOLDER WITH NAME MYCONTACTS.
            if (m_CustomFolder == null)
            {
                m_CustomFolder = contactsFolder.Folders.Add(argFolder, Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
            }
        }

        /// <summary>
        /// ENABLING AND DISABLING THE CUSTOM PROPERTY CONTROLS.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void chkAdd_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAddCustomProperty.Checked)
            {
                txtProp1.Enabled = true;
                chkUpdateExistingContact.Enabled = true;
            }
            else
            {
                txtProp1.Enabled = false;
                chkUpdateExistingContact.Enabled = false;
            }
        }
    } // class frmNewContact : Form
} // namespace OutlookContactsAnalyser
