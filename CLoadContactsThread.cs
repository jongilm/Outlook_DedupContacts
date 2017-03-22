using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Outlook;

using OutLook1 = Microsoft.Office.Interop.Outlook;
using Contact1 = Microsoft.Office.Interop.Outlook._ContactItem;

namespace OutlookContactsAnalyser
{

  public class CLoadContactsThread
  {
    #region MemberVariables
    ManualResetEvent m_EventStop;    // Main thread sets this event to stop worker thread
    ManualResetEvent m_EventStopped; // Worker thread sets this event when it is stopped
    MainForm m_form;                 // Reference to main form used to make syncronous user interface calls
    String m_folderName;
    int m_MaxContactsToRead;
    //public int m_ContactsInOutlook = 0;
    //int m_NumberOfContactsRead = 0;
    //List<MyContact> m_ContactList = null;
    #endregion // MemberVariables

    #region Functions

    //////////////////////////////////////
    // Constructor
    //////////////////////////////////////
    public CLoadContactsThread(ManualResetEvent eventStop,
                               ManualResetEvent eventStopped,
                               MainForm form,
                               string folderName,
                               int iMaxContactstoRead)
    {
      m_EventStop = eventStop;
      m_EventStopped = eventStopped;
      m_form = form;
      m_folderName = folderName;
      m_MaxContactsToRead = iMaxContactstoRead;
    }

    //////////////////////////////////////
    // Function runs in worker thread and emulates long process.
    //////////////////////////////////////
    public void Run()
    {
      String s;

      object missing = System.Reflection.Missing.Value;
      int ii;

      // Create an instance of Outlook Application and Outlook Contacts folder.
      //try
      {
        //Monitor.Enter(this);
        
        s = "Opening Folder \"" + m_folderName + "\"";
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AppendToLog, new Object[] {s});

        OutLook1.MAPIFolder   fldContacts = null;
        OutLook1._Application outlookObj  = new OutLook1.Application();

        //m_ContactList = new List<MyContact>();

        if (m_folderName == "Default")
        {
          fldContacts = (OutLook1.MAPIFolder)outlookObj.Session.GetDefaultFolder(OutLook1.OlDefaultFolders.olFolderContacts);
        }
        else
        {
          OutLook1.MAPIFolder contactsFolder = (OutLook1.MAPIFolder)
          outlookObj.Session.GetDefaultFolder(OutLook1.OlDefaultFolders.olFolderContacts);

          // Verify the custom folder in Outlook.
          foreach (OutLook1.MAPIFolder subFolder in contactsFolder.Folders)
          {
            if (subFolder.Name == m_folderName)
            {
              fldContacts = subFolder;
              break;
            }
          }
        }

        /////////////////////////////////
        // CountFields
        /////////////////////////////////
        //MyContact contact = new MyContact();
        MyFields flds = new MyFields(m_form);
        flds.ClearFieldCount();
    
        /////////////////////////////////
        // Load Contacts
        /////////////////////////////////
        s = "Load contacts from Folder \"" + m_folderName + "\"";
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AppendToLog, new Object[] {s});

        // Loop through contact in specified folder.
        ii = 0;
        foreach (Microsoft.Office.Interop.Outlook._ContactItem OutlookContact in fldContacts.Items)
        {
          ii++;
          if (ii > m_MaxContactsToRead)
            break;

          // Count used fields
          flds.AccumulateFields(OutlookContact);

          // Add contact to list
          MyContact contact = new MyContact();
          contact.Clear();
          contact.StoreContact(contact, ii, OutlookContact);

          // Make synchronous call to main form.
          // MainForm.AppendToLog function runs in main thread.
          // (To make asynchronous call use BeginInvoke)

          // Add line to log
          //s = m_folderName + "[" + m_MaxContactsToRead + "] " + ii.ToString() + " read";
          //m_form.Invoke(m_form.m_DelegateLoadContactsThreadAppendToLog, new Object[] {s});

          // Update counter on main form
          m_form.Invoke(m_form.m_Delegate_LoadContactsThread_UpdateCounter, new Object[] {ii});

          // Build Contact List
          m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AddOneContact, new Object[] {contact});

          // Check if thread is cancelled
          if ( m_EventStop.WaitOne(0, true) )
          {
            // Clean-up operations may be placed here
            // ...

            // Inform main thread that this thread stopped
            m_EventStopped.Set();
            return;
          }
        }
        // Build Field Occurance List
        flds.PrintFieldCount();

        s = "Done";
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AppendToLog, new Object[] {s});

        // Inform the user of the status
        //m_NumberOfContactsRead = m_ContactList.Count;
        //m_ContactsInOutlook = m_ContactList.Count;
        //txtContactsInOutlook.Text = m_ContactList.Count.ToString();

        // Make synchronous call to main form
        // to inform it that thread finished
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_Finished, null);

        //Monitor.Exit(this);
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}

    }

    #endregion // Functions
  }

} // namespace OutlookContactsAnalyser

