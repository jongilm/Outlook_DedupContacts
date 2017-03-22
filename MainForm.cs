
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Outlook;

using OutLook = Microsoft.Office.Interop.Outlook;

namespace OutlookContactsAnalyser
{
  // #########################################################################
  // # region ClassTypes
  // #########################################################################
  #region ClassTypes
  // Delegates used to call MainForm functions from worker thread
  public delegate void Delegate_LoadContactsThread_AppendToLog_t(String s);
  public delegate void Delegate_LoadContactsThread_UpdateCounter_t(int i);
  public delegate void Delegate_LoadContactsThread_AddOneContact_t(MyContact contact);
  public delegate void Delegate_LoadContactsThread_AppendToListbox_t(String s);
  public delegate void Delegate_LoadContactsThread_AppendToFieldList_t(MyField fld);
  public delegate void Delegate_LoadContactsThread_Finished_t();

  // Delegates used to call MainForm functions from worker thread
  public delegate void Delegate_CountFieldsThread_AppendToLog_t(String s);
  public delegate void Delegate_CountFieldsThread_UpdateCounter_t(int i);
  public delegate void Delegate_CountFieldsThread_Finished_t();
  #endregion // ClassTypes

  //////////////////////////////////////
  // class MainForm
  //////////////////////////////////////
  public partial class MainForm : Form
  {
    // #########################################################################
    // # region MemberVariables
    // #########################################################################
    #region MemberVariables

    //////////////////////////////////////
    // General Members
    //////////////////////////////////////
    // Counters
    int m_NumberOfDuplicateContacts             = 0;
    int m_NumberOfSimilarContacts               = 0;
    int m_NumberOfVaguelySimilarContacts        = 0;
    int m_NumberOfSameNameContacts              = 0;
    int m_NumberOfSameNameAndSuffixContacts     = 0;
    int m_NumberOfTruncatedNotes                = 0;
    int m_NumberOfDislikedContent               = 0;
    int m_NumberOfDuplicateAddresses            = 0;
    int m_NumberOfDuplicatePhoneNumbers         = 0;
    int m_NumberOfDuplicateEmailAddress         = 0;

    int m_SourceOfMailingAddressUnknown         = 0;
    int m_SourceOfMailingAddressKnown           = 0;
    int m_MailingAddressNotSet                  = 0;
    int m_NumberOfMismatchedAddressComponents   = 0;
    int m_NumberOfMinimalInfo                   = 0;

    int m_MaxNotesLength           = 0;
    int m_MaxNotesNdx              = 0;
    int m_NumberOfModifiedContacts = 0;

    string MyClipboard = "";

    // Lists
    List<MyContact> DummyDataSource              = null;
    public List<MyContact> m_ContactList         = null;
    public List<MyContact> m_ContactListOrig     = null;
    List<MyContact> m_ContactListDuplicates      = null;
    List<MyContact> m_ContactListSimilar         = null;
    List<MyContact> m_ContactListVaguelySimilar  = null;
    List<MyContact> m_ContactListSameName        = null;
    List<MyContact> m_ContactListSameNameAndSuffix = null;
    List<MyContact> m_ContactListTruncatedNotes  = null;
    List<MyContact> m_ContactListDislikedContent = null;
    List<MyContact> m_ContactListDuplicateAddresses           = null;
    List<MyContact> m_ContactListDuplicatePhoneNumbers        = null;
    List<MyContact> m_ContactListDuplicateEmailAddress        = null;
    List<MyContact> m_ContactListUnlinkedMailingAddresses     = null;
    List<MyContact> m_ContactListLinkedMailingAddresses       = null;
    List<MyContact> m_ContactListEmptyMailingAddresses        = null;
    List<MyContact> m_ContactListMismatchedAddressComponents  = null;
    //List<MyContact> m_ContactListMatchedAddressComponents     = null;
    List<MyContact> m_ContactListMinimalInfo                  = null;

    List<MyField> m_FieldList                    = null;
    List<MyField> m_ForbiddenFieldList           = null;

    //////////////////////////////////////
    // Stuff for LoadContacts Thread
    //////////////////////////////////////
    Thread m_LoadContactsThread;                       // Worker thread
    ManualResetEvent m_EventStopLoadContactsThread;    // Event used to stop worker thread
    ManualResetEvent m_EventLoadContactsThreadStopped; // Event used to stop worker thread

    string m_folderName;
    int m_iMaxContactsToRead;

    // Delegate instances used to call user interface functions from worker thread:
    public Delegate_LoadContactsThread_AppendToLog_t       m_Delegate_LoadContactsThread_AppendToLog;
    public Delegate_LoadContactsThread_UpdateCounter_t     m_Delegate_LoadContactsThread_UpdateCounter;
    public Delegate_LoadContactsThread_AddOneContact_t     m_Delegate_LoadContactsThread_AddOneContact;
    public Delegate_LoadContactsThread_AppendToListbox_t   m_Delegate_LoadContactsThread_AppendToListbox;
    public Delegate_LoadContactsThread_AppendToFieldList_t m_Delegate_LoadContactsThread_AppendToFieldList;
    public Delegate_LoadContactsThread_Finished_t          m_Delegate_LoadContactsThread_Finished;

    //////////////////////////////////////
    // Stuff for CountFields Thread
    //////////////////////////////////////
    Thread m_CountFieldsThread;                       // Worker thread
    ManualResetEvent m_EventStopCountFieldsThread;    // Event used to stop worker thread
    ManualResetEvent m_EventCountFieldsThreadStopped; // Event used to stop worker thread

    // Delegate instances used to call user interface functions from worker thread:
    public Delegate_CountFieldsThread_AppendToLog_t        m_Delegate_CountFieldsThread_AppendToLog;
    public Delegate_CountFieldsThread_UpdateCounter_t      m_Delegate_CountFieldsThread_UpdateCounter;
    public Delegate_CountFieldsThread_Finished_t           m_Delegate_CountFieldsThread_Finished;

    #endregion // MemberVariables

    // #########################################################################
    // # region MainFormFunctions
    // #########################################################################
    #region MainFormFunctions
    //////////////////////////////////////
    // Main Form constructor
    //////////////////////////////////////
    public MainForm()
    {
      InitializeComponent();

      LoadContactsThread_Initialize();

      m_folderName = "";
      m_iMaxContactsToRead = 0;

      CountFieldsThread_Initialize();

      // Load the folders in Outlook
      LoadContactFoldersCombo();
    }

    //////////////////////////////////////
    // When the form is loaded.
    //////////////////////////////////////
    private void MainForm_Load(object sender, EventArgs e)
    {
      DummyDataSource               = new List<MyContact>();
      m_ContactList                 = new List<MyContact>();
      m_ContactListOrig             = new List<MyContact>();
      m_ContactListDuplicates       = new List<MyContact>();
      m_ContactListSimilar          = new List<MyContact>();
      m_ContactListVaguelySimilar   = new List<MyContact>();
      m_ContactListSameName         = new List<MyContact>();
      m_ContactListSameNameAndSuffix = new List<MyContact>();
      m_ContactListTruncatedNotes   = new List<MyContact>();
      m_ContactListDislikedContent  = new List<MyContact>();
      m_ContactListDuplicateAddresses = new List<MyContact>();
      m_ContactListDuplicatePhoneNumbers = new List<MyContact>();
      m_ContactListDuplicateEmailAddress = new List<MyContact>();
      
      m_ContactListUnlinkedMailingAddresses      = new List<MyContact>();
      m_ContactListLinkedMailingAddresses        = new List<MyContact>();
      m_ContactListEmptyMailingAddresses         = new List<MyContact>();

      m_ContactListMismatchedAddressComponents   = new List<MyContact>();
      //m_ContactListMatchedAddressComponents      = new List<MyContact>();
      m_ContactListMinimalInfo                   = new List<MyContact>();

      m_FieldList                   = new List<MyField>();
      m_ForbiddenFieldList          = new List<MyField>();

      InitialiseForbiddenFields();

      txtLog.AppendText("MainForm Load\r\n");

      this.Cursor = Cursors.WaitCursor;
      dgvContacts.DataSource = DummyDataSource;
      this.Cursor = Cursors.Arrow;
    }

    //////////////////////////////////////
    // Resize main form
    //////////////////////////////////////
    private void MainForm_Resize(object sender, EventArgs e)
    {
      //tlpMain.Width = this.Width - 20; // MainForm.Width - 10;
      //dgvContacts.Width = tlpMain.Width - 10;
      //txtLog.Width = tlpMain.Width - 10;

      //txtLog.AppendText("MainForm Resize\r\n");
    }

    #endregion // MainFormFunctions

    // #########################################################################
    // # region GetContactsThreadFunctions
    // #########################################################################
    #region GetContactsThreadFunctions

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void LoadContactsThread_Initialize()
    {
      // Initialize delegates
      m_Delegate_LoadContactsThread_AppendToLog       = new Delegate_LoadContactsThread_AppendToLog_t       (this.LoadContactsThread_AppendToLog);
      m_Delegate_LoadContactsThread_UpdateCounter     = new Delegate_LoadContactsThread_UpdateCounter_t     (this.LoadContactsThread_UpdateCounter);
      m_Delegate_LoadContactsThread_AddOneContact     = new Delegate_LoadContactsThread_AddOneContact_t     (this.LoadContactsThread_AddOneContact);
      m_Delegate_LoadContactsThread_AppendToListbox   = new Delegate_LoadContactsThread_AppendToListbox_t   (this.LoadContactsThread_AppendToListbox);
      m_Delegate_LoadContactsThread_AppendToFieldList = new Delegate_LoadContactsThread_AppendToFieldList_t (this.LoadContactsThread_AppendToFieldList);
      m_Delegate_LoadContactsThread_Finished          = new Delegate_LoadContactsThread_Finished_t          (this.LoadContactsThread_Finished);

      // Initialize events
      m_EventStopLoadContactsThread = new ManualResetEvent(false);
      m_EventLoadContactsThreadStopped = new ManualResetEvent(false);

      fileStopGetContactsThreadToolStripMenuItem.Enabled = false;
    }

    //////////////////////////////////////
    // Get the contacts in the specified the folder.
    //////////////////////////////////////
    private void LoadContacts(string folderName, int iMaxContactsToRead)
    {
      ThreadStart job;

      //Console.WriteLine("Command: Get Contacts from Folder - Start Thread");
      txtLog.AppendText("GetContacts - Start Thread\r\n");

      // Reset events
      m_EventStopLoadContactsThread.Reset();
      m_EventLoadContactsThreadStopped.Reset();

      m_folderName = folderName;
      m_iMaxContactsToRead = iMaxContactsToRead;

      txtLog.AppendText("GetContacts from \"" + folderName + "\"\r\n");

      // Create worker thread instance
      job = new ThreadStart(this.LoadContactsThread_Function);

      m_LoadContactsThread = new Thread(job);
      m_LoadContactsThread.Name = "LoadContactsThread";   // looks nice in Output window
      m_LoadContactsThread.Start();
    }

    //////////////////////////////////////
    // Worker thread function.
    // Called indirectly from btnStartThread_Click
    //////////////////////////////////////
    private void LoadContactsThread_Function()
    {
      CLoadContactsThread loadcontacts_thread;

      loadcontacts_thread = new CLoadContactsThread(m_EventStopLoadContactsThread,
                                                    m_EventLoadContactsThreadStopped,
                                                    this,
                                                    m_folderName,
                                                    m_iMaxContactsToRead);
      loadcontacts_thread.Run();
    }

    //////////////////////////////////////
    // Stop worker thread if it is running.
    // Called when user presses Stop button or form is closed.
    //////////////////////////////////////
    private void LoadContactsThread_Stop()
    {
      txtLog.AppendText("GetContacts - Stopping Thread...\r\n");

      if ( m_LoadContactsThread != null  &&  m_LoadContactsThread.IsAlive )  // thread is active
      {
        // Set event "Stop"
        m_EventStopLoadContactsThread.Set();

        // Wait when thread  will stop or finish
        while (m_LoadContactsThread.IsAlive)
        {
          // We cannot use here infinite wait because our thread
          // makes syncronous calls to main form, this will cause deadlock.
          // Instead of this we wait for event some appropriate time
          // (and by the way give time to worker thread) and
          // process events. These events may contain Invoke calls.

          if ( WaitHandle.WaitAll((new ManualResetEvent[] {m_EventLoadContactsThreadStopped}), 100, true) )
          {
            break;
          }
          System.Windows.Forms.Application.DoEvents();
        }
        LoadContactsThread_Stopped();
      }
    }

    //////////////////////////////////////
    // Add string to list box.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void LoadContactsThread_AppendToLog(String s)
    {
      //listBox1.Items.Add(s);
      txtLog.AppendText("LoadContactsThread - " + s + "\r\n");
    }

    //////////////////////////////////////
    // Update Counter.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void LoadContactsThread_UpdateCounter(int i)
    {
      radioViewAll.Text = "All Contacts: " + i.ToString();
    }

    //////////////////////////////////////
    // Add One Contact.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void LoadContactsThread_AddOneContact(MyContact contact)
    {
      m_ContactList.Add(contact);
      MyContact x = MyContact.Copy(contact);
      m_ContactListOrig.Add(x);
    }

    //////////////////////////////////////
    // Add string to list box.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void LoadContactsThread_AppendToListbox(String s)
    {
      //listBox1.Items.Add(s);
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void LoadContactsThread_AppendToFieldList(MyField fld)
    {
      m_FieldList.Add(fld);
    }

    //////////////////////////////////////
    // Set initial state of controls.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void LoadContactsThread_Finished()
    {
      txtLog.AppendText("GetContacts - Thread finished...\r\n");
      LoadContactsThread_Cleanup();
    }

    //////////////////////////////////////
    // Set initial state of controls.
    //////////////////////////////////////
    private void LoadContactsThread_Stopped()
    {
      txtLog.AppendText("GetContacts - Thread stopped...\r\n");
      LoadContactsThread_Cleanup();
    }

    //////////////////////////////////////
    // Set initial state of controls.
    //////////////////////////////////////
    private void LoadContactsThread_Cleanup()
    {
      if (m_ContactList.Count > 0)
      {
        this.Cursor = Cursors.WaitCursor;
        dgvContacts.DataSource = m_ContactList;
        this.Cursor = Cursors.Arrow;
      }

      btnGetContacts.Enabled = true;
      fileStopGetContactsThreadToolStripMenuItem.Enabled = false;
      dgvContacts.Enabled = true;

      // Enable all Action buttons and menu items
      EnableControls(true);

      //MessageBox.Show(m_ContactList.Count + " contacts found in this folder");
      txtLog.AppendText("GetContacts - " + m_ContactList.Count + " contacts found\r\n");
    }
    #endregion // GetContactsThreadFunctions

    // #########################################################################
    // # region CountFieldsThreadFunctions
    // #########################################################################
    #region CountFieldsThreadFunctions

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void CountFieldsThread_Initialize()
    {
      // Initialize delegates
      m_Delegate_CountFieldsThread_AppendToLog   = new Delegate_CountFieldsThread_AppendToLog_t  (this.CountFieldsThread_AppendToLog);
      m_Delegate_CountFieldsThread_UpdateCounter = new Delegate_CountFieldsThread_UpdateCounter_t(this.CountFieldsThread_UpdateCounter);
      m_Delegate_CountFieldsThread_Finished      = new Delegate_CountFieldsThread_Finished_t     (this.CountFieldsThread_Finished);

      // Initialize events
      m_EventStopCountFieldsThread    = new ManualResetEvent(false);
      m_EventCountFieldsThreadStopped = new ManualResetEvent(false);
    }

    //////////////////////////////////////
    // Get the contacts in the specified the folder.
    //////////////////////////////////////
    private void CountFields(string folderName, int iMaxContactsToRead)
    {
      ThreadStart job;

      //Console.WriteLine("Command: CountFields - Start Thread");
      txtLog.AppendText("CountFields - Start Thread\r\n");

      // Reset events
      m_EventStopCountFieldsThread.Reset();
      m_EventCountFieldsThreadStopped.Reset();

      m_folderName = folderName;
      m_iMaxContactsToRead = iMaxContactsToRead;

      txtLog.AppendText("CountFields \"" + folderName + "\"\r\n");

      // Create worker thread instance
      job = new ThreadStart(this.CountFieldsThread_Function);

      m_CountFieldsThread = new Thread(job);
      m_CountFieldsThread.Name = "CountFieldsThread";   // looks nice in Output window
      m_CountFieldsThread.Start();
    }

    //////////////////////////////////////
    // Worker thread function.
    // Called indirectly from btnStartThread_Click
    //////////////////////////////////////
    private void CountFieldsThread_Function()
    {
      CCountFieldsThread countfields_thread;

      countfields_thread = new CCountFieldsThread  (m_EventStopCountFieldsThread,
                                                    m_EventCountFieldsThreadStopped,
                                                    this,
                                                    m_folderName,
                                                    m_iMaxContactsToRead);
      countfields_thread.Run();
    }

    //////////////////////////////////////
    // Stop worker thread if it is running.
    // Called when user presses Stop button or form is closed.
    //////////////////////////////////////
    private void CountFieldsThread_Stop()
    {
      txtLog.AppendText("CountFields - Stopping Thread...\r\n");

      if ( m_CountFieldsThread != null  &&  m_CountFieldsThread.IsAlive )  // thread is active
      {
        // Set event "Stop"
        m_EventStopCountFieldsThread.Set();

        // Wait when thread  will stop or finish
        while (m_CountFieldsThread.IsAlive)
        {
          // We cannot use here infinite wait because our thread
          // makes syncronous calls to main form, this will cause deadlock.
          // Instead of this we wait for event some appropriate time
          // (and by the way give time to worker thread) and
          // process events. These events may contain Invoke calls.

          if ( WaitHandle.WaitAll((new ManualResetEvent[] {m_EventCountFieldsThreadStopped}), 100, true) )
          {
            break;
          }
          System.Windows.Forms.Application.DoEvents();
        }
        CountFieldsThread_Stopped();
      }
    }

    //////////////////////////////////////
    // Add string to list box.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void CountFieldsThread_AppendToLog(String s)
    {
      txtLog.AppendText("CountFieldsThread - \"" + s + "\"\r\n");
    }

    //////////////////////////////////////
    // Update Counter.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void CountFieldsThread_UpdateCounter(int i)
    {
      radioViewAll.Text = "All Contacts: " + i.ToString();
    }

    //////////////////////////////////////
    // Set initial state of controls.
    // Called from worker thread using delegate and Control.Invoke
    //////////////////////////////////////
    private void CountFieldsThread_Finished()
    {
      txtLog.AppendText("CountFields - Thread finished...\r\n");
      CountFieldsThread_Cleanup();
    }

    //////////////////////////////////////
    // Set initial state of controls.
    //////////////////////////////////////
    private void CountFieldsThread_Stopped()
    {
      txtLog.AppendText("CountFields - Thread stopped...\r\n");
      CountFieldsThread_Cleanup();
    }

    //////////////////////////////////////
    // Set initial state of controls.
    //////////////////////////////////////
    private void CountFieldsThread_Cleanup()
    {
      //HideUnusedColumns();
      btnGetContacts.Enabled = true;
      fileStopGetContactsThreadToolStripMenuItem.Enabled = false;

      dgvContacts.Enabled = true;
      if (m_ContactList.Count > 0)
      {
        this.Cursor = Cursors.WaitCursor;
        dgvContacts.DataSource = m_ContactList;
        this.Cursor = Cursors.Arrow;
      }
      MessageBox.Show(m_ContactList.Count + " contacts found in this folder");
      txtLog.AppendText("CountFields - " + m_ContactList.Count + " contacts processed\r\n");
    }
    #endregion // CountFieldsThreadFunctions

    // #########################################################################
    // # region LocalFunctions
    // #########################################################################
    #region LocalFunctions

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private int StrLen(string Str)
    {
      if (Str != null)
        return Str.Length;
      return 0;   
    }  

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private bool IsMarkedForDeletion(MyContact contact1)
    {
      if (contact1 != null)
        if (contact1.Suffix != null)
          if (contact1.Suffix.Contains("deleteme"))
            return true;
      return false;
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private int SourceOfMailingAddress (MyContact contact1)
    {
      int MailingAddressSource = 0; // None

      if (StrLen(contact1.BusinessAddress) > 0)
      {
        if (contact1.BusinessAddress == contact1.MailingAddress)
          MailingAddressSource = 1; // BusinessAddress
      }
      if (StrLen(contact1.HomeAddress) > 0)
      {
        if (contact1.HomeAddress == contact1.MailingAddress)
          MailingAddressSource = 2; // HomeAddress
      }
      if (StrLen(contact1.OtherAddress) > 0)
      {
        if (contact1.OtherAddress == contact1.MailingAddress)
          MailingAddressSource = 3; // OtherAddress
      }
      return MailingAddressSource;
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private bool HasAnAddressOtherThanMailingAddress (MyContact contact1)
    {
      bool AddressExists = false;
      if (StrLen(contact1.BusinessAddress) > 0)
      {
        AddressExists = true;
      }
      if (StrLen(contact1.HomeAddress) > 0)
      {
        AddressExists = true;
      }
      if (StrLen(contact1.OtherAddress) > 0)
      {
        AddressExists = true;
      }
      return AddressExists;
    }
            
    //////////////////////////////////////
    //
    //////////////////////////////////////
    private string BuildAddressFromComponents(string AddressStreet,
                                              string AddressState,
                                              string AddressCity,
                                              string AddressPostalCode,
                                              string AddressCountry)
    {
      string tmpstr = "";

      if (StrLen(AddressStreet)>0)
      {
          if (tmpstr.Length > 0)
              tmpstr += "\\r\\n";
          tmpstr += AddressStreet;
      }
      if (StrLen(AddressState)>0)
      {
          if (tmpstr.Length > 0)
              tmpstr += "\\r\\n";
          tmpstr += AddressState;
      }
      if (StrLen(AddressCity)>0)
      {
          if (tmpstr.Length > 0)
              tmpstr += "\\r\\n";
          tmpstr += AddressCity;
      }
      if (StrLen(AddressPostalCode)>0)
      {
          if (tmpstr.Length > 0)
              tmpstr += "\\r\\n";
          tmpstr += AddressPostalCode;
      }
      if (StrLen(AddressCountry)>0)
      {
          if (tmpstr.Length > 0)
              tmpstr += "\\r\\n";
          tmpstr += AddressCountry;
      }
      return tmpstr;
    }
    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void EnableControls(bool bEnabled)
    {
      btnRefreshFolders.Enabled                         = bEnabled;
      btnGetContacts.Enabled                            = bEnabled;
      txtMaxContactstoRead.Enabled                      = bEnabled;
      btnAnalyse.Enabled                                = bEnabled;
      btnContactDetails.Enabled                         = bEnabled;
      btnViewFields.Enabled                             = bEnabled;

      radioViewAll.Enabled                              = bEnabled;
      radioViewAllByName.Enabled                        = bEnabled;
      radioViewDuplicates.Enabled                       = bEnabled;
      radioViewSimilar.Enabled                          = bEnabled;
      radioViewVaguelySimilar.Enabled                   = bEnabled;
      radioViewSameName.Enabled                         = bEnabled;
      radioViewSameNameAndSuffix.Enabled                = bEnabled;
      radioViewTruncatedNotes.Enabled                   = bEnabled;
      radioViewDislikedContents.Enabled                 = bEnabled;
      radioViewDuplicateAddresses.Enabled               = bEnabled;
      radioViewDuplicatePhoneNumbers.Enabled            = bEnabled;
      radioViewDuplicateEmails.Enabled                  = bEnabled;
      radioViewUnlinkedMailingAddresses.Enabled         = bEnabled;
      radioViewLinkedMailingAddresses.Enabled           = bEnabled;
      radioViewEmptyMailingAddresses.Enabled            = bEnabled;
      radioViewMismatchedAddrAndComp.Enabled            = bEnabled;
      radioViewMinimalInfo.Enabled                      = bEnabled;

    //fileToolStripMenuItem.Enabled                           = bEnabled;
      fileNewToolStripMenuItem.Enabled                        = bEnabled;
      fileOpenToolStripMenuItem.Enabled                       = bEnabled;
      filePrintPreviewToolStripMenuItem.Enabled               = bEnabled;
      filePrintToolStripMenuItem.Enabled                      = bEnabled;
      fileSaveToolStripMenuItem.Enabled                       = bEnabled;
      fileSaveAsToolStripMenuItem.Enabled                     = bEnabled;
    //fileExitToolStripMenuItem.Enabled.Enabled               = bEnabled;
      fileRefreshFoldersToolStripMenuItem.Enabled             = bEnabled;
      fileLoadContactsToolStripMenuItem.Enabled               = bEnabled;
      fileSaveChangedContactsToolStripMenuItem.Enabled        = bEnabled;
      fileStopGetContactsThreadToolStripMenuItem.Enabled      = bEnabled;

    //editToolStripMenuItem.Enabled                           = bEnabled;
    //editUndoToolStripMenuItem.Enabled                       = bEnabled;
    //editCopyToolStripMenuItem.Enabled                       = bEnabled;
    //editCutToolStripMenuItem.Enabled                        = bEnabled;
    //editPasteToolStripMenuItem.Enabled                      = bEnabled;
    //editRedoToolStripMenuItem.Enabled                       = bEnabled;
    //editSelectAllToolStripMenuItem.Enabled                  = bEnabled;

    //viewToolStripMenuItem.Enabled                           = bEnabled;
      viewShowAllContactsToolStripMenuItem.Enabled            = bEnabled;
      viewShowDuplicateContactsToolStripMenuItem.Enabled      = bEnabled;
      viewShowSimilarContactsToolStripMenuItem.Enabled        = bEnabled;
      viewShowVaguelySimilarContactsToolStripMenuItem.Enabled = bEnabled;
      viewShowWithDislikedContentsToolStripMenuItem.Enabled   = bEnabled;
      viewShowWithTheSameNameToolStripMenuItem.Enabled        = bEnabled;
      viewShowWithTheSameNameAndSuffixToolStripMenuItem.Enabled = bEnabled;
      viewShowWithTruncatedNotesToolStripMenuItem.Enabled     = bEnabled;
      viewShowWithDuplicateAddressesToolStripMenuItem.Enabled = bEnabled;
      viewHideUnusedColumnsToolStripMenuItem.Enabled          = bEnabled;

    //toolsToolStripMenuItem.Enabled                          = bEnabled;
    //toolsCustomizeToolStripMenuItem.Enabled                 = bEnabled;
    //toolsOptionsToolStripMenuItem.Enabled                   = bEnabled;
      toolsAnalyseContactsToolStripMenuItem.Enabled           = bEnabled;
      toolsEditContactDetailsToolStripMenuItem.Enabled        = bEnabled;
      toolsFieldsToolStripMenuItem.Enabled                    = bEnabled;
      toolsNewContactToolStripMenuItem.Enabled                = bEnabled;
    }

    //////////////////////////////////////
    // Insert a contact into a secondary list, if it is not there already
    //////////////////////////////////////
    private bool InsertContactIntoSecondaryList(List<MyContact> argWhichList, MyContact argContact)
    {
      // Create shallow copy of the contact to add to the secondary list
      MyContact newContact = new MyContact();
      newContact = argContact;
      if (!argWhichList.Contains(newContact))
      {
        argWhichList.Add(newContact);
        return true;
      }
      return false;
    }

    //////////////////////////////////////
    // Analyse the contacts.
    //////////////////////////////////////
    private void AnalyseContacts()
    {
      //try
      {
        txtLog.AppendText("Analysing Contacts...\r\n");
        this.Cursor = Cursors.WaitCursor;

        tlpMain.Width = this.Width - 20; // MainForm.Width - 10;
        dgvContacts.Width = tlpMain.Width - 10;

        ClearSecondaryLists();

        foreach (MyContact contact1 in m_ContactList)
        {
          foreach (MyContact contact2 in m_ContactList)
          {
            // Don't test a contact against itself
            if (contact1.Ndx == contact2.Ndx)
              continue;

            if (contact1.IsIdentical(contact1, contact2, MyContact.COMPARING_DIFFERENT_ITEMS))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListDuplicates,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListDuplicates,contact2) == true) )
              {
                m_NumberOfDuplicateContacts++;
                //txtLog.AppendText("Duplicate: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
            if (contact1.IsSimilar(contact1, contact2, true))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListSimilar,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListSimilar,contact2) == true) )
              {
                m_NumberOfSimilarContacts++;
                //txtLog.AppendText("Similar: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
            if (contact1.IsSimilar(contact1, contact2, false))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListVaguelySimilar,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListVaguelySimilar,contact2) == true) )
              {
                m_NumberOfVaguelySimilarContacts++;
                //txtLog.AppendText("VaguelySimilar: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
            if (contact1.HasSameName(contact1, contact2, false))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListSameName,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListSameName,contact2) == true) )
              {
                m_NumberOfSameNameContacts++;
                //txtLog.AppendText("SameName: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
            if (contact1.HasSameName(contact1, contact2, true))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListSameNameAndSuffix,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListSameNameAndSuffix,contact2) == true) )
              {
                m_NumberOfSameNameAndSuffixContacts++;
                //txtLog.AppendText("SameNameAndSuffix: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
            if (contact1.HasTruncatedNotes(contact1,contact2))
            {
              if ((InsertContactIntoSecondaryList(m_ContactListTruncatedNotes,contact1) == true) ||
                  (InsertContactIntoSecondaryList(m_ContactListTruncatedNotes,contact2) == true) )
              {
                m_NumberOfTruncatedNotes++;
                //txtLog.AppendText("TruncatedNotes: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
              }
            }
          }
          if (contact1.HasDislikedContent(contact1))
          {
            if (InsertContactIntoSecondaryList(m_ContactListDislikedContent,contact1) == true)
            {
              m_NumberOfDislikedContent++;
              //txtLog.AppendText("DislikedContent: " + "&" + "\r\n");
            }
          }
          if (contact1.HasDuplicateAddress(contact1))
          {
            if (InsertContactIntoSecondaryList(m_ContactListDuplicateAddresses,contact1) == true)
            {
              m_NumberOfDuplicateAddresses++;
              //txtLog.AppendText("DuplicateAddresses: " + "&" + "\r\n");
            }
          }
          if (contact1.HasDuplicatePhoneNumbers(contact1))
          {
            if (InsertContactIntoSecondaryList(m_ContactListDuplicatePhoneNumbers,contact1) == true)
            {
              m_NumberOfDuplicatePhoneNumbers++;
              //txtLog.AppendText("DuplicatePhoneNumbers: " + "&" + "\r\n");
            }
          }
          if (contact1.HasDuplicateEmailAddress(contact1))
          {
            if (InsertContactIntoSecondaryList(m_ContactListDuplicateEmailAddress,contact1) == true)
            {
              m_NumberOfDuplicateEmailAddress++;
              //txtLog.AppendText("DuplicateEmailAddress: " + "&" + "\r\n");
            }
          }
          if (contact1.HasMinimalInfo(contact1))
          {
            if (InsertContactIntoSecondaryList(m_ContactListMinimalInfo,contact1) == true)
            {
              m_NumberOfMinimalInfo++;
              //txtLog.AppendText("DuplicatePhoneNumbers: " + "&" + "\r\n");
            }
          }
          
          CheckForUnlinkedMailingAddresses1(contact1);
          CheckForMismatchedAddressComponents(contact1);

          // Gather Statistics
          if (contact1.Body != null &&
              contact1.Body.ToString().Length > m_MaxNotesLength)
          {
            m_MaxNotesLength = contact1.Body.ToString().Length;
            m_MaxNotesNdx = contact1.Ndx;
          }
        }
        UpdateStatisticsFields();

        this.Cursor = Cursors.Arrow;
        txtLog.AppendText("Analysis Complete\r\n");
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void CheckForUnlinkedMailingAddresses1(MyContact contact1)
    {
      if (StrLen(contact1.MailingAddress) > 0)
      {
        // if MailingAddress is not a copy Business, Home or Other Address
        if (SourceOfMailingAddress(contact1) == 0)
        {
          if (InsertContactIntoSecondaryList(m_ContactListUnlinkedMailingAddresses, contact1) == true)
          {
            m_SourceOfMailingAddressUnknown++;
          }
        }
        else
        {
          if (InsertContactIntoSecondaryList(m_ContactListLinkedMailingAddresses, contact1) == true)
          {
            m_SourceOfMailingAddressKnown++;
          }
        }
      }
      else
      {
        if (HasAnAddressOtherThanMailingAddress(contact1))
        {
          if (InsertContactIntoSecondaryList(m_ContactListEmptyMailingAddresses, contact1) == true)
          {
            m_MailingAddressNotSet++;
          }
        }
      }
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void CheckForMismatchedAddressComponents1(string AddressStreet,
                                                      string AddressState,
                                                      string AddressCity,
                                                      string AddressPostalCode,
                                                      string AddressCountry,
                                                      MyContact contact1 )
    {
      string tmpstr = "";

      if (StrLen(contact1.MailingAddress) > 0)
      {
        tmpstr = BuildAddressFromComponents(AddressStreet,
                                            AddressState,
                                            AddressCity,
                                            AddressPostalCode,
                                            AddressCountry);
        if (contact1.MailingAddress != tmpstr)
        {
          if (InsertContactIntoSecondaryList(m_ContactListMismatchedAddressComponents, contact1) == true)
          {
            m_NumberOfMismatchedAddressComponents++;
          }
        }
      }
      else // Mailing Address is empty
      {
        if (StrLen(AddressStreet    )>0 ||
            StrLen(AddressState     )>0 ||
            StrLen(AddressCity      )>0 ||
            StrLen(AddressPostalCode)>0 ||
            StrLen(AddressCountry   )>0 )
        {
          if (InsertContactIntoSecondaryList(m_ContactListMismatchedAddressComponents, contact1) == true)
          {
            m_NumberOfMismatchedAddressComponents++;
          }
        }
      }
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void CheckForMismatchedAddressComponents(MyContact contact1)
    {
      CheckForMismatchedAddressComponents1(contact1.MailingAddressStreet,
                                           contact1.MailingAddressState,
                                           contact1.MailingAddressCity,
                                           contact1.MailingAddressPostalCode,
                                           contact1.MailingAddressCountry,
                                           contact1);
      CheckForMismatchedAddressComponents1(contact1.BusinessAddressStreet,
                                           contact1.BusinessAddressState,
                                           contact1.BusinessAddressCity,
                                           contact1.BusinessAddressPostalCode,
                                           contact1.BusinessAddressCountry,
                                           contact1);
      CheckForMismatchedAddressComponents1(contact1.HomeAddressStreet,
                                           contact1.HomeAddressState,
                                           contact1.HomeAddressCity,
                                           contact1.HomeAddressPostalCode,
                                           contact1.HomeAddressCountry,
                                           contact1);
    }

    //////////////////////////////////////
    // MarkDuplicatesForDeletion
    //////////////////////////////////////
    private void DoActionMarkDuplicatesForDeletion()
    {
      int ContactsMarkedForDeletion = 0;
      //try
      {
        txtLog.AppendText("Mark Duplicates For Deletion...\r\n");
        this.Cursor = Cursors.WaitCursor;

        foreach (MyContact contact1 in m_ContactList)
        {
          // Skip over contacts that are already identical with some other contact
          if (IsMarkedForDeletion(contact1))
            continue;
          foreach (MyContact contact2 in m_ContactList)
          {
            // Don't test a contact against itself
            if (contact1.Ndx == contact2.Ndx)
              continue;

            // Skip over contacts that are already identical with some other contact
            if (IsMarkedForDeletion(contact1))
              continue;

            if (contact1.IsIdentical(contact1, contact2, MyContact.COMPARING_DIFFERENT_ITEMS))
            {
              ContactsMarkedForDeletion++;
              txtLog.AppendText("Contact " + contact2.Ndx.ToString() + " is identical to " + contact1.Ndx.ToString() + ". Marking " + contact2.Ndx.ToString() + " for deletion\r\n");
              contact2.Suffix = "deleteme" + " " + contact2.Suffix;
            }
          }
        }

        this.Cursor = Cursors.Arrow;
        txtLog.AppendText("Contacts marked for deletion: " + ContactsMarkedForDeletion.ToString() + " (each name Suffix prefixed with \"deleteme\")\r\n");
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }

    private string PutSpaceAfterSubString(string Str, string Substring)
    {
      string SubstringWithSpace = Substring + " ";
      string SubstringWithTwoSpaces = Substring + "  ";
      if (Str.StartsWith(Substring))
      {
        if (!Str.StartsWith(SubstringWithSpace))
        {
          Str = Str.Replace(Substring,SubstringWithSpace);
          Str = Str.Replace(SubstringWithTwoSpaces,SubstringWithSpace);
        }
      }
      return Str;
    }

    private string FixPhoneNumber(int Ndx, string PhoneNumber, ref int Changes)
    {
      string OldValue = PhoneNumber;
      string NewValue = PhoneNumber;

      if (PhoneNumber == null)
        return OldValue;

      NewValue = NewValue.Trim();
      //NewValue = NewValue.Replace("("," ");
      //NewValue = NewValue.Replace(")"," ");
      NewValue = NewValue.Replace("-","");
      NewValue = NewValue.Replace("  "," ");
      NewValue = NewValue.Replace("  "," ");
      NewValue = NewValue.Replace("  "," ");
      NewValue = NewValue.Replace("  "," ");
      NewValue = PutSpaceAfterSubString(NewValue,"+1"); // USA
      NewValue = PutSpaceAfterSubString(NewValue,"+1 425");
      NewValue = PutSpaceAfterSubString(NewValue,"+1 760");
      NewValue = PutSpaceAfterSubString(NewValue,"+263"); // Zim
      NewValue = PutSpaceAfterSubString(NewValue,"+263 4"); // Sby?
      NewValue = PutSpaceAfterSubString(NewValue,"+263 7"); // Mob?
      NewValue = PutSpaceAfterSubString(NewValue,"+263 91");
      NewValue = PutSpaceAfterSubString(NewValue,"+27"); // RSA
      NewValue = PutSpaceAfterSubString(NewValue,"+27 11"); // Jhb
      NewValue = PutSpaceAfterSubString(NewValue,"+27 12"); // Pta
      NewValue = PutSpaceAfterSubString(NewValue,"+27 15");
      NewValue = PutSpaceAfterSubString(NewValue,"+27 21"); // CT
      NewValue = PutSpaceAfterSubString(NewValue,"+27 31"); // Durban
      NewValue = PutSpaceAfterSubString(NewValue,"+27 32");
      NewValue = PutSpaceAfterSubString(NewValue,"+27 33");
      NewValue = PutSpaceAfterSubString(NewValue,"+27 41");
      NewValue = PutSpaceAfterSubString(NewValue,"+27 464");
      NewValue = PutSpaceAfterSubString(NewValue,"+27 82"); // Cell
      NewValue = PutSpaceAfterSubString(NewValue,"+27 83"); // Cell
      NewValue = PutSpaceAfterSubString(NewValue,"+27 72"); // Cell
      NewValue = PutSpaceAfterSubString(NewValue,"+27 73"); // Cell
      NewValue = PutSpaceAfterSubString(NewValue,"+30"); // Greece?
      NewValue = PutSpaceAfterSubString(NewValue,"+31");
      NewValue = PutSpaceAfterSubString(NewValue,"+33");
      NewValue = PutSpaceAfterSubString(NewValue,"+353"); // Ireland
      NewValue = PutSpaceAfterSubString(NewValue,"+353 1"); // Dublin?
      NewValue = PutSpaceAfterSubString(NewValue,"+359"); // Bg
      NewValue = PutSpaceAfterSubString(NewValue,"+359 2"); // Sophia?
      NewValue = PutSpaceAfterSubString(NewValue,"+39"); // Italy
      NewValue = PutSpaceAfterSubString(NewValue,"+39 03");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 03 46");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 03 47");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 03 48");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 04");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 0546");
      NewValue = PutSpaceAfterSubString(NewValue,"+39 55");
      NewValue = PutSpaceAfterSubString(NewValue,"+43");
      NewValue = PutSpaceAfterSubString(NewValue,"+43 732");
      NewValue = PutSpaceAfterSubString(NewValue,"+44"); // UK
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1144");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1173");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1179");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1189");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1229");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1232");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1252");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1264");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1273");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1293");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1327");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1328");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1334");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1342");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1344");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1364"); // Ashburton
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1382");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1383");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1388");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1392");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1403");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1483");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1634");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1642");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1647");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1708");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1737"); // Reigate/Redhill
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1883"); // Caterham
      NewValue = PutSpaceAfterSubString(NewValue,"+44 1959");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 20"); // London
      NewValue = PutSpaceAfterSubString(NewValue,"+44 226");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 232");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 2380");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 2890");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 44");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 44 247");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 44 815");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 57");
      NewValue = PutSpaceAfterSubString(NewValue,"+44 635");

      NewValue = PutSpaceAfterSubString(NewValue,"+46"); // Sweden
      NewValue = PutSpaceAfterSubString(NewValue,"+49"); // Germany
      NewValue = PutSpaceAfterSubString(NewValue,"+49 89"); // Munich
      NewValue = PutSpaceAfterSubString(NewValue,"+55"); // Brazil
      NewValue = PutSpaceAfterSubString(NewValue,"+55 21"); // Rio?
      NewValue = PutSpaceAfterSubString(NewValue,"+61"); // Australia
      NewValue = PutSpaceAfterSubString(NewValue,"+64"); // New Zealand
      NewValue = PutSpaceAfterSubString(NewValue,"+64 3"); // Christchurch
      NewValue = PutSpaceAfterSubString(NewValue,"+82"); // S. Korea
      NewValue = PutSpaceAfterSubString(NewValue,"+82 0");
      NewValue = PutSpaceAfterSubString(NewValue,"+82 11");
      NewValue = PutSpaceAfterSubString(NewValue,"+82 2");
      NewValue = PutSpaceAfterSubString(NewValue,"+91"); // India

      if (NewValue != OldValue)
      {
        Changes++;
        txtLog.AppendText("Phone number " + Ndx.ToString() + " changed from \"" + OldValue + "\" to \"" + NewValue + "\"\r\n");
        return NewValue;
      }
      return OldValue;
    }

    //////////////////////////////////////
    // DoActionFixPhoneNumbers
    //////////////////////////////////////
    private void DoActionFixPhoneNumbers()
    {
      int Changes = 0;
      //try
      {
        txtLog.AppendText("Fix Phone Numbers...\r\n");
        this.Cursor = Cursors.WaitCursor;

        foreach (MyContact contact1 in m_ContactList)
        {
          // Skip over contacts that are already identical with some other contact
          if (IsMarkedForDeletion(contact1))
            continue;
          contact1.AssistantTelephoneNumber   = FixPhoneNumber(contact1.Ndx, contact1.AssistantTelephoneNumber  , ref Changes);
          contact1.BusinessTelephoneNumber    = FixPhoneNumber(contact1.Ndx, contact1.BusinessTelephoneNumber   , ref Changes);
          contact1.Business2TelephoneNumber   = FixPhoneNumber(contact1.Ndx, contact1.Business2TelephoneNumber  , ref Changes);
          contact1.CallbackTelephoneNumber    = FixPhoneNumber(contact1.Ndx, contact1.CallbackTelephoneNumber   , ref Changes);
          contact1.CarTelephoneNumber         = FixPhoneNumber(contact1.Ndx, contact1.CarTelephoneNumber        , ref Changes);
          contact1.CompanyMainTelephoneNumber = FixPhoneNumber(contact1.Ndx, contact1.CompanyMainTelephoneNumber, ref Changes);
          contact1.HomeTelephoneNumber        = FixPhoneNumber(contact1.Ndx, contact1.HomeTelephoneNumber       , ref Changes);
          contact1.Home2TelephoneNumber       = FixPhoneNumber(contact1.Ndx, contact1.Home2TelephoneNumber      , ref Changes);
          contact1.MobileTelephoneNumber      = FixPhoneNumber(contact1.Ndx, contact1.MobileTelephoneNumber     , ref Changes);
          contact1.OtherTelephoneNumber       = FixPhoneNumber(contact1.Ndx, contact1.OtherTelephoneNumber      , ref Changes);
          contact1.PrimaryTelephoneNumber     = FixPhoneNumber(contact1.Ndx, contact1.PrimaryTelephoneNumber    , ref Changes);
          contact1.RadioTelephoneNumber       = FixPhoneNumber(contact1.Ndx, contact1.RadioTelephoneNumber      , ref Changes);
          contact1.TTYTDDTelephoneNumber      = FixPhoneNumber(contact1.Ndx, contact1.TTYTDDTelephoneNumber     , ref Changes);
        }

        this.Cursor = Cursors.Arrow;
        txtLog.AppendText("Phone Numbers fixed: " + Changes.ToString() + "\r\n");
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }

    //////////////////////////////////////
    // DoActionMergeDuplicateContacts
    //////////////////////////////////////
    private void DoActionMergeDuplicateContacts()
    {
      int TotalContactsMerged = 0;
      int TotalFieldsMerged = 0;

      if (MessageBox.Show("This is a bit dangerous. It will loose whatever individuality may exist in each similar contact. Continue?","Merge Duplicate Contacts...",MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
      {
        return;
      }

      //try
      {
        txtLog.AppendText("Merge Duplicate Contacts...\r\n");
        this.Cursor = Cursors.WaitCursor;

        tlpMain.Width = this.Width - 20; // MainForm.Width - 10;
        dgvContacts.Width = tlpMain.Width - 10;

        foreach (MyContact contact1 in m_ContactList)
        {
          foreach (MyContact contact2 in m_ContactList)
          {
            // Don't test a contact against itself
            if (contact1.Ndx == contact2.Ndx)
              continue;

            if (contact1.HasSameName(contact1, contact2, true))
            {
              int FieldsMerged;
              FieldsMerged = contact1.MergeContacts(contact1, contact2);
              if (FieldsMerged > 0) 
              {
                txtLog.AppendText("FieldsMerged (" + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + ") = " + FieldsMerged.ToString() + "\r\n");
                TotalContactsMerged++;
                TotalFieldsMerged += FieldsMerged;
              }
            }
          }
          //if (contact1.HasDuplicateAddress(contact1))
          //{
          //  //txtLog.AppendText("DuplicateAddresses: " + "&" + "\r\n");
          //}
        }

        this.Cursor = Cursors.Arrow;

        txtLog.AppendText("Total Contacts Merged = " + TotalContactsMerged.ToString() + "\r\n");
        txtLog.AppendText("Total Fields   Merged = " + TotalFieldsMerged.ToString() + "\r\n");
        txtLog.AppendText("Merge Duplicate Contacts Complete\r\n");
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }

    //////////////////////////////////////
    // Load the contact folders from outlook application
    //////////////////////////////////////
    private void LoadContactFoldersCombo()
    {
      OutLook._Application outlookObj;
      OutLook.MAPIFolder contactsFolder;

      try
      {
        txtLog.AppendText("Refresh Folders\r\n");

        outlookObj = new OutLook.Application();
        contactsFolder = (OutLook.MAPIFolder)outlookObj.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);

        if (!cbFolder.Items.Contains("Default"))
        {
          cbFolder.Items.Add("Default");
        }

        // Verify the custom folder in Outlook
        foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
        {
          if (!cbFolder.Items.Contains(subFolder.Name))
          {
            cbFolder.Items.Add(subFolder.Name);
          }
        }
      }
      catch (System.Exception ex)
      {
        txtLog.AppendText("Load Outlook Contact Folders failed with \"" + ex.Message + "\"\r\n");
      }
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void InitialiseForbiddenFields()
    // I can never find this function, so add some words as alternate searches...
    // i.e. disallowed, not permitted, illegal
    {
      m_ForbiddenFieldList.Clear();

      AddForbiddenField("Account",0);
      AddForbiddenField("AssistantName",0);
      AddForbiddenField("AssistantTelephoneNumber",0);
      AddForbiddenField("BillingInformation",0);
      AddForbiddenField("Business2TelephoneNumber",0);
      AddForbiddenField("BusinessAddressPostOfficeBox",0);
      AddForbiddenField("CallbackTelephoneNumber",0);
      AddForbiddenField("CarTelephoneNumber",0);
      AddForbiddenField("Children",0);
      AddForbiddenField("Companies",0);
      AddForbiddenField("CompanyMainTelephoneNumber",0);
      AddForbiddenField("ComputerNetworkName",0);
      AddForbiddenField("CustomerID",0);
      AddForbiddenField("Department",0);
      AddForbiddenField("FTPSite",0);
      AddForbiddenField("Gender",0);
      AddForbiddenField("GovernmentIDNumber",0);
      AddForbiddenField("Hobby",0);
      AddForbiddenField("Home2TelephoneNumber",0);
      AddForbiddenField("HomeFaxNumber",0);
      AddForbiddenField("Importance",0);
      AddForbiddenField("ISDNNumber",0);
      AddForbiddenField("Journal",0);
      AddForbiddenField("Language",0);
      AddForbiddenField("MailingAddressPostOfficeBox",0);
      AddForbiddenField("ManagerName",0);
      AddForbiddenField("Mileage",0);
      AddForbiddenField("NoAging",0);
      AddForbiddenField("OfficeLocation",0);
      AddForbiddenField("OrganizationalIDNumber",0);
      AddForbiddenField("OtherAddress",0);
      AddForbiddenField("OtherAddressCity",0);
      AddForbiddenField("OtherAddressCountry",0);
      AddForbiddenField("OtherAddressPostalCode",0);
      AddForbiddenField("OtherAddressPostOfficeBox",0);
      AddForbiddenField("OtherAddressState",0);
      AddForbiddenField("OtherAddressStreet",0);
      AddForbiddenField("OtherFaxNumber",0);
      AddForbiddenField("PagerNumber",0);
      AddForbiddenField("PersonalHomePage",0);
      AddForbiddenField("PrimaryTelephoneNumber",0);
      AddForbiddenField("Profession",0);
      AddForbiddenField("RadioTelephoneNumber",0);
      AddForbiddenField("ReferredBy",0);
      AddForbiddenField("Saved",0);
      //AddForbiddenField("Suffix",0);
      AddForbiddenField("TelexNumber",0);
      AddForbiddenField("TTYTDDTelephoneNumber",0);
      AddForbiddenField("UnRead",0);
      AddForbiddenField("User1",0);
      AddForbiddenField("User2",0);
      AddForbiddenField("User3",0);
      AddForbiddenField("User4",0);
      AddForbiddenField("UserCertificate",0);
      AddForbiddenField("YomiCompanyName",0);
      AddForbiddenField("YomiFirstName",0);
      AddForbiddenField("YomiLastName",0);
    }

    //////////////////////////////////////
    // Fields should be never used...
    //////////////////////////////////////
    private bool IsForbiddenField(string FieldName)
    {
      foreach (MyField fld in m_ForbiddenFieldList)
      {
        if (FieldName == fld.FieldName)
          return true;
      }
      return false;
    }

    //////////////////////////////////////
    // Hide unused columns
    //////////////////////////////////////
    private void HideUnusedColumns()
    {
      int NumberOfHiddenColumns = 0;
      Object SaveDataSource;

      txtLog.AppendText("Hiding Columns...\r\n");
      this.Cursor = Cursors.WaitCursor;

      SaveDataSource = dgvContacts.DataSource;
      dgvContacts.DataSource = DummyDataSource;
      dgvContacts.Enabled = false;
      dgvContacts.SuspendLayout();

      // Hide unused DataGridView::Columns
      foreach (DataGridViewColumn col in dgvContacts.Columns)
      {
        col.Visible = true;
        // 01/01/4501
        if (
            //col.HeaderText == "Account" ||
            col.HeaderText == "EntryID" ||
            //col.HeaderText == "Journal" ||
            col.HeaderText == "MessageClass" ||
            //col.HeaderText == "NoAging" ||
            col.HeaderText == "OutlookInternalVersion" ||
            col.HeaderText == "OutlookVersion" ||
            //col.HeaderText == "UnRead" ||
            //col.HeaderText == "OutlookNdx" ||
            false
           )
        {
          if (col.Visible)
          {
            NumberOfHiddenColumns++;
            col.Visible = false;
          }
          continue;
        }

        if (col.Visible)
        {
          foreach (MyField fld in m_FieldList)
          {
            if (fld.FieldName == col.HeaderText && fld.Occurances==0)
            {
              NumberOfHiddenColumns++;
              col.Visible = false;
              break;
            }
          }
        }
      }

      dgvContacts.ResumeLayout();
      dgvContacts.Enabled = true;
      dgvContacts.DataSource = SaveDataSource;

      this.Cursor = Cursors.Arrow;
      txtLog.AppendText("Number Of Hidden Columns = " + NumberOfHiddenColumns + "\r\n");
    }

    //////////////////////////////////////
    // Clear all items from the Secondary Lists
    //////////////////////////////////////
    private void ClearPrimaryLists()
    {
      dgvContacts.Enabled = false;

      // Clear all primary lists
      m_ContactList.Clear();
      m_ContactListOrig.Clear();
      m_FieldList.Clear();
    }

    //////////////////////////////////////
    // Clear all items from the Secondary Lists
    //////////////////////////////////////
    private void ClearSecondaryLists()
    { 
     // Turn the DataGrid away from any one of these lists
     if (dgvContacts.DataSource == m_ContactList ||
         dgvContacts.DataSource == m_ContactListDuplicates ||
         dgvContacts.DataSource == m_ContactListSimilar ||
         dgvContacts.DataSource == m_ContactListVaguelySimilar ||
         dgvContacts.DataSource == m_ContactListSameName ||
         dgvContacts.DataSource == m_ContactListSameNameAndSuffix ||
         dgvContacts.DataSource == m_ContactListTruncatedNotes ||
         dgvContacts.DataSource == m_ContactListDislikedContent )
      {
         this.Cursor = Cursors.WaitCursor;
         if (m_ContactList.Count > 0)
         {
           // Set the DataGrid to display primary list
           dgvContacts.Enabled = true;
           dgvContacts.DataSource = m_ContactList;
         }
         else
         {
           // Reset the DataGrid
           dgvContacts.Enabled = false;
           dgvContacts.DataSource = DummyDataSource;
         }
         this.Cursor = Cursors.Arrow;
      }
      
      // Clear all secondary lists
      m_ContactListDuplicates.Clear();
      m_ContactListSimilar.Clear();
      m_ContactListVaguelySimilar.Clear();
      m_ContactListSameName.Clear();
      m_ContactListSameNameAndSuffix.Clear();
      m_ContactListTruncatedNotes.Clear();
      m_ContactListDislikedContent.Clear();
      m_ContactListDuplicateAddresses.Clear();
      m_ContactListDuplicatePhoneNumbers.Clear();
      m_ContactListDuplicateEmailAddress.Clear();
      m_ContactListUnlinkedMailingAddresses.Clear();
      m_ContactListLinkedMailingAddresses.Clear();
      m_ContactListEmptyMailingAddresses.Clear();
      m_ContactListMismatchedAddressComponents.Clear();
      //m_ContactListMatchedAddressComponents.Clear();
      m_ContactListMinimalInfo.Clear();

      m_NumberOfDuplicateContacts           = 0;
      m_NumberOfSameNameContacts            = 0;
      m_NumberOfSameNameAndSuffixContacts   = 0;
      m_NumberOfSimilarContacts             = 0;
      m_NumberOfVaguelySimilarContacts      = 0;
      m_NumberOfTruncatedNotes              = 0;
      m_NumberOfDislikedContent             = 0;
      m_NumberOfDuplicateAddresses          = 0;
      m_NumberOfDuplicatePhoneNumbers       = 0;
      m_NumberOfDuplicateEmailAddress       = 0;
      m_SourceOfMailingAddressUnknown       = 0;
      m_SourceOfMailingAddressKnown         = 0;
      m_MailingAddressNotSet                = 0;
      m_NumberOfMismatchedAddressComponents = 0;
      m_NumberOfMinimalInfo                 = 0;

      m_MaxNotesLength                 = 0;
      m_MaxNotesNdx                    = 0;
      UpdateStatisticsFields();
    }

    private void UpdateStatisticsFields()
    {
        radioViewDuplicates.Text                  = "Duplicates: "         + m_NumberOfDuplicateContacts.ToString();
        radioViewSimilar.Text                     = "Similar: "             + m_NumberOfSimilarContacts.ToString();
        radioViewVaguelySimilar.Text              = "Vaguely Similar: "     + m_NumberOfVaguelySimilarContacts.ToString();
        radioViewSameName.Text                    = "SameName: "            + m_NumberOfSameNameContacts.ToString();
        radioViewSameNameAndSuffix.Text           = "SameNameSuffix: "      + m_NumberOfSameNameAndSuffixContacts.ToString();
        radioViewTruncatedNotes.Text              = "Truncated Notes: "     + m_NumberOfTruncatedNotes.ToString();
        radioViewDislikedContents.Text            = "Ampersands: "          + m_NumberOfDislikedContent.ToString();
        radioViewDuplicateAddresses.Text          = "Dupl Addresses: "      + m_NumberOfDuplicateAddresses.ToString();
        radioViewDuplicatePhoneNumbers.Text       = "Dupl Phones: "         + m_NumberOfDuplicatePhoneNumbers.ToString();
        radioViewDuplicateEmails.Text             = "Dupl Emails: "         + m_NumberOfDuplicateEmailAddress.ToString();
        radioViewUnlinkedMailingAddresses.Text    = "MailAddr!=AnyAddr: "   + m_SourceOfMailingAddressUnknown.ToString();
        radioViewLinkedMailingAddresses.Text      = "MailAddr Set: "        + m_SourceOfMailingAddressKnown.ToString();
        radioViewEmptyMailingAddresses.Text       = "MailAddr NotSet: "     + m_MailingAddressNotSet.ToString();
        radioViewMismatchedAddrAndComp.Text       = "Addr!=AddrFlds: "      + m_NumberOfMismatchedAddressComponents.ToString();
        radioViewMinimalInfo.Text                 = "Minimal Info: "        + m_NumberOfMinimalInfo.ToString();
        

        String TempStr;
        TempStr = "Ndx=" + m_MaxNotesNdx.ToString() + ", Len=" + m_MaxNotesLength.ToString();
        txtStatusLine.AppendText("MaxNotesLen=");
        txtStatusLine.AppendText(TempStr);
        txtStatusLine.AppendText(";");
    }

    //////////////////////////////////////
    // Refresh the contacts.
    //////////////////////////////////////
    private void DoActionGetContacts()
    {
      string cntctFldrName = string.Empty;
      int iMaxContactsToRead;

      txtLog.AppendText("Retrieving contacts from Outlook...\r\n");

      btnGetContacts.Enabled = false;
      fileStopGetContactsThreadToolStripMenuItem.Enabled = true;

      if (cbFolder.SelectedItem == null)
      {
        cntctFldrName = "Default";
        cbFolder.SelectedItem = "Default";
      }
      else
      {
        cntctFldrName = cbFolder.SelectedItem.ToString();
      }

      if (txtMaxContactstoRead.TextLength>0)
        iMaxContactsToRead = System.Convert.ToInt32(txtMaxContactstoRead.Text);
      else
        iMaxContactsToRead = 1000000;

      ClearPrimaryLists();
      ClearSecondaryLists();
      LoadContacts(cntctFldrName,iMaxContactsToRead);
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void AddForbiddenField(string FieldName, int Occurances)
    {
      MyField fld = new MyField();
      fld.SetField(FieldName,Occurances);
      m_ForbiddenFieldList.Add(fld);
    }

    //////////////////////////////////////
    // Add a new contact outlook.
    //////////////////////////////////////
    private void DoActionNewContact()
    {
      txtLog.AppendText("Add Contact\r\n");
      frmNewContact nc = new frmNewContact();
      nc.FormBorderStyle = FormBorderStyle.FixedToolWindow;
      nc.StartPosition = FormStartPosition.CenterParent;
      nc.ShowDialog();
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void DoActionViewFields()
    {
      int ForbiddenFieldsFound = 0;
      String sWarning = "";

      txtLog.AppendText("View Fields\r\n");
      frmFields form = new frmFields();
      form.FormBorderStyle = FormBorderStyle.FixedToolWindow;
      form.StartPosition = FormStartPosition.CenterParent;
      form.m_FieldList.Clear();
      foreach (MyField fld in m_FieldList)
      {
        form.m_FieldList.Add(fld);
        if (fld.Occurances>0 && IsForbiddenField(fld.FieldName))
        {
          txtLog.AppendText("Forbidden field \"" + fld.FieldName + "\" contains data\r\n");
          // If there is already an item in the warning string, append a comma
          if (sWarning.Length > 0)
            sWarning += ", ";
          sWarning += fld.FieldName;
          ForbiddenFieldsFound++;
        }
      }
      form.m_ContactList_Count = m_ContactList.Count;
      if (ForbiddenFieldsFound>0)
        MessageBox.Show(ForbiddenFieldsFound + " Forbidden Fields contain data (" + sWarning + ")");
      form.ShowDialog();
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void DoActionContactDetails(int WhichNdx)
    {
      txtLog.AppendText("View Fields\r\n");
      frmContactDetails form = new frmContactDetails();
      //form.FormBorderStyle = frmContactDetails.FixedToolWindow;
      //form.StartPosition = frmContactDetails.CenterParent;
      //form.m_FieldList.Clear();
      //foreach (MyField fld in m_FieldList)
      //{
      //  form.m_FieldList.Add(fld);
      //}
      //form.m_ContactList_Count = m_ContactList.Count;
      //form.myContactBindingSource = m_ContactList;

      form.m_MainForm = this;
      form.ShowDialog();
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private OutLook._ContactItem FindContactItem(string keyEntryID, /*MyContact contact, */OutLook.MAPIFolder folder)
    {
      object missing = System.Reflection.Missing.Value;
      foreach (OutLook._ContactItem OutlookContact in folder.Items)
      {
        //OutLook.UserProperty Prop = OutlookContact.UserProperties.Find("myPetName", missing);
        //OutLook.OutlookNdx Prop = OutlookContact.OutlookNdx.Find("OutlookNdx", missing);
        //if (Prop != null)
        //{
        //  if (Prop.Value.Equals(key))
        //    return OutlookContact;
        //}
        if (OutlookContact.EntryID == keyEntryID)
          return OutlookContact;
      }
      return null;
    }


    //////////////////////////////////////
    //
    //////////////////////////////////////
    private Microsoft.Office.Interop.Outlook.MAPIFolder GetCustomFolder(String argFolder)
    {
      OutLook._Application outlookApplication = new OutLook.Application();
      OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);

      if (argFolder == "Default")
      {
        return contactsFolder;
      }
      else
      {
        foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
        {
          if (subFolder.Name == argFolder)
          {
              return subFolder;
          }
        }
      }
      return null;
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private bool CheckCustomFolderExists(String argFolder)
    {
      OutLook._Application outlookApplication = new OutLook.Application();
      OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);

      // Verifying the mycontacts (custom) sub folder in contacts folder in Outlook.
      foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
      {
        if (subFolder.Name == argFolder)
        {
          //m_CustomFolder = subFolder;
          return true;
        }
      }
      return false;
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private Microsoft.Office.Interop.Outlook.MAPIFolder CreateCustomFolder(String argFolder)
    {
      OutLook._Application outlookApplication = new OutLook.Application();
      OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)outlookApplication.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
      Microsoft.Office.Interop.Outlook.MAPIFolder tmpCustomFolder = null;

      // Verifying the mycontacts(custom) folder in Outlook.
      tmpCustomFolder = null;
      foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
      {
        if (subFolder.Name == argFolder)
        {
          tmpCustomFolder = subFolder;
          break;
        }
      }
      // If myContacts folder does not exist create a new folder with name myContacts.
      if (tmpCustomFolder == null)
      {
        tmpCustomFolder = contactsFolder.Folders.Add(argFolder, Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
      }
      return tmpCustomFolder;
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private bool SaveContact(MyContact contact)
    {
      try
      {
        string cntctFldrName = string.Empty;
        //MyContact contact = new MyContact();
        Microsoft.Office.Interop.Outlook.MAPIFolder m_CustomFolder = null;
        //Microsoft.Office.Interop.Outlook.Results
        //HRESULT hResult;

        cntctFldrName = cbFolder.SelectedItem.ToString();
        m_CustomFolder = GetCustomFolder(cntctFldrName);
        if (m_CustomFolder == null)
        {
          //CreateCustomFolder(cntctFldrName);
          //MessageBox("Folder not found");
          txtLog.AppendText("Folder not found: \"" +cntctFldrName + "\"\r\n");
          return false;
        }
        OutLook.ContactItem existingContact = (OutLook.ContactItem)FindContactItem( contact.EntryID, /*contact, */ m_CustomFolder);

        // The values we can get from web services or data base or class.
        // We have to assign the values to Outlook contact item object .
        if (existingContact == null)
        {
          txtLog.AppendText("Contact not found: \"" +contact.OutlookNdx.ToString() + "\"\r\n");
          return false;
        }
        contact.UnStoreContact(contact,contact.Ndx,existingContact);
        //if (contact.FirstName                != null) existingContact.FirstName                = contact.FirstName.Trim().ToString();
        //if (contact.LastName                 != null) existingContact.LastName                 = contact.LastName.Trim().ToString();
        //if (contact.Email1Address            != null) existingContact.Email1Address            = contact.Email1Address.Trim().ToString();
        //if (contact.Business2TelephoneNumber != null) existingContact.Business2TelephoneNumber = contact.Business2TelephoneNumber.Trim().ToString();
        //if (contact.BusinessAddress          != null) existingContact.BusinessAddress          = contact.BusinessAddress.Trim().ToString();
        existingContact.Save();
        //existingContact.MAPIOBJECT.
        //if (!hResult)
        //{
        //  txtLog.AppendText("Save Contact " + contact.Ndx + " failed with result " + hResult.ToString() + "\r\n");
        //  return false;
        //}

        //this.Close();
        txtLog.AppendText("Save Contact " + contact.Ndx + " OK\r\n");
        return true;
      }
      //catch (Exception ex)
      catch (System.Exception ex)
      //catch (Microsoft.Office.Interop.Outlook.Exception ex)
      //catch (Outlook.Exception ex)
      {
        //Debug.WriteLine(ex.Message);
        txtLog.AppendText("Save Contact " + contact.Ndx + " Failed with \"" + ex.Message + "\"\r\n");
        //string errorInfo = (string)ex.Message.Substring(0, 11);
        //if (errorInfo == "Cannot save")
        //{
        //  MessageBox.Show(@"Create Folder C:\TestFileSave");
        //}
        return false;
      }
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void DoActionSaveContacts()
    {
      //try
      {
        txtLog.AppendText("Save Contacts...\r\n");

        this.Cursor = Cursors.WaitCursor;

        m_NumberOfModifiedContacts = 0;

        foreach (MyContact contact1 in m_ContactList)
        {
          // Ndx starts at 1
          MyContact contact2 = m_ContactListOrig[contact1.Ndx-1];
          if (!contact1.IsIdentical(contact1, contact2, MyContact.COMPARING_SAME_ITEMS))
          {
            m_NumberOfModifiedContacts++;
            if (SaveContact(contact1))
            {
              // Deep Copy from m_ContactList to m_ContactListOrig
              //contact2 = contact1;
              //MyContact x = MyContact.Copy(contact);
              //m_ContactListOrig.Add(x);
              contact2 = MyContact.Copy(contact1);
            }
            else
            {
              // Oops. Dunno why. Let's try to keep going...
            }
            //txtLog.AppendText("Duplicate: " + contact1.Ndx.ToString() + " and " + contact2.Ndx.ToString() + "\r\n");
          }
        }

        //txtNumberOfModifiedContacts.Text        = m_NumberOfModifiedContacts.ToString();
        txtLog.AppendText("Modified Contacts: " + m_NumberOfModifiedContacts.ToString() + "\r\n");

        this.Cursor = Cursors.Arrow;
        txtLog.AppendText("Save Complete\r\n");
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private static int CompareStrings (string x, string y)
    {
      if ((x == null) && (y == null))
      {
        // If x is null and y is null, they're equal.
        return 0;
      }
      else if ((x == null) && (y != null))
      {
        // If x is null and y is not null, y is greater.
        return -1;
      }
      else if ((x != null) && (y == null))
      {
        // If x is not null and y is null, x is greater.
        return 1;
      }
      else
      {
        // if x is not null and y is not null, sort them with ordinary string comparison.
        return x.CompareTo(y);
        /*
        // if x is not null and y is not null, compare the lengths of the two strings.
        int retval = x.Length.CompareTo(y.Length);
        if (retval != 0)
        {
          // If the strings are not of equal length, the longer string is greater.
          return retval;
        }
        else
        {
          // If the strings are of equal length, sort them with ordinary string comparison.
          return x.CompareTo(y);
        }
        */
      }
    }

#if false
    //////////////////////////////////////
    // Based upon...
    //    public delegate int Comparison<MyContact> (MyContact x, MyContact y)
    //////////////////////////////////////
    private static int CompareContactsByName (MyContact x, MyContact y)
    //int IComparer<OutlookContactsAnalyser.MyContact> argSortFunction (MyContact x, MyContact y)
    //int IComparer<OutlookContactsAnalyser.MyContact>.Compare argSortFunction (MyContact x, MyContact y)
    {
      int rc;

      rc = CompareStrings (x.LastName, y.LastName);
      if (rc != 0)
        return rc; // If the strings are not the same, return the result of the compare

      rc = CompareStrings (x.FirstName, y.FirstName);
      if (rc != 0)
        return rc; // If the strings are not the same, return the result of the compare

      rc = CompareStrings (x.CompanyName, y.CompanyName);
      if (rc != 0)
        return rc; // If the strings are not the same, return the result of the compare

      // If everything above are identical, sort by Ndx
      return x.Ndx.CompareTo(y.Ndx);
    }
#endif

#if false
    //public class SamplesArrayList
    //{
      public class myReverserClass : IComparer
      {
        // Calls CaseInsensitiveComparer.Compare with the parameters reversed.
        int IComparer.Compare( Object x, Object y )
        {
          return( (new CaseInsensitiveComparer()).Compare( y, x ) );
        }
      }
    //}
#endif

    public class myCompareContactsByNameClass : IComparer<MyContact>
    {
      int IComparer<MyContact>.Compare( MyContact x, MyContact y )
      {
        int rc;

        rc = CompareStrings (x.LastName, y.LastName);
        if (rc != 0)
          return rc; // If the strings are not the same, return the result of the compare

        rc = CompareStrings (x.FirstName, y.FirstName);
        if (rc != 0)
          return rc; // If the strings are not the same, return the result of the compare

        rc = CompareStrings (x.CompanyName, y.CompanyName);
        if (rc != 0)
          return rc; // If the strings are not the same, return the result of the compare

        // If everything above are identical, sort by Ndx
        return x.Ndx.CompareTo(y.Ndx);
      }
    }
    
    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void DoActionCleanNotesFields()
    {
      int Notes_fixed = 0;
      foreach (MyContact contact1 in m_ContactList)
      {
        string oldstr = contact1.Body;
        string newstr;

        newstr = contact1.CleanField(oldstr);
        if (newstr != oldstr)
          Notes_fixed = Notes_fixed + 1;
        contact1.Body = newstr;

      }
      txtLog.AppendText("Notes fields cleaned = " + Notes_fixed + "\r\n");
    }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private void ShowView(string argTitle, List<MyContact> argWhichList, IComparer<MyContact> argSortFunction)
    //Error	1	The best overloaded method match for 'System.Collections.Generic.List<OutlookContactsAnalyser.MyContact>.Sort(System.Collections.Generic.IComparer<OutlookContactsAnalyser.MyContact>)' has some invalid arguments	C:\home\dev\Outlook_DedupContacts\cs_2008\MainForm.cs	1934	10	OutlookContactsAnalyser
    {
      txtLog.AppendText(argTitle);
      txtLog.AppendText("\r\n");
      
      this.Cursor = Cursors.WaitCursor;
      if (argSortFunction != null)
      {
         argWhichList.Sort(argSortFunction);
         //eg m_ContactListTruncatedNotes.Sort(CompareContactsByName);
      }

      dgvContacts.DataSource = argWhichList;

      // (Attempt to get the Grid to do the sort - but I didn't try too hard)
      //DataGridViewColumn sort_col;
      //foreach (DataGridViewColumn col in dgvContacts.Column)
      //{
      //  if (col.Name == "Surname")
      //    sort_col = col;
      //}
      //dgvContacts.Sort(col,0);

      this.Cursor = Cursors.Arrow;
    }

    #endregion // LocalFunctions

    // #########################################################################
    // # region EventHandlerFunctions
    // #########################################################################
    #region EventHandlerFunctions

    private void btnRefreshFolders_Click(object sender, EventArgs e)
    {
      LoadContactFoldersCombo();
    }

    private void btnGetContacts_Click(object sender, EventArgs e)
    {
      DoActionGetContacts();
    }

    private void btnAnalyse_Click(object sender, EventArgs e)
    {
      AnalyseContacts();
    }

    private void btnViewFields_Click(object sender, EventArgs e)
    {
      DoActionViewFields();
    }

    private void btnContactDetails_Click(object sender, EventArgs e)
    {
      DoActionContactDetails(0);
    }

    #endregion // EventHandlerFunctions

    // #########################################################################
    // # region MenuHandlerFunctions
    // #########################################################################
    #region MenuHandlerFunctions

    // File
    private void fileRefreshFoldersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      LoadContactFoldersCombo();
    }

    private void fileLoadContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      txtLog.AppendText("Get Contacts\r\n");
      DoActionGetContacts();
    }

    private void fileSaveChangedContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionSaveContacts();
    }

    private void fileStopGetContactsThreadToolStripMenuItem_Click(object sender, EventArgs e)
    {
      LoadContactsThread_Stop();
    }

    // View
    private void viewShowAllContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      ShowView("View All Contacts", m_ContactList, null);
    }

    private void viewShowDuplicateContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View Duplicate Contacts", m_ContactListDuplicates, myComparer);
    }

    private void viewShowSimilarContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View Similar Contacts", m_ContactListSimilar, myComparer);
    }

    private void viewShowVaguelySimilarContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View Vaguely Similar Contacts",m_ContactListVaguelySimilar, myComparer);
    }

    private void viewShowWithTheSameNameToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View SameName Contacts", m_ContactListSameName, myComparer);
    }

    private void viewShowWithTheSameNameAndSuffixToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View SameName Contacts", m_ContactListSameNameAndSuffix, myComparer);
    }

    private void viewShowWithTruncatedNotesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View TruncatedNotes Contacts", m_ContactListTruncatedNotes, myComparer);
    }

    private void viewShowWithDislikedContentsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View DislikedContent Contacts", m_ContactListDislikedContent, myComparer);
    }

    private void viewShowWithDuplicateAddressesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View DuplicateAddresses Contacts", m_ContactListDuplicateAddresses, myComparer);
    }

    private void viewShowWithDuplicatePhoneNumbersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View PhoneNumbers Contacts", m_ContactListDuplicatePhoneNumbers, myComparer);
    }

    private void viewShowWithDuplicateEmailAddressToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View EmailAddress Contacts", m_ContactListDuplicateEmailAddress, myComparer);
    }

    private void viewHideUnusedColumnsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      HideUnusedColumns();
    }

    private void viewShowUnlinkedMailingAddressesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View unlinkedMailingAddresses Contacts", m_ContactListUnlinkedMailingAddresses, myComparer);
    }

    private void viewShowLinkedMailingAddressesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View linkedMailingAddresses Contacts", m_ContactListLinkedMailingAddresses, myComparer);
    }

    private void viewShowEmptyMailingAddressesToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View EmptyMailingAddresses Contacts", m_ContactListEmptyMailingAddresses, myComparer);
    }

    private void viewShowMismatchedAddressComponentsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View mismatchedAddressComponents Contacts", m_ContactListMismatchedAddressComponents, myComparer);
    }

    //private void viewShowMatchedAddressComponentsToolStripMenuItem_Click(object sender, EventArgs e)
    //{
    //  IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
    //  ShowView("View matchedAddressComponents Contacts", m_ContactListMatchedAddressComponents, myComparer);
    //}

    private void viewShowWithMinimalInfoToolStripMenuItem_Click(object sender, EventArgs e)
    {
      IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      ShowView("View MinimalInfo Contacts", m_ContactListMinimalInfo, myComparer);
    }

    // Tools
    private void toolsAnalyseContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      AnalyseContacts();
    }

    private void toolsEditContactDetailsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionContactDetails(0);
    }

    private void toolsFieldsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionViewFields();
    }

    private void toolsNewContactToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionNewContact();
    }

    private void cleanNotesFieldsToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionCleanNotesFields();
    }

    private void markDuplicatesForDeletionToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionMarkDuplicatesForDeletion();
    }

    private void fixPhoneNumbersToolStripMenuItem_Click(object sender, EventArgs e)
    {
      DoActionFixPhoneNumbers();
    }

    private void mergeDuplicateContactsToolStripMenuItem_Click(object sender, EventArgs e)
    {
       DoActionMergeDuplicateContacts();
    }

    #endregion // MenuHandlerFunctions

    private void dgvContacts_MouseDoubleClick(object sender, MouseEventArgs e)
    {
       //int x = dgvContacts.SelectedRows.Count;
       //if (e.
       //dgvContacts.CurrentRow.Cells.ToString();
       //dgvContacts.CurrentCell.ToString();
    }

    private void dgvContacts_KeyUp(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.F4)
      {
        MyContact x = (MyContact)dgvContacts.CurrentRow.DataBoundItem;
        if (x != null)
        {
          //MessageBox.Show("Ndx = " + x.Ndx.ToString());
          DoActionContactDetails(x.Ndx);
        }

      }
      else if (e.KeyCode == Keys.F5)
        MyClipboard = dgvContacts.CurrentCell.Value.ToString();
      else if (e.KeyCode == Keys.F6)
        dgvContacts.CurrentCell.Value = MyClipboard.ToString();
    }

    private void radioView_SelectView(object sender, EventArgs e)
    {
      RadioButton radiobtn = (RadioButton)sender;

      if (radiobtn.Checked)
      {
        //txtLog.AppendText("SelectView: " + radiobtn.Name + " = Checked\r\n");
      }
      else
      {
        //txtLog.AppendText("SelectView: " + radiobtn.Name + " = Unchecked\r\n");
        return;
      }

      if (radiobtn.Name=="radioViewAll")                                                                      
      {
        ShowView("View All Contacts", m_ContactList, null);
      }
      else if (radiobtn.Name=="radioViewAllByName")                                                                      
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View All Contacts by Name", m_ContactList, myComparer);
      }
      else if (radiobtn.Name=="radioViewDuplicates")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Duplicate Contacts", m_ContactListDuplicates, myComparer);
      }
      else if (radiobtn.Name=="radioViewSimilar")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Similar Contacts", m_ContactListSimilar, myComparer);
      }
      else if (radiobtn.Name=="radioViewVaguelySimilar")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Vaguely Similar Contacts", m_ContactListVaguelySimilar, myComparer);
      }
      else if (radiobtn.Name=="radioViewSameName")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View SameName Contacts", m_ContactListSameName, myComparer);
      }
      else if (radiobtn.Name=="radioViewSameNameAndSuffix")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View SameNameSuffix Contacts", m_ContactListSameNameAndSuffix, myComparer);
      }
      else if (radiobtn.Name=="radioViewTruncatedNotes")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View TruncatedNotes Contacts", m_ContactListTruncatedNotes, myComparer);
      }
      else if (radiobtn.Name=="radioViewDislikedContents")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View DislikedContent Contacts", m_ContactListDislikedContent, myComparer);
      }
      else if (radiobtn.Name=="radioViewDuplicateAddresses")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Contacts with Duplicate Addresses", m_ContactListDuplicateAddresses, myComparer);
      }
      else if (radiobtn.Name=="radioViewDuplicatePhoneNumbers")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Contacts with Duplicate PhoneNumbers", m_ContactListDuplicatePhoneNumbers, myComparer);
      }
      else if (radiobtn.Name=="radioViewDuplicateEmails")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Contacts with Duplicate EmailAddress", m_ContactListDuplicateEmailAddress, myComparer);
      }
      else if (radiobtn.Name=="radioViewUnlinkedMailingAddresses")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View UnlinkedMailingAddresses Contacts", m_ContactListUnlinkedMailingAddresses, myComparer);
      }
      else if (radiobtn.Name=="radioViewLinkedMailingAddresses")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View LinkedMailingAddresses Contacts", m_ContactListLinkedMailingAddresses, myComparer);
      }
      else if (radiobtn.Name=="radioViewEmptyMailingAddresses")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View EmptyMailingAddresses Contacts", m_ContactListEmptyMailingAddresses, myComparer);
      }
      else if (radiobtn.Name=="radioViewMismatchedAddrAndComp")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View MismatchedAddressComponents Contacts", m_ContactListMismatchedAddressComponents, myComparer);
      }
      //else if (radiobtn.Name=="radioViewMatchedAddressComponents")
      //{
      //  IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
      //  ShowView("View MatchedAddressComponents Contacts", m_ContactListMatchedAddressComponents, myComparer);
      //}
      else if (radiobtn.Name=="radioViewMinimalInfo")
      {
        IComparer<MyContact> myComparer = new myCompareContactsByNameClass();
        ShowView("View Contacts with Minimal Info", m_ContactListMinimalInfo, myComparer);
      }
      else
         MessageBox.Show(radiobtn.Name);
    }

  } // class MainForm

} // namespace OutlookContactsAnalyser

