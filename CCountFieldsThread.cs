using System;
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
  public class CCountFieldsThread
  {
    #region MemberVariables
    private int _Account                      ; public int Account                      { get { return _Account                     ; } set { _Account                      = value; } }
    private int _Anniversary                  ; public int Anniversary                  { get { return _Anniversary                 ; } set { _Anniversary                  = value; } }
    private int _AssistantName                ; public int AssistantName                { get { return _AssistantName               ; } set { _AssistantName                = value; } }
    private int _AssistantTelephoneNumber     ; public int AssistantTelephoneNumber     { get { return _AssistantTelephoneNumber    ; } set { _AssistantTelephoneNumber     = value; } }
    private int _BillingInformation           ; public int BillingInformation           { get { return _BillingInformation          ; } set { _BillingInformation           = value; } }
    private int _Birthday                     ; public int Birthday                     { get { return _Birthday                    ; } set { _Birthday                     = value; } }
    private int _Body                         ; public int Body                         { get { return _Body                        ; } set { _Body                         = value; } }
    private int _Business2TelephoneNumber     ; public int Business2TelephoneNumber     { get { return _Business2TelephoneNumber    ; } set { _Business2TelephoneNumber     = value; } }
    private int _BusinessAddress              ; public int BusinessAddress              { get { return _BusinessAddress             ; } set { _BusinessAddress              = value; } }
    private int _BusinessAddressCity          ; public int BusinessAddressCity          { get { return _BusinessAddressCity         ; } set { _BusinessAddressCity          = value; } }
    private int _BusinessAddressCountry       ; public int BusinessAddressCountry       { get { return _BusinessAddressCountry      ; } set { _BusinessAddressCountry       = value; } }
    private int _BusinessAddressPostalCode    ; public int BusinessAddressPostalCode    { get { return _BusinessAddressPostalCode   ; } set { _BusinessAddressPostalCode    = value; } }
    private int _BusinessAddressPostOfficeBox ; public int BusinessAddressPostOfficeBox { get { return _BusinessAddressPostOfficeBox; } set { _BusinessAddressPostOfficeBox = value; } }
    private int _BusinessAddressState         ; public int BusinessAddressState         { get { return _BusinessAddressState        ; } set { _BusinessAddressState         = value; } }
    private int _BusinessAddressStreet        ; public int BusinessAddressStreet        { get { return _BusinessAddressStreet       ; } set { _BusinessAddressStreet        = value; } }
    private int _BusinessFaxNumber            ; public int BusinessFaxNumber            { get { return _BusinessFaxNumber           ; } set { _BusinessFaxNumber            = value; } }
    private int _BusinessHomePage             ; public int BusinessHomePage             { get { return _BusinessHomePage            ; } set { _BusinessHomePage             = value; } }
    private int _BusinessTelephoneNumber      ; public int BusinessTelephoneNumber      { get { return _BusinessTelephoneNumber     ; } set { _BusinessTelephoneNumber      = value; } }
    private int _CallbackTelephoneNumber      ; public int CallbackTelephoneNumber      { get { return _CallbackTelephoneNumber     ; } set { _CallbackTelephoneNumber      = value; } }
    private int _CarTelephoneNumber           ; public int CarTelephoneNumber           { get { return _CarTelephoneNumber          ; } set { _CarTelephoneNumber           = value; } }
    private int _Categories                   ; public int Categories                   { get { return _Categories                  ; } set { _Categories                   = value; } }
    private int _Children                     ; public int Children                     { get { return _Children                    ; } set { _Children                     = value; } }
    private int _Companies                    ; public int Companies                    { get { return _Companies                   ; } set { _Companies                    = value; } }
    private int _CompanyAndFullName           ; public int CompanyAndFullName           { get { return _CompanyAndFullName          ; } set { _CompanyAndFullName           = value; } }
    private int _CompanyMainTelephoneNumber   ; public int CompanyMainTelephoneNumber   { get { return _CompanyMainTelephoneNumber  ; } set { _CompanyMainTelephoneNumber   = value; } }
    private int _CompanyName                  ; public int CompanyName                  { get { return _CompanyName                 ; } set { _CompanyName                  = value; } }
    private int _ComputerNetworkName          ; public int ComputerNetworkName          { get { return _ComputerNetworkName         ; } set { _ComputerNetworkName          = value; } }
    private int _CreationTime                 ; public int CreationTime                 { get { return _CreationTime                ; } set { _CreationTime                 = value; } }
    private int _CustomerID                   ; public int CustomerID                   { get { return _CustomerID                  ; } set { _CustomerID                   = value; } }
    private int _Department                   ; public int Department                   { get { return _Department                  ; } set { _Department                   = value; } }
    private int _Email1Address                ; public int Email1Address                { get { return _Email1Address               ; } set { _Email1Address                = value; } }
    private int _Email1AddressType            ; public int Email1AddressType            { get { return _Email1AddressType           ; } set { _Email1AddressType            = value; } }
    private int _Email1DisplayName            ; public int Email1DisplayName            { get { return _Email1DisplayName           ; } set { _Email1DisplayName            = value; } }
    private int _Email2Address                ; public int Email2Address                { get { return _Email2Address               ; } set { _Email2Address                = value; } }
    private int _Email2AddressType            ; public int Email2AddressType            { get { return _Email2AddressType           ; } set { _Email2AddressType            = value; } }
    private int _Email2DisplayName            ; public int Email2DisplayName            { get { return _Email2DisplayName           ; } set { _Email2DisplayName            = value; } }
    private int _Email3Address                ; public int Email3Address                { get { return _Email3Address               ; } set { _Email3Address                = value; } }
    private int _Email3AddressType            ; public int Email3AddressType            { get { return _Email3AddressType           ; } set { _Email3AddressType            = value; } }
    private int _Email3DisplayName            ; public int Email3DisplayName            { get { return _Email3DisplayName           ; } set { _Email3DisplayName            = value; } }
    private int _EntryID                      ; public int EntryID                      { get { return _EntryID                     ; } set { _EntryID                      = value; } }
    private int _FileAs                       ; public int FileAs                       { get { return _FileAs                      ; } set { _FileAs                       = value; } }
    private int _FirstName                    ; public int FirstName                    { get { return _FirstName                   ; } set { _FirstName                    = value; } }
    private int _FTPSite                      ; public int FTPSite                      { get { return _FTPSite                     ; } set { _FTPSite                      = value; } }
    private int _FullName                     ; public int FullName                     { get { return _FullName                    ; } set { _FullName                     = value; } }
    private int _FullNameAndCompany           ; public int FullNameAndCompany           { get { return _FullNameAndCompany          ; } set { _FullNameAndCompany           = value; } }
    private int _GovernmentIDNumber           ; public int GovernmentIDNumber           { get { return _GovernmentIDNumber          ; } set { _GovernmentIDNumber           = value; } }
    private int _Hobby                        ; public int Hobby                        { get { return _Hobby                       ; } set { _Hobby                        = value; } }
    private int _Home2TelephoneNumber         ; public int Home2TelephoneNumber         { get { return _Home2TelephoneNumber        ; } set { _Home2TelephoneNumber         = value; } }
    private int _HomeAddress                  ; public int HomeAddress                  { get { return _HomeAddress                 ; } set { _HomeAddress                  = value; } }
    private int _HomeAddressCity              ; public int HomeAddressCity              { get { return _HomeAddressCity             ; } set { _HomeAddressCity              = value; } }
    private int _HomeAddressCountry           ; public int HomeAddressCountry           { get { return _HomeAddressCountry          ; } set { _HomeAddressCountry           = value; } }
    private int _HomeAddressPostalCode        ; public int HomeAddressPostalCode        { get { return _HomeAddressPostalCode       ; } set { _HomeAddressPostalCode        = value; } }
    private int _HomeAddressPostOfficeBox     ; public int HomeAddressPostOfficeBox     { get { return _HomeAddressPostOfficeBox    ; } set { _HomeAddressPostOfficeBox     = value; } }
    private int _HomeAddressState             ; public int HomeAddressState             { get { return _HomeAddressState            ; } set { _HomeAddressState             = value; } }
    private int _HomeAddressStreet            ; public int HomeAddressStreet            { get { return _HomeAddressStreet           ; } set { _HomeAddressStreet            = value; } }
    private int _HomeFaxNumber                ; public int HomeFaxNumber                { get { return _HomeFaxNumber               ; } set { _HomeFaxNumber                = value; } }
    private int _HomeTelephoneNumber          ; public int HomeTelephoneNumber          { get { return _HomeTelephoneNumber         ; } set { _HomeTelephoneNumber          = value; } }
    private int _Initials                     ; public int Initials                     { get { return _Initials                    ; } set { _Initials                     = value; } }
    private int _ISDNNumber                   ; public int ISDNNumber                   { get { return _ISDNNumber                  ; } set { _ISDNNumber                   = value; } }
    private int _JobTitle                     ; public int JobTitle                     { get { return _JobTitle                    ; } set { _JobTitle                     = value; } }
    private int _Journal                      ; public int Journal                      { get { return _Journal                     ; } set { _Journal                      = value; } }
    private int _Language                     ; public int Language                     { get { return _Language                    ; } set { _Language                     = value; } }
    private int _LastModificationTime         ; public int LastModificationTime         { get { return _LastModificationTime        ; } set { _LastModificationTime         = value; } }
    private int _LastName                     ; public int LastName                     { get { return _LastName                    ; } set { _LastName                     = value; } }
    private int _LastNameAndFirstName         ; public int LastNameAndFirstName         { get { return _LastNameAndFirstName        ; } set { _LastNameAndFirstName         = value; } }
    private int _MailingAddress               ; public int MailingAddress               { get { return _MailingAddress              ; } set { _MailingAddress               = value; } }
    private int _MailingAddressCity           ; public int MailingAddressCity           { get { return _MailingAddressCity          ; } set { _MailingAddressCity           = value; } }
    private int _MailingAddressCountry        ; public int MailingAddressCountry        { get { return _MailingAddressCountry       ; } set { _MailingAddressCountry        = value; } }
    private int _MailingAddressPostalCode     ; public int MailingAddressPostalCode     { get { return _MailingAddressPostalCode    ; } set { _MailingAddressPostalCode     = value; } }
    private int _MailingAddressPostOfficeBox  ; public int MailingAddressPostOfficeBox  { get { return _MailingAddressPostOfficeBox ; } set { _MailingAddressPostOfficeBox  = value; } }
    private int _MailingAddressState          ; public int MailingAddressState          { get { return _MailingAddressState         ; } set { _MailingAddressState          = value; } }
    private int _MailingAddressStreet         ; public int MailingAddressStreet         { get { return _MailingAddressStreet        ; } set { _MailingAddressStreet         = value; } }
    private int _ManagerName                  ; public int ManagerName                  { get { return _ManagerName                 ; } set { _ManagerName                  = value; } }
    private int _MessageClass                 ; public int MessageClass                 { get { return _MessageClass                ; } set { _MessageClass                 = value; } }
    private int _MiddleName                   ; public int MiddleName                   { get { return _MiddleName                  ; } set { _MiddleName                   = value; } }
    private int _Mileage                      ; public int Mileage                      { get { return _Mileage                     ; } set { _Mileage                      = value; } }
    private int _MobileTelephoneNumber        ; public int MobileTelephoneNumber        { get { return _MobileTelephoneNumber       ; } set { _MobileTelephoneNumber        = value; } }
    private int _NickName                     ; public int NickName                     { get { return _NickName                    ; } set { _NickName                     = value; } }
    private int _NoAging                      ; public int NoAging                      { get { return _NoAging                     ; } set { _NoAging                      = value; } }
    private int _OfficeLocation               ; public int OfficeLocation               { get { return _OfficeLocation              ; } set { _OfficeLocation               = value; } }
    private int _OrganizationalIDNumber       ; public int OrganizationalIDNumber       { get { return _OrganizationalIDNumber      ; } set { _OrganizationalIDNumber       = value; } }
    private int _OtherAddress                 ; public int OtherAddress                 { get { return _OtherAddress                ; } set { _OtherAddress                 = value; } }
    private int _OtherAddressCity             ; public int OtherAddressCity             { get { return _OtherAddressCity            ; } set { _OtherAddressCity             = value; } }
    private int _OtherAddressCountry          ; public int OtherAddressCountry          { get { return _OtherAddressCountry         ; } set { _OtherAddressCountry          = value; } }
    private int _OtherAddressPostalCode       ; public int OtherAddressPostalCode       { get { return _OtherAddressPostalCode      ; } set { _OtherAddressPostalCode       = value; } }
    private int _OtherAddressPostOfficeBox    ; public int OtherAddressPostOfficeBox    { get { return _OtherAddressPostOfficeBox   ; } set { _OtherAddressPostOfficeBox    = value; } }
    private int _OtherAddressState            ; public int OtherAddressState            { get { return _OtherAddressState           ; } set { _OtherAddressState            = value; } }
    private int _OtherAddressStreet           ; public int OtherAddressStreet           { get { return _OtherAddressStreet          ; } set { _OtherAddressStreet           = value; } }
    private int _OtherFaxNumber               ; public int OtherFaxNumber               { get { return _OtherFaxNumber              ; } set { _OtherFaxNumber               = value; } }
    private int _OtherTelephoneNumber         ; public int OtherTelephoneNumber         { get { return _OtherTelephoneNumber        ; } set { _OtherTelephoneNumber         = value; } }
    private int _OutlookInternalVersion       ; public int OutlookInternalVersion       { get { return _OutlookInternalVersion      ; } set { _OutlookInternalVersion       = value; } }
    private int _OutlookVersion               ; public int OutlookVersion               { get { return _OutlookVersion              ; } set { _OutlookVersion               = value; } }
    private int _PagerNumber                  ; public int PagerNumber                  { get { return _PagerNumber                 ; } set { _PagerNumber                  = value; } }
    private int _PersonalHomePage             ; public int PersonalHomePage             { get { return _PersonalHomePage            ; } set { _PersonalHomePage             = value; } }
    private int _PrimaryTelephoneNumber       ; public int PrimaryTelephoneNumber       { get { return _PrimaryTelephoneNumber      ; } set { _PrimaryTelephoneNumber       = value; } }
    private int _Profession                   ; public int Profession                   { get { return _Profession                  ; } set { _Profession                   = value; } }
    private int _RadioTelephoneNumber         ; public int RadioTelephoneNumber         { get { return _RadioTelephoneNumber        ; } set { _RadioTelephoneNumber         = value; } }
    private int _ReferredBy                   ; public int ReferredBy                   { get { return _ReferredBy                  ; } set { _ReferredBy                   = value; } }
    private int _Saved                        ; public int Saved                        { get { return _Saved                       ; } set { _Saved                        = value; } }
    private int _Size                         ; public int Size                         { get { return _Size                        ; } set { _Size                         = value; } }
    private int _Spouse                       ; public int Spouse                       { get { return _Spouse                      ; } set { _Spouse                       = value; } }
    private int _Subject                      ; public int Subject                      { get { return _Subject                     ; } set { _Subject                      = value; } }
    private int _Suffix                       ; public int Suffix                       { get { return _Suffix                      ; } set { _Suffix                       = value; } }
    private int _TelexNumber                  ; public int TelexNumber                  { get { return _TelexNumber                 ; } set { _TelexNumber                  = value; } }
    private int _Title                        ; public int Title                        { get { return _Title                       ; } set { _Title                        = value; } }
    private int _TTYTDDTelephoneNumber        ; public int TTYTDDTelephoneNumber        { get { return _TTYTDDTelephoneNumber       ; } set { _TTYTDDTelephoneNumber        = value; } }
    private int _UnRead                       ; public int UnRead                       { get { return _UnRead                      ; } set { _UnRead                       = value; } }
    private int _User1                        ; public int User1                        { get { return _User1                       ; } set { _User1                        = value; } }
    private int _User2                        ; public int User2                        { get { return _User2                       ; } set { _User2                        = value; } }
    private int _User3                        ; public int User3                        { get { return _User3                       ; } set { _User3                        = value; } }
    private int _User4                        ; public int User4                        { get { return _User4                       ; } set { _User4                        = value; } }
    private int _UserCertificate              ; public int UserCertificate              { get { return _UserCertificate             ; } set { _UserCertificate              = value; } }
    private int _WebPage                      ; public int WebPage                      { get { return _WebPage                     ; } set { _WebPage                      = value; } }
    private int _YomiCompanyName              ; public int YomiCompanyName              { get { return _YomiCompanyName             ; } set { _YomiCompanyName              = value; } }
    private int _YomiFirstName                ; public int YomiFirstName                { get { return _YomiFirstName               ; } set { _YomiFirstName                = value; } }
    private int _YomiLastName                 ; public int YomiLastName                 { get { return _YomiLastName                ; } set { _YomiLastName                 = value; } }

    private int _Gender                       ; public int Gender                       { get { return _Gender                      ; } set { _Gender                       = value; } }
    private int _Importance                   ; public int Importance                   { get { return _Importance                  ; } set { _Importance                   = value; } }
    private int _SelectedMailingAddress       ; public int SelectedMailingAddress       { get { return _SelectedMailingAddress      ; } set { _SelectedMailingAddress       = value; } }
    private int _Sensitivity                  ; public int Sensitivity                  { get { return _Sensitivity                 ; } set { _Sensitivity                  = value; } }

  //private int _ContactIndex                 ; public int ContactIndex                 { get { return _ContactIndex                ; } set { _ContactIndex                 = value; } }
  //private int _Actions                      ; public int Actions                      { get { return _Actions                     ; } set { _Actions                      = value; } }
  //private int _Application                  ; public int Application                  { get { return _Application                 ; } set { _Application                  = value; } }
  //private int _Attachments                  ; public int Attachments                  { get { return _Attachments                 ; } set { _Attachments                  = value; } }
  //private int _Email1EntryID                ; public int Email1EntryID                { get { return _Email1EntryID               ; } set { _Email1EntryID                = value; } }
  //private int _Email2EntryID                ; public int Email2EntryID                { get { return _Email2EntryID               ; } set { _Email2EntryID                = value; } }
  //private int _Email3EntryID                ; public int Email3EntryID                { get { return _Email3EntryID               ; } set { _Email3EntryID                = value; } }
  //private int _FormDescription              ; public int FormDescription              { get { return _FormDescription             ; } set { _FormDescription              = value; } }
  //private int _GetInspector                 ; public int GetInspector                 { get { return _GetInspector                ; } set { _GetInspector                 = value; } }
  //private int _Parent                       ; public int Parent                       { get { return _Parent                      ; } set { _Parent                       = value; } }
  //private int _UserProperties               ; public int UserProperties               { get { return _UserProperties              ; } set { _UserProperties               = value; } }
    #endregion // MemberVariables

    #region MemberVariables
    ManualResetEvent m_EventStop;    // Main thread sets this event to stop worker thread
    ManualResetEvent m_EventStopped; // Worker thread sets this event when it is stopped
    MainForm m_form;                 // Reference to main form used to make syncronous user interface calls
    String m_folderName;
    int m_MaxContactsToRead;
    #endregion // MemberVariables

    #region Functions

    //////////////////////////////////////
    // Constructor
    //////////////////////////////////////
    public CCountFieldsThread (ManualResetEvent eventStop,
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
        s = "Count fields in Folder \"" + m_folderName + "\"";
        m_form.Invoke(m_form.m_Delegate_CountFieldsThread_AppendToLog, new Object[] {s});

        OutLook.MAPIFolder   fldContacts = null;
        OutLook._Application outlookObj  = new OutLook.Application();

        if (m_folderName == "Default")
        {
          fldContacts = (OutLook.MAPIFolder)outlookObj.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
        }
        else
        {
          OutLook.MAPIFolder contactsFolder = (OutLook.MAPIFolder)
          outlookObj.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);

          // Verify the custom folder in Outlook.
          foreach (OutLook.MAPIFolder subFolder in contactsFolder.Folders)
          {
            if (subFolder.Name == m_folderName)
            {
              fldContacts = subFolder;
              break;
            }
          }
        }

        ClearFieldCount();
    
        // Loop through contact in specified folder.
        ii = 0;
        foreach (Microsoft.Office.Interop.Outlook._ContactItem OutlookContact in fldContacts.Items)
        {
          ii++;
          AccumulateFields(OutlookContact);

          if (ii >= m_MaxContactsToRead)
            break;

          // Make synchronous call to main form.
          // MainForm.AppendToLog function runs in main thread.
          // (To make asynchronous call use BeginInvoke)
          //s = m_folderName + "[" + m_MaxContactsToRead + "] " + ii.ToString() + " read";
          //m_form.Invoke(m_form.m_DelegateLoadContactsThreadAppendToLog, new Object[] {s});
          m_form.Invoke(m_form.m_Delegate_CountFieldsThread_UpdateCounter, new Object[] {ii});

          // check if thread is cancelled
          if ( m_EventStop.WaitOne(0, true) )
          {
            // Inform main thread that this thread stopped
            m_EventStopped.Set();
            return;
          }
        }

        // Make synchronous call to main form
        // to inform it that thread finished
        m_form.Invoke(m_form.m_Delegate_CountFieldsThread_Finished, null);
      }

      //catch (System.Exception ex)
      //{
      //  throw new ApplicationException(ex.Message);
      //}
    }
    
    public void ClearFieldCount()
    {
      Account                      = 0;
      Anniversary                  = 0;
      AssistantName                = 0;
      AssistantTelephoneNumber     = 0;
      BillingInformation           = 0;
      Birthday                     = 0;
      Body                         = 0;
      Business2TelephoneNumber     = 0;
      BusinessAddress              = 0;
      BusinessAddressCity          = 0;
      BusinessAddressCountry       = 0;
      BusinessAddressPostalCode    = 0;
      BusinessAddressPostOfficeBox = 0;
      BusinessAddressState         = 0;
      BusinessAddressStreet        = 0;
      BusinessFaxNumber            = 0;
      BusinessHomePage             = 0;
      BusinessTelephoneNumber      = 0;
      CallbackTelephoneNumber      = 0;
      CarTelephoneNumber           = 0;
      Categories                   = 0;
      Children                     = 0;
      Companies                    = 0;
      CompanyAndFullName           = 0;
      CompanyMainTelephoneNumber   = 0;
      CompanyName                  = 0;
      ComputerNetworkName          = 0;
      CreationTime                 = 0;
      CustomerID                   = 0;
      Department                   = 0;
      Email1Address                = 0;
      Email1AddressType            = 0;
      Email1DisplayName            = 0;
      Email2Address                = 0;
      Email2AddressType            = 0;
      Email2DisplayName            = 0;
      Email3Address                = 0;
      Email3AddressType            = 0;
      Email3DisplayName            = 0;
      EntryID                      = 0;
      FileAs                       = 0;
      FirstName                    = 0;
      FTPSite                      = 0;
      FullName                     = 0;
      FullNameAndCompany           = 0;
      Gender                       = 0;
      GovernmentIDNumber           = 0;
      Hobby                        = 0;
      Home2TelephoneNumber         = 0;
      HomeAddress                  = 0;
      HomeAddressCity              = 0;
      HomeAddressCountry           = 0;
      HomeAddressPostalCode        = 0;
      HomeAddressPostOfficeBox     = 0;
      HomeAddressState             = 0;
      HomeAddressStreet            = 0;
      HomeFaxNumber                = 0;
      HomeTelephoneNumber          = 0;
      Importance                   = 0;
      Initials                     = 0;
      ISDNNumber                   = 0;
      JobTitle                     = 0;
      Journal                      = 0;
      Language                     = 0;
      LastModificationTime         = 0;
      LastName                     = 0;
      LastNameAndFirstName         = 0;
      MailingAddress               = 0;
      MailingAddressCity           = 0;
      MailingAddressCountry        = 0;
      MailingAddressPostalCode     = 0;
      MailingAddressPostOfficeBox  = 0;
      MailingAddressState          = 0;
      MailingAddressStreet         = 0;
      ManagerName                  = 0;
      MessageClass                 = 0;
      MiddleName                   = 0;
      Mileage                      = 0;
      MobileTelephoneNumber        = 0;
      NickName                     = 0;
      NoAging                      = 0;
      OfficeLocation               = 0;
      OrganizationalIDNumber       = 0;
      OtherAddress                 = 0;
      OtherAddressCity             = 0;
      OtherAddressCountry          = 0;
      OtherAddressPostalCode       = 0;
      OtherAddressPostOfficeBox    = 0;
      OtherAddressState            = 0;
      OtherAddressStreet           = 0;
      OtherFaxNumber               = 0;
      OtherTelephoneNumber         = 0;
      OutlookInternalVersion       = 0;
      OutlookVersion               = 0;
      PagerNumber                  = 0;
      PersonalHomePage             = 0;
      PrimaryTelephoneNumber       = 0;
      Profession                   = 0;
      RadioTelephoneNumber         = 0;
      ReferredBy                   = 0;
      Saved                        = 0;
      SelectedMailingAddress       = 0;
      Sensitivity                  = 0;
      Size                         = 0;
      Spouse                       = 0;
      Subject                      = 0;
      Suffix                       = 0;
      TelexNumber                  = 0;
      Title                        = 0;
      TTYTDDTelephoneNumber        = 0;
      UnRead                       = 0;
      User1                        = 0;
      User2                        = 0;
      User3                        = 0;
      User4                        = 0;
      UserCertificate              = 0;
      WebPage                      = 0;
      YomiCompanyName              = 0;
      YomiFirstName                = 0;
      YomiLastName                 = 0;
    //ContactIndex                 = 0;
    //Actions                      = 0;
    //Application                  = 0;
    //Attachments                  = 0;
    //Email1EntryID                = 0;
    //Email2EntryID                = 0;
    //Email3EntryID                = 0;
    //FormDescription              = 0;
    //GetInspector                 = 0;
    //Parent                       = 0;
    //UserProperties               = 0;
    }
    
    public void AccumulateFields(Microsoft.Office.Interop.Outlook._ContactItem OutlookContact)
    {
      if (OutlookContact.Account                      != "") Account                      ++;
      if (OutlookContact.Anniversary                  != MyFields.EmptyDate) Anniversary  ++;
      if (OutlookContact.AssistantName                != "") AssistantName                ++;
      if (OutlookContact.AssistantTelephoneNumber     != "") AssistantTelephoneNumber     ++;
      if (OutlookContact.BillingInformation           != "") BillingInformation           ++;
      if (OutlookContact.Birthday                     != MyFields.EmptyDate) Birthday     ++;
      if (OutlookContact.Body                         != "") Body                         ++;
      if (OutlookContact.Business2TelephoneNumber     != "") Business2TelephoneNumber     ++;
      if (OutlookContact.BusinessAddress              != "") BusinessAddress              ++;
      if (OutlookContact.BusinessAddressCity          != "") BusinessAddressCity          ++;
      if (OutlookContact.BusinessAddressCountry       != "") BusinessAddressCountry       ++;
      if (OutlookContact.BusinessAddressPostalCode    != "") BusinessAddressPostalCode    ++;
      if (OutlookContact.BusinessAddressPostOfficeBox != "") BusinessAddressPostOfficeBox ++;
      if (OutlookContact.BusinessAddressState         != "") BusinessAddressState         ++;
      if (OutlookContact.BusinessAddressStreet        != "") BusinessAddressStreet        ++;
      if (OutlookContact.BusinessFaxNumber            != "") BusinessFaxNumber            ++;
      if (OutlookContact.BusinessHomePage             != "") BusinessHomePage             ++;
      if (OutlookContact.BusinessTelephoneNumber      != "") BusinessTelephoneNumber      ++;
      if (OutlookContact.CallbackTelephoneNumber      != "") CallbackTelephoneNumber      ++;
      if (OutlookContact.CarTelephoneNumber           != "") CarTelephoneNumber           ++;
      if (OutlookContact.Categories                   != "") Categories                   ++;
      if (OutlookContact.Children                     != "") Children                     ++;
      if (OutlookContact.Companies                    != "") Companies                    ++;
      if (OutlookContact.CompanyAndFullName           != "") CompanyAndFullName           ++;
      if (OutlookContact.CompanyMainTelephoneNumber   != "") CompanyMainTelephoneNumber   ++;
      if (OutlookContact.CompanyName                  != "") CompanyName                  ++;
      if (OutlookContact.ComputerNetworkName          != "") ComputerNetworkName          ++;
      if (OutlookContact.CreationTime                 != MyFields.EmptyDate) CreationTime ++;
      if (OutlookContact.CustomerID                   != "") CustomerID                   ++;
      if (OutlookContact.Department                   != "") Department                   ++;
      if (OutlookContact.Email1Address                != "") Email1Address                ++;
      if (OutlookContact.Email1AddressType            != "") Email1AddressType            ++;
      if (OutlookContact.Email1DisplayName            != "") Email1DisplayName            ++;
      if (OutlookContact.Email2Address                != "") Email2Address                ++;
      if (OutlookContact.Email2AddressType            != "") Email2AddressType            ++;
      if (OutlookContact.Email2DisplayName            != "") Email2DisplayName            ++;
      if (OutlookContact.Email3Address                != "") Email3Address                ++;
      if (OutlookContact.Email3AddressType            != "") Email3AddressType            ++;
      if (OutlookContact.Email3DisplayName            != "") Email3DisplayName            ++;
      if (OutlookContact.EntryID                      != "") EntryID                      ++;
      if (OutlookContact.FileAs                       != "") FileAs                       ++;
      if (OutlookContact.FirstName                    != "") FirstName                    ++;
      if (OutlookContact.FTPSite                      != "") FTPSite                      ++;
      if (OutlookContact.FullName                     != "") FullName                     ++;
      if (OutlookContact.FullNameAndCompany           != "") FullNameAndCompany           ++;
      if (OutlookContact.Gender                       != MyFields.EmptyGender) Gender     ++;
      if (OutlookContact.GovernmentIDNumber           != "") GovernmentIDNumber           ++;
      if (OutlookContact.Hobby                        != "") Hobby                        ++;
      if (OutlookContact.Home2TelephoneNumber         != "") Home2TelephoneNumber         ++;
      if (OutlookContact.HomeAddress                  != "") HomeAddress                  ++;
      if (OutlookContact.HomeAddressCity              != "") HomeAddressCity              ++;
      if (OutlookContact.HomeAddressCountry           != "") HomeAddressCountry           ++;
      if (OutlookContact.HomeAddressPostalCode        != "") HomeAddressPostalCode        ++;
      if (OutlookContact.HomeAddressPostOfficeBox     != "") HomeAddressPostOfficeBox     ++;
      if (OutlookContact.HomeAddressState             != "") HomeAddressState             ++;
      if (OutlookContact.HomeAddressStreet            != "") HomeAddressStreet            ++;
      if (OutlookContact.HomeFaxNumber                != "") HomeFaxNumber                ++;
      if (OutlookContact.HomeTelephoneNumber          != "") HomeTelephoneNumber          ++;
      if (OutlookContact.Importance                   != MyFields.EmptyImportance) Importance++;
      if (OutlookContact.Initials                     != "") Initials                     ++;
      if (OutlookContact.ISDNNumber                   != "") ISDNNumber                   ++;
      if (OutlookContact.JobTitle                     != "") JobTitle                     ++;
      if (OutlookContact.Journal                      != MyFields.EmptyJournal) Journal   ++;
      if (OutlookContact.Language                     != "") Language                     ++;
      if (OutlookContact.LastModificationTime         != MyFields.EmptyDate) LastModificationTime++;
      if (OutlookContact.LastName                     != "") LastName                     ++;
      if (OutlookContact.LastNameAndFirstName         != "") LastNameAndFirstName         ++;
      if (OutlookContact.MailingAddress               != "") MailingAddress               ++;
      if (OutlookContact.MailingAddressCity           != "") MailingAddressCity           ++;
      if (OutlookContact.MailingAddressCountry        != "") MailingAddressCountry        ++;
      if (OutlookContact.MailingAddressPostalCode     != "") MailingAddressPostalCode     ++;
      if (OutlookContact.MailingAddressPostOfficeBox  != "") MailingAddressPostOfficeBox  ++;
      if (OutlookContact.MailingAddressState          != "") MailingAddressState          ++;
      if (OutlookContact.MailingAddressStreet         != "") MailingAddressStreet         ++;
      if (OutlookContact.ManagerName                  != "") ManagerName                  ++;
      if (OutlookContact.MessageClass                 != "") MessageClass                 ++;
      if (OutlookContact.MiddleName                   != "") MiddleName                   ++;
      if (OutlookContact.Mileage                      != "") Mileage                      ++;
      if (OutlookContact.MobileTelephoneNumber        != "") MobileTelephoneNumber        ++;
      if (OutlookContact.NickName                     != "") NickName                     ++;
      if (OutlookContact.NoAging                      != MyFields.EmptyNoAging) NoAging   ++;
      if (OutlookContact.OfficeLocation               != "") OfficeLocation               ++;
      if (OutlookContact.OrganizationalIDNumber       != "") OrganizationalIDNumber       ++;
      if (OutlookContact.OtherAddress                 != "") OtherAddress                 ++;
      if (OutlookContact.OtherAddressCity             != "") OtherAddressCity             ++;
      if (OutlookContact.OtherAddressCountry          != "") OtherAddressCountry          ++;
      if (OutlookContact.OtherAddressPostalCode       != "") OtherAddressPostalCode       ++;
      if (OutlookContact.OtherAddressPostOfficeBox    != "") OtherAddressPostOfficeBox    ++;
      if (OutlookContact.OtherAddressState            != "") OtherAddressState            ++;
      if (OutlookContact.OtherAddressStreet           != "") OtherAddressStreet           ++;
      if (OutlookContact.OtherFaxNumber               != "") OtherFaxNumber               ++;
      if (OutlookContact.OtherTelephoneNumber         != "") OtherTelephoneNumber         ++;
      if (OutlookContact.OutlookInternalVersion       > 0  ) OutlookInternalVersion       ++;
      if (OutlookContact.OutlookVersion               != "") OutlookVersion               ++;
      if (OutlookContact.PagerNumber                  != "") PagerNumber                  ++;
      if (OutlookContact.PersonalHomePage             != "") PersonalHomePage             ++;
      if (OutlookContact.PrimaryTelephoneNumber       != "") PrimaryTelephoneNumber       ++;
      if (OutlookContact.Profession                   != "") Profession                   ++;
      if (OutlookContact.RadioTelephoneNumber         != "") RadioTelephoneNumber         ++;
      if (OutlookContact.ReferredBy                   != "") ReferredBy                   ++;
      if (OutlookContact.Saved                        != MyFields.EmptySaved) Saved       ++;
      if (OutlookContact.SelectedMailingAddress       != MyFields.EmptyMailingAddress) SelectedMailingAddress++;
      if (OutlookContact.Sensitivity                  != MyFields.EmptySensitivity) Sensitivity++;
      if (OutlookContact.Size                         != 0 ) Size                         ++;
      if (OutlookContact.Spouse                       != "") Spouse                       ++;
      if (OutlookContact.Subject                      != "") Subject                      ++;
      if (OutlookContact.Suffix                       != "") Suffix                       ++;
      if (OutlookContact.TelexNumber                  != "") TelexNumber                  ++;
      if (OutlookContact.Title                        != "") Title                        ++;
      if (OutlookContact.TTYTDDTelephoneNumber        != "") TTYTDDTelephoneNumber        ++;
      if (OutlookContact.UnRead                       != MyFields.EmptyUnRead) UnRead     ++;
      if (OutlookContact.User1                        != "") User1                        ++;
      if (OutlookContact.User2                        != "") User2                        ++;
      if (OutlookContact.User3                        != "") User3                        ++;
      if (OutlookContact.User4                        != "") User4                        ++;
      if (OutlookContact.UserCertificate              != "") UserCertificate              ++;
      if (OutlookContact.WebPage                      != "") WebPage                      ++;
      if (OutlookContact.YomiCompanyName              != "") YomiCompanyName              ++;
      if (OutlookContact.YomiFirstName                != "") YomiFirstName                ++;
      if (OutlookContact.YomiLastName                 != "") YomiLastName                 ++;
    //if (OutlookContact.ContactIndex                 != "") ContactIndex                 ++;
    //if (OutlookContact.Actions                      != "") Actions                      ++;
    //if (OutlookContact.Application                  != "") Application                  ++;
    //if (OutlookContact.Attachments                  != "") Attachments                  ++;
    //if (OutlookContact.Email1EntryID                != "") Email1EntryID                ++;
    //if (OutlookContact.Email2EntryID                != "") Email2EntryID                ++;
    //if (OutlookContact.Email3EntryID                != "") Email3EntryID                ++;
    //if (OutlookContact.FormDescription              != "") FormDescription              ++;
    //if (OutlookContact.GetInspector                 != "") GetInspector                 ++;
    //if (OutlookContact.Parent                       != "") Parent                       ++;
    //if (OutlookContact.UserProperties               != "") UserProperties               ++;
    }

    #endregion // Functions
  }
}
