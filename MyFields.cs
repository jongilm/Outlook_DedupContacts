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
  public class MyField
  {
    // Member Variables
    static  int    _NextIndex = 0;
    private int    _Ndx       ; public int    Ndx        { get { return _Ndx       ; } set { _Ndx        = value; } }
    private string _FieldName ; public string FieldName  { get { return _FieldName ; } set { _FieldName  = value; } }
    private int    _Occurances; public int    Occurances { get { return _Occurances; } set { _Occurances = value; } }

    // Constructor
    public MyField() 
    {
      _Ndx = _NextIndex++;
      _FieldName = ""; 
      _Occurances = 0;
    }

    // Constructor
    public MyField(string FieldName, int Occurances) 
    {
      SetField(FieldName, Occurances);
    }

    // AddField
    public void SetField(string FieldName, int Occurances) 
    {
      _Ndx = _NextIndex++;
      _FieldName = FieldName; 
      _Occurances = Occurances;
    }

    // Reset NextIndex
    public void ResetIndex() 
    {
      _NextIndex = 0;
    }
  }

  public class MyFields
  {
    MainForm m_form;                 // Reference to main form used to make syncronous user interface calls
    int m_ContactsProcessed;
  	
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

    public static DateTime         EmptyDate           = DateTime.MaxValue; // 01/01/4501
    public static OlGender         EmptyGender         = OlGender.olUnspecified;
    public static OlImportance     EmptyImportance     = OlImportance.olImportanceNormal;
    public static OlMailingAddress EmptyMailingAddress = OlMailingAddress.olNone;
    public static OlSensitivity    EmptySensitivity    = OlSensitivity.olNormal;
    public static bool             EmptyJournal        = false;
    public static bool             EmptyNoAging        = false;
    public static bool             EmptySaved          = true;
    public static bool             EmptyUnRead         = false;

    
    #endregion // MemberVariables

    #region Functions

    //////////////////////////////////////
    // Constructor
    //////////////////////////////////////
    public MyFields ( MainForm form )
    {
      m_form = form;
      ClearFieldCount();
    }
    
    //////////////////////////////////////
    // ClearFieldCount
    //////////////////////////////////////
    public void ClearFieldCount()
    {
      m_ContactsProcessed = 0;
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
    
    //////////////////////////////////////
    // AccumulateFields
    //////////////////////////////////////
    public void AccumulateFields(Microsoft.Office.Interop.Outlook._ContactItem OutlookContact)
    {
      m_ContactsProcessed++;

      if (OutlookContact.Account                      != null               ) Account                      ++;
      if (OutlookContact.Anniversary                  != EmptyDate          ) Anniversary                  ++;
      //if (OutlookContact.AssistantName                != ""                 ) AssistantName                ++;
      //if (OutlookContact.AssistantName.Length         > 0                   ) AssistantName                ++;
      if (OutlookContact.AssistantName                != null               ) AssistantName                ++;
      if (OutlookContact.AssistantTelephoneNumber     != null               ) AssistantTelephoneNumber     ++;
      if (OutlookContact.BillingInformation           != null               ) BillingInformation           ++;
      if (OutlookContact.Birthday                     != EmptyDate          ) Birthday                     ++;
      if (OutlookContact.Body                         != null               ) Body                         ++;
      if (OutlookContact.Business2TelephoneNumber     != null               ) Business2TelephoneNumber     ++;
      if (OutlookContact.BusinessAddress              != null               ) BusinessAddress              ++;
      if (OutlookContact.BusinessAddressCity          != null               ) BusinessAddressCity          ++;
      if (OutlookContact.BusinessAddressCountry       != null               ) BusinessAddressCountry       ++;
      if (OutlookContact.BusinessAddressPostalCode    != null               ) BusinessAddressPostalCode    ++;
      if (OutlookContact.BusinessAddressPostOfficeBox != null               ) BusinessAddressPostOfficeBox ++;
      if (OutlookContact.BusinessAddressState         != null               ) BusinessAddressState         ++;
      if (OutlookContact.BusinessAddressStreet        != null               ) BusinessAddressStreet        ++;
      if (OutlookContact.BusinessFaxNumber            != null               ) BusinessFaxNumber            ++;
      if (OutlookContact.BusinessHomePage             != null               ) BusinessHomePage             ++;
      if (OutlookContact.BusinessTelephoneNumber      != null               ) BusinessTelephoneNumber      ++;
      if (OutlookContact.CallbackTelephoneNumber      != null               ) CallbackTelephoneNumber      ++;
      if (OutlookContact.CarTelephoneNumber           != null               ) CarTelephoneNumber           ++;
      if (OutlookContact.Categories                   != null               ) Categories                   ++;
      if (OutlookContact.Children                     != null               ) Children                     ++;
      if (OutlookContact.Companies                    != null               ) Companies                    ++;
      if (OutlookContact.CompanyAndFullName           != null               ) CompanyAndFullName           ++;
      if (OutlookContact.CompanyMainTelephoneNumber   != null               ) CompanyMainTelephoneNumber   ++;
      if (OutlookContact.CompanyName                  != null               ) CompanyName                  ++;
      if (OutlookContact.ComputerNetworkName          != null               ) ComputerNetworkName          ++;
      if (OutlookContact.CreationTime                 != EmptyDate          ) CreationTime                 ++;
      if (OutlookContact.CustomerID                   != null               ) CustomerID                   ++;
      if (OutlookContact.Department                   != null               ) Department                   ++;
      if (OutlookContact.Email1Address                != null               ) Email1Address                ++;
      if (OutlookContact.Email1AddressType            != null               ) Email1AddressType            ++;
      if (OutlookContact.Email1DisplayName            != null               ) Email1DisplayName            ++;
      if (OutlookContact.Email2Address                != null               ) Email2Address                ++;
      if (OutlookContact.Email2AddressType            != null               ) Email2AddressType            ++;
      if (OutlookContact.Email2DisplayName            != null               ) Email2DisplayName            ++;
      if (OutlookContact.Email3Address                != null               ) Email3Address                ++;
      if (OutlookContact.Email3AddressType            != null               ) Email3AddressType            ++;
      if (OutlookContact.Email3DisplayName            != null               ) Email3DisplayName            ++;
      if (OutlookContact.EntryID                      != null               ) EntryID                      ++;
      if (OutlookContact.FileAs                       != null               ) FileAs                       ++;
      if (OutlookContact.FirstName                    != null               ) FirstName                    ++;
      if (OutlookContact.FTPSite                      != null               ) FTPSite                      ++;
      if (OutlookContact.FullName                     != null               ) FullName                     ++;
      if (OutlookContact.FullNameAndCompany           != null               ) FullNameAndCompany           ++;
      if (OutlookContact.Gender                       != EmptyGender        ) Gender                       ++;
      if (OutlookContact.GovernmentIDNumber           != null               ) GovernmentIDNumber           ++;
      if (OutlookContact.Hobby                        != null               ) Hobby                        ++;
      if (OutlookContact.Home2TelephoneNumber         != null               ) Home2TelephoneNumber         ++;
      if (OutlookContact.HomeAddress                  != null               ) HomeAddress                  ++;
      if (OutlookContact.HomeAddressCity              != null               ) HomeAddressCity              ++;
      if (OutlookContact.HomeAddressCountry           != null               ) HomeAddressCountry           ++;
      if (OutlookContact.HomeAddressPostalCode        != null               ) HomeAddressPostalCode        ++;
      if (OutlookContact.HomeAddressPostOfficeBox     != null               ) HomeAddressPostOfficeBox     ++;
      if (OutlookContact.HomeAddressState             != null               ) HomeAddressState             ++;
      if (OutlookContact.HomeAddressStreet            != null               ) HomeAddressStreet            ++;
      if (OutlookContact.HomeFaxNumber                != null               ) HomeFaxNumber                ++;
      if (OutlookContact.HomeTelephoneNumber          != null               ) HomeTelephoneNumber          ++;
      if (OutlookContact.Importance                   != EmptyImportance    ) Importance                   ++;
      if (OutlookContact.Initials                     != null               ) Initials                     ++;
      if (OutlookContact.ISDNNumber                   != null               ) ISDNNumber                   ++;
      if (OutlookContact.JobTitle                     != null               ) JobTitle                     ++;
      if (OutlookContact.Journal                      != false              ) Journal                      ++;
      if (OutlookContact.Language                     != null               ) Language                     ++;
      if (OutlookContact.LastModificationTime         != EmptyDate          ) LastModificationTime         ++;
      if (OutlookContact.LastName                     != null               ) LastName                     ++;
      if (OutlookContact.LastNameAndFirstName         != null               ) LastNameAndFirstName         ++;
      if (OutlookContact.MailingAddress               != null               ) MailingAddress               ++;
      if (OutlookContact.MailingAddressCity           != null               ) MailingAddressCity           ++;
      if (OutlookContact.MailingAddressCountry        != null               ) MailingAddressCountry        ++;
      if (OutlookContact.MailingAddressPostalCode     != null               ) MailingAddressPostalCode     ++;
      if (OutlookContact.MailingAddressPostOfficeBox  != null               ) MailingAddressPostOfficeBox  ++;
      if (OutlookContact.MailingAddressState          != null               ) MailingAddressState          ++;
      if (OutlookContact.MailingAddressStreet         != null               ) MailingAddressStreet         ++;
      if (OutlookContact.ManagerName                  != null               ) ManagerName                  ++;
      if (OutlookContact.MessageClass                 != null               ) MessageClass                 ++;
      if (OutlookContact.MiddleName                   != null               ) MiddleName                   ++;
      if (OutlookContact.Mileage                      != null               ) Mileage                      ++;
      if (OutlookContact.MobileTelephoneNumber        != null               ) MobileTelephoneNumber        ++;
      if (OutlookContact.NickName                     != null               ) NickName                     ++;
      if (OutlookContact.NoAging                      != EmptyNoAging       ) NoAging                      ++;
      if (OutlookContact.OfficeLocation               != null               ) OfficeLocation               ++;
      if (OutlookContact.OrganizationalIDNumber       != null               ) OrganizationalIDNumber       ++;
      if (OutlookContact.OtherAddress                 != null               ) OtherAddress                 ++;
      if (OutlookContact.OtherAddressCity             != null               ) OtherAddressCity             ++;
      if (OutlookContact.OtherAddressCountry          != null               ) OtherAddressCountry          ++;
      if (OutlookContact.OtherAddressPostalCode       != null               ) OtherAddressPostalCode       ++;
      if (OutlookContact.OtherAddressPostOfficeBox    != null               ) OtherAddressPostOfficeBox    ++;
      if (OutlookContact.OtherAddressState            != null               ) OtherAddressState            ++;
      if (OutlookContact.OtherAddressStreet           != null               ) OtherAddressStreet           ++;
      if (OutlookContact.OtherFaxNumber               != null               ) OtherFaxNumber               ++;
      if (OutlookContact.OtherTelephoneNumber         != null               ) OtherTelephoneNumber         ++;
      if (OutlookContact.OutlookInternalVersion       > 0                   ) OutlookInternalVersion       ++;
      if (OutlookContact.OutlookVersion               != null               ) OutlookVersion               ++;
      if (OutlookContact.PagerNumber                  != null               ) PagerNumber                  ++;
      if (OutlookContact.PersonalHomePage             != null               ) PersonalHomePage             ++;
      if (OutlookContact.PrimaryTelephoneNumber       != null               ) PrimaryTelephoneNumber       ++;
      if (OutlookContact.Profession                   != null               ) Profession                   ++;
      if (OutlookContact.RadioTelephoneNumber         != null               ) RadioTelephoneNumber         ++;
      if (OutlookContact.ReferredBy                   != null               ) ReferredBy                   ++;
      if (OutlookContact.Saved                        != EmptySaved         ) Saved                        ++;
      if (OutlookContact.SelectedMailingAddress       != EmptyMailingAddress) SelectedMailingAddress       ++;
      if (OutlookContact.Sensitivity                  != EmptySensitivity   ) Sensitivity                  ++;
      if (OutlookContact.Size                         != 0                  ) Size                         ++;
      if (OutlookContact.Spouse                       != null               ) Spouse                       ++;
      if (OutlookContact.Subject                      != null               ) Subject                      ++;
      if (OutlookContact.Suffix                       != null               ) Suffix                       ++;
      if (OutlookContact.TelexNumber                  != null               ) TelexNumber                  ++;
      if (OutlookContact.Title                        != null               ) Title                        ++;
      if (OutlookContact.TTYTDDTelephoneNumber        != null               ) TTYTDDTelephoneNumber        ++;
      if (OutlookContact.UnRead                       != EmptyUnRead        ) UnRead                       ++;
      if (OutlookContact.User1                        != null               ) User1                        ++;
      if (OutlookContact.User2                        != null               ) User2                        ++;
      if (OutlookContact.User3                        != null               ) User3                        ++;
      if (OutlookContact.User4                        != null               ) User4                        ++;
      if (OutlookContact.UserCertificate              != null               ) UserCertificate              ++;
      if (OutlookContact.WebPage                      != null               ) WebPage                      ++;
      if (OutlookContact.YomiCompanyName              != null               ) YomiCompanyName              ++;
      if (OutlookContact.YomiFirstName                != null               ) YomiFirstName                ++;
      if (OutlookContact.YomiLastName                 != null               ) YomiLastName                 ++;
    //if (OutlookContact.ContactIndex                 != null               ) ContactIndex                 ++;
    //if (OutlookContact.Actions                      != null               ) Actions                      ++;
    //if (OutlookContact.Application                  != null               ) Application                  ++;
    //if (OutlookContact.Attachments                  != null               ) Attachments                  ++;
    //if (OutlookContact.Email1EntryID                != null               ) Email1EntryID                ++;
    //if (OutlookContact.Email2EntryID                != null               ) Email2EntryID                ++;
    //if (OutlookContact.Email3EntryID                != null               ) Email3EntryID                ++;
    //if (OutlookContact.FormDescription              != null               ) FormDescription              ++;
    //if (OutlookContact.GetInspector                 != null               ) GetInspector                 ++;
    //if (OutlookContact.Parent                       != null               ) Parent                       ++;
    //if (OutlookContact.UserProperties               != null               ) UserProperties               ++;
    }

    //////////////////////////////////////
    // PrintSingleFieldCount
    //////////////////////////////////////
    private void PrintSingleFieldCount(string FieldName, int FieldCount)
    {
      if (FieldCount != m_ContactsProcessed)
      {
        string s;
        s = FieldName.Trim();
        s.PadRight(28);
        s += (" : " + FieldCount + "\r\n");
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AppendToListbox, new Object[] {s});
      }
      {
        MyField fld = new MyField();
        fld.FieldName = FieldName.Trim();
        fld.Occurances = FieldCount;
        m_form.Invoke(m_form.m_Delegate_LoadContactsThread_AppendToFieldList, new Object[] {fld});
      }
    }

    //////////////////////////////////////
    // PrintFieldCount
    //////////////////////////////////////
    public void PrintFieldCount()
    {
      PrintSingleFieldCount("Account                     ",Account                       );
      PrintSingleFieldCount("Anniversary                 ",Anniversary                   );
      PrintSingleFieldCount("AssistantName               ",AssistantName                 );
      PrintSingleFieldCount("AssistantTelephoneNumber    ",AssistantTelephoneNumber      );
      PrintSingleFieldCount("BillingInformation          ",BillingInformation            );
      PrintSingleFieldCount("Birthday                    ",Birthday                      );
      PrintSingleFieldCount("Body                        ",Body                          );
      PrintSingleFieldCount("Business2TelephoneNumber    ",Business2TelephoneNumber      );
      PrintSingleFieldCount("BusinessAddress             ",BusinessAddress               );
      PrintSingleFieldCount("BusinessAddressCity         ",BusinessAddressCity           );
      PrintSingleFieldCount("BusinessAddressCountry      ",BusinessAddressCountry        );
      PrintSingleFieldCount("BusinessAddressPostalCode   ",BusinessAddressPostalCode     );
      PrintSingleFieldCount("BusinessAddressPostOfficeBox",BusinessAddressPostOfficeBox  );
      PrintSingleFieldCount("BusinessAddressState        ",BusinessAddressState          );
      PrintSingleFieldCount("BusinessAddressStreet       ",BusinessAddressStreet         );
      PrintSingleFieldCount("BusinessFaxNumber           ",BusinessFaxNumber             );
      PrintSingleFieldCount("BusinessHomePage            ",BusinessHomePage              );
      PrintSingleFieldCount("BusinessTelephoneNumber     ",BusinessTelephoneNumber       );
      PrintSingleFieldCount("CallbackTelephoneNumber     ",CallbackTelephoneNumber       );
      PrintSingleFieldCount("CarTelephoneNumber          ",CarTelephoneNumber            );
      PrintSingleFieldCount("Categories                  ",Categories                    );
      PrintSingleFieldCount("Children                    ",Children                      );
      PrintSingleFieldCount("Companies                   ",Companies                     );
      PrintSingleFieldCount("CompanyAndFullName          ",CompanyAndFullName            );
      PrintSingleFieldCount("CompanyMainTelephoneNumber  ",CompanyMainTelephoneNumber    );
      PrintSingleFieldCount("CompanyName                 ",CompanyName                   );
      PrintSingleFieldCount("ComputerNetworkName         ",ComputerNetworkName           );
      PrintSingleFieldCount("CreationTime                ",CreationTime                  );
      PrintSingleFieldCount("CustomerID                  ",CustomerID                    );
      PrintSingleFieldCount("Department                  ",Department                    );
      PrintSingleFieldCount("Email1Address               ",Email1Address                 );
      PrintSingleFieldCount("Email1AddressType           ",Email1AddressType             );
      PrintSingleFieldCount("Email1DisplayName           ",Email1DisplayName             );
      PrintSingleFieldCount("Email2Address               ",Email2Address                 );
      PrintSingleFieldCount("Email2AddressType           ",Email2AddressType             );
      PrintSingleFieldCount("Email2DisplayName           ",Email2DisplayName             );
      PrintSingleFieldCount("Email3Address               ",Email3Address                 );
      PrintSingleFieldCount("Email3AddressType           ",Email3AddressType             );
      PrintSingleFieldCount("Email3DisplayName           ",Email3DisplayName             );
      PrintSingleFieldCount("EntryID                     ",EntryID                       );
      PrintSingleFieldCount("FileAs                      ",FileAs                        );
      PrintSingleFieldCount("FirstName                   ",FirstName                     );
      PrintSingleFieldCount("FTPSite                     ",FTPSite                       );
      PrintSingleFieldCount("FullName                    ",FullName                      );
      PrintSingleFieldCount("FullNameAndCompany          ",FullNameAndCompany            );
      PrintSingleFieldCount("Gender                      ",Gender                        );
      PrintSingleFieldCount("GovernmentIDNumber          ",GovernmentIDNumber            );
      PrintSingleFieldCount("Hobby                       ",Hobby                         );
      PrintSingleFieldCount("Home2TelephoneNumber        ",Home2TelephoneNumber          );
      PrintSingleFieldCount("HomeAddress                 ",HomeAddress                   );
      PrintSingleFieldCount("HomeAddressCity             ",HomeAddressCity               );
      PrintSingleFieldCount("HomeAddressCountry          ",HomeAddressCountry            );
      PrintSingleFieldCount("HomeAddressPostalCode       ",HomeAddressPostalCode         );
      PrintSingleFieldCount("HomeAddressPostOfficeBox    ",HomeAddressPostOfficeBox      );
      PrintSingleFieldCount("HomeAddressState            ",HomeAddressState              );
      PrintSingleFieldCount("HomeAddressStreet           ",HomeAddressStreet             );
      PrintSingleFieldCount("HomeFaxNumber               ",HomeFaxNumber                 );
      PrintSingleFieldCount("HomeTelephoneNumber         ",HomeTelephoneNumber           );
      PrintSingleFieldCount("Importance                  ",Importance                    );
      PrintSingleFieldCount("Initials                    ",Initials                      );
      PrintSingleFieldCount("ISDNNumber                  ",ISDNNumber                    );
      PrintSingleFieldCount("JobTitle                    ",JobTitle                      );
      PrintSingleFieldCount("Journal                     ",Journal                       );
      PrintSingleFieldCount("Language                    ",Language                      );
      PrintSingleFieldCount("LastModificationTime        ",LastModificationTime          );
      PrintSingleFieldCount("LastName                    ",LastName                      );
      PrintSingleFieldCount("LastNameAndFirstName        ",LastNameAndFirstName          );
      PrintSingleFieldCount("MailingAddress              ",MailingAddress                );
      PrintSingleFieldCount("MailingAddressCity          ",MailingAddressCity            );
      PrintSingleFieldCount("MailingAddressCountry       ",MailingAddressCountry         );
      PrintSingleFieldCount("MailingAddressPostalCode    ",MailingAddressPostalCode      );
      PrintSingleFieldCount("MailingAddressPostOfficeBox ",MailingAddressPostOfficeBox   );
      PrintSingleFieldCount("MailingAddressState         ",MailingAddressState           );
      PrintSingleFieldCount("MailingAddressStreet        ",MailingAddressStreet          );
      PrintSingleFieldCount("ManagerName                 ",ManagerName                   );
      PrintSingleFieldCount("MessageClass                ",MessageClass                  );
      PrintSingleFieldCount("MiddleName                  ",MiddleName                    );
      PrintSingleFieldCount("Mileage                     ",Mileage                       );
      PrintSingleFieldCount("MobileTelephoneNumber       ",MobileTelephoneNumber         );
      PrintSingleFieldCount("NickName                    ",NickName                      );
      PrintSingleFieldCount("NoAging                     ",NoAging                       );
      PrintSingleFieldCount("OfficeLocation              ",OfficeLocation                );
      PrintSingleFieldCount("OrganizationalIDNumber      ",OrganizationalIDNumber        );
      PrintSingleFieldCount("OtherAddress                ",OtherAddress                  );
      PrintSingleFieldCount("OtherAddressCity            ",OtherAddressCity              );
      PrintSingleFieldCount("OtherAddressCountry         ",OtherAddressCountry           );
      PrintSingleFieldCount("OtherAddressPostalCode      ",OtherAddressPostalCode        );
      PrintSingleFieldCount("OtherAddressPostOfficeBox   ",OtherAddressPostOfficeBox     );
      PrintSingleFieldCount("OtherAddressState           ",OtherAddressState             );
      PrintSingleFieldCount("OtherAddressStreet          ",OtherAddressStreet            );
      PrintSingleFieldCount("OtherFaxNumber              ",OtherFaxNumber                );
      PrintSingleFieldCount("OtherTelephoneNumber        ",OtherTelephoneNumber          );
      PrintSingleFieldCount("OutlookInternalVersion      ",OutlookInternalVersion        );
      PrintSingleFieldCount("OutlookVersion              ",OutlookVersion                );
      PrintSingleFieldCount("PagerNumber                 ",PagerNumber                   );
      PrintSingleFieldCount("PersonalHomePage            ",PersonalHomePage              );
      PrintSingleFieldCount("PrimaryTelephoneNumber      ",PrimaryTelephoneNumber        );
      PrintSingleFieldCount("Profession                  ",Profession                    );
      PrintSingleFieldCount("RadioTelephoneNumber        ",RadioTelephoneNumber          );
      PrintSingleFieldCount("ReferredBy                  ",ReferredBy                    );
      PrintSingleFieldCount("Saved                       ",Saved                         );
      PrintSingleFieldCount("SelectedMailingAddress      ",SelectedMailingAddress        );
      PrintSingleFieldCount("Sensitivity                 ",Sensitivity                   );
      PrintSingleFieldCount("Size                        ",Size                          );
      PrintSingleFieldCount("Spouse                      ",Spouse                        );
      PrintSingleFieldCount("Subject                     ",Subject                       );
      PrintSingleFieldCount("Suffix                      ",Suffix                        );
      PrintSingleFieldCount("TelexNumber                 ",TelexNumber                   );
      PrintSingleFieldCount("Title                       ",Title                         );
      PrintSingleFieldCount("TTYTDDTelephoneNumber       ",TTYTDDTelephoneNumber         );
      PrintSingleFieldCount("UnRead                      ",UnRead                        );
      PrintSingleFieldCount("User1                       ",User1                         );
      PrintSingleFieldCount("User2                       ",User2                         );
      PrintSingleFieldCount("User3                       ",User3                         );
      PrintSingleFieldCount("User4                       ",User4                         );
      PrintSingleFieldCount("UserCertificate             ",UserCertificate               );
      PrintSingleFieldCount("WebPage                     ",WebPage                       );
      PrintSingleFieldCount("YomiCompanyName             ",YomiCompanyName               );
      PrintSingleFieldCount("YomiFirstName               ",YomiFirstName                 );
      PrintSingleFieldCount("YomiLastName                ",YomiLastName                  );
    //PrintSingleFieldCount("ContactIndex                ",ContactIndex                  );
    //PrintSingleFieldCount("Actions                     ",Actions                       );
    //PrintSingleFieldCount("Application                 ",Application                   );
    //PrintSingleFieldCount("Attachments                 ",Attachments                   );
    //PrintSingleFieldCount("Email1EntryID               ",Email1EntryID                 );
    //PrintSingleFieldCount("Email2EntryID               ",Email2EntryID                 );
    //PrintSingleFieldCount("Email3EntryID               ",Email3EntryID                 );
    //PrintSingleFieldCount("FormDescription             ",FormDescription               );
    //PrintSingleFieldCount("GetInspector                ",GetInspector                  );
    //PrintSingleFieldCount("Parent                      ",Parent                        );
    //PrintSingleFieldCount("UserProperties              ",UserProperties                );
    }
    #endregion // Functions
  }

} // namespace OutlookContactsAnalyser
