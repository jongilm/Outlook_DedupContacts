using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Outlook;

using OutLook1 = Microsoft.Office.Interop.Outlook;
// Microsoft.Office.Interop.Outlook._ContactItem.LastName

namespace OutlookContactsAnalyser
{
  public class MyContact
  {
    public static bool m_PerformThoroughIdenticalChecks = false;
    public static int COMPARING_SAME_ITEMS = 1;
    public static int COMPARING_DIFFERENT_ITEMS = 2;

    private int      _Ndx                          ; public int      Ndx                          { get { return _Ndx                         ; } set { _Ndx                          = value; } }
    private int      _OutlookNdx                   ; public int      OutlookNdx                   { get { return _OutlookNdx                  ; } set { _OutlookNdx                   = value; } }
    //public List<int> SameNameList;
    
    private string   _Account                      ; public string   Account                      { get { return _Account                     ; } set { _Account                      = value; } }
    private DateTime _Anniversary                  ; public DateTime Anniversary                  { get { return _Anniversary                 ; } set { _Anniversary                  = value; } }
    private string   _AssistantName                ; public string   AssistantName                { get { return _AssistantName               ; } set { _AssistantName                = value; } }
    private string   _AssistantTelephoneNumber     ; public string   AssistantTelephoneNumber     { get { return _AssistantTelephoneNumber    ; } set { _AssistantTelephoneNumber     = value; } }
    private string   _BillingInformation           ; public string   BillingInformation           { get { return _BillingInformation          ; } set { _BillingInformation           = value; } }
    private DateTime _Birthday                     ; public DateTime Birthday                     { get { return _Birthday                    ; } set { _Birthday                     = value; } }
    private string   _Body                         ; public string   Body                         { get { return _Body                        ; } set { _Body                         = value; } }
    private string   _Business2TelephoneNumber     ; public string   Business2TelephoneNumber     { get { return _Business2TelephoneNumber    ; } set { _Business2TelephoneNumber     = value; } }
    private string   _BusinessAddress              ; public string   BusinessAddress              { get { return _BusinessAddress             ; } set { _BusinessAddress              = value; } }
    private string   _BusinessAddressCity          ; public string   BusinessAddressCity          { get { return _BusinessAddressCity         ; } set { _BusinessAddressCity          = value; } }
    private string   _BusinessAddressCountry       ; public string   BusinessAddressCountry       { get { return _BusinessAddressCountry      ; } set { _BusinessAddressCountry       = value; } }
    private string   _BusinessAddressPostalCode    ; public string   BusinessAddressPostalCode    { get { return _BusinessAddressPostalCode   ; } set { _BusinessAddressPostalCode    = value; } }
    private string   _BusinessAddressPostOfficeBox ; public string   BusinessAddressPostOfficeBox { get { return _BusinessAddressPostOfficeBox; } set { _BusinessAddressPostOfficeBox = value; } }
    private string   _BusinessAddressState         ; public string   BusinessAddressState         { get { return _BusinessAddressState        ; } set { _BusinessAddressState         = value; } }
    private string   _BusinessAddressStreet        ; public string   BusinessAddressStreet        { get { return _BusinessAddressStreet       ; } set { _BusinessAddressStreet        = value; } }
    private string   _BusinessFaxNumber            ; public string   BusinessFaxNumber            { get { return _BusinessFaxNumber           ; } set { _BusinessFaxNumber            = value; } }
    private string   _BusinessHomePage             ; public string   BusinessHomePage             { get { return _BusinessHomePage            ; } set { _BusinessHomePage             = value; } }
    private string   _BusinessTelephoneNumber      ; public string   BusinessTelephoneNumber      { get { return _BusinessTelephoneNumber     ; } set { _BusinessTelephoneNumber      = value; } }
    private string   _CallbackTelephoneNumber      ; public string   CallbackTelephoneNumber      { get { return _CallbackTelephoneNumber     ; } set { _CallbackTelephoneNumber      = value; } }
    private string   _CarTelephoneNumber           ; public string   CarTelephoneNumber           { get { return _CarTelephoneNumber          ; } set { _CarTelephoneNumber           = value; } }
    private string   _Categories                   ; public string   Categories                   { get { return _Categories                  ; } set { _Categories                   = value; } }
    private string   _Children                     ; public string   Children                     { get { return _Children                    ; } set { _Children                     = value; } }
    private string   _Companies                    ; public string   Companies                    { get { return _Companies                   ; } set { _Companies                    = value; } }
    private string   _CompanyAndFullName           ; public string   CompanyAndFullName           { get { return _CompanyAndFullName          ; } set { _CompanyAndFullName           = value; } }
    private string   _CompanyMainTelephoneNumber   ; public string   CompanyMainTelephoneNumber   { get { return _CompanyMainTelephoneNumber  ; } set { _CompanyMainTelephoneNumber   = value; } }
    private string   _CompanyName                  ; public string   CompanyName                  { get { return _CompanyName                 ; } set { _CompanyName                  = value; } }
    private string   _ComputerNetworkName          ; public string   ComputerNetworkName          { get { return _ComputerNetworkName         ; } set { _ComputerNetworkName          = value; } }
    private DateTime _CreationTime                 ; public DateTime CreationTime                 { get { return _CreationTime                ; } set { _CreationTime                 = value; } }
    private string   _CustomerID                   ; public string   CustomerID                   { get { return _CustomerID                  ; } set { _CustomerID                   = value; } }
    private string   _Department                   ; public string   Department                   { get { return _Department                  ; } set { _Department                   = value; } }
    private string   _Email1Address                ; public string   Email1Address                { get { return _Email1Address               ; } set { _Email1Address                = value; } }
    private string   _Email1AddressType            ; public string   Email1AddressType            { get { return _Email1AddressType           ; } set { _Email1AddressType            = value; } }
    private string   _Email1DisplayName            ; public string   Email1DisplayName            { get { return _Email1DisplayName           ; } set { _Email1DisplayName            = value; } }
    private string   _Email2Address                ; public string   Email2Address                { get { return _Email2Address               ; } set { _Email2Address                = value; } }
    private string   _Email2AddressType            ; public string   Email2AddressType            { get { return _Email2AddressType           ; } set { _Email2AddressType            = value; } }
    private string   _Email2DisplayName            ; public string   Email2DisplayName            { get { return _Email2DisplayName           ; } set { _Email2DisplayName            = value; } }
    private string   _Email3Address                ; public string   Email3Address                { get { return _Email3Address               ; } set { _Email3Address                = value; } }
    private string   _Email3AddressType            ; public string   Email3AddressType            { get { return _Email3AddressType           ; } set { _Email3AddressType            = value; } }
    private string   _Email3DisplayName            ; public string   Email3DisplayName            { get { return _Email3DisplayName           ; } set { _Email3DisplayName            = value; } }
    private string   _EntryID                      ; public string   EntryID                      { get { return _EntryID                     ; } set { _EntryID                      = value; } }
    private string   _FileAs                       ; public string   FileAs                       { get { return _FileAs                      ; } set { _FileAs                       = value; } }
    private string   _FirstName                    ; public string   FirstName                    { get { return _FirstName                   ; } set { _FirstName                    = value; } }
    private string   _FTPSite                      ; public string   FTPSite                      { get { return _FTPSite                     ; } set { _FTPSite                      = value; } }
    private string   _FullName                     ; public string   FullName                     { get { return _FullName                    ; } set { _FullName                     = value; } }
    private string   _FullNameAndCompany           ; public string   FullNameAndCompany           { get { return _FullNameAndCompany          ; } set { _FullNameAndCompany           = value; } }
    private string   _GovernmentIDNumber           ; public string   GovernmentIDNumber           { get { return _GovernmentIDNumber          ; } set { _GovernmentIDNumber           = value; } }
    private string   _Hobby                        ; public string   Hobby                        { get { return _Hobby                       ; } set { _Hobby                        = value; } }
    private string   _Home2TelephoneNumber         ; public string   Home2TelephoneNumber         { get { return _Home2TelephoneNumber        ; } set { _Home2TelephoneNumber         = value; } }
    private string   _HomeAddress                  ; public string   HomeAddress                  { get { return _HomeAddress                 ; } set { _HomeAddress                  = value; } }
    private string   _HomeAddressCity              ; public string   HomeAddressCity              { get { return _HomeAddressCity             ; } set { _HomeAddressCity              = value; } }
    private string   _HomeAddressCountry           ; public string   HomeAddressCountry           { get { return _HomeAddressCountry          ; } set { _HomeAddressCountry           = value; } }
    private string   _HomeAddressPostalCode        ; public string   HomeAddressPostalCode        { get { return _HomeAddressPostalCode       ; } set { _HomeAddressPostalCode        = value; } }
    private string   _HomeAddressPostOfficeBox     ; public string   HomeAddressPostOfficeBox     { get { return _HomeAddressPostOfficeBox    ; } set { _HomeAddressPostOfficeBox     = value; } }
    private string   _HomeAddressState             ; public string   HomeAddressState             { get { return _HomeAddressState            ; } set { _HomeAddressState             = value; } }
    private string   _HomeAddressStreet            ; public string   HomeAddressStreet            { get { return _HomeAddressStreet           ; } set { _HomeAddressStreet            = value; } }
    private string   _HomeFaxNumber                ; public string   HomeFaxNumber                { get { return _HomeFaxNumber               ; } set { _HomeFaxNumber                = value; } }
    private string   _HomeTelephoneNumber          ; public string   HomeTelephoneNumber          { get { return _HomeTelephoneNumber         ; } set { _HomeTelephoneNumber          = value; } }
    private string   _Initials                     ; public string   Initials                     { get { return _Initials                    ; } set { _Initials                     = value; } }
    private string   _ISDNNumber                   ; public string   ISDNNumber                   { get { return _ISDNNumber                  ; } set { _ISDNNumber                   = value; } }
    private string   _JobTitle                     ; public string   JobTitle                     { get { return _JobTitle                    ; } set { _JobTitle                     = value; } }
    private bool     _Journal                      ; public bool     Journal                      { get { return _Journal                     ; } set { _Journal                      = value; } }
    private string   _Language                     ; public string   Language                     { get { return _Language                    ; } set { _Language                     = value; } }
    private DateTime _LastModificationTime         ; public DateTime LastModificationTime         { get { return _LastModificationTime        ; } set { _LastModificationTime         = value; } }
    private string   _LastName                     ; public string   LastName                     { get { return _LastName                    ; } set { _LastName                     = value; } }
    private string   _LastNameAndFirstName         ; public string   LastNameAndFirstName         { get { return _LastNameAndFirstName        ; } set { _LastNameAndFirstName         = value; } }
    private string   _MailingAddress               ; public string   MailingAddress               { get { return _MailingAddress              ; } set { _MailingAddress               = value; } }
    private string   _MailingAddressCity           ; public string   MailingAddressCity           { get { return _MailingAddressCity          ; } set { _MailingAddressCity           = value; } }
    private string   _MailingAddressCountry        ; public string   MailingAddressCountry        { get { return _MailingAddressCountry       ; } set { _MailingAddressCountry        = value; } }
    private string   _MailingAddressPostalCode     ; public string   MailingAddressPostalCode     { get { return _MailingAddressPostalCode    ; } set { _MailingAddressPostalCode     = value; } }
    private string   _MailingAddressPostOfficeBox  ; public string   MailingAddressPostOfficeBox  { get { return _MailingAddressPostOfficeBox ; } set { _MailingAddressPostOfficeBox  = value; } }
    private string   _MailingAddressState          ; public string   MailingAddressState          { get { return _MailingAddressState         ; } set { _MailingAddressState          = value; } }
    private string   _MailingAddressStreet         ; public string   MailingAddressStreet         { get { return _MailingAddressStreet        ; } set { _MailingAddressStreet         = value; } }
    private string   _ManagerName                  ; public string   ManagerName                  { get { return _ManagerName                 ; } set { _ManagerName                  = value; } }
    private string   _MessageClass                 ; public string   MessageClass                 { get { return _MessageClass                ; } set { _MessageClass                 = value; } }
    private string   _MiddleName                   ; public string   MiddleName                   { get { return _MiddleName                  ; } set { _MiddleName                   = value; } }
    private string   _Mileage                      ; public string   Mileage                      { get { return _Mileage                     ; } set { _Mileage                      = value; } }
    private string   _MobileTelephoneNumber        ; public string   MobileTelephoneNumber        { get { return _MobileTelephoneNumber       ; } set { _MobileTelephoneNumber        = value; } }
    private string   _NickName                     ; public string   NickName                     { get { return _NickName                    ; } set { _NickName                     = value; } }
    private bool     _NoAging                      ; public bool     NoAging                      { get { return _NoAging                     ; } set { _NoAging                      = value; } }
    private string   _OfficeLocation               ; public string   OfficeLocation               { get { return _OfficeLocation              ; } set { _OfficeLocation               = value; } }
    private string   _OrganizationalIDNumber       ; public string   OrganizationalIDNumber       { get { return _OrganizationalIDNumber      ; } set { _OrganizationalIDNumber       = value; } }
    private string   _OtherAddress                 ; public string   OtherAddress                 { get { return _OtherAddress                ; } set { _OtherAddress                 = value; } }
    private string   _OtherAddressCity             ; public string   OtherAddressCity             { get { return _OtherAddressCity            ; } set { _OtherAddressCity             = value; } }
    private string   _OtherAddressCountry          ; public string   OtherAddressCountry          { get { return _OtherAddressCountry         ; } set { _OtherAddressCountry          = value; } }
    private string   _OtherAddressPostalCode       ; public string   OtherAddressPostalCode       { get { return _OtherAddressPostalCode      ; } set { _OtherAddressPostalCode       = value; } }
    private string   _OtherAddressPostOfficeBox    ; public string   OtherAddressPostOfficeBox    { get { return _OtherAddressPostOfficeBox   ; } set { _OtherAddressPostOfficeBox    = value; } }
    private string   _OtherAddressState            ; public string   OtherAddressState            { get { return _OtherAddressState           ; } set { _OtherAddressState            = value; } }
    private string   _OtherAddressStreet           ; public string   OtherAddressStreet           { get { return _OtherAddressStreet          ; } set { _OtherAddressStreet           = value; } }
    private string   _OtherFaxNumber               ; public string   OtherFaxNumber               { get { return _OtherFaxNumber              ; } set { _OtherFaxNumber               = value; } }
    private string   _OtherTelephoneNumber         ; public string   OtherTelephoneNumber         { get { return _OtherTelephoneNumber        ; } set { _OtherTelephoneNumber         = value; } }
    private int      _OutlookInternalVersion       ; public int      OutlookInternalVersion       { get { return _OutlookInternalVersion      ; } set { _OutlookInternalVersion       = value; } }
    private string   _OutlookVersion               ; public string   OutlookVersion               { get { return _OutlookVersion              ; } set { _OutlookVersion               = value; } }
    private string   _PagerNumber                  ; public string   PagerNumber                  { get { return _PagerNumber                 ; } set { _PagerNumber                  = value; } }
    private string   _PersonalHomePage             ; public string   PersonalHomePage             { get { return _PersonalHomePage            ; } set { _PersonalHomePage             = value; } }
    private string   _PrimaryTelephoneNumber       ; public string   PrimaryTelephoneNumber       { get { return _PrimaryTelephoneNumber      ; } set { _PrimaryTelephoneNumber       = value; } }
    private string   _Profession                   ; public string   Profession                   { get { return _Profession                  ; } set { _Profession                   = value; } }
    private string   _RadioTelephoneNumber         ; public string   RadioTelephoneNumber         { get { return _RadioTelephoneNumber        ; } set { _RadioTelephoneNumber         = value; } }
    private string   _ReferredBy                   ; public string   ReferredBy                   { get { return _ReferredBy                  ; } set { _ReferredBy                   = value; } }
    private bool     _Saved                        ; public bool     Saved                        { get { return _Saved                       ; } set { _Saved                        = value; } }
    private int      _Size                         ; public int      Size                         { get { return _Size                        ; } set { _Size                         = value; } }
    private string   _Spouse                       ; public string   Spouse                       { get { return _Spouse                      ; } set { _Spouse                       = value; } }
    private string   _Subject                      ; public string   Subject                      { get { return _Subject                     ; } set { _Subject                      = value; } }
    private string   _Suffix                       ; public string   Suffix                       { get { return _Suffix                      ; } set { _Suffix                       = value; } }
    private string   _TelexNumber                  ; public string   TelexNumber                  { get { return _TelexNumber                 ; } set { _TelexNumber                  = value; } }
    private string   _Title                        ; public string   Title                        { get { return _Title                       ; } set { _Title                        = value; } }
    private string   _TTYTDDTelephoneNumber        ; public string   TTYTDDTelephoneNumber        { get { return _TTYTDDTelephoneNumber       ; } set { _TTYTDDTelephoneNumber        = value; } }
    private bool     _UnRead                       ; public bool     UnRead                       { get { return _UnRead                      ; } set { _UnRead                       = value; } }
    private string   _User1                        ; public string   User1                        { get { return _User1                       ; } set { _User1                        = value; } }
    private string   _User2                        ; public string   User2                        { get { return _User2                       ; } set { _User2                        = value; } }
    private string   _User3                        ; public string   User3                        { get { return _User3                       ; } set { _User3                        = value; } }
    private string   _User4                        ; public string   User4                        { get { return _User4                       ; } set { _User4                        = value; } }
    private string   _UserCertificate              ; public string   UserCertificate              { get { return _UserCertificate             ; } set { _UserCertificate              = value; } }
    private string   _WebPage                      ; public string   WebPage                      { get { return _WebPage                     ; } set { _WebPage                      = value; } }
    private string   _YomiCompanyName              ; public string   YomiCompanyName              { get { return _YomiCompanyName             ; } set { _YomiCompanyName              = value; } }
    private string   _YomiFirstName                ; public string   YomiFirstName                { get { return _YomiFirstName               ; } set { _YomiFirstName                = value; } }
    private string   _YomiLastName                 ; public string   YomiLastName                 { get { return _YomiLastName                ; } set { _YomiLastName                 = value; } }

    private Microsoft.Office.Interop.Outlook.OlGender         _Gender                       ; public Microsoft.Office.Interop.Outlook.OlGender         Gender                       { get { return _Gender                      ; } set { _Gender                       = value; } }
    private Microsoft.Office.Interop.Outlook.OlImportance     _Importance                   ; public Microsoft.Office.Interop.Outlook.OlImportance     Importance                   { get { return _Importance                  ; } set { _Importance                   = value; } }
    private Microsoft.Office.Interop.Outlook.OlMailingAddress _SelectedMailingAddress       ; public Microsoft.Office.Interop.Outlook.OlMailingAddress SelectedMailingAddress       { get { return _SelectedMailingAddress      ; } set { _SelectedMailingAddress       = value; } }
    private Microsoft.Office.Interop.Outlook.OlSensitivity    _Sensitivity                  ; public Microsoft.Office.Interop.Outlook.OlSensitivity    Sensitivity                  { get { return _Sensitivity                 ; } set { _Sensitivity                  = value; } }

  //private string   _ContactIndex                 ; public string   ContactIndex                 { get { return _ContactIndex                ; } set { _ContactIndex                 = value; } }
  //private string   _Actions                      ; public string   Actions                      { get { return _Actions                     ; } set { _Actions                      = value; } }
  //private string   _Application                  ; public string   Application                  { get { return _Application                 ; } set { _Application                  = value; } }
  //private string   _Attachments                  ; public string   Attachments                  { get { return _Attachments                 ; } set { _Attachments                  = value; } }
  //private string   _Email1EntryID                ; public string   Email1EntryID                { get { return _Email1EntryID               ; } set { _Email1EntryID                = value; } }
  //private string   _Email2EntryID                ; public string   Email2EntryID                { get { return _Email2EntryID               ; } set { _Email2EntryID                = value; } }
  //private string   _Email3EntryID                ; public string   Email3EntryID                { get { return _Email3EntryID               ; } set { _Email3EntryID                = value; } }
  //private string   _FormDescription              ; public string   FormDescription              { get { return _FormDescription             ; } set { _FormDescription              = value; } }
  //private string   _GetInspector                 ; public string   GetInspector                 { get { return _GetInspector                ; } set { _GetInspector                 = value; } }
  //private string   _Parent                       ; public string   Parent                       { get { return _Parent                      ; } set { _Parent                       = value; } }
  //private string   _UserProperties               ; public string   UserProperties               { get { return _UserProperties              ; } set { _UserProperties               = value; } }

    //////////////////////////////////////
    //
    //////////////////////////////////////
    private int StrLen(string Str)
    {
      if (Str != null)
        return Str.Length;
      return 0;   
    }  

    public void Clear()
    {
      Ndx        = 0;
      OutlookNdx = 0;
        
      Account                       = "";
      Anniversary                   = MyFields.EmptyDate;
      AssistantName                 = "";
      AssistantTelephoneNumber      = "";
      BillingInformation            = "";
      Birthday                      = MyFields.EmptyDate;
      Body                          = "";
      Business2TelephoneNumber      = "";
      BusinessAddress               = "";
      BusinessAddressCity           = "";
      BusinessAddressCountry        = "";
      BusinessAddressPostalCode     = "";
      BusinessAddressPostOfficeBox  = "";
      BusinessAddressState          = "";
      BusinessAddressStreet         = "";
      BusinessFaxNumber             = "";
      BusinessHomePage              = "";
      BusinessTelephoneNumber       = "";
      CallbackTelephoneNumber       = "";
      CarTelephoneNumber            = "";
      Categories                    = "";
      Children                      = "";
      Companies                     = "";
      CompanyAndFullName            = "";
      CompanyMainTelephoneNumber    = "";
      CompanyName                   = "";
      ComputerNetworkName           = "";
      CreationTime                  = MyFields.EmptyDate;
      CustomerID                    = "";
      Department                    = "";
      Email1Address                 = "";
      Email1AddressType             = "";
      Email1DisplayName             = "";
      Email2Address                 = "";
      Email2AddressType             = "";
      Email2DisplayName             = "";
      Email3Address                 = "";
      Email3AddressType             = "";
      Email3DisplayName             = "";
      EntryID                       = "";
      FileAs                        = "";
      FirstName                     = "";
      FTPSite                       = "";
      FullName                      = "";
      FullNameAndCompany            = "";
      Gender                        = MyFields.EmptyGender;
      GovernmentIDNumber            = "";
      Hobby                         = "";
      Home2TelephoneNumber          = "";
      HomeAddress                   = "";
      HomeAddressCity               = "";
      HomeAddressCountry            = "";
      HomeAddressPostalCode         = "";
      HomeAddressPostOfficeBox      = "";
      HomeAddressState              = "";
      HomeAddressStreet             = "";
      HomeFaxNumber                 = "";
      HomeTelephoneNumber           = "";
      Importance                    = MyFields.EmptyImportance;
      Initials                      = "";
      ISDNNumber                    = "";
      JobTitle                      = "";
      Journal                       = MyFields.EmptyJournal;
      Language                      = "";
      LastModificationTime          = MyFields.EmptyDate;
      LastName                      = "";
      LastNameAndFirstName          = "";
      MailingAddress                = "";
      MailingAddressCity            = "";
      MailingAddressCountry         = "";
      MailingAddressPostalCode      = "";
      MailingAddressPostOfficeBox   = "";
      MailingAddressState           = "";
      MailingAddressStreet          = "";
      ManagerName                   = "";
      MessageClass                  = "";
      MiddleName                    = "";
      Mileage                       = "";
      MobileTelephoneNumber         = "";
      NickName                      = "";
      NoAging                       = MyFields.EmptyNoAging;
      OfficeLocation                = "";
      OrganizationalIDNumber        = "";
      OtherAddress                  = "";
      OtherAddressCity              = "";
      OtherAddressCountry           = "";
      OtherAddressPostalCode        = "";
      OtherAddressPostOfficeBox     = "";
      OtherAddressState             = "";
      OtherAddressStreet            = "";
      OtherFaxNumber                = "";
      OtherTelephoneNumber          = "";
      OutlookInternalVersion        = 0;
      OutlookVersion                = "";
      PagerNumber                   = "";
      PersonalHomePage              = "";
      PrimaryTelephoneNumber        = "";
      Profession                    = "";
      RadioTelephoneNumber          = "";
      ReferredBy                    = "";
      Saved                         = MyFields.EmptySaved;
      SelectedMailingAddress        = MyFields.EmptyMailingAddress;
      Sensitivity                   = MyFields.EmptySensitivity;
      Size                          = 0;
      Spouse                        = "";
      Subject                       = "";
      Suffix                        = "";
      TelexNumber                   = "";
      Title                         = "";
      TTYTDDTelephoneNumber         = "";
      UnRead                        = MyFields.EmptyUnRead;
      User1                         = "";
      User2                         = "";
      User3                         = "";
      User4                         = "";
      UserCertificate               = "";
      WebPage                       = "";
      YomiCompanyName               = "";
      YomiFirstName                 = "";
      YomiLastName                  = "";
      //ContactIndex                  = "";
      //Actions                       = "";
      //Application                   = "";
      //Attachments                   = "";
      //Email1EntryID                 = "";
      //Email2EntryID                 = "";
      //Email3EntryID                 = "";
      //FormDescription               = "";
      //GetInspector                  = "";
      //Parent                        = "";
      //UserProperties                = "";
    }

    public string CleanField(string s1)
    {
      //char[] szTrimChars = {'\r','\n','\t'};
      if (s1 == null)
        return null;

      s1 = s1.Trim();

      // Trim all leading and trailing CRLFs and TABs
      //s1 = s1.Trim(szTrimChars);


      // Remove Leading CRLFs
      for(;;)
      {
        if (s1.StartsWith("\\r\\n"))
          s1 = s1.Substring(4);
        else
          break;
      }
      // Remove Trailing CRLFs
      for(;;)
      {
        if (s1.EndsWith("\\r\\n"))
          s1 = s1.Substring(0,s1.Length-4);
        else
          break;
      }

      // Ensure all Body (Notes) fields end with a single CRLF
      if (!s1.EndsWith("\\r\\n"))
        s1 = s1 + "\\r\\n";

      s1 = s1.Replace("&", "and");

      return s1;
    }

    private string Escape(string s1)
    {
      string s2;
      if (s1 == null)
        return null;

      // Backslash replacement must be done first
      s2 = s1.Replace("\\","\\\\").Replace("\r","\\r").Replace("\n","\\n").Replace("\t","\\t");
      return s2;
    }
    private string UnEscape(string s1)
    {
      string s2;
      if (s1 == null)
        return null;
      // Backslash replacement must be done last
      s2 = s1.Replace("\\t","\t").Replace("\\n","\n").Replace("\\r","\r").Replace("\\\\","\\");
      return s2;
    }
    public void StoreContact(MyContact contact, int Ndx, Microsoft.Office.Interop.Outlook._ContactItem OutlookContact)
    {
        contact.Ndx = Ndx;
        contact.OutlookNdx = 0;

        contact.Account                       = OutlookContact.Account                      ;
        contact.Anniversary                   = OutlookContact.Anniversary                  ;
        contact.AssistantName                 = OutlookContact.AssistantName                ;
        contact.AssistantTelephoneNumber      = OutlookContact.AssistantTelephoneNumber     ;
        contact.BillingInformation            = OutlookContact.BillingInformation           ;
        contact.Birthday                      = OutlookContact.Birthday                     ;
        contact.Body                          = Escape(OutlookContact.Body)                 ;
        contact.Business2TelephoneNumber      = OutlookContact.Business2TelephoneNumber     ;
        contact.BusinessAddress               = Escape(OutlookContact.BusinessAddress)      ;
        contact.BusinessAddressCity           = OutlookContact.BusinessAddressCity          ;
        contact.BusinessAddressCountry        = OutlookContact.BusinessAddressCountry       ;
        contact.BusinessAddressPostalCode     = OutlookContact.BusinessAddressPostalCode    ;
        contact.BusinessAddressPostOfficeBox  = OutlookContact.BusinessAddressPostOfficeBox ;
        contact.BusinessAddressState          = OutlookContact.BusinessAddressState         ;
        contact.BusinessAddressStreet         = Escape(OutlookContact.BusinessAddressStreet);
        contact.BusinessFaxNumber             = OutlookContact.BusinessFaxNumber            ;
        contact.BusinessHomePage              = OutlookContact.BusinessHomePage             ;
        contact.BusinessTelephoneNumber       = OutlookContact.BusinessTelephoneNumber      ;
        contact.CallbackTelephoneNumber       = OutlookContact.CallbackTelephoneNumber      ;
        contact.CarTelephoneNumber            = OutlookContact.CarTelephoneNumber           ;
        contact.Categories                    = OutlookContact.Categories                   ;
        contact.Children                      = OutlookContact.Children                     ;
        contact.Companies                     = OutlookContact.Companies                    ;
        contact.CompanyAndFullName            = OutlookContact.CompanyAndFullName           ;
        contact.CompanyMainTelephoneNumber    = OutlookContact.CompanyMainTelephoneNumber   ;
        contact.CompanyName                   = OutlookContact.CompanyName                  ;
        contact.ComputerNetworkName           = OutlookContact.ComputerNetworkName          ;
        contact.CreationTime                  = OutlookContact.CreationTime                 ;
        contact.CustomerID                    = OutlookContact.CustomerID                   ;
        contact.Department                    = OutlookContact.Department                   ;
        contact.Email1Address                 = OutlookContact.Email1Address                ;
        contact.Email1AddressType             = OutlookContact.Email1AddressType            ;
        contact.Email1DisplayName             = OutlookContact.Email1DisplayName            ;
        contact.Email2Address                 = OutlookContact.Email2Address                ;
        contact.Email2AddressType             = OutlookContact.Email2AddressType            ;
        contact.Email2DisplayName             = OutlookContact.Email2DisplayName            ;
        contact.Email3Address                 = OutlookContact.Email3Address                ;
        contact.Email3AddressType             = OutlookContact.Email3AddressType            ;
        contact.Email3DisplayName             = OutlookContact.Email3DisplayName            ;
        contact.EntryID                       = OutlookContact.EntryID                      ;
        contact.FileAs                        = OutlookContact.FileAs                       ;
        contact.FirstName                     = OutlookContact.FirstName                    ;
        contact.FTPSite                       = OutlookContact.FTPSite                      ;
        contact.FullName                      = OutlookContact.FullName                     ;
        contact.FullNameAndCompany            = OutlookContact.FullNameAndCompany           ;
        contact.Gender                        = OutlookContact.Gender                       ;
        contact.GovernmentIDNumber            = OutlookContact.GovernmentIDNumber           ;
        contact.Hobby                         = OutlookContact.Hobby                        ;
        contact.Home2TelephoneNumber          = OutlookContact.Home2TelephoneNumber         ;
        contact.HomeAddress                   = Escape(OutlookContact.HomeAddress)          ;
        contact.HomeAddressCity               = OutlookContact.HomeAddressCity              ;
        contact.HomeAddressCountry            = OutlookContact.HomeAddressCountry           ;
        contact.HomeAddressPostalCode         = OutlookContact.HomeAddressPostalCode        ;
        contact.HomeAddressPostOfficeBox      = OutlookContact.HomeAddressPostOfficeBox     ;
        contact.HomeAddressState              = OutlookContact.HomeAddressState             ;
        contact.HomeAddressStreet             = Escape(OutlookContact.HomeAddressStreet)    ;
        contact.HomeFaxNumber                 = OutlookContact.HomeFaxNumber                ;
        contact.HomeTelephoneNumber           = OutlookContact.HomeTelephoneNumber          ;
        contact.Importance                    = OutlookContact.Importance                   ;
        contact.Initials                      = OutlookContact.Initials                     ;
        contact.ISDNNumber                    = OutlookContact.ISDNNumber                   ;
        contact.JobTitle                      = OutlookContact.JobTitle                     ;
        contact.Journal                       = OutlookContact.Journal                      ;
        contact.Language                      = OutlookContact.Language                     ;
        contact.LastModificationTime          = OutlookContact.LastModificationTime         ;
        contact.LastName                      = OutlookContact.LastName                     ;
        contact.LastNameAndFirstName          = OutlookContact.LastNameAndFirstName         ;
        contact.MailingAddress                = Escape(OutlookContact.MailingAddress)       ;
        contact.MailingAddressCity            = OutlookContact.MailingAddressCity           ;
        contact.MailingAddressCountry         = OutlookContact.MailingAddressCountry        ;
        contact.MailingAddressPostalCode      = OutlookContact.MailingAddressPostalCode     ;
        contact.MailingAddressPostOfficeBox   = OutlookContact.MailingAddressPostOfficeBox  ;
        contact.MailingAddressState           = OutlookContact.MailingAddressState          ;
        contact.MailingAddressStreet          = Escape(OutlookContact.MailingAddressStreet) ;
        contact.ManagerName                   = OutlookContact.ManagerName                  ;
        contact.MessageClass                  = OutlookContact.MessageClass                 ;
        contact.MiddleName                    = OutlookContact.MiddleName                   ;
        contact.Mileage                       = OutlookContact.Mileage                      ;
        contact.MobileTelephoneNumber         = OutlookContact.MobileTelephoneNumber        ;
        contact.NickName                      = OutlookContact.NickName                     ;
        contact.NoAging                       = OutlookContact.NoAging                      ;
        contact.OfficeLocation                = OutlookContact.OfficeLocation               ;
        contact.OrganizationalIDNumber        = OutlookContact.OrganizationalIDNumber       ;
        contact.OtherAddress                  = Escape(OutlookContact.OtherAddress)         ;
        contact.OtherAddressCity              = OutlookContact.OtherAddressCity             ;
        contact.OtherAddressCountry           = OutlookContact.OtherAddressCountry          ;
        contact.OtherAddressPostalCode        = OutlookContact.OtherAddressPostalCode       ;
        contact.OtherAddressPostOfficeBox     = OutlookContact.OtherAddressPostOfficeBox    ;
        contact.OtherAddressState             = OutlookContact.OtherAddressState            ;
        contact.OtherAddressStreet            = Escape(OutlookContact.OtherAddressStreet)   ;
        contact.OtherFaxNumber                = OutlookContact.OtherFaxNumber               ;
        contact.OtherTelephoneNumber          = OutlookContact.OtherTelephoneNumber         ;
        contact.OutlookInternalVersion        = OutlookContact.OutlookInternalVersion       ;
        contact.OutlookVersion                = OutlookContact.OutlookVersion               ;
        contact.PagerNumber                   = OutlookContact.PagerNumber                  ;
        contact.PersonalHomePage              = OutlookContact.PersonalHomePage             ;
        contact.PrimaryTelephoneNumber        = OutlookContact.PrimaryTelephoneNumber       ;
        contact.Profession                    = OutlookContact.Profession                   ;
        contact.RadioTelephoneNumber          = OutlookContact.RadioTelephoneNumber         ;
        contact.ReferredBy                    = OutlookContact.ReferredBy                   ;
        contact.Saved                         = OutlookContact.Saved                        ;
        contact.SelectedMailingAddress        = OutlookContact.SelectedMailingAddress       ;
        contact.Sensitivity                   = OutlookContact.Sensitivity                  ;
        contact.Size                          = OutlookContact.Size                         ;
        contact.Spouse                        = OutlookContact.Spouse                       ;
        contact.Subject                       = OutlookContact.Subject                      ;
        contact.Suffix                        = OutlookContact.Suffix                       ;
        contact.TelexNumber                   = OutlookContact.TelexNumber                  ;
        contact.Title                         = OutlookContact.Title                        ;
        contact.TTYTDDTelephoneNumber         = OutlookContact.TTYTDDTelephoneNumber        ;
        contact.UnRead                        = OutlookContact.UnRead                       ;
        contact.User1                         = OutlookContact.User1                        ;
        contact.User2                         = OutlookContact.User2                        ;
        contact.User3                         = OutlookContact.User3                        ;
        contact.User4                         = OutlookContact.User4                        ;
        contact.UserCertificate               = OutlookContact.UserCertificate              ;
        contact.WebPage                       = OutlookContact.WebPage                      ;
        contact.YomiCompanyName               = OutlookContact.YomiCompanyName              ;
        contact.YomiFirstName                 = OutlookContact.YomiFirstName                ;
        contact.YomiLastName                  = OutlookContact.YomiLastName                 ;
      //contact.ContactIndex                  = OutlookContact.ContactIndex                 ;
      //contact.Actions                       = OutlookContact.Actions                      ;
      //contact.Application                   = OutlookContact.Application                  ;
      //contact.Attachments                   = OutlookContact.Attachments                  ;
      //contact.Email1EntryID                 = OutlookContact.Email1EntryID                ;
      //contact.Email2EntryID                 = OutlookContact.Email2EntryID                ;
      //contact.Email3EntryID                 = OutlookContact.Email3EntryID                ;
      //contact.FormDescription               = OutlookContact.FormDescription              ;
      //contact.GetInspector                  = OutlookContact.GetInspector                 ;
      //contact.Parent                        = OutlookContact.Parent                       ;
      //contact.UserProperties                = OutlookContact.UserProperties               ;
    }

    public void UnStoreContact(MyContact contact, int Ndx, Microsoft.Office.Interop.Outlook._ContactItem OutlookContact)
    {
        if (contact.Account                       != null ) OutlookContact.Account                       = contact.Account                      ;
        if (contact.Anniversary                   != null ) OutlookContact.Anniversary                   = contact.Anniversary                  ;
        if (contact.AssistantName                 != null ) OutlookContact.AssistantName                 = contact.AssistantName                ;
        if (contact.AssistantTelephoneNumber      != null ) OutlookContact.AssistantTelephoneNumber      = contact.AssistantTelephoneNumber     ;
        if (contact.BillingInformation            != null ) OutlookContact.BillingInformation            = contact.BillingInformation           ;
        if (contact.Birthday                      != null ) OutlookContact.Birthday                      = contact.Birthday                     ;
        if (contact.Body                          != null ) OutlookContact.Body                          = UnEscape(contact.Body)               ;
        if (contact.Business2TelephoneNumber      != null ) OutlookContact.Business2TelephoneNumber      = contact.Business2TelephoneNumber     ;
        if (contact.BusinessAddress               != null ) OutlookContact.BusinessAddress               = UnEscape(contact.BusinessAddress)    ;
        if (contact.BusinessAddressCity           != null ) OutlookContact.BusinessAddressCity           = contact.BusinessAddressCity          ;
        if (contact.BusinessAddressCountry        != null ) OutlookContact.BusinessAddressCountry        = contact.BusinessAddressCountry       ;
        if (contact.BusinessAddressPostalCode     != null ) OutlookContact.BusinessAddressPostalCode     = contact.BusinessAddressPostalCode    ;
        if (contact.BusinessAddressPostOfficeBox  != null ) OutlookContact.BusinessAddressPostOfficeBox  = contact.BusinessAddressPostOfficeBox ;
        if (contact.BusinessAddressState          != null ) OutlookContact.BusinessAddressState          = contact.BusinessAddressState         ;
        if (contact.BusinessAddressStreet         != null ) OutlookContact.BusinessAddressStreet         = UnEscape(contact.BusinessAddressStreet);
        if (contact.BusinessFaxNumber             != null ) OutlookContact.BusinessFaxNumber             = contact.BusinessFaxNumber            ;
        if (contact.BusinessHomePage              != null ) OutlookContact.BusinessHomePage              = contact.BusinessHomePage             ;
        if (contact.BusinessTelephoneNumber       != null ) OutlookContact.BusinessTelephoneNumber       = contact.BusinessTelephoneNumber      ;
        if (contact.CallbackTelephoneNumber       != null ) OutlookContact.CallbackTelephoneNumber       = contact.CallbackTelephoneNumber      ;
        if (contact.CarTelephoneNumber            != null ) OutlookContact.CarTelephoneNumber            = contact.CarTelephoneNumber           ;
        if (contact.Categories                    != null ) OutlookContact.Categories                    = contact.Categories                   ;
        if (contact.Children                      != null ) OutlookContact.Children                      = contact.Children                     ;
        if (contact.Companies                     != null ) OutlookContact.Companies                     = contact.Companies                    ;
      //if (contact.CompanyAndFullName            != null ) OutlookContact.CompanyAndFullName            = contact.CompanyAndFullName           ;
        if (contact.CompanyMainTelephoneNumber    != null ) OutlookContact.CompanyMainTelephoneNumber    = contact.CompanyMainTelephoneNumber   ;
        if (contact.CompanyName                   != null ) OutlookContact.CompanyName                   = contact.CompanyName                  ;
        if (contact.ComputerNetworkName           != null ) OutlookContact.ComputerNetworkName           = contact.ComputerNetworkName          ;
      //if (contact.CreationTime                  != null ) OutlookContact.CreationTime                  = contact.CreationTime                 ;
        if (contact.CustomerID                    != null ) OutlookContact.CustomerID                    = contact.CustomerID                   ;
        if (contact.Department                    != null ) OutlookContact.Department                    = contact.Department                   ;
        if (contact.Email1Address                 != null ) OutlookContact.Email1Address                 = contact.Email1Address                ;
        if (contact.Email1AddressType             != null ) OutlookContact.Email1AddressType             = contact.Email1AddressType            ;
        if (contact.Email1DisplayName             != null ) OutlookContact.Email1DisplayName             = contact.Email1DisplayName            ;
        if (contact.Email2Address                 != null ) OutlookContact.Email2Address                 = contact.Email2Address                ;
        if (contact.Email2AddressType             != null ) OutlookContact.Email2AddressType             = contact.Email2AddressType            ;
        if (contact.Email2DisplayName             != null ) OutlookContact.Email2DisplayName             = contact.Email2DisplayName            ;
        if (contact.Email3Address                 != null ) OutlookContact.Email3Address                 = contact.Email3Address                ;
        if (contact.Email3AddressType             != null ) OutlookContact.Email3AddressType             = contact.Email3AddressType            ;
        if (contact.Email3DisplayName             != null ) OutlookContact.Email3DisplayName             = contact.Email3DisplayName            ;
      //if (contact.EntryID                       != null ) OutlookContact.EntryID                       = contact.EntryID                      ;
        if (contact.FileAs                        != null ) OutlookContact.FileAs                        = contact.FileAs                       ;
        if (contact.FirstName                     != null ) OutlookContact.FirstName                     = contact.FirstName                    ;
        if (contact.FTPSite                       != null ) OutlookContact.FTPSite                       = contact.FTPSite                      ;
        if (contact.FullName                      != null ) OutlookContact.FullName                      = contact.FullName                     ;
      //if (contact.FullNameAndCompany            != null ) OutlookContact.FullNameAndCompany            = contact.FullNameAndCompany           ;
                                                            OutlookContact.Gender                        = contact.Gender                       ;
        if (contact.GovernmentIDNumber            != null ) OutlookContact.GovernmentIDNumber            = contact.GovernmentIDNumber           ;
        if (contact.Hobby                         != null ) OutlookContact.Hobby                         = contact.Hobby                        ;
        if (contact.Home2TelephoneNumber          != null ) OutlookContact.Home2TelephoneNumber          = contact.Home2TelephoneNumber         ;
        if (contact.HomeAddress                   != null ) OutlookContact.HomeAddress                   = UnEscape(contact.HomeAddress)        ;
        if (contact.HomeAddressCity               != null ) OutlookContact.HomeAddressCity               = contact.HomeAddressCity              ;
        if (contact.HomeAddressCountry            != null ) OutlookContact.HomeAddressCountry            = contact.HomeAddressCountry           ;
        if (contact.HomeAddressPostalCode         != null ) OutlookContact.HomeAddressPostalCode         = contact.HomeAddressPostalCode        ;
        if (contact.HomeAddressPostOfficeBox      != null ) OutlookContact.HomeAddressPostOfficeBox      = contact.HomeAddressPostOfficeBox     ;
        if (contact.HomeAddressState              != null ) OutlookContact.HomeAddressState              = contact.HomeAddressState             ;
        if (contact.HomeAddressStreet             != null ) OutlookContact.HomeAddressStreet             = UnEscape(contact.HomeAddressStreet)  ;
        if (contact.HomeFaxNumber                 != null ) OutlookContact.HomeFaxNumber                 = contact.HomeFaxNumber                ;
        if (contact.HomeTelephoneNumber           != null ) OutlookContact.HomeTelephoneNumber           = contact.HomeTelephoneNumber          ;
                                                            OutlookContact.Importance                    = contact.Importance                   ;
        if (contact.Initials                      != null ) OutlookContact.Initials                      = contact.Initials                     ;
        if (contact.ISDNNumber                    != null ) OutlookContact.ISDNNumber                    = contact.ISDNNumber                   ;
        if (contact.JobTitle                      != null ) OutlookContact.JobTitle                      = contact.JobTitle                     ;
                                                            OutlookContact.Journal                       = contact.Journal                      ;
        if (contact.Language                      != null ) OutlookContact.Language                      = contact.Language                     ;
      //if (contact.LastModificationTime          != null ) OutlookContact.LastModificationTime          = contact.LastModificationTime         ;
        if (contact.LastName                      != null ) OutlookContact.LastName                      = contact.LastName                     ;
      //if (contact.LastNameAndFirstName          != null ) OutlookContact.LastNameAndFirstName          = contact.LastNameAndFirstName         ;
        if (contact.MailingAddress                != null ) OutlookContact.MailingAddress                = UnEscape(contact.MailingAddress)     ;
        if (contact.MailingAddressCity            != null ) OutlookContact.MailingAddressCity            = contact.MailingAddressCity           ;
        if (contact.MailingAddressCountry         != null ) OutlookContact.MailingAddressCountry         = contact.MailingAddressCountry        ;
        if (contact.MailingAddressPostalCode      != null ) OutlookContact.MailingAddressPostalCode      = contact.MailingAddressPostalCode     ;
        if (contact.MailingAddressPostOfficeBox   != null ) OutlookContact.MailingAddressPostOfficeBox   = contact.MailingAddressPostOfficeBox  ;
        if (contact.MailingAddressState           != null ) OutlookContact.MailingAddressState           = contact.MailingAddressState          ;
        if (contact.MailingAddressStreet          != null ) OutlookContact.MailingAddressStreet          = UnEscape(contact.MailingAddressStreet);
        if (contact.ManagerName                   != null ) OutlookContact.ManagerName                   = contact.ManagerName                  ;
        if (contact.MessageClass                  != null ) OutlookContact.MessageClass                  = contact.MessageClass                 ;
        if (contact.MiddleName                    != null ) OutlookContact.MiddleName                    = contact.MiddleName                   ;
        if (contact.Mileage                       != null ) OutlookContact.Mileage                       = contact.Mileage                      ;
        if (contact.MobileTelephoneNumber         != null ) OutlookContact.MobileTelephoneNumber         = contact.MobileTelephoneNumber        ;
        if (contact.NickName                      != null ) OutlookContact.NickName                      = contact.NickName                     ;
                                                            OutlookContact.NoAging                       = contact.NoAging                      ;
        if (contact.OfficeLocation                != null ) OutlookContact.OfficeLocation                = contact.OfficeLocation               ;
        if (contact.OrganizationalIDNumber        != null ) OutlookContact.OrganizationalIDNumber        = contact.OrganizationalIDNumber       ;
        if (contact.OtherAddress                  != null ) OutlookContact.OtherAddress                  = UnEscape(contact.OtherAddress)       ;
        if (contact.OtherAddressCity              != null ) OutlookContact.OtherAddressCity              = contact.OtherAddressCity             ;
        if (contact.OtherAddressCountry           != null ) OutlookContact.OtherAddressCountry           = contact.OtherAddressCountry          ;
        if (contact.OtherAddressPostalCode        != null ) OutlookContact.OtherAddressPostalCode        = contact.OtherAddressPostalCode       ;
        if (contact.OtherAddressPostOfficeBox     != null ) OutlookContact.OtherAddressPostOfficeBox     = contact.OtherAddressPostOfficeBox    ;
        if (contact.OtherAddressState             != null ) OutlookContact.OtherAddressState             = contact.OtherAddressState            ;
        if (contact.OtherAddressStreet            != null ) OutlookContact.OtherAddressStreet            = UnEscape(contact.OtherAddressStreet) ;
        if (contact.OtherFaxNumber                != null ) OutlookContact.OtherFaxNumber                = contact.OtherFaxNumber               ;
        if (contact.OtherTelephoneNumber          != null ) OutlookContact.OtherTelephoneNumber          = contact.OtherTelephoneNumber         ;
      //if (contact.OutlookInternalVersion        != null ) OutlookContact.OutlookInternalVersion        = contact.OutlookInternalVersion       ;
      //if (contact.OutlookVersion                != null ) OutlookContact.OutlookVersion                = contact.OutlookVersion               ;
        if (contact.PagerNumber                   != null ) OutlookContact.PagerNumber                   = contact.PagerNumber                  ;
        if (contact.PersonalHomePage              != null ) OutlookContact.PersonalHomePage              = contact.PersonalHomePage             ;
        if (contact.PrimaryTelephoneNumber        != null ) OutlookContact.PrimaryTelephoneNumber        = contact.PrimaryTelephoneNumber       ;
        if (contact.Profession                    != null ) OutlookContact.Profession                    = contact.Profession                   ;
        if (contact.RadioTelephoneNumber          != null ) OutlookContact.RadioTelephoneNumber          = contact.RadioTelephoneNumber         ;
        if (contact.ReferredBy                    != null ) OutlookContact.ReferredBy                    = contact.ReferredBy                   ;
      //if (contact.Saved                         != null ) OutlookContact.Saved                         = contact.Saved                        ;
                                                            OutlookContact.SelectedMailingAddress        = contact.SelectedMailingAddress       ;
                                                            OutlookContact.Sensitivity                   = contact.Sensitivity                  ;
      //if (contact.Size                          != null ) OutlookContact.Size                          = contact.Size                         ;
        if (contact.Spouse                        != null ) OutlookContact.Spouse                        = contact.Spouse                       ;
        if (contact.Subject                       != null ) OutlookContact.Subject                       = contact.Subject                      ;
        if (contact.Suffix                        != null ) OutlookContact.Suffix                        = contact.Suffix                       ;
        if (contact.TelexNumber                   != null ) OutlookContact.TelexNumber                   = contact.TelexNumber                  ;
        if (contact.Title                         != null ) OutlookContact.Title                         = contact.Title                        ;
        if (contact.TTYTDDTelephoneNumber         != null ) OutlookContact.TTYTDDTelephoneNumber         = contact.TTYTDDTelephoneNumber        ;
      //if (contact.UnRead                        != null ) OutlookContact.UnRead                        = contact.UnRead                       ;
        if (contact.User1                         != null ) OutlookContact.User1                         = contact.User1                        ;
        if (contact.User2                         != null ) OutlookContact.User2                         = contact.User2                        ;
        if (contact.User3                         != null ) OutlookContact.User3                         = contact.User3                        ;
        if (contact.User4                         != null ) OutlookContact.User4                         = contact.User4                        ;
        if (contact.UserCertificate               != null ) OutlookContact.UserCertificate               = contact.UserCertificate              ;
        if (contact.WebPage                       != null ) OutlookContact.WebPage                       = contact.WebPage                      ;
        if (contact.YomiCompanyName               != null ) OutlookContact.YomiCompanyName               = contact.YomiCompanyName              ;
        if (contact.YomiFirstName                 != null ) OutlookContact.YomiFirstName                 = contact.YomiFirstName                ;
        if (contact.YomiLastName                  != null ) OutlookContact.YomiLastName                  = contact.YomiLastName                 ;
      //if (contact.ContactIndex                  != null ) OutlookContact.ContactIndex                  = contact.ContactIndex                 ;
      //if (contact.Actions                       != null ) OutlookContact.Actions                       = contact.Actions                      ;
      //if (contact.Application                   != null ) OutlookContact.Application                   = contact.Application                  ;
      //if (contact.Attachments                   != null ) OutlookContact.Attachments                   = contact.Attachments                  ;
      //if (contact.Email1EntryID                 != null ) OutlookContact.Email1EntryID                 = contact.Email1EntryID                ;
      //if (contact.Email2EntryID                 != null ) OutlookContact.Email2EntryID                 = contact.Email2EntryID                ;
      //if (contact.Email3EntryID                 != null ) OutlookContact.Email3EntryID                 = contact.Email3EntryID                ;
      //if (contact.FormDescription               != null ) OutlookContact.FormDescription               = contact.FormDescription              ;
      //if (contact.GetInspector                  != null ) OutlookContact.GetInspector                  = contact.GetInspector                 ;
      //if (contact.Parent                        != null ) OutlookContact.Parent                        = contact.Parent                       ;
      //if (contact.UserProperties                != null ) OutlookContact.UserProperties                = contact.UserProperties               ;
    }

    public static MyContact Copy(MyContact contact)
    {
      return (MyContact) contact.MemberwiseClone();
    }

//    public MyContact MakeDuplicate(MyContact src)
//    {
//      MyContact dst = new MyContact();
      
//      //dst.Ndx = Ndx;
//      dst.OutlookNdx = 0;
//      dst.Account = src.Account;
        
//      dst.Anniversary                   = src.Anniversary                  ;
//      dst.AssistantName                 = src.AssistantName                ;
//      dst.AssistantTelephoneNumber      = src.AssistantTelephoneNumber     ;
//      dst.BillingInformation            = src.BillingInformation           ;
//      dst.Birthday                      = src.Birthday                     ;
//      dst.Body                          = src.Body                         ;
//      dst.Business2TelephoneNumber      = src.Business2TelephoneNumber     ;
//      dst.BusinessAddress               = src.BusinessAddress              ;
//      dst.BusinessAddressCity           = src.BusinessAddressCity          ;
//      dst.BusinessAddressCountry        = src.BusinessAddressCountry       ;
//      dst.BusinessAddressPostalCode     = src.BusinessAddressPostalCode    ;
//      dst.BusinessAddressPostOfficeBox  = src.BusinessAddressPostOfficeBox ;
//      dst.BusinessAddressState          = src.BusinessAddressState         ;
//      dst.BusinessAddressStreet         = src.BusinessAddressStreet        ;
//      dst.BusinessFaxNumber             = src.BusinessFaxNumber            ;
//      dst.BusinessHomePage              = src.BusinessHomePage             ;
//      dst.BusinessTelephoneNumber       = src.BusinessTelephoneNumber      ;
//      dst.CallbackTelephoneNumber       = src.CallbackTelephoneNumber      ;
//      dst.CarTelephoneNumber            = src.CarTelephoneNumber           ;
//      dst.Categories                    = src.Categories                   ;
//      dst.Children                      = src.Children                     ;
//      dst.Companies                     = src.Companies                    ;
//      dst.CompanyAndFullName            = src.CompanyAndFullName           ;
//      dst.CompanyMainTelephoneNumber    = src.CompanyMainTelephoneNumber   ;
//      dst.CompanyName                   = src.CompanyName                  ;
//      dst.ComputerNetworkName           = src.ComputerNetworkName          ;
//      dst.CreationTime                  = src.CreationTime                 ;
//      dst.CustomerID                    = src.CustomerID                   ;
//      dst.Department                    = src.Department                   ;
//      dst.Email1Address                 = src.Email1Address                ;
//      dst.Email1AddressType             = src.Email1AddressType            ;
//      dst.Email1DisplayName             = src.Email1DisplayName            ;
//      dst.Email2Address                 = src.Email2Address                ;
//      dst.Email2AddressType             = src.Email2AddressType            ;
//      dst.Email2DisplayName             = src.Email2DisplayName            ;
//      dst.Email3Address                 = src.Email3Address                ;
//      dst.Email3AddressType             = src.Email3AddressType            ;
//      dst.Email3DisplayName             = src.Email3DisplayName            ;
//      dst.EntryID                       = src.EntryID                      ;
//      dst.FileAs                        = src.FileAs                       ;
//      dst.FirstName                     = src.FirstName                    ;
//      dst.FTPSite                       = src.FTPSite                      ;
//      dst.FullName                      = src.FullName                     ;
//      dst.FullNameAndCompany            = src.FullNameAndCompany           ;
//      dst.Gender                        = src.Gender                       ;
//      dst.GovernmentIDNumber            = src.GovernmentIDNumber           ;
//      dst.Hobby                         = src.Hobby                        ;
//      dst.Home2TelephoneNumber          = src.Home2TelephoneNumber         ;
//      dst.HomeAddress                   = src.HomeAddress                  ;
//      dst.HomeAddressCity               = src.HomeAddressCity              ;
//      dst.HomeAddressCountry            = src.HomeAddressCountry           ;
//      dst.HomeAddressPostalCode         = src.HomeAddressPostalCode        ;
//      dst.HomeAddressPostOfficeBox      = src.HomeAddressPostOfficeBox     ;
//      dst.HomeAddressState              = src.HomeAddressState             ;
//      dst.HomeAddressStreet             = src.HomeAddressStreet            ;
//      dst.HomeFaxNumber                 = src.HomeFaxNumber                ;
//      dst.HomeTelephoneNumber           = src.HomeTelephoneNumber          ;
//      dst.Importance                    = src.Importance                   ;
//      dst.Initials                      = src.Initials                     ;
//      dst.ISDNNumber                    = src.ISDNNumber                   ;
//      dst.JobTitle                      = src.JobTitle                     ;
//      dst.Journal                       = src.Journal                      ;
//      dst.Language                      = src.Language                     ;
//      dst.LastModificationTime          = src.LastModificationTime         ;
//      dst.LastName                      = src.LastName                     ;
//      dst.LastNameAndFirstName          = src.LastNameAndFirstName         ;
//      dst.MailingAddress                = src.MailingAddress               ;
//      dst.MailingAddressCity            = src.MailingAddressCity           ;
//      dst.MailingAddressCountry         = src.MailingAddressCountry        ;
//      dst.MailingAddressPostalCode      = src.MailingAddressPostalCode     ;
//      dst.MailingAddressPostOfficeBox   = src.MailingAddressPostOfficeBox  ;
//      dst.MailingAddressState           = src.MailingAddressState          ;
//      dst.MailingAddressStreet          = src.MailingAddressStreet         ;
//      dst.ManagerName                   = src.ManagerName                  ;
//      dst.MessageClass                  = src.MessageClass                 ;
//      dst.MiddleName                    = src.MiddleName                   ;
//      dst.Mileage                       = src.Mileage                      ;
//      dst.MobileTelephoneNumber         = src.MobileTelephoneNumber        ;
//      dst.NickName                      = src.NickName                     ;
//      dst.NoAging                       = src.NoAging                      ;
//      dst.OfficeLocation                = src.OfficeLocation               ;
//      dst.OrganizationalIDNumber        = src.OrganizationalIDNumber       ;
//      dst.OtherAddress                  = src.OtherAddress                 ;
//      dst.OtherAddressCity              = src.OtherAddressCity             ;
//      dst.OtherAddressCountry           = src.OtherAddressCountry          ;
//      dst.OtherAddressPostalCode        = src.OtherAddressPostalCode       ;
//      dst.OtherAddressPostOfficeBox     = src.OtherAddressPostOfficeBox    ;
//      dst.OtherAddressState             = src.OtherAddressState            ;
//      dst.OtherAddressStreet            = src.OtherAddressStreet           ;
//      dst.OtherFaxNumber                = src.OtherFaxNumber               ;
//      dst.OtherTelephoneNumber          = src.OtherTelephoneNumber         ;
//      dst.OutlookInternalVersion        = src.OutlookInternalVersion       ;
//      dst.OutlookVersion                = src.OutlookVersion               ;
//      dst.PagerNumber                   = src.PagerNumber                  ;
//      dst.PersonalHomePage              = src.PersonalHomePage             ;
//      dst.PrimaryTelephoneNumber        = src.PrimaryTelephoneNumber       ;
//      dst.Profession                    = src.Profession                   ;
//      dst.RadioTelephoneNumber          = src.RadioTelephoneNumber         ;
//      dst.ReferredBy                    = src.ReferredBy                   ;
//      dst.Saved                         = src.Saved                        ;
//      dst.SelectedMailingAddress        = src.SelectedMailingAddress       ;
//      dst.Sensitivity                   = src.Sensitivity                  ;
//      dst.Size                          = src.Size                         ;
//      dst.Spouse                        = src.Spouse                       ;
//      dst.Subject                       = src.Subject                      ;
//      dst.Suffix                        = src.Suffix                       ;
//      dst.TelexNumber                   = src.TelexNumber                  ;
//      dst.Title                         = src.Title                        ;
//      dst.TTYTDDTelephoneNumber         = src.TTYTDDTelephoneNumber        ;
//      dst.UnRead                        = src.UnRead                       ;
//      dst.User1                         = src.User1                        ;
//      dst.User2                         = src.User2                        ;
//      dst.User3                         = src.User3                        ;
//      dst.User4                         = src.User4                        ;
//      dst.UserCertificate               = src.UserCertificate              ;
//      dst.WebPage                       = src.WebPage                      ;
//      dst.YomiCompanyName               = src.YomiCompanyName              ;
//      dst.YomiFirstName                 = src.YomiFirstName                ;
//      dst.YomiLastName                  = src.YomiLastName                 ;
//#if ZERO      
//      dst.ContactIndex                  = src.ContactIndex                 ;
//      dst.Actions                       = src.Actions                      ;
//      dst.Application                   = src.Application                  ;
//      dst.Attachments                   = src.Attachments                  ;
//      dst.Email1EntryID                 = src.Email1EntryID                ;
//      dst.Email2EntryID                 = src.Email2EntryID                ;
//      dst.Email3EntryID                 = src.Email3EntryID                ;
//      dst.FormDescription               = src.FormDescription              ;
//      dst.GetInspector                  = src.GetInspector                 ;
//      dst.Parent                        = src.Parent                       ;
//      dst.UserProperties                = src.UserProperties               ;
//#endif    
//      return dst;
//    }

    public bool IsIdentical_Business(MyContact contact1, MyContact contact2)
    {
      if (contact1.AssistantName                 != contact2.AssistantName                ) return false;
      if (contact1.BillingInformation            != contact2.BillingInformation           ) return false;
      if (contact1.Business2TelephoneNumber      != contact2.Business2TelephoneNumber     ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.BusinessAddress               != contact2.BusinessAddress              ) return false;
      }
      if (contact1.BusinessAddressCity           != contact2.BusinessAddressCity          ) return false;
      if (contact1.BusinessAddressCountry        != contact2.BusinessAddressCountry       ) return false;
      if (contact1.BusinessAddressPostalCode     != contact2.BusinessAddressPostalCode    ) return false;
      if (contact1.BusinessAddressPostOfficeBox  != contact2.BusinessAddressPostOfficeBox ) return false;
      if (contact1.BusinessAddressState          != contact2.BusinessAddressState         ) return false;
      if (contact1.BusinessAddressStreet         != contact2.BusinessAddressStreet        ) return false;
      if (contact1.BusinessFaxNumber             != contact2.BusinessFaxNumber            ) return false;
      if (contact1.BusinessHomePage              != contact2.BusinessHomePage             ) return false;
      if (contact1.BusinessTelephoneNumber       != contact2.BusinessTelephoneNumber      ) return false;
      if (contact1.Companies                     != contact2.Companies                    ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.CompanyAndFullName            != contact2.CompanyAndFullName           ) return false;
      }
      if (contact1.CompanyMainTelephoneNumber    != contact2.CompanyMainTelephoneNumber   ) return false;
      if (contact1.CompanyName                   != contact2.CompanyName                  ) return false;
      if (contact1.CustomerID                    != contact2.CustomerID                   ) return false;
      if (contact1.Department                    != contact2.Department                   ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.FullNameAndCompany            != contact2.FullNameAndCompany           ) return false;
      }
      if (contact1.GovernmentIDNumber            != contact2.GovernmentIDNumber           ) return false;
      if (contact1.JobTitle                      != contact2.JobTitle                     ) return false;
      if (contact1.ManagerName                   != contact2.ManagerName                  ) return false;
      if (contact1.Mileage                       != contact2.Mileage                      ) return false;
      if (contact1.OfficeLocation                != contact2.OfficeLocation               ) return false;
      if (contact1.OrganizationalIDNumber        != contact2.OrganizationalIDNumber       ) return false;
      if (contact1.Profession                    != contact2.Profession                   ) return false;
      if (contact1.ReferredBy                    != contact2.ReferredBy                   ) return false;
      if (contact1.YomiCompanyName               != contact2.YomiCompanyName              ) return false;
      if (contact1.YomiFirstName                 != contact2.YomiFirstName                ) return false;
      if (contact1.YomiLastName                  != contact2.YomiLastName                 ) return false;
      return true;
    }

    public bool IsIdentical_Email(MyContact contact1, MyContact contact2)
    {
      if (contact1.Email1Address                 != contact2.Email1Address                ) return false;
      if (contact1.Email1AddressType             != contact2.Email1AddressType            ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.Email1DisplayName             != contact2.Email1DisplayName            ) return false;
      }
      if (contact1.Email2Address                 != contact2.Email2Address                ) return false;
      if (contact1.Email2AddressType             != contact2.Email2AddressType            ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.Email2DisplayName             != contact2.Email2DisplayName            ) return false;
      }
      if (contact1.Email3Address                 != contact2.Email3Address                ) return false;
      if (contact1.Email3AddressType             != contact2.Email3AddressType            ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.Email3DisplayName             != contact2.Email3DisplayName            ) return false;
      }
      return true;
    }

    private bool PhoneNumbersAreTheSame(string number1, string number2)
    {
      string temp1;
      string temp2;
      if (number1 == null && number2==null)
        return true;
      if (number1 == null || number2==null)
        return false;
      temp1 = number1.Replace(" ","").Replace("(","").Replace(")","").Replace("-","");
      temp2 = number2.Replace(" ","").Replace("(","").Replace(")","").Replace("-","");
      if (temp1 != temp2) 
        return false;
      return true;
    }

    public bool IsIdentical_Phone(MyContact contact1, MyContact contact2)
    {
      if (!PhoneNumbersAreTheSame(contact1.AssistantTelephoneNumber , contact2.AssistantTelephoneNumber)) return false;
      if (!PhoneNumbersAreTheSame(contact1.CallbackTelephoneNumber  , contact2.CallbackTelephoneNumber )) return false;
      if (!PhoneNumbersAreTheSame(contact1.CarTelephoneNumber       , contact2.CarTelephoneNumber      )) return false;
      if (!PhoneNumbersAreTheSame(contact1.Home2TelephoneNumber     , contact2.Home2TelephoneNumber    )) return false;
      if (!PhoneNumbersAreTheSame(contact1.HomeFaxNumber            , contact2.HomeFaxNumber           )) return false;
      if (!PhoneNumbersAreTheSame(contact1.HomeTelephoneNumber      , contact2.HomeTelephoneNumber     )) return false;
      if (!PhoneNumbersAreTheSame(contact1.ISDNNumber               , contact2.ISDNNumber              )) return false;
      if (!PhoneNumbersAreTheSame(contact1.MobileTelephoneNumber    , contact2.MobileTelephoneNumber   )) return false;
      if (!PhoneNumbersAreTheSame(contact1.OtherFaxNumber           , contact2.OtherFaxNumber          )) return false;
      if (!PhoneNumbersAreTheSame(contact1.OtherTelephoneNumber     , contact2.OtherTelephoneNumber    )) return false;
      if (!PhoneNumbersAreTheSame(contact1.PagerNumber              , contact2.PagerNumber             )) return false;
      if (!PhoneNumbersAreTheSame(contact1.PrimaryTelephoneNumber   , contact2.PrimaryTelephoneNumber  )) return false;
      if (!PhoneNumbersAreTheSame(contact1.RadioTelephoneNumber     , contact2.RadioTelephoneNumber    )) return false;
      if (!PhoneNumbersAreTheSame(contact1.TelexNumber              , contact2.TelexNumber             )) return false;
      if (!PhoneNumbersAreTheSame(contact1.TTYTDDTelephoneNumber    , contact2.TTYTDDTelephoneNumber   )) return false;
      return true;                                                 
    }

    public bool IsIdentical_Names(MyContact contact1, MyContact contact2)
    {
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.FileAs                        != contact2.FileAs                       ) return false;
      }
      if (contact1.FirstName                     != contact2.FirstName                    ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.FullName                      != contact2.FullName                     ) return false;
      }
      if (contact1.Initials                      != contact2.Initials                     ) return false;
      if (contact1.LastName                      != contact2.LastName                     ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.LastNameAndFirstName          != contact2.LastNameAndFirstName         ) return false;
      }
      if (contact1.MiddleName                    != contact2.MiddleName                   ) return false;
      if (contact1.NickName                      != contact2.NickName                     ) return false;
      if (contact1.Suffix                        != contact2.Suffix                       ) return false;
      if (contact1.Title                         != contact2.Title                        ) return false;
      return true;
    }

    public bool IsIdentical_HomeAddress(MyContact contact1, MyContact contact2)
    {
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.HomeAddress                   != contact2.HomeAddress                  ) return false;
      }
      if (contact1.HomeAddressCity               != contact2.HomeAddressCity              ) return false;
      if (contact1.HomeAddressCountry            != contact2.HomeAddressCountry           ) return false;
      if (contact1.HomeAddressPostalCode         != contact2.HomeAddressPostalCode        ) return false;
      if (contact1.HomeAddressPostOfficeBox      != contact2.HomeAddressPostOfficeBox     ) return false;
      if (contact1.HomeAddressState              != contact2.HomeAddressState             ) return false;
      if (contact1.HomeAddressStreet             != contact2.HomeAddressStreet            ) return false;
      return true;
    }

    public bool IsIdentical_MailingAddress(MyContact contact1, MyContact contact2)
    {
      if (m_PerformThoroughIdenticalChecks)
      {
	      if (contact1.MailingAddress                != contact2.MailingAddress               ) return false;
	      if (contact1.MailingAddressCity            != contact2.MailingAddressCity           ) return false;
	      if (contact1.MailingAddressCountry         != contact2.MailingAddressCountry        ) return false;
	      if (contact1.MailingAddressPostalCode      != contact2.MailingAddressPostalCode     ) return false;
	      if (contact1.MailingAddressPostOfficeBox   != contact2.MailingAddressPostOfficeBox  ) return false;
	      if (contact1.MailingAddressState           != contact2.MailingAddressState          ) return false;
	      if (contact1.MailingAddressStreet          != contact2.MailingAddressStreet         ) return false;
	    }
      return true;
    }

    public bool IsIdentical_OtherAddress(MyContact contact1, MyContact contact2)
    {
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.OtherAddress                  != contact2.OtherAddress                 ) return false;
      }
      if (contact1.OtherAddressCity              != contact2.OtherAddressCity             ) return false;
      if (contact1.OtherAddressCountry           != contact2.OtherAddressCountry          ) return false;
      if (contact1.OtherAddressPostalCode        != contact2.OtherAddressPostalCode       ) return false;
      if (contact1.OtherAddressPostOfficeBox     != contact2.OtherAddressPostOfficeBox    ) return false;
      if (contact1.OtherAddressState             != contact2.OtherAddressState            ) return false;
      if (contact1.OtherAddressStreet            != contact2.OtherAddressStreet           ) return false;
      return true;
    }

    public bool IsIdentical_PersInfo(MyContact contact1, MyContact contact2)
    {
      if (contact1.Anniversary                   != contact2.Anniversary                  ) return false;
      if (contact1.Birthday                      != contact2.Birthday                     ) return false;
      if (contact1.Children                      != contact2.Children                     ) return false;
      if (contact1.Gender                        != contact2.Gender                       ) return false;
      if (contact1.Hobby                         != contact2.Hobby                        ) return false;
      if (contact1.Language                      != contact2.Language                     ) return false;
      if (contact1.Spouse                        != contact2.Spouse                       ) return false;
      return true;
    }

    public bool IsIdentical_WebPage(MyContact contact1, MyContact contact2)
    {
      if (contact1.ComputerNetworkName           != contact2.ComputerNetworkName          ) return false;
      if (contact1.FTPSite                       != contact2.FTPSite                      ) return false;
      if (contact1.PersonalHomePage              != contact2.PersonalHomePage             ) return false;
      if (contact1.WebPage                       != contact2.WebPage                      ) return false;
      return true;
    }

    public bool IsIdentical_HouseKeeping(MyContact contact1, MyContact contact2)
    {
      if (contact1.Body                          != contact2.Body                         ) return false;
      if (contact1.Categories                    != contact2.Categories                   ) return false;
      if (contact1.Importance                    != contact2.Importance                   ) return false;
      if (contact1.MessageClass                  != contact2.MessageClass                 ) return false;
      if (contact1.NoAging                       != contact2.NoAging                      ) return false;
      if (contact1.Saved                         != contact2.Saved                        ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
      if (contact1.SelectedMailingAddress        != contact2.SelectedMailingAddress       ) return false;
      }
      if (contact1.Sensitivity                   != contact2.Sensitivity                  ) return false;
      //if (contact1.Size                          != contact2.Size                         ) return false;
      if (m_PerformThoroughIdenticalChecks)
      {
        if (contact1.Subject                       != contact2.Subject                      ) return false;
      }
      if (contact1.UnRead                        != contact2.UnRead                       ) return false;
      if (contact1.User1                         != contact2.User1                        ) return false;
      if (contact1.User2                         != contact2.User2                        ) return false;
      if (contact1.User3                         != contact2.User3                        ) return false;
      if (contact1.User4                         != contact2.User4                        ) return false;
      if (contact1.UserCertificate               != contact2.UserCertificate              ) return false;

      return true;
    }

    public bool IsIdentical_Admin(MyContact contact1, MyContact contact2)
    {
      if (contact1.Size                          != contact2.Size                         ) return false;

      if (contact1.OutlookNdx                    != contact2.OutlookNdx                   ) return false;
      if (contact1.Account                       != contact2.Account                      ) return false;
      if (contact1.CreationTime                  != contact2.CreationTime                 ) return false;
      if (contact1.EntryID                       != contact2.EntryID                      ) return false;
      if (contact1.Journal                       != contact2.Journal                      ) return false;
      if (contact1.LastModificationTime          != contact2.LastModificationTime         ) return false;
      if (contact1.OutlookInternalVersion        != contact2.OutlookInternalVersion       ) return false;
      if (contact1.OutlookVersion                != contact2.OutlookVersion               ) return false;

      //if (contact1.ContactIndex                  != contact2.ContactIndex                 ) return false;
      //if (contact1.Actions                       != contact2.Actions                      ) return false;
      //if (contact1.Application                   != contact2.Application                  ) return false;
      //if (contact1.Attachments                   != contact2.Attachments                  ) return false;
      //if (contact1.Email1EntryID                 != contact2.Email1EntryID                ) return false;
      //if (contact1.Email2EntryID                 != contact2.Email2EntryID                ) return false;
      //if (contact1.Email3EntryID                 != contact2.Email3EntryID                ) return false;
      //if (contact1.FormDescription               != contact2.FormDescription              ) return false;
      //if (contact1.GetInspector                  != contact2.GetInspector                 ) return false;
      //if (contact1.Parent                        != contact2.Parent                       ) return false;
      //if (contact1.UserProperties                != contact2.UserProperties               ) return false;

      return true;
    }

    public bool IsIdentical(MyContact contact1, MyContact contact2, int CheckType)
    {
      if (CheckType == COMPARING_DIFFERENT_ITEMS)
      {
         if (contact1.Ndx == contact2.Ndx) 
           return false;
      }

      if (!IsIdentical_Business      (contact1, contact2)) return false;
      if (!IsIdentical_Email         (contact1, contact2)) return false;
      if (!IsIdentical_Phone         (contact1, contact2)) return false;
      if (!IsIdentical_Names         (contact1, contact2)) return false;
      if (!IsIdentical_HomeAddress   (contact1, contact2)) return false;
      if (!IsIdentical_MailingAddress(contact1, contact2)) return false;
      if (!IsIdentical_OtherAddress  (contact1, contact2)) return false;
      if (!IsIdentical_PersInfo      (contact1, contact2)) return false;
      if (!IsIdentical_WebPage       (contact1, contact2)) return false;
      if (!IsIdentical_HouseKeeping  (contact1, contact2)) return false;

      if (CheckType == COMPARING_SAME_ITEMS)
      {
        if (!IsIdentical_Admin (contact1, contact2)) 
          return false;
      }
      return true;
    }


    private bool StringsAreSimilar(string str1,string str2,bool Precise) 
    {
      if ((str1 == "") || (str1 == null))
        return false;
      if ((str2 == "") || (str2 == null))
        return false;

      if (Precise)
      {
        if (str1 == str2)
          return true;
      }
      else
      {
        if (str1 == str2)
          return false;

        if (str1.Length>1 && str2.Length>1 && str1.Contains(str2))
          return true;
        if (str1.Length>1 && str2.Length>1 && str2.Contains(str1))
          return true;
      }

      return false;
    }

    public bool IsSimilar(MyContact contact1, MyContact contact2, bool Precise)
    {
      if (contact1.Ndx == contact2.Ndx) return false;

      //if (!IsIdentical_Email         (contact1, contact2)) return false;
      //if (!IsIdentical_Phone         (contact1, contact2)) return false;
      //if (!IsIdentical_Names         (contact1, contact2)) return false;


      if (StringsAreSimilar(contact1.FirstName                     ,contact2.FirstName                    ,Precise))
        return true;
      if (StringsAreSimilar(contact1.FirstName                     ,contact2.LastName                     ,Precise))
        return true;
      if (StringsAreSimilar(contact1.FirstName                     ,contact2.CompanyName                  ,Precise))
        return true;
      if (StringsAreSimilar(contact1.FirstName                     ,contact2.MiddleName                   ,Precise))
        return true;
      if (StringsAreSimilar(contact1.LastName                      ,contact2.FirstName                    ,Precise))
        return true;
    //if (StringsAreSimilar(contact1.LastName                      ,contact2.LastName                     ,Precise))
    //  return true;
      if (StringsAreSimilar(contact1.LastName                      ,contact2.CompanyName                  ,Precise))
        return true;
      if (StringsAreSimilar(contact1.LastName                      ,contact2.MiddleName                   ,Precise))
        return true;
      if (StringsAreSimilar(contact1.CompanyName                   ,contact2.FirstName                    ,Precise))
        return true;
      if (StringsAreSimilar(contact1.CompanyName                   ,contact2.LastName                     ,Precise))
        return true;
    //if (StringsAreSimilar(contact1.CompanyName                   ,contact2.CompanyName                  ,Precise))
    //  return true;
      if (StringsAreSimilar(contact1.CompanyName                   ,contact2.MiddleName                   ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MiddleName                    ,contact2.FirstName                    ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MiddleName                    ,contact2.LastName                     ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MiddleName                    ,contact2.CompanyName                  ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MiddleName                    ,contact2.MiddleName                   ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Home2TelephoneNumber          ,contact2.Home2TelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Home2TelephoneNumber          ,contact2.HomeTelephoneNumber          ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Home2TelephoneNumber          ,contact2.MobileTelephoneNumber        ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Home2TelephoneNumber          ,contact2.OtherTelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Home2TelephoneNumber          ,contact2.PrimaryTelephoneNumber       ,Precise))
        return true;
      if (StringsAreSimilar(contact1.HomeTelephoneNumber           ,contact2.Home2TelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.HomeTelephoneNumber           ,contact2.HomeTelephoneNumber          ,Precise))
        return true;
      if (StringsAreSimilar(contact1.HomeTelephoneNumber           ,contact2.MobileTelephoneNumber        ,Precise))
        return true;
      if (StringsAreSimilar(contact1.HomeTelephoneNumber           ,contact2.OtherTelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.HomeTelephoneNumber           ,contact2.PrimaryTelephoneNumber       ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MobileTelephoneNumber         ,contact2.Home2TelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MobileTelephoneNumber         ,contact2.HomeTelephoneNumber          ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MobileTelephoneNumber         ,contact2.MobileTelephoneNumber        ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MobileTelephoneNumber         ,contact2.OtherTelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.MobileTelephoneNumber         ,contact2.PrimaryTelephoneNumber       ,Precise))
        return true;
      if (StringsAreSimilar(contact1.OtherTelephoneNumber          ,contact2.Home2TelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.OtherTelephoneNumber          ,contact2.HomeTelephoneNumber          ,Precise))
        return true;
      if (StringsAreSimilar(contact1.OtherTelephoneNumber          ,contact2.MobileTelephoneNumber        ,Precise))
        return true;
      if (StringsAreSimilar(contact1.OtherTelephoneNumber          ,contact2.OtherTelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.OtherTelephoneNumber          ,contact2.PrimaryTelephoneNumber       ,Precise))
        return true;
      if (StringsAreSimilar(contact1.PrimaryTelephoneNumber        ,contact2.Home2TelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.PrimaryTelephoneNumber        ,contact2.HomeTelephoneNumber          ,Precise))
        return true;
      if (StringsAreSimilar(contact1.PrimaryTelephoneNumber        ,contact2.MobileTelephoneNumber        ,Precise))
        return true;
      if (StringsAreSimilar(contact1.PrimaryTelephoneNumber        ,contact2.OtherTelephoneNumber         ,Precise))
        return true;
      if (StringsAreSimilar(contact1.PrimaryTelephoneNumber        ,contact2.PrimaryTelephoneNumber       ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email1Address                 ,contact2.Email1Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email1Address                 ,contact2.Email2Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email1Address                 ,contact2.Email3Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email2Address                 ,contact2.Email1Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email2Address                 ,contact2.Email2Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email2Address                 ,contact2.Email3Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email3Address                 ,contact2.Email1Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email3Address                 ,contact2.Email2Address                ,Precise))
        return true;
      if (StringsAreSimilar(contact1.Email3Address                 ,contact2.Email3Address                ,Precise))
        return true;

      return false;
    }

    public bool HasSameName(MyContact contact1, MyContact contact2, bool AlsoSameSuffix)
    {
      if (contact1.Ndx == contact2.Ndx) 
         return false;

      bool contact1_FirstName_IsSet   = (contact1.FirstName   != null && contact1.FirstName.Length   > 0);
      bool contact2_FirstName_IsSet   = (contact2.FirstName   != null && contact2.FirstName.Length   > 0);
      bool contact1_LastName_IsSet    = (contact1.LastName    != null && contact1.LastName.Length    > 0);
      bool contact2_LastName_IsSet    = (contact2.LastName    != null && contact2.LastName.Length    > 0);
      bool contact1_CompanyName_IsSet = (contact1.CompanyName != null && contact1.CompanyName.Length > 0);
      bool contact2_CompanyName_IsSet = (contact2.CompanyName != null && contact2.CompanyName.Length > 0);
      bool contact1_Suffix_IsSet = (contact1.Suffix != null && contact1.Suffix.Length > 0);
      bool contact2_Suffix_IsSet = (contact2.Suffix != null && contact2.Suffix.Length > 0);
      bool SuffixesAreTheSame = false;

      if (contact1_Suffix_IsSet && contact2_Suffix_IsSet && 
          contact1.Suffix == contact2.Suffix)
        SuffixesAreTheSame = true;
      else if (!contact1_Suffix_IsSet && !contact2_Suffix_IsSet)
        SuffixesAreTheSame = true;

      // If we are ignoring Suffix, then assume they are the same
      if (!AlsoSameSuffix)
        SuffixesAreTheSame = true;

      // If both First and Last names are set, then check if they are identical
      if (contact1_FirstName_IsSet && contact2_FirstName_IsSet &&
          contact1_LastName_IsSet  && contact2_LastName_IsSet  )
      {
        if (contact1.FirstName == contact2.FirstName &&
            contact1.LastName  == contact2.LastName  ) 
        {
          if (SuffixesAreTheSame)
            return true;
        }
      }
      // If only First names are set, then check if they are identical
      else if ( contact1_FirstName_IsSet &&  contact2_FirstName_IsSet &&
               !contact1_FirstName_IsSet && !contact2_FirstName_IsSet )
      {
        if (contact1.FirstName == contact2.FirstName)
          if (SuffixesAreTheSame)
            return true;
      }
      // If only Last names are set, then check if they are identical
      else if ( !contact1_FirstName_IsSet && !contact2_FirstName_IsSet &&
                 contact1_FirstName_IsSet &&  contact2_FirstName_IsSet )
      {
        if (contact1.LastName == contact2.LastName)
          if (SuffixesAreTheSame)
            return true;
      }
      // If First and Last names are not set, but Company Names are set, then check if these are identical
      else if ( !contact1_FirstName_IsSet   && !contact2_FirstName_IsSet &&
                !contact1_FirstName_IsSet   && !contact2_FirstName_IsSet &&
                 contact1_CompanyName_IsSet &&  contact2_CompanyName_IsSet )
      {
        if (contact1.CompanyName == contact2.CompanyName)
          if (SuffixesAreTheSame)
            return true;
      }
      return false;
    }

    public bool HasTruncatedNotes(MyContact contact1, MyContact contact2)
    {
    	string Temp1;
    	string Temp2;
    	
      if (contact1.Ndx == contact2.Ndx) return false;

      if (!HasSameName(contact1,contact2, false))
        return false;

      if (contact1.Body == null) return false;
      if (contact2.Body == null) return false;

      if (contact1.Body.Length == 0) return false;
      if (contact2.Body.Length == 0) return false;
      if (contact1.Body.ToString() == contact2.Body.ToString()) return false;

      // Before compare, remove trailing CRLFs
      Temp1 = contact1.Body;
      Temp2 = contact2.Body;
      if (Temp1.EndsWith("\\r\\n"))
        Temp1 = Temp1.Substring(0,Temp1.Length-4);
      if (Temp2.EndsWith("\\r\\n"))
        Temp2 = Temp2.Substring(0,Temp2.Length-4);

      // They might now be identical (we could copy one to the other)
      if (Temp1.ToString() == Temp2.ToString()) return false;

      // Now check if one is a subset of the other  (we could copy one to the other)
      if (Temp1.StartsWith(Temp2)) return true;
      if (Temp2.StartsWith(Temp1)) return true;

      //if (contact1.Body.StartsWith(contact2.Body)) return true;
      //if (contact2.Body.StartsWith(contact1.Body)) return true;

      return false;
    }

    public bool HasDislikedContent(MyContact contact1)
    {
      if (contact1.CompanyName != null && contact1.CompanyName.Contains("&")) return true;
      if (contact1.FirstName   != null && contact1.FirstName.Contains("&")) return true;
      if (contact1.LastName    != null && contact1.LastName.Contains("&")) return true;
      if (contact1.NickName    != null && contact1.NickName.Contains("&")) return true;
      if (contact1.Suffix      != null && contact1.Suffix.Contains("&")) return true;
      if (contact1.Title       != null && contact1.Title.Contains("&")) return true;
      //if (contact1.Body        != null && contact1.Body.Contains("&")) return true;

      return false;
    }

    public bool StringsAreIdentical(string x, string y)
    {
      if (StrLen(x)>0 &&
          StrLen(y)>0)
      {
        if (x == y)
          return true;
      }
      return false;
    }

    public bool HasDuplicateAddress(MyContact contact1)
    {
      if (StringsAreIdentical(contact1.BusinessAddress,contact1.HomeAddress))
          return true;
      //if (StringsAreIdentical(contact1.BusinessAddress,contact1.MailingAddress))
      //    return true;
      if (StringsAreIdentical(contact1.BusinessAddress,contact1.OtherAddress))
          return true;
      //if (StringsAreIdentical(contact1.HomeAddress,contact1.MailingAddress))
      //    return true;
      if (StringsAreIdentical(contact1.HomeAddress,contact1.OtherAddress))
          return true;
      //if (StringsAreIdentical(contact1.MailingAddress,contact1.OtherAddress))
      //    return true;

      return false;
    }

    public bool HasDuplicatePhoneNumbers(MyContact contact1)
    {
      //PrimaryTelephoneNumber
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.HomeTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.Home2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.CompanyMainTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.BusinessTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.Business2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.PrimaryTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //HomeTelephoneNumber
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.Home2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.CompanyMainTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.BusinessTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.Business2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.HomeTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //Home2TelephoneNumber
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.CompanyMainTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.BusinessTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.Business2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.Home2TelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //CompanyMainTelephoneNumber
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.BusinessTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.Business2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.CompanyMainTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //BusinessTelephoneNumber
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.Business2TelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //Business2TelephoneNumber
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.MobileTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.Business2TelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //MobileTelephoneNumber
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.OtherTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.MobileTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //OtherTelephoneNumber
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.HomeFaxNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.OtherTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //HomeFaxNumber
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.BusinessFaxNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.HomeFaxNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //BusinessFaxNumber
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.OtherFaxNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.BusinessFaxNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //OtherFaxNumber
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.AssistantTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.OtherFaxNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //AssistantTelephoneNumber
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.CallbackTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.AssistantTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //CallbackTelephoneNumber
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.CarTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.CallbackTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //CarTelephoneNumber
      if (StringsAreIdentical(contact1.CarTelephoneNumber,contact1.ISDNNumber)) return true;
      if (StringsAreIdentical(contact1.CarTelephoneNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.CarTelephoneNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.CarTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.CarTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //ISDNNumber
      if (StringsAreIdentical(contact1.ISDNNumber,contact1.PagerNumber)) return true;
      if (StringsAreIdentical(contact1.ISDNNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.ISDNNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.ISDNNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //PagerNumber
      if (StringsAreIdentical(contact1.PagerNumber,contact1.RadioTelephoneNumber)) return true;
      if (StringsAreIdentical(contact1.PagerNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.PagerNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //RadioTelephoneNumber
      if (StringsAreIdentical(contact1.RadioTelephoneNumber,contact1.TelexNumber)) return true;
      if (StringsAreIdentical(contact1.RadioTelephoneNumber,contact1.TTYTDDTelephoneNumber)) return true;

      //TelexNumber
      if (StringsAreIdentical(contact1.TelexNumber,contact1.TTYTDDTelephoneNumber)) return true;

      return false;
    }

    public bool HasDuplicateEmailAddress(MyContact contact1)
    {
      if (StringsAreIdentical(contact1.Email1Address,contact1.Email2Address)) return true;
      if (StringsAreIdentical(contact1.Email1Address,contact1.Email3Address)) return true;
      if (StringsAreIdentical(contact1.Email2Address,contact1.Email3Address)) return true;

      if (StringsAreIdentical(contact1.Email1DisplayName,contact1.Email2DisplayName)) return true;
      if (StringsAreIdentical(contact1.Email1DisplayName,contact1.Email3DisplayName)) return true;
      if (StringsAreIdentical(contact1.Email2DisplayName,contact1.Email3DisplayName)) return true;
      return false;
    }

    public bool HasMinimalInfo(MyContact contact1)
    {
      if (
          //StrLen(contact1.FirstName                     )==0 &&
          //StrLen(contact1.LastName                      )==0 &&
          //StrLen(contact1.CompanyName                   )==0 &&
          //StrLen(contact1.MiddleName                    )==0 &&
          StrLen(contact1.Home2TelephoneNumber          )==0 &&
          StrLen(contact1.HomeTelephoneNumber           )==0 &&
          StrLen(contact1.MobileTelephoneNumber         )==0 &&
          StrLen(contact1.BusinessTelephoneNumber       )==0 &&
          StrLen(contact1.OtherTelephoneNumber          )==0 &&
          StrLen(contact1.PrimaryTelephoneNumber        )==0 &&
          StrLen(contact1.Email1Address                 )==0 &&
          StrLen(contact1.Email2Address                 )==0 &&
          StrLen(contact1.Email3Address                 )==0 &&
          //StrLen(contact1.HomeAddress                   )==0 &&
          StrLen(contact1.HomeAddressStreet             )==0 &&
          StrLen(contact1.HomeAddressCity               )==0 &&
          StrLen(contact1.HomeAddressPostalCode         )==0 &&
          StrLen(contact1.HomeAddressState              )==0 &&
          //StrLen(contact1.BusinessAddress               )==0 &&
          StrLen(contact1.BusinessAddressStreet         )==0 &&
          StrLen(contact1.BusinessAddressCity           )==0 &&
          StrLen(contact1.BusinessAddressPostalCode     )==0 &&
          StrLen(contact1.BusinessAddressState          )==0 &&
          StrLen(contact1.Body                          )==0 )
        return true;

      return false;
    }

    public bool HasSame_MobileTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.MobileTelephoneNumber         != contact2.MobileTelephoneNumber        ) return false;
      return true;
    }
    public bool HasSame_HomeTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.HomeTelephoneNumber           != contact2.HomeTelephoneNumber          ) return false;
      return true;
    }
    public bool HasSame_Home2TelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.Home2TelephoneNumber          != contact2.Home2TelephoneNumber         ) return false;
      return true;
    }
    public bool HasSame_PrimaryTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.PrimaryTelephoneNumber        != contact2.PrimaryTelephoneNumber       ) return false;
      return true;
    }
    public bool HasSame_OtherTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.OtherTelephoneNumber          != contact2.OtherTelephoneNumber         ) return false;
      return true;
    }
    public bool HasSame_AssistantTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.AssistantTelephoneNumber      != contact2.AssistantTelephoneNumber     ) return false;
      return true;
    }
    public bool HasSame_CallbackTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.CallbackTelephoneNumber       != contact2.CallbackTelephoneNumber      ) return false;
      return true;
    }
    public bool HasSame_CarTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.CarTelephoneNumber            != contact2.CarTelephoneNumber           ) return false;
      return true;
    }
    public bool HasSame_HomeFaxNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.HomeFaxNumber                 != contact2.HomeFaxNumber                ) return false;
      return true;
    }
    public bool HasSame_ISDNNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.ISDNNumber                    != contact2.ISDNNumber                   ) return false;
      return true;
    }
    public bool HasSame_OtherFaxNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.OtherFaxNumber                != contact2.OtherFaxNumber               ) return false;
      return true;
    }
    public bool HasSame_PagerNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.PagerNumber                   != contact2.PagerNumber                  ) return false;
      return true;
    }
    public bool HasSame_RadioTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.RadioTelephoneNumber          != contact2.RadioTelephoneNumber         ) return false;
      return true;
    }
    public bool HasSame_TelexNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.TelexNumber                   != contact2.TelexNumber                  ) return false;
      return true;
    }
    public bool HasSame_TTYTDDTelephoneNumber(MyContact contact1, MyContact contact2)
    {
      if (contact1.Ndx == contact2.Ndx) return false;
      if (contact1.TTYTDDTelephoneNumber         != contact2.TTYTDDTelephoneNumber        ) return false;
      return true;
    }

    public int MergeEmptyFields(MyContact contact1, MyContact contact2)
    {
      int FieldsMerged = 0;
      
      if (contact1.Account                       != null && contact2.Account                       == null ) { contact2.Account                       = contact1.Account                      ; FieldsMerged++; }
      if (contact1.Anniversary                   != null && contact2.Anniversary                   == null ) { contact2.Anniversary                   = contact1.Anniversary                  ; FieldsMerged++; }
      if (contact1.AssistantName                 != null && contact2.AssistantName                 == null ) { contact2.AssistantName                 = contact1.AssistantName                ; FieldsMerged++; }
      if (contact1.AssistantTelephoneNumber      != null && contact2.AssistantTelephoneNumber      == null ) { contact2.AssistantTelephoneNumber      = contact1.AssistantTelephoneNumber     ; FieldsMerged++; }
      if (contact1.BillingInformation            != null && contact2.BillingInformation            == null ) { contact2.BillingInformation            = contact1.BillingInformation           ; FieldsMerged++; }
      if (contact1.Birthday                      != null && contact2.Birthday                      == null ) { contact2.Birthday                      = contact1.Birthday                     ; FieldsMerged++; }
      if (contact1.Body                          != null && contact2.Body                          == null ) { contact2.Body                          = contact1.Body                         ; FieldsMerged++; }
      if (contact1.Business2TelephoneNumber      != null && contact2.Business2TelephoneNumber      == null ) { contact2.Business2TelephoneNumber      = contact1.Business2TelephoneNumber     ; FieldsMerged++; }
      if (contact1.BusinessAddress               != null && contact2.BusinessAddress               == null ) { contact2.BusinessAddress               = contact1.BusinessAddress              ; FieldsMerged++; }
      if (contact1.BusinessAddressCity           != null && contact2.BusinessAddressCity           == null ) { contact2.BusinessAddressCity           = contact1.BusinessAddressCity          ; FieldsMerged++; }
      if (contact1.BusinessAddressCountry        != null && contact2.BusinessAddressCountry        == null ) { contact2.BusinessAddressCountry        = contact1.BusinessAddressCountry       ; FieldsMerged++; }
      if (contact1.BusinessAddressPostalCode     != null && contact2.BusinessAddressPostalCode     == null ) { contact2.BusinessAddressPostalCode     = contact1.BusinessAddressPostalCode    ; FieldsMerged++; }
      if (contact1.BusinessAddressPostOfficeBox  != null && contact2.BusinessAddressPostOfficeBox  == null ) { contact2.BusinessAddressPostOfficeBox  = contact1.BusinessAddressPostOfficeBox ; FieldsMerged++; }
      if (contact1.BusinessAddressState          != null && contact2.BusinessAddressState          == null ) { contact2.BusinessAddressState          = contact1.BusinessAddressState         ; FieldsMerged++; }
      if (contact1.BusinessAddressStreet         != null && contact2.BusinessAddressStreet         == null ) { contact2.BusinessAddressStreet         = contact1.BusinessAddressStreet        ; FieldsMerged++; }
      if (contact1.BusinessFaxNumber             != null && contact2.BusinessFaxNumber             == null ) { contact2.BusinessFaxNumber             = contact1.BusinessFaxNumber            ; FieldsMerged++; }
      if (contact1.BusinessHomePage              != null && contact2.BusinessHomePage              == null ) { contact2.BusinessHomePage              = contact1.BusinessHomePage             ; FieldsMerged++; }
      if (contact1.BusinessTelephoneNumber       != null && contact2.BusinessTelephoneNumber       == null ) { contact2.BusinessTelephoneNumber       = contact1.BusinessTelephoneNumber      ; FieldsMerged++; }
      if (contact1.CallbackTelephoneNumber       != null && contact2.CallbackTelephoneNumber       == null ) { contact2.CallbackTelephoneNumber       = contact1.CallbackTelephoneNumber      ; FieldsMerged++; }
      if (contact1.CarTelephoneNumber            != null && contact2.CarTelephoneNumber            == null ) { contact2.CarTelephoneNumber            = contact1.CarTelephoneNumber           ; FieldsMerged++; }
      if (contact1.Categories                    != null && contact2.Categories                    == null ) { contact2.Categories                    = contact1.Categories                   ; FieldsMerged++; }
      if (contact1.Children                      != null && contact2.Children                      == null ) { contact2.Children                      = contact1.Children                     ; FieldsMerged++; }
      if (contact1.Companies                     != null && contact2.Companies                     == null ) { contact2.Companies                     = contact1.Companies                    ; FieldsMerged++; }
      if (contact1.CompanyMainTelephoneNumber    != null && contact2.CompanyMainTelephoneNumber    == null ) { contact2.CompanyMainTelephoneNumber    = contact1.CompanyMainTelephoneNumber   ; FieldsMerged++; }
      if (contact1.CompanyName                   != null && contact2.CompanyName                   == null ) { contact2.CompanyName                   = contact1.CompanyName                  ; FieldsMerged++; }
      if (contact1.ComputerNetworkName           != null && contact2.ComputerNetworkName           == null ) { contact2.ComputerNetworkName           = contact1.ComputerNetworkName          ; FieldsMerged++; }
      if (contact1.CustomerID                    != null && contact2.CustomerID                    == null ) { contact2.CustomerID                    = contact1.CustomerID                   ; FieldsMerged++; }
      if (contact1.Department                    != null && contact2.Department                    == null ) { contact2.Department                    = contact1.Department                   ; FieldsMerged++; }
      if (contact1.Email1Address                 != null && contact2.Email1Address                 == null ) { contact2.Email1Address                 = contact1.Email1Address                ; FieldsMerged++; }
      if (contact1.Email1AddressType             != null && contact2.Email1AddressType             == null ) { contact2.Email1AddressType             = contact1.Email1AddressType            ; FieldsMerged++; }
      if (contact1.Email1DisplayName             != null && contact2.Email1DisplayName             == null ) { contact2.Email1DisplayName             = contact1.Email1DisplayName            ; FieldsMerged++; }
      if (contact1.Email2Address                 != null && contact2.Email2Address                 == null ) { contact2.Email2Address                 = contact1.Email2Address                ; FieldsMerged++; }
      if (contact1.Email2AddressType             != null && contact2.Email2AddressType             == null ) { contact2.Email2AddressType             = contact1.Email2AddressType            ; FieldsMerged++; }
      if (contact1.Email2DisplayName             != null && contact2.Email2DisplayName             == null ) { contact2.Email2DisplayName             = contact1.Email2DisplayName            ; FieldsMerged++; }
      if (contact1.Email3Address                 != null && contact2.Email3Address                 == null ) { contact2.Email3Address                 = contact1.Email3Address                ; FieldsMerged++; }
      if (contact1.Email3AddressType             != null && contact2.Email3AddressType             == null ) { contact2.Email3AddressType             = contact1.Email3AddressType            ; FieldsMerged++; }
      if (contact1.Email3DisplayName             != null && contact2.Email3DisplayName             == null ) { contact2.Email3DisplayName             = contact1.Email3DisplayName            ; FieldsMerged++; }
      if (contact1.FileAs                        != null && contact2.FileAs                        == null ) { contact2.FileAs                        = contact1.FileAs                       ; FieldsMerged++; }
      if (contact1.FirstName                     != null && contact2.FirstName                     == null ) { contact2.FirstName                     = contact1.FirstName                    ; FieldsMerged++; }
      if (contact1.FTPSite                       != null && contact2.FTPSite                       == null ) { contact2.FTPSite                       = contact1.FTPSite                      ; FieldsMerged++; }
      if (contact1.FullName                      != null && contact2.FullName                      == null ) { contact2.FullName                      = contact1.FullName                     ; FieldsMerged++; }
      if (contact1.GovernmentIDNumber            != null && contact2.GovernmentIDNumber            == null ) { contact2.GovernmentIDNumber            = contact1.GovernmentIDNumber           ; FieldsMerged++; }
      if (contact1.Hobby                         != null && contact2.Hobby                         == null ) { contact2.Hobby                         = contact1.Hobby                        ; FieldsMerged++; }
      if (contact1.Home2TelephoneNumber          != null && contact2.Home2TelephoneNumber          == null ) { contact2.Home2TelephoneNumber          = contact1.Home2TelephoneNumber         ; FieldsMerged++; }
      if (contact1.HomeAddress                   != null && contact2.HomeAddress                   == null ) { contact2.HomeAddress                   = contact1.HomeAddress                  ; FieldsMerged++; }
      if (contact1.HomeAddressCity               != null && contact2.HomeAddressCity               == null ) { contact2.HomeAddressCity               = contact1.HomeAddressCity              ; FieldsMerged++; }
      if (contact1.HomeAddressCountry            != null && contact2.HomeAddressCountry            == null ) { contact2.HomeAddressCountry            = contact1.HomeAddressCountry           ; FieldsMerged++; }
      if (contact1.HomeAddressPostalCode         != null && contact2.HomeAddressPostalCode         == null ) { contact2.HomeAddressPostalCode         = contact1.HomeAddressPostalCode        ; FieldsMerged++; }
      if (contact1.HomeAddressPostOfficeBox      != null && contact2.HomeAddressPostOfficeBox      == null ) { contact2.HomeAddressPostOfficeBox      = contact1.HomeAddressPostOfficeBox     ; FieldsMerged++; }
      if (contact1.HomeAddressState              != null && contact2.HomeAddressState              == null ) { contact2.HomeAddressState              = contact1.HomeAddressState             ; FieldsMerged++; }
      if (contact1.HomeAddressStreet             != null && contact2.HomeAddressStreet             == null ) { contact2.HomeAddressStreet             = contact1.HomeAddressStreet            ; FieldsMerged++; }
      if (contact1.HomeFaxNumber                 != null && contact2.HomeFaxNumber                 == null ) { contact2.HomeFaxNumber                 = contact1.HomeFaxNumber                ; FieldsMerged++; }
      if (contact1.HomeTelephoneNumber           != null && contact2.HomeTelephoneNumber           == null ) { contact2.HomeTelephoneNumber           = contact1.HomeTelephoneNumber          ; FieldsMerged++; }
      if (contact1.Initials                      != null && contact2.Initials                      == null ) { contact2.Initials                      = contact1.Initials                     ; FieldsMerged++; }
      if (contact1.ISDNNumber                    != null && contact2.ISDNNumber                    == null ) { contact2.ISDNNumber                    = contact1.ISDNNumber                   ; FieldsMerged++; }
      if (contact1.JobTitle                      != null && contact2.JobTitle                      == null ) { contact2.JobTitle                      = contact1.JobTitle                     ; FieldsMerged++; }
      if (contact1.Language                      != null && contact2.Language                      == null ) { contact2.Language                      = contact1.Language                     ; FieldsMerged++; }
      if (contact1.LastName                      != null && contact2.LastName                      == null ) { contact2.LastName                      = contact1.LastName                     ; FieldsMerged++; }
      if (contact1.ManagerName                   != null && contact2.ManagerName                   == null ) { contact2.ManagerName                   = contact1.ManagerName                  ; FieldsMerged++; }
      if (contact1.MessageClass                  != null && contact2.MessageClass                  == null ) { contact2.MessageClass                  = contact1.MessageClass                 ; FieldsMerged++; }
      if (contact1.MiddleName                    != null && contact2.MiddleName                    == null ) { contact2.MiddleName                    = contact1.MiddleName                   ; FieldsMerged++; }
      if (contact1.Mileage                       != null && contact2.Mileage                       == null ) { contact2.Mileage                       = contact1.Mileage                      ; FieldsMerged++; }
      if (contact1.MobileTelephoneNumber         != null && contact2.MobileTelephoneNumber         == null ) { contact2.MobileTelephoneNumber         = contact1.MobileTelephoneNumber        ; FieldsMerged++; }
      if (contact1.NickName                      != null && contact2.NickName                      == null ) { contact2.NickName                      = contact1.NickName                     ; FieldsMerged++; }
      if (contact1.OfficeLocation                != null && contact2.OfficeLocation                == null ) { contact2.OfficeLocation                = contact1.OfficeLocation               ; FieldsMerged++; }
      if (contact1.OrganizationalIDNumber        != null && contact2.OrganizationalIDNumber        == null ) { contact2.OrganizationalIDNumber        = contact1.OrganizationalIDNumber       ; FieldsMerged++; }
      if (contact1.OtherAddress                  != null && contact2.OtherAddress                  == null ) { contact2.OtherAddress                  = contact1.OtherAddress                 ; FieldsMerged++; }
      if (contact1.OtherAddressCity              != null && contact2.OtherAddressCity              == null ) { contact2.OtherAddressCity              = contact1.OtherAddressCity             ; FieldsMerged++; }
      if (contact1.OtherAddressCountry           != null && contact2.OtherAddressCountry           == null ) { contact2.OtherAddressCountry           = contact1.OtherAddressCountry          ; FieldsMerged++; }
      if (contact1.OtherAddressPostalCode        != null && contact2.OtherAddressPostalCode        == null ) { contact2.OtherAddressPostalCode        = contact1.OtherAddressPostalCode       ; FieldsMerged++; }
      if (contact1.OtherAddressPostOfficeBox     != null && contact2.OtherAddressPostOfficeBox     == null ) { contact2.OtherAddressPostOfficeBox     = contact1.OtherAddressPostOfficeBox    ; FieldsMerged++; }
      if (contact1.OtherAddressState             != null && contact2.OtherAddressState             == null ) { contact2.OtherAddressState             = contact1.OtherAddressState            ; FieldsMerged++; }
      if (contact1.OtherAddressStreet            != null && contact2.OtherAddressStreet            == null ) { contact2.OtherAddressStreet            = contact1.OtherAddressStreet           ; FieldsMerged++; }
      if (contact1.OtherFaxNumber                != null && contact2.OtherFaxNumber                == null ) { contact2.OtherFaxNumber                = contact1.OtherFaxNumber               ; FieldsMerged++; }
      if (contact1.OtherTelephoneNumber          != null && contact2.OtherTelephoneNumber          == null ) { contact2.OtherTelephoneNumber          = contact1.OtherTelephoneNumber         ; FieldsMerged++; }
      if (contact1.PagerNumber                   != null && contact2.PagerNumber                   == null ) { contact2.PagerNumber                   = contact1.PagerNumber                  ; FieldsMerged++; }
      if (contact1.PersonalHomePage              != null && contact2.PersonalHomePage              == null ) { contact2.PersonalHomePage              = contact1.PersonalHomePage             ; FieldsMerged++; }
      if (contact1.PrimaryTelephoneNumber        != null && contact2.PrimaryTelephoneNumber        == null ) { contact2.PrimaryTelephoneNumber        = contact1.PrimaryTelephoneNumber       ; FieldsMerged++; }
      if (contact1.Profession                    != null && contact2.Profession                    == null ) { contact2.Profession                    = contact1.Profession                   ; FieldsMerged++; }
      if (contact1.RadioTelephoneNumber          != null && contact2.RadioTelephoneNumber          == null ) { contact2.RadioTelephoneNumber          = contact1.RadioTelephoneNumber         ; FieldsMerged++; }
      if (contact1.ReferredBy                    != null && contact2.ReferredBy                    == null ) { contact2.ReferredBy                    = contact1.ReferredBy                   ; FieldsMerged++; }
      if (contact1.Spouse                        != null && contact2.Spouse                        == null ) { contact2.Spouse                        = contact1.Spouse                       ; FieldsMerged++; }
      if (contact1.Subject                       != null && contact2.Subject                       == null ) { contact2.Subject                       = contact1.Subject                      ; FieldsMerged++; }
      if (contact1.Suffix                        != null && contact2.Suffix                        == null ) { contact2.Suffix                        = contact1.Suffix                       ; FieldsMerged++; }
      if (contact1.TelexNumber                   != null && contact2.TelexNumber                   == null ) { contact2.TelexNumber                   = contact1.TelexNumber                  ; FieldsMerged++; }
      if (contact1.Title                         != null && contact2.Title                         == null ) { contact2.Title                         = contact1.Title                        ; FieldsMerged++; }
      if (contact1.TTYTDDTelephoneNumber         != null && contact2.TTYTDDTelephoneNumber         == null ) { contact2.TTYTDDTelephoneNumber         = contact1.TTYTDDTelephoneNumber        ; FieldsMerged++; }
      if (contact1.User1                         != null && contact2.User1                         == null ) { contact2.User1                         = contact1.User1                        ; FieldsMerged++; }
      if (contact1.User2                         != null && contact2.User2                         == null ) { contact2.User2                         = contact1.User2                        ; FieldsMerged++; }
      if (contact1.User3                         != null && contact2.User3                         == null ) { contact2.User3                         = contact1.User3                        ; FieldsMerged++; }
      if (contact1.User4                         != null && contact2.User4                         == null ) { contact2.User4                         = contact1.User4                        ; FieldsMerged++; }
      if (contact1.UserCertificate               != null && contact2.UserCertificate               == null ) { contact2.UserCertificate               = contact1.UserCertificate              ; FieldsMerged++; }
      if (contact1.WebPage                       != null && contact2.WebPage                       == null ) { contact2.WebPage                       = contact1.WebPage                      ; FieldsMerged++; }
      if (contact1.YomiCompanyName               != null && contact2.YomiCompanyName               == null ) { contact2.YomiCompanyName               = contact1.YomiCompanyName              ; FieldsMerged++; }
      if (contact1.YomiFirstName                 != null && contact2.YomiFirstName                 == null ) { contact2.YomiFirstName                 = contact1.YomiFirstName                ; FieldsMerged++; }
      if (contact1.YomiLastName                  != null && contact2.YomiLastName                  == null ) { contact2.YomiLastName                  = contact1.YomiLastName                 ; FieldsMerged++; }
/*
      if (contact1.MailingAddress                != null && contact1.MailingAddress                == null ) contact2.MailingAddress                = contact1.MailingAddress               ;
      if (contact1.MailingAddressCity            != null && contact1.MailingAddressCity            == null ) contact2.MailingAddressCity            = contact1.MailingAddressCity           ;
      if (contact1.MailingAddressCountry         != null && contact1.MailingAddressCountry         == null ) contact2.MailingAddressCountry         = contact1.MailingAddressCountry        ;
      if (contact1.MailingAddressPostalCode      != null && contact1.MailingAddressPostalCode      == null ) contact2.MailingAddressPostalCode      = contact1.MailingAddressPostalCode     ;
      if (contact1.MailingAddressPostOfficeBox   != null && contact1.MailingAddressPostOfficeBox   == null ) contact2.MailingAddressPostOfficeBox   = contact1.MailingAddressPostOfficeBox  ;
      if (contact1.MailingAddressState           != null && contact1.MailingAddressState           == null ) contact2.MailingAddressState           = contact1.MailingAddressState          ;
      if (contact1.MailingAddressStreet          != null && contact1.MailingAddressStreet          == null ) contact2.MailingAddressStreet          = contact1.MailingAddressStreet         ;
                                                          contact2.Gender                        = contact1.Gender                       ;
                                                          contact2.Importance                    = contact1.Importance                   ;
                                                          contact2.Journal                       = contact1.Journal                      ;
                                                          contact2.NoAging                       = contact1.NoAging                      ;
                                                          contact2.SelectedMailingAddress        = contact1.SelectedMailingAddress       ;
                                                          contact2.Sensitivity                   = contact1.Sensitivity                  ;
    //if (contact1.CompanyAndFullName            != null ) contact2.CompanyAndFullName            = contact1.CompanyAndFullName           ;
    //if (contact1.CreationTime                  != null ) contact2.CreationTime                  = contact1.CreationTime                 ;
    //if (contact1.EntryID                       != null ) contact2.EntryID                       = contact1.EntryID                      ;
    //if (contact1.FullNameAndCompany            != null ) contact2.FullNameAndCompany            = contact1.FullNameAndCompany           ;
    //if (contact1.LastModificationTime          != null ) contact2.LastModificationTime          = contact1.LastModificationTime         ;
    //if (contact1.LastNameAndFirstName          != null ) contact2.LastNameAndFirstName          = contact1.LastNameAndFirstName         ;
    //if (contact1.OutlookInternalVersion        != null ) contact2.OutlookInternalVersion        = contact1.OutlookInternalVersion       ;
    //if (contact1.OutlookVersion                != null ) contact2.OutlookVersion                = contact1.OutlookVersion               ;
    //if (contact1.Saved                         != null ) contact2.Saved                         = contact1.Saved                        ;
    //if (contact1.Size                          != null ) contact2.Size                          = contact1.Size                         ;
    //if (contact1.UnRead                        != null ) contact2.UnRead                        = contact1.UnRead                       ;
    //if (contact1.ContactIndex                  != null ) contact2.ContactIndex                  = contact1.ContactIndex                 ;
    //if (contact1.Actions                       != null ) contact2.Actions                       = contact1.Actions                      ;
    //if (contact1.Application                   != null ) contact2.Application                   = contact1.Application                  ;
    //if (contact1.Attachments                   != null ) contact2.Attachments                   = contact1.Attachments                  ;
    //if (contact1.Email1EntryID                 != null ) contact2.Email1EntryID                 = contact1.Email1EntryID                ;
    //if (contact1.Email2EntryID                 != null ) contact2.Email2EntryID                 = contact1.Email2EntryID                ;
    //if (contact1.Email3EntryID                 != null ) contact2.Email3EntryID                 = contact1.Email3EntryID                ;
    //if (contact1.FormDescription               != null ) contact2.FormDescription               = contact1.FormDescription              ;
    //if (contact1.GetInspector                  != null ) contact2.GetInspector                  = contact1.GetInspector                 ;
    //if (contact1.Parent                        != null ) contact2.Parent                        = contact1.Parent                       ;
    //if (contact1.UserProperties                != null ) contact2.UserProperties                = contact1.UserProperties               ;
*/      
      return FieldsMerged;
    }

    public int MergeContacts(MyContact contact1, MyContact contact2)
    {
      int FieldsMerged = 0;
      FieldsMerged += MergeEmptyFields(contact1, contact2);
      FieldsMerged += MergeEmptyFields(contact2, contact1);
      return FieldsMerged;
    }
  }
}
