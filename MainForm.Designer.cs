namespace OutlookContactsAnalyser
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
            this.dgvContacts = new System.Windows.Forms.DataGridView();
            this.RowSelect = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnViewFields = new System.Windows.Forms.Button();
            this.btnContactDetails = new System.Windows.Forms.Button();
            this.btnAnalyse = new System.Windows.Forms.Button();
            this.btnRefreshFolders = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.cbFolder = new System.Windows.Forms.ComboBox();
            this.btnGetContacts = new System.Windows.Forms.Button();
            this.txtMaxContactstoRead = new System.Windows.Forms.TextBox();
            this.lblMaxContactstoRead = new System.Windows.Forms.Label();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBoxViewRadioButtons = new System.Windows.Forms.GroupBox();
            this.radioViewMinimalInfo = new System.Windows.Forms.RadioButton();
            this.radioViewDislikedContents = new System.Windows.Forms.RadioButton();
            this.radioViewMismatchedAddrAndComp = new System.Windows.Forms.RadioButton();
            this.radioViewEmptyMailingAddresses = new System.Windows.Forms.RadioButton();
            this.radioViewLinkedMailingAddresses = new System.Windows.Forms.RadioButton();
            this.radioViewUnlinkedMailingAddresses = new System.Windows.Forms.RadioButton();
            this.radioViewTruncatedNotes = new System.Windows.Forms.RadioButton();
            this.radioViewDuplicateEmails = new System.Windows.Forms.RadioButton();
            this.radioViewDuplicatePhoneNumbers = new System.Windows.Forms.RadioButton();
            this.radioViewDuplicateAddresses = new System.Windows.Forms.RadioButton();
            this.radioViewSameNameAndSuffix = new System.Windows.Forms.RadioButton();
            this.radioViewSameName = new System.Windows.Forms.RadioButton();
            this.radioViewVaguelySimilar = new System.Windows.Forms.RadioButton();
            this.radioViewSimilar = new System.Windows.Forms.RadioButton();
            this.radioViewDuplicates = new System.Windows.Forms.RadioButton();
            this.radioViewAllByName = new System.Windows.Forms.RadioButton();
            this.radioViewAll = new System.Windows.Forms.RadioButton();
            this.txtStatusLine = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileNewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileOpenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.fileSaveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileSaveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.filePrintToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.filePrintPreviewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.fileExitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileRefreshFoldersToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileLoadContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileSaveChangedContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fileStopGetContactsThreadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowAllContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowDuplicateContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowSimilarContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowVaguelySimilarContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowWithTheSameNameToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowWithTruncatedNotesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowWithDislikedContentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowWithDuplicateAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showWithDuplicatePhoneNumbersToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.showWithDuplicateEmailAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.viewHideUnusedColumnsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mailingAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowLinkedMailingAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowEmptyMailingAddressesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewShowMismatchedAddressComponentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editUndoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editRedoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.editCutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editCopyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.editPasteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
            this.editSelectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsCustomizeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsOptionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsAnalyseContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsEditContactDetailsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsFieldsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsNewContactToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cleanNotesFieldsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.markDuplicatesForDeletionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fixPhoneNumbersToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mergeDuplicateContactsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpContentsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpIndexToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpSearchToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
            this.helpAboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ndxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fullNameAndCompanyDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.titleDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lastNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.middleNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.firstNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.suffixDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.nickNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.companyNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fullNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lastNameAndFirstNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.companyAndFullNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.primaryTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mobileTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.home2TelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.business2TelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeFaxNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessFaxNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherFaxNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.callbackTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.carTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pagerNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.iSDNNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.telexNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.accountDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.anniversaryDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.assistantNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.assistantTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.billingInformationDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.birthdayDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressCityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressCountryDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressPostalCodeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressStateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessAddressStreetDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.businessHomePageDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.categoriesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.childrenDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.companiesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.computerNetworkNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.creationTimeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.customerIDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.departmentDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email1AddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email1AddressTypeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email1DisplayNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email2AddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email2AddressTypeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email2DisplayNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email3AddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email3AddressTypeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.email3DisplayNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.entryIDDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fileAsDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fTPSiteDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.governmentIDNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.hobbyDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressCityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressCountryDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressPostalCodeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressStateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.homeAddressStreetDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.initialsDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.jobTitleDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.journalDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.languageDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lastModificationTimeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressCityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressCountryDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressStateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mailingAddressStreetDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.managerNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.messageClassDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mileageDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.noAgingDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.officeLocationDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.organizationalIDNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressCityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressCountryDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressPostalCodeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressStateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.otherAddressStreetDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.outlookInternalVersionDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.outlookVersionDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.personalHomePageDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.professionDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.radioTelephoneNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.referredByDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.savedDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.sizeDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.spouseDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.subjectDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.unReadDataGridViewCheckBoxColumn = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.user1DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.user2DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.user3DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.user4DataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.userCertificateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.webPageDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yomiCompanyNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yomiFirstNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yomiLastNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.genderDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.importanceDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.selectedMailingAddressDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sensitivityDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.outlookNdxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bodyDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.myContactBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.tlpMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvContacts)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.groupBoxViewRadioButtons.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.myContactBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // tlpMain
            // 
            this.tlpMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tlpMain.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tlpMain.ColumnCount = 1;
            this.tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlpMain.Controls.Add(this.dgvContacts, 0, 1);
            this.tlpMain.Controls.Add(this.tableLayoutPanel1, 0, 0);
            this.tlpMain.Controls.Add(this.tableLayoutPanel2, 0, 2);
            this.tlpMain.Controls.Add(this.txtStatusLine, 0, 3);
            this.tlpMain.Location = new System.Drawing.Point(12, 27);
            this.tlpMain.Name = "tlpMain";
            this.tlpMain.RowCount = 4;
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7.829978F));
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 52.75459F));
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 34.39065F));
            this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 4.778761F));
            this.tlpMain.Size = new System.Drawing.Size(857, 599);
            this.tlpMain.TabIndex = 0;
            // 
            // dgvContacts
            // 
            this.dgvContacts.AllowUserToAddRows = false;
            this.dgvContacts.AllowUserToDeleteRows = false;
            this.dgvContacts.AllowUserToOrderColumns = true;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.dgvContacts.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvContacts.AutoGenerateColumns = false;
            this.dgvContacts.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgvContacts.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvContacts.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.NavajoWhite;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvContacts.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvContacts.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvContacts.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ndxDataGridViewTextBoxColumn,
            this.fullNameAndCompanyDataGridViewTextBoxColumn,
            this.RowSelect,
            this.titleDataGridViewTextBoxColumn,
            this.lastNameDataGridViewTextBoxColumn,
            this.middleNameDataGridViewTextBoxColumn,
            this.firstNameDataGridViewTextBoxColumn,
            this.suffixDataGridViewTextBoxColumn,
            this.nickNameDataGridViewTextBoxColumn,
            this.companyNameDataGridViewTextBoxColumn,
            this.fullNameDataGridViewTextBoxColumn,
            this.lastNameAndFirstNameDataGridViewTextBoxColumn,
            this.companyAndFullNameDataGridViewTextBoxColumn,
            this.primaryTelephoneNumberDataGridViewTextBoxColumn,
            this.mobileTelephoneNumberDataGridViewTextBoxColumn,
            this.homeTelephoneNumberDataGridViewTextBoxColumn,
            this.home2TelephoneNumberDataGridViewTextBoxColumn,
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn,
            this.businessTelephoneNumberDataGridViewTextBoxColumn,
            this.business2TelephoneNumberDataGridViewTextBoxColumn,
            this.otherTelephoneNumberDataGridViewTextBoxColumn,
            this.homeFaxNumberDataGridViewTextBoxColumn,
            this.businessFaxNumberDataGridViewTextBoxColumn,
            this.otherFaxNumberDataGridViewTextBoxColumn,
            this.callbackTelephoneNumberDataGridViewTextBoxColumn,
            this.carTelephoneNumberDataGridViewTextBoxColumn,
            this.pagerNumberDataGridViewTextBoxColumn,
            this.iSDNNumberDataGridViewTextBoxColumn,
            this.telexNumberDataGridViewTextBoxColumn,
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn,
            this.accountDataGridViewTextBoxColumn,
            this.anniversaryDataGridViewTextBoxColumn,
            this.assistantNameDataGridViewTextBoxColumn,
            this.assistantTelephoneNumberDataGridViewTextBoxColumn,
            this.billingInformationDataGridViewTextBoxColumn,
            this.birthdayDataGridViewTextBoxColumn,
            this.businessAddressDataGridViewTextBoxColumn,
            this.businessAddressCityDataGridViewTextBoxColumn,
            this.businessAddressCountryDataGridViewTextBoxColumn,
            this.businessAddressPostalCodeDataGridViewTextBoxColumn,
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn,
            this.businessAddressStateDataGridViewTextBoxColumn,
            this.businessAddressStreetDataGridViewTextBoxColumn,
            this.businessHomePageDataGridViewTextBoxColumn,
            this.categoriesDataGridViewTextBoxColumn,
            this.childrenDataGridViewTextBoxColumn,
            this.companiesDataGridViewTextBoxColumn,
            this.computerNetworkNameDataGridViewTextBoxColumn,
            this.creationTimeDataGridViewTextBoxColumn,
            this.customerIDDataGridViewTextBoxColumn,
            this.departmentDataGridViewTextBoxColumn,
            this.email1AddressDataGridViewTextBoxColumn,
            this.email1AddressTypeDataGridViewTextBoxColumn,
            this.email1DisplayNameDataGridViewTextBoxColumn,
            this.email2AddressDataGridViewTextBoxColumn,
            this.email2AddressTypeDataGridViewTextBoxColumn,
            this.email2DisplayNameDataGridViewTextBoxColumn,
            this.email3AddressDataGridViewTextBoxColumn,
            this.email3AddressTypeDataGridViewTextBoxColumn,
            this.email3DisplayNameDataGridViewTextBoxColumn,
            this.entryIDDataGridViewTextBoxColumn,
            this.fileAsDataGridViewTextBoxColumn,
            this.fTPSiteDataGridViewTextBoxColumn,
            this.governmentIDNumberDataGridViewTextBoxColumn,
            this.hobbyDataGridViewTextBoxColumn,
            this.homeAddressDataGridViewTextBoxColumn,
            this.homeAddressCityDataGridViewTextBoxColumn,
            this.homeAddressCountryDataGridViewTextBoxColumn,
            this.homeAddressPostalCodeDataGridViewTextBoxColumn,
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn,
            this.homeAddressStateDataGridViewTextBoxColumn,
            this.homeAddressStreetDataGridViewTextBoxColumn,
            this.initialsDataGridViewTextBoxColumn,
            this.jobTitleDataGridViewTextBoxColumn,
            this.journalDataGridViewCheckBoxColumn,
            this.languageDataGridViewTextBoxColumn,
            this.lastModificationTimeDataGridViewTextBoxColumn,
            this.mailingAddressDataGridViewTextBoxColumn,
            this.mailingAddressCityDataGridViewTextBoxColumn,
            this.mailingAddressCountryDataGridViewTextBoxColumn,
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn,
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn,
            this.mailingAddressStateDataGridViewTextBoxColumn,
            this.mailingAddressStreetDataGridViewTextBoxColumn,
            this.managerNameDataGridViewTextBoxColumn,
            this.messageClassDataGridViewTextBoxColumn,
            this.mileageDataGridViewTextBoxColumn,
            this.noAgingDataGridViewCheckBoxColumn,
            this.officeLocationDataGridViewTextBoxColumn,
            this.organizationalIDNumberDataGridViewTextBoxColumn,
            this.otherAddressDataGridViewTextBoxColumn,
            this.otherAddressCityDataGridViewTextBoxColumn,
            this.otherAddressCountryDataGridViewTextBoxColumn,
            this.otherAddressPostalCodeDataGridViewTextBoxColumn,
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn,
            this.otherAddressStateDataGridViewTextBoxColumn,
            this.otherAddressStreetDataGridViewTextBoxColumn,
            this.outlookInternalVersionDataGridViewTextBoxColumn,
            this.outlookVersionDataGridViewTextBoxColumn,
            this.personalHomePageDataGridViewTextBoxColumn,
            this.professionDataGridViewTextBoxColumn,
            this.radioTelephoneNumberDataGridViewTextBoxColumn,
            this.referredByDataGridViewTextBoxColumn,
            this.savedDataGridViewCheckBoxColumn,
            this.sizeDataGridViewTextBoxColumn,
            this.spouseDataGridViewTextBoxColumn,
            this.subjectDataGridViewTextBoxColumn,
            this.unReadDataGridViewCheckBoxColumn,
            this.user1DataGridViewTextBoxColumn,
            this.user2DataGridViewTextBoxColumn,
            this.user3DataGridViewTextBoxColumn,
            this.user4DataGridViewTextBoxColumn,
            this.userCertificateDataGridViewTextBoxColumn,
            this.webPageDataGridViewTextBoxColumn,
            this.yomiCompanyNameDataGridViewTextBoxColumn,
            this.yomiFirstNameDataGridViewTextBoxColumn,
            this.yomiLastNameDataGridViewTextBoxColumn,
            this.genderDataGridViewTextBoxColumn,
            this.importanceDataGridViewTextBoxColumn,
            this.selectedMailingAddressDataGridViewTextBoxColumn,
            this.sensitivityDataGridViewTextBoxColumn,
            this.outlookNdxDataGridViewTextBoxColumn,
            this.bodyDataGridViewTextBoxColumn});
            this.dgvContacts.DataSource = this.myContactBindingSource;
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Info;
            dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvContacts.DefaultCellStyle = dataGridViewCellStyle3;
            this.dgvContacts.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvContacts.Location = new System.Drawing.Point(3, 50);
            this.dgvContacts.Name = "dgvContacts";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvContacts.RowHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dgvContacts.RowHeadersVisible = false;
            this.dgvContacts.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvContacts.Size = new System.Drawing.Size(851, 310);
            this.dgvContacts.TabIndex = 0;
            this.dgvContacts.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dgvContacts_KeyUp);
            // 
            // RowSelect
            // 
            this.RowSelect.Frozen = true;
            this.RowSelect.HeaderText = "";
            this.RowSelect.Name = "RowSelect";
            this.RowSelect.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.RowSelect.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.RowSelect.Width = 19;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 9;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.625F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 59.375F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 107F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 88F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 65F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 112F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 91F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 94F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102F));
            this.tableLayoutPanel1.Controls.Add(this.btnViewFields, 8, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnContactDetails, 7, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnAnalyse, 6, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnRefreshFolders, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.cbFolder, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnGetContacts, 5, 0);
            this.tableLayoutPanel1.Controls.Add(this.txtMaxContactstoRead, 4, 0);
            this.tableLayoutPanel1.Controls.Add(this.lblMaxContactstoRead, 3, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 35.48387F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(851, 26);
            this.tableLayoutPanel1.TabIndex = 15;
            // 
            // btnViewFields
            // 
            this.btnViewFields.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnViewFields.Enabled = false;
            this.btnViewFields.Location = new System.Drawing.Point(752, 3);
            this.btnViewFields.Name = "btnViewFields";
            this.btnViewFields.Size = new System.Drawing.Size(96, 20);
            this.btnViewFields.TabIndex = 3;
            this.btnViewFields.Text = "Fields";
            this.btnViewFields.UseVisualStyleBackColor = true;
            this.btnViewFields.Click += new System.EventHandler(this.btnViewFields_Click);
            // 
            // btnContactDetails
            // 
            this.btnContactDetails.Enabled = false;
            this.btnContactDetails.Location = new System.Drawing.Point(658, 3);
            this.btnContactDetails.Name = "btnContactDetails";
            this.btnContactDetails.Size = new System.Drawing.Size(88, 20);
            this.btnContactDetails.TabIndex = 19;
            this.btnContactDetails.Text = "Contact Details";
            this.btnContactDetails.UseVisualStyleBackColor = true;
            this.btnContactDetails.Click += new System.EventHandler(this.btnContactDetails_Click);
            // 
            // btnAnalyse
            // 
            this.btnAnalyse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAnalyse.BackColor = System.Drawing.SystemColors.Control;
            this.btnAnalyse.Enabled = false;
            this.btnAnalyse.Location = new System.Drawing.Point(567, 3);
            this.btnAnalyse.Name = "btnAnalyse";
            this.btnAnalyse.Size = new System.Drawing.Size(85, 20);
            this.btnAnalyse.TabIndex = 2;
            this.btnAnalyse.Text = "Analyse";
            this.btnAnalyse.UseVisualStyleBackColor = false;
            this.btnAnalyse.Click += new System.EventHandler(this.btnAnalyse_Click);
            // 
            // btnRefreshFolders
            // 
            this.btnRefreshFolders.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRefreshFolders.BackColor = System.Drawing.SystemColors.Control;
            this.btnRefreshFolders.Location = new System.Drawing.Point(195, 3);
            this.btnRefreshFolders.Name = "btnRefreshFolders";
            this.btnRefreshFolders.Size = new System.Drawing.Size(101, 20);
            this.btnRefreshFolders.TabIndex = 1;
            this.btnRefreshFolders.Text = "Refresh Folders";
            this.btnRefreshFolders.UseVisualStyleBackColor = false;
            this.btnRefreshFolders.Click += new System.EventHandler(this.btnRefreshFolders_Click);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Folder";
            // 
            // cbFolder
            // 
            this.cbFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.cbFolder.FormattingEnabled = true;
            this.cbFolder.Location = new System.Drawing.Point(81, 3);
            this.cbFolder.Name = "cbFolder";
            this.cbFolder.Size = new System.Drawing.Size(108, 21);
            this.cbFolder.TabIndex = 0;
            // 
            // btnGetContacts
            // 
            this.btnGetContacts.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGetContacts.BackColor = System.Drawing.SystemColors.Control;
            this.btnGetContacts.Location = new System.Drawing.Point(455, 3);
            this.btnGetContacts.Name = "btnGetContacts";
            this.btnGetContacts.Size = new System.Drawing.Size(106, 20);
            this.btnGetContacts.TabIndex = 0;
            this.btnGetContacts.Text = "Get Contacts";
            this.btnGetContacts.UseVisualStyleBackColor = false;
            this.btnGetContacts.Click += new System.EventHandler(this.btnGetContacts_Click);
            // 
            // txtMaxContactstoRead
            // 
            this.txtMaxContactstoRead.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMaxContactstoRead.Location = new System.Drawing.Point(390, 3);
            this.txtMaxContactstoRead.Name = "txtMaxContactstoRead";
            this.txtMaxContactstoRead.Size = new System.Drawing.Size(59, 20);
            this.txtMaxContactstoRead.TabIndex = 1;
            // 
            // lblMaxContactstoRead
            // 
            this.lblMaxContactstoRead.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.lblMaxContactstoRead.AutoSize = true;
            this.lblMaxContactstoRead.Location = new System.Drawing.Point(302, 6);
            this.lblMaxContactstoRead.Name = "lblMaxContactstoRead";
            this.lblMaxContactstoRead.Size = new System.Drawing.Size(82, 13);
            this.lblMaxContactstoRead.TabIndex = 10;
            this.lblMaxContactstoRead.Text = "Max Contacts";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 3;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 350F));
            this.tableLayoutPanel2.Controls.Add(this.txtLog, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.groupBoxViewRadioButtons, 2, 0);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(3, 366);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 1;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(851, 200);
            this.tableLayoutPanel2.TabIndex = 8;
            // 
            // txtLog
            // 
            this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLog.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLog.Location = new System.Drawing.Point(34, 3);
            this.txtLog.Multiline = true;
            this.txtLog.Name = "txtLog";
            this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLog.Size = new System.Drawing.Size(464, 194);
            this.txtLog.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Log";
            // 
            // groupBoxViewRadioButtons
            // 
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewMinimalInfo);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewDislikedContents);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewMismatchedAddrAndComp);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewEmptyMailingAddresses);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewLinkedMailingAddresses);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewUnlinkedMailingAddresses);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewTruncatedNotes);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewDuplicateEmails);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewDuplicatePhoneNumbers);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewDuplicateAddresses);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewSameNameAndSuffix);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewSameName);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewVaguelySimilar);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewSimilar);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewDuplicates);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewAllByName);
            this.groupBoxViewRadioButtons.Controls.Add(this.radioViewAll);
            this.groupBoxViewRadioButtons.Location = new System.Drawing.Point(504, 3);
            this.groupBoxViewRadioButtons.Name = "groupBoxViewRadioButtons";
            this.groupBoxViewRadioButtons.Size = new System.Drawing.Size(344, 194);
            this.groupBoxViewRadioButtons.TabIndex = 5;
            this.groupBoxViewRadioButtons.TabStop = false;
            this.groupBoxViewRadioButtons.Text = "View";
            // 
            // radioViewMinimalInfo
            // 
            this.radioViewMinimalInfo.AutoSize = true;
            this.radioViewMinimalInfo.Location = new System.Drawing.Point(7, 155);
            this.radioViewMinimalInfo.Name = "radioViewMinimalInfo";
            this.radioViewMinimalInfo.Size = new System.Drawing.Size(81, 17);
            this.radioViewMinimalInfo.TabIndex = 20;
            this.radioViewMinimalInfo.TabStop = true;
            this.radioViewMinimalInfo.Text = "&Minimal Info";
            this.radioViewMinimalInfo.UseVisualStyleBackColor = true;
            this.radioViewMinimalInfo.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewDislikedContents
            // 
            this.radioViewDislikedContents.AutoSize = true;
            this.radioViewDislikedContents.Location = new System.Drawing.Point(159, 138);
            this.radioViewDislikedContents.Name = "radioViewDislikedContents";
            this.radioViewDislikedContents.Size = new System.Drawing.Size(102, 17);
            this.radioViewDislikedContents.TabIndex = 19;
            this.radioViewDislikedContents.TabStop = true;
            this.radioViewDislikedContents.Text = "&Disliked Content";
            this.radioViewDislikedContents.UseVisualStyleBackColor = true;
            this.radioViewDislikedContents.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewMismatchedAddrAndComp
            // 
            this.radioViewMismatchedAddrAndComp.AutoSize = true;
            this.radioViewMismatchedAddrAndComp.Location = new System.Drawing.Point(159, 121);
            this.radioViewMismatchedAddrAndComp.Name = "radioViewMismatchedAddrAndComp";
            this.radioViewMismatchedAddrAndComp.Size = new System.Drawing.Size(137, 17);
            this.radioViewMismatchedAddrAndComp.TabIndex = 18;
            this.radioViewMismatchedAddrAndComp.TabStop = true;
            this.radioViewMismatchedAddrAndComp.Text = "&Address != Components";
            this.radioViewMismatchedAddrAndComp.UseVisualStyleBackColor = true;
            this.radioViewMismatchedAddrAndComp.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewEmptyMailingAddresses
            // 
            this.radioViewEmptyMailingAddresses.AutoSize = true;
            this.radioViewEmptyMailingAddresses.Location = new System.Drawing.Point(159, 104);
            this.radioViewEmptyMailingAddresses.Name = "radioViewEmptyMailingAddresses";
            this.radioViewEmptyMailingAddresses.Size = new System.Drawing.Size(138, 17);
            this.radioViewEmptyMailingAddresses.TabIndex = 17;
            this.radioViewEmptyMailingAddresses.TabStop = true;
            this.radioViewEmptyMailingAddresses.Text = "&Mailing Address Not Set";
            this.radioViewEmptyMailingAddresses.UseVisualStyleBackColor = true;
            this.radioViewEmptyMailingAddresses.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewLinkedMailingAddresses
            // 
            this.radioViewLinkedMailingAddresses.AutoSize = true;
            this.radioViewLinkedMailingAddresses.Location = new System.Drawing.Point(159, 87);
            this.radioViewLinkedMailingAddresses.Name = "radioViewLinkedMailingAddresses";
            this.radioViewLinkedMailingAddresses.Size = new System.Drawing.Size(118, 17);
            this.radioViewLinkedMailingAddresses.TabIndex = 16;
            this.radioViewLinkedMailingAddresses.TabStop = true;
            this.radioViewLinkedMailingAddresses.Text = "&Mailing Address Set";
            this.radioViewLinkedMailingAddresses.UseVisualStyleBackColor = true;
            this.radioViewLinkedMailingAddresses.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewUnlinkedMailingAddresses
            // 
            this.radioViewUnlinkedMailingAddresses.AutoSize = true;
            this.radioViewUnlinkedMailingAddresses.Location = new System.Drawing.Point(159, 70);
            this.radioViewUnlinkedMailingAddresses.Name = "radioViewUnlinkedMailingAddresses";
            this.radioViewUnlinkedMailingAddresses.Size = new System.Drawing.Size(161, 17);
            this.radioViewUnlinkedMailingAddresses.TabIndex = 15;
            this.radioViewUnlinkedMailingAddresses.TabStop = true;
            this.radioViewUnlinkedMailingAddresses.Text = "&Mailing Address != Any Other";
            this.radioViewUnlinkedMailingAddresses.UseVisualStyleBackColor = true;
            this.radioViewUnlinkedMailingAddresses.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewTruncatedNotes
            // 
            this.radioViewTruncatedNotes.AutoSize = true;
            this.radioViewTruncatedNotes.Location = new System.Drawing.Point(159, 53);
            this.radioViewTruncatedNotes.Name = "radioViewTruncatedNotes";
            this.radioViewTruncatedNotes.Size = new System.Drawing.Size(105, 17);
            this.radioViewTruncatedNotes.TabIndex = 14;
            this.radioViewTruncatedNotes.TabStop = true;
            this.radioViewTruncatedNotes.Text = "&Truncated Notes";
            this.radioViewTruncatedNotes.UseVisualStyleBackColor = true;
            this.radioViewTruncatedNotes.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewDuplicateEmails
            // 
            this.radioViewDuplicateEmails.AutoSize = true;
            this.radioViewDuplicateEmails.Location = new System.Drawing.Point(159, 36);
            this.radioViewDuplicateEmails.Name = "radioViewDuplicateEmails";
            this.radioViewDuplicateEmails.Size = new System.Drawing.Size(100, 17);
            this.radioViewDuplicateEmails.TabIndex = 13;
            this.radioViewDuplicateEmails.TabStop = true;
            this.radioViewDuplicateEmails.Text = "&DuplicateEmails";
            this.radioViewDuplicateEmails.UseVisualStyleBackColor = true;
            this.radioViewDuplicateEmails.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewDuplicatePhoneNumbers
            // 
            this.radioViewDuplicatePhoneNumbers.AutoSize = true;
            this.radioViewDuplicatePhoneNumbers.Location = new System.Drawing.Point(7, 138);
            this.radioViewDuplicatePhoneNumbers.Name = "radioViewDuplicatePhoneNumbers";
            this.radioViewDuplicatePhoneNumbers.Size = new System.Drawing.Size(106, 17);
            this.radioViewDuplicatePhoneNumbers.TabIndex = 12;
            this.radioViewDuplicatePhoneNumbers.TabStop = true;
            this.radioViewDuplicatePhoneNumbers.Text = "&DuplicatePhones";
            this.radioViewDuplicatePhoneNumbers.UseVisualStyleBackColor = true;
            this.radioViewDuplicatePhoneNumbers.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewDuplicateAddresses
            // 
            this.radioViewDuplicateAddresses.AutoSize = true;
            this.radioViewDuplicateAddresses.Location = new System.Drawing.Point(7, 121);
            this.radioViewDuplicateAddresses.Name = "radioViewDuplicateAddresses";
            this.radioViewDuplicateAddresses.Size = new System.Drawing.Size(119, 17);
            this.radioViewDuplicateAddresses.TabIndex = 11;
            this.radioViewDuplicateAddresses.TabStop = true;
            this.radioViewDuplicateAddresses.Text = "&DuplicateAddresses";
            this.radioViewDuplicateAddresses.UseVisualStyleBackColor = true;
            this.radioViewDuplicateAddresses.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewSameNameAndSuffix
            // 
            this.radioViewSameNameAndSuffix.AutoSize = true;
            this.radioViewSameNameAndSuffix.Location = new System.Drawing.Point(7, 104);
            this.radioViewSameNameAndSuffix.Name = "radioViewSameNameAndSuffix";
            this.radioViewSameNameAndSuffix.Size = new System.Drawing.Size(125, 17);
            this.radioViewSameNameAndSuffix.TabIndex = 10;
            this.radioViewSameNameAndSuffix.TabStop = true;
            this.radioViewSameNameAndSuffix.Text = "&SameNameAndSuffix";
            this.radioViewSameNameAndSuffix.UseVisualStyleBackColor = true;
            this.radioViewSameNameAndSuffix.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewSameName
            // 
            this.radioViewSameName.AutoSize = true;
            this.radioViewSameName.Location = new System.Drawing.Point(7, 87);
            this.radioViewSameName.Name = "radioViewSameName";
            this.radioViewSameName.Size = new System.Drawing.Size(80, 17);
            this.radioViewSameName.TabIndex = 9;
            this.radioViewSameName.TabStop = true;
            this.radioViewSameName.Text = "&SameName";
            this.radioViewSameName.UseVisualStyleBackColor = true;
            this.radioViewSameName.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewVaguelySimilar
            // 
            this.radioViewVaguelySimilar.AutoSize = true;
            this.radioViewVaguelySimilar.Location = new System.Drawing.Point(7, 70);
            this.radioViewVaguelySimilar.Name = "radioViewVaguelySimilar";
            this.radioViewVaguelySimilar.Size = new System.Drawing.Size(96, 17);
            this.radioViewVaguelySimilar.TabIndex = 8;
            this.radioViewVaguelySimilar.TabStop = true;
            this.radioViewVaguelySimilar.Text = "&Vaguely Similar";
            this.radioViewVaguelySimilar.UseVisualStyleBackColor = true;
            this.radioViewVaguelySimilar.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewSimilar
            // 
            this.radioViewSimilar.AutoSize = true;
            this.radioViewSimilar.Location = new System.Drawing.Point(7, 53);
            this.radioViewSimilar.Name = "radioViewSimilar";
            this.radioViewSimilar.Size = new System.Drawing.Size(55, 17);
            this.radioViewSimilar.TabIndex = 7;
            this.radioViewSimilar.TabStop = true;
            this.radioViewSimilar.Text = "&Similar";
            this.radioViewSimilar.UseVisualStyleBackColor = true;
            this.radioViewSimilar.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewDuplicates
            // 
            this.radioViewDuplicates.AutoSize = true;
            this.radioViewDuplicates.Location = new System.Drawing.Point(7, 36);
            this.radioViewDuplicates.Name = "radioViewDuplicates";
            this.radioViewDuplicates.Size = new System.Drawing.Size(75, 17);
            this.radioViewDuplicates.TabIndex = 6;
            this.radioViewDuplicates.TabStop = true;
            this.radioViewDuplicates.Text = "&Duplicates";
            this.radioViewDuplicates.UseVisualStyleBackColor = true;
            this.radioViewDuplicates.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewAllByName
            // 
            this.radioViewAllByName.AutoSize = true;
            this.radioViewAllByName.Location = new System.Drawing.Point(159, 19);
            this.radioViewAllByName.Name = "radioViewAllByName";
            this.radioViewAllByName.Size = new System.Drawing.Size(81, 17);
            this.radioViewAllByName.TabIndex = 5;
            this.radioViewAllByName.TabStop = true;
            this.radioViewAllByName.Text = "All by &Name";
            this.radioViewAllByName.UseVisualStyleBackColor = true;
            this.radioViewAllByName.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // radioViewAll
            // 
            this.radioViewAll.AutoSize = true;
            this.radioViewAll.Location = new System.Drawing.Point(7, 19);
            this.radioViewAll.Name = "radioViewAll";
            this.radioViewAll.Size = new System.Drawing.Size(36, 17);
            this.radioViewAll.TabIndex = 5;
            this.radioViewAll.TabStop = true;
            this.radioViewAll.Text = "&All";
            this.radioViewAll.UseVisualStyleBackColor = true;
            this.radioViewAll.CheckedChanged += new System.EventHandler(this.radioView_SelectView);
            // 
            // txtStatusLine
            // 
            this.txtStatusLine.BackColor = System.Drawing.SystemColors.ControlLight;
            this.txtStatusLine.Location = new System.Drawing.Point(3, 572);
            this.txtStatusLine.Name = "txtStatusLine";
            this.txtStatusLine.Size = new System.Drawing.Size(851, 20);
            this.txtStatusLine.TabIndex = 16;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.viewToolStripMenuItem,
            this.editToolStripMenuItem,
            this.toolsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(881, 24);
            this.menuStrip1.TabIndex = 1;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileNewToolStripMenuItem,
            this.fileOpenToolStripMenuItem,
            this.toolStripSeparator,
            this.fileSaveToolStripMenuItem,
            this.fileSaveAsToolStripMenuItem,
            this.toolStripSeparator4,
            this.filePrintToolStripMenuItem,
            this.filePrintPreviewToolStripMenuItem,
            this.toolStripSeparator5,
            this.fileExitToolStripMenuItem,
            this.fileRefreshFoldersToolStripMenuItem,
            this.fileLoadContactsToolStripMenuItem,
            this.fileSaveChangedContactsToolStripMenuItem,
            this.fileStopGetContactsThreadToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(35, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // fileNewToolStripMenuItem
            // 
            this.fileNewToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("fileNewToolStripMenuItem.Image")));
            this.fileNewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.fileNewToolStripMenuItem.Name = "fileNewToolStripMenuItem";
            this.fileNewToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.fileNewToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileNewToolStripMenuItem.Text = "&New";
            // 
            // fileOpenToolStripMenuItem
            // 
            this.fileOpenToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("fileOpenToolStripMenuItem.Image")));
            this.fileOpenToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.fileOpenToolStripMenuItem.Name = "fileOpenToolStripMenuItem";
            this.fileOpenToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.fileOpenToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileOpenToolStripMenuItem.Text = "&Open";
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(194, 6);
            // 
            // fileSaveToolStripMenuItem
            // 
            this.fileSaveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("fileSaveToolStripMenuItem.Image")));
            this.fileSaveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.fileSaveToolStripMenuItem.Name = "fileSaveToolStripMenuItem";
            this.fileSaveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.fileSaveToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileSaveToolStripMenuItem.Text = "&Save";
            // 
            // fileSaveAsToolStripMenuItem
            // 
            this.fileSaveAsToolStripMenuItem.Name = "fileSaveAsToolStripMenuItem";
            this.fileSaveAsToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileSaveAsToolStripMenuItem.Text = "Save &As";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(194, 6);
            // 
            // filePrintToolStripMenuItem
            // 
            this.filePrintToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("filePrintToolStripMenuItem.Image")));
            this.filePrintToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.filePrintToolStripMenuItem.Name = "filePrintToolStripMenuItem";
            this.filePrintToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.filePrintToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.filePrintToolStripMenuItem.Text = "&Print";
            // 
            // filePrintPreviewToolStripMenuItem
            // 
            this.filePrintPreviewToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("filePrintPreviewToolStripMenuItem.Image")));
            this.filePrintPreviewToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.filePrintPreviewToolStripMenuItem.Name = "filePrintPreviewToolStripMenuItem";
            this.filePrintPreviewToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.filePrintPreviewToolStripMenuItem.Text = "Print Pre&view";
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(194, 6);
            // 
            // fileExitToolStripMenuItem
            // 
            this.fileExitToolStripMenuItem.Name = "fileExitToolStripMenuItem";
            this.fileExitToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileExitToolStripMenuItem.Text = "E&xit";
            // 
            // fileRefreshFoldersToolStripMenuItem
            // 
            this.fileRefreshFoldersToolStripMenuItem.Name = "fileRefreshFoldersToolStripMenuItem";
            this.fileRefreshFoldersToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileRefreshFoldersToolStripMenuItem.Text = "&Refresh Folders";
            this.fileRefreshFoldersToolStripMenuItem.Click += new System.EventHandler(this.fileRefreshFoldersToolStripMenuItem_Click);
            // 
            // fileLoadContactsToolStripMenuItem
            // 
            this.fileLoadContactsToolStripMenuItem.Name = "fileLoadContactsToolStripMenuItem";
            this.fileLoadContactsToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileLoadContactsToolStripMenuItem.Text = "&Load Contacts";
            this.fileLoadContactsToolStripMenuItem.Click += new System.EventHandler(this.fileLoadContactsToolStripMenuItem_Click);
            // 
            // fileSaveChangedContactsToolStripMenuItem
            // 
            this.fileSaveChangedContactsToolStripMenuItem.Enabled = false;
            this.fileSaveChangedContactsToolStripMenuItem.Name = "fileSaveChangedContactsToolStripMenuItem";
            this.fileSaveChangedContactsToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileSaveChangedContactsToolStripMenuItem.Text = "&Save changed contacts";
            this.fileSaveChangedContactsToolStripMenuItem.Click += new System.EventHandler(this.fileSaveChangedContactsToolStripMenuItem_Click);
            // 
            // fileStopGetContactsThreadToolStripMenuItem
            // 
            this.fileStopGetContactsThreadToolStripMenuItem.Enabled = false;
            this.fileStopGetContactsThreadToolStripMenuItem.Name = "fileStopGetContactsThreadToolStripMenuItem";
            this.fileStopGetContactsThreadToolStripMenuItem.Size = new System.Drawing.Size(197, 22);
            this.fileStopGetContactsThreadToolStripMenuItem.Text = "&Stop Get Contacts thread";
            this.fileStopGetContactsThreadToolStripMenuItem.Click += new System.EventHandler(this.fileStopGetContactsThreadToolStripMenuItem_Click);
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.viewShowAllContactsToolStripMenuItem,
            this.viewShowDuplicateContactsToolStripMenuItem,
            this.viewShowSimilarContactsToolStripMenuItem,
            this.viewShowVaguelySimilarContactsToolStripMenuItem,
            this.viewShowWithTheSameNameToolStripMenuItem,
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem,
            this.viewShowWithTruncatedNotesToolStripMenuItem,
            this.viewShowWithDislikedContentsToolStripMenuItem,
            this.viewShowWithDuplicateAddressesToolStripMenuItem,
            this.showWithDuplicatePhoneNumbersToolStripMenuItem,
            this.showWithDuplicateEmailAddressesToolStripMenuItem,
            this.toolStripSeparator1,
            this.viewHideUnusedColumnsToolStripMenuItem,
            this.mailingAddressesToolStripMenuItem});
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(41, 20);
            this.viewToolStripMenuItem.Text = "&View";
            // 
            // viewShowAllContactsToolStripMenuItem
            // 
            this.viewShowAllContactsToolStripMenuItem.Enabled = false;
            this.viewShowAllContactsToolStripMenuItem.Name = "viewShowAllContactsToolStripMenuItem";
            this.viewShowAllContactsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowAllContactsToolStripMenuItem.Text = "Show &All Contacts ";
            this.viewShowAllContactsToolStripMenuItem.Click += new System.EventHandler(this.viewShowAllContactsToolStripMenuItem_Click);
            // 
            // viewShowDuplicateContactsToolStripMenuItem
            // 
            this.viewShowDuplicateContactsToolStripMenuItem.Enabled = false;
            this.viewShowDuplicateContactsToolStripMenuItem.Name = "viewShowDuplicateContactsToolStripMenuItem";
            this.viewShowDuplicateContactsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowDuplicateContactsToolStripMenuItem.Text = "Show &Duplicate Contacts";
            this.viewShowDuplicateContactsToolStripMenuItem.Click += new System.EventHandler(this.viewShowDuplicateContactsToolStripMenuItem_Click);
            // 
            // viewShowSimilarContactsToolStripMenuItem
            // 
            this.viewShowSimilarContactsToolStripMenuItem.Enabled = false;
            this.viewShowSimilarContactsToolStripMenuItem.Name = "viewShowSimilarContactsToolStripMenuItem";
            this.viewShowSimilarContactsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowSimilarContactsToolStripMenuItem.Text = "Show &Similar Contacts";
            this.viewShowSimilarContactsToolStripMenuItem.Click += new System.EventHandler(this.viewShowSimilarContactsToolStripMenuItem_Click);
            // 
            // viewShowVaguelySimilarContactsToolStripMenuItem
            // 
            this.viewShowVaguelySimilarContactsToolStripMenuItem.Enabled = false;
            this.viewShowVaguelySimilarContactsToolStripMenuItem.Name = "viewShowVaguelySimilarContactsToolStripMenuItem";
            this.viewShowVaguelySimilarContactsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowVaguelySimilarContactsToolStripMenuItem.Text = "Show &Vaguely Similar Contacts";
            this.viewShowVaguelySimilarContactsToolStripMenuItem.Click += new System.EventHandler(this.viewShowVaguelySimilarContactsToolStripMenuItem_Click);
            // 
            // viewShowWithTheSameNameToolStripMenuItem
            // 
            this.viewShowWithTheSameNameToolStripMenuItem.Enabled = false;
            this.viewShowWithTheSameNameToolStripMenuItem.Name = "viewShowWithTheSameNameToolStripMenuItem";
            this.viewShowWithTheSameNameToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowWithTheSameNameToolStripMenuItem.Text = "Show with the same &Name";
            this.viewShowWithTheSameNameToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithTheSameNameToolStripMenuItem_Click);
            // 
            // viewShowWithTheSameNameAndSuffixToolStripMenuItem
            // 
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem.Enabled = false;
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem.Name = "viewShowWithTheSameNameAndSuffixToolStripMenuItem";
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem.Text = "Show with the same name and Suffix";
            this.viewShowWithTheSameNameAndSuffixToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithTheSameNameAndSuffixToolStripMenuItem_Click);
            // 
            // viewShowWithTruncatedNotesToolStripMenuItem
            // 
            this.viewShowWithTruncatedNotesToolStripMenuItem.Enabled = false;
            this.viewShowWithTruncatedNotesToolStripMenuItem.Name = "viewShowWithTruncatedNotesToolStripMenuItem";
            this.viewShowWithTruncatedNotesToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowWithTruncatedNotesToolStripMenuItem.Text = "Show with &Truncated Notes";
            this.viewShowWithTruncatedNotesToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithTruncatedNotesToolStripMenuItem_Click);
            // 
            // viewShowWithDislikedContentsToolStripMenuItem
            // 
            this.viewShowWithDislikedContentsToolStripMenuItem.Enabled = false;
            this.viewShowWithDislikedContentsToolStripMenuItem.Name = "viewShowWithDislikedContentsToolStripMenuItem";
            this.viewShowWithDislikedContentsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowWithDislikedContentsToolStripMenuItem.Text = "Show with &Disliked contents";
            this.viewShowWithDislikedContentsToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithDislikedContentsToolStripMenuItem_Click);
            // 
            // viewShowWithDuplicateAddressesToolStripMenuItem
            // 
            this.viewShowWithDuplicateAddressesToolStripMenuItem.Enabled = false;
            this.viewShowWithDuplicateAddressesToolStripMenuItem.Name = "viewShowWithDuplicateAddressesToolStripMenuItem";
            this.viewShowWithDuplicateAddressesToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewShowWithDuplicateAddressesToolStripMenuItem.Text = "Show with Duplicate Addresses";
            this.viewShowWithDuplicateAddressesToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithDuplicateAddressesToolStripMenuItem_Click);
            // 
            // showWithDuplicatePhoneNumbersToolStripMenuItem
            // 
            this.showWithDuplicatePhoneNumbersToolStripMenuItem.Enabled = false;
            this.showWithDuplicatePhoneNumbersToolStripMenuItem.Name = "showWithDuplicatePhoneNumbersToolStripMenuItem";
            this.showWithDuplicatePhoneNumbersToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.showWithDuplicatePhoneNumbersToolStripMenuItem.Text = "Show with Duplicate Phone Numbers";
            this.showWithDuplicatePhoneNumbersToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithDuplicatePhoneNumbersToolStripMenuItem_Click);
            // 
            // showWithDuplicateEmailAddressesToolStripMenuItem
            // 
            this.showWithDuplicateEmailAddressesToolStripMenuItem.Enabled = false;
            this.showWithDuplicateEmailAddressesToolStripMenuItem.Name = "showWithDuplicateEmailAddressesToolStripMenuItem";
            this.showWithDuplicateEmailAddressesToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.showWithDuplicateEmailAddressesToolStripMenuItem.Text = "Show with Duplicate Email Addresses";
            this.showWithDuplicateEmailAddressesToolStripMenuItem.Click += new System.EventHandler(this.viewShowWithDuplicateEmailAddressToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(248, 6);
            // 
            // viewHideUnusedColumnsToolStripMenuItem
            // 
            this.viewHideUnusedColumnsToolStripMenuItem.Enabled = false;
            this.viewHideUnusedColumnsToolStripMenuItem.Name = "viewHideUnusedColumnsToolStripMenuItem";
            this.viewHideUnusedColumnsToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.viewHideUnusedColumnsToolStripMenuItem.Text = "&Hide unused columns";
            this.viewHideUnusedColumnsToolStripMenuItem.Click += new System.EventHandler(this.viewHideUnusedColumnsToolStripMenuItem_Click);
            // 
            // mailingAddressesToolStripMenuItem
            // 
            this.mailingAddressesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem,
            this.viewShowLinkedMailingAddressesToolStripMenuItem,
            this.viewShowEmptyMailingAddressesToolStripMenuItem,
            this.viewShowMismatchedAddressComponentsToolStripMenuItem});
            this.mailingAddressesToolStripMenuItem.Name = "mailingAddressesToolStripMenuItem";
            this.mailingAddressesToolStripMenuItem.Size = new System.Drawing.Size(251, 22);
            this.mailingAddressesToolStripMenuItem.Text = "Addresses";
            // 
            // viewShowUnlinkedMailingAddressesToolStripMenuItem
            // 
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem.Name = "viewShowUnlinkedMailingAddressesToolStripMenuItem";
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem.Size = new System.Drawing.Size(235, 22);
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem.Text = "Unlinked Mailing Addresses";
            this.viewShowUnlinkedMailingAddressesToolStripMenuItem.Click += new System.EventHandler(this.viewShowUnlinkedMailingAddressesToolStripMenuItem_Click);
            // 
            // viewShowLinkedMailingAddressesToolStripMenuItem
            // 
            this.viewShowLinkedMailingAddressesToolStripMenuItem.Name = "viewShowLinkedMailingAddressesToolStripMenuItem";
            this.viewShowLinkedMailingAddressesToolStripMenuItem.Size = new System.Drawing.Size(235, 22);
            this.viewShowLinkedMailingAddressesToolStripMenuItem.Text = "Linked Mailing Addresses";
            this.viewShowLinkedMailingAddressesToolStripMenuItem.Click += new System.EventHandler(this.viewShowLinkedMailingAddressesToolStripMenuItem_Click);
            // 
            // viewShowEmptyMailingAddressesToolStripMenuItem
            // 
            this.viewShowEmptyMailingAddressesToolStripMenuItem.Name = "viewShowEmptyMailingAddressesToolStripMenuItem";
            this.viewShowEmptyMailingAddressesToolStripMenuItem.Size = new System.Drawing.Size(235, 22);
            this.viewShowEmptyMailingAddressesToolStripMenuItem.Text = "Empty Mailing Addresses";
            this.viewShowEmptyMailingAddressesToolStripMenuItem.Click += new System.EventHandler(this.viewShowEmptyMailingAddressesToolStripMenuItem_Click);
            // 
            // viewShowMismatchedAddressComponentsToolStripMenuItem
            // 
            this.viewShowMismatchedAddressComponentsToolStripMenuItem.Name = "viewShowMismatchedAddressComponentsToolStripMenuItem";
            this.viewShowMismatchedAddressComponentsToolStripMenuItem.Size = new System.Drawing.Size(235, 22);
            this.viewShowMismatchedAddressComponentsToolStripMenuItem.Text = "Mismatched Address Components";
            this.viewShowMismatchedAddressComponentsToolStripMenuItem.Click += new System.EventHandler(this.viewShowMismatchedAddressComponentsToolStripMenuItem_Click);
            // 
            // editToolStripMenuItem
            // 
            this.editToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.editUndoToolStripMenuItem,
            this.editRedoToolStripMenuItem,
            this.toolStripSeparator6,
            this.editCutToolStripMenuItem,
            this.editCopyToolStripMenuItem,
            this.editPasteToolStripMenuItem,
            this.toolStripSeparator7,
            this.editSelectAllToolStripMenuItem});
            this.editToolStripMenuItem.Name = "editToolStripMenuItem";
            this.editToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.editToolStripMenuItem.Text = "&Edit";
            // 
            // editUndoToolStripMenuItem
            // 
            this.editUndoToolStripMenuItem.Name = "editUndoToolStripMenuItem";
            this.editUndoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Z)));
            this.editUndoToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editUndoToolStripMenuItem.Text = "&Undo";
            // 
            // editRedoToolStripMenuItem
            // 
            this.editRedoToolStripMenuItem.Name = "editRedoToolStripMenuItem";
            this.editRedoToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Y)));
            this.editRedoToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editRedoToolStripMenuItem.Text = "&Redo";
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(136, 6);
            // 
            // editCutToolStripMenuItem
            // 
            this.editCutToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("editCutToolStripMenuItem.Image")));
            this.editCutToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.editCutToolStripMenuItem.Name = "editCutToolStripMenuItem";
            this.editCutToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.X)));
            this.editCutToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editCutToolStripMenuItem.Text = "Cu&t";
            // 
            // editCopyToolStripMenuItem
            // 
            this.editCopyToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("editCopyToolStripMenuItem.Image")));
            this.editCopyToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.editCopyToolStripMenuItem.Name = "editCopyToolStripMenuItem";
            this.editCopyToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.C)));
            this.editCopyToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editCopyToolStripMenuItem.Text = "&Copy";
            // 
            // editPasteToolStripMenuItem
            // 
            this.editPasteToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("editPasteToolStripMenuItem.Image")));
            this.editPasteToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.editPasteToolStripMenuItem.Name = "editPasteToolStripMenuItem";
            this.editPasteToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.V)));
            this.editPasteToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editPasteToolStripMenuItem.Text = "&Paste";
            // 
            // toolStripSeparator7
            // 
            this.toolStripSeparator7.Name = "toolStripSeparator7";
            this.toolStripSeparator7.Size = new System.Drawing.Size(136, 6);
            // 
            // editSelectAllToolStripMenuItem
            // 
            this.editSelectAllToolStripMenuItem.Name = "editSelectAllToolStripMenuItem";
            this.editSelectAllToolStripMenuItem.Size = new System.Drawing.Size(139, 22);
            this.editSelectAllToolStripMenuItem.Text = "Select &All";
            // 
            // toolsToolStripMenuItem
            // 
            this.toolsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolsCustomizeToolStripMenuItem,
            this.toolsOptionsToolStripMenuItem,
            this.toolsAnalyseContactsToolStripMenuItem,
            this.toolsEditContactDetailsToolStripMenuItem,
            this.toolsFieldsToolStripMenuItem,
            this.toolsNewContactToolStripMenuItem,
            this.cleanNotesFieldsToolStripMenuItem,
            this.markDuplicatesForDeletionToolStripMenuItem,
            this.fixPhoneNumbersToolStripMenuItem,
            this.mergeDuplicateContactsToolStripMenuItem});
            this.toolsToolStripMenuItem.Name = "toolsToolStripMenuItem";
            this.toolsToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.toolsToolStripMenuItem.Text = "&Tools";
            // 
            // toolsCustomizeToolStripMenuItem
            // 
            this.toolsCustomizeToolStripMenuItem.Name = "toolsCustomizeToolStripMenuItem";
            this.toolsCustomizeToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsCustomizeToolStripMenuItem.Text = "&Customize";
            // 
            // toolsOptionsToolStripMenuItem
            // 
            this.toolsOptionsToolStripMenuItem.Name = "toolsOptionsToolStripMenuItem";
            this.toolsOptionsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsOptionsToolStripMenuItem.Text = "&Options";
            // 
            // toolsAnalyseContactsToolStripMenuItem
            // 
            this.toolsAnalyseContactsToolStripMenuItem.Enabled = false;
            this.toolsAnalyseContactsToolStripMenuItem.Name = "toolsAnalyseContactsToolStripMenuItem";
            this.toolsAnalyseContactsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsAnalyseContactsToolStripMenuItem.Text = "&Analyse contacts";
            this.toolsAnalyseContactsToolStripMenuItem.Click += new System.EventHandler(this.toolsAnalyseContactsToolStripMenuItem_Click);
            // 
            // toolsEditContactDetailsToolStripMenuItem
            // 
            this.toolsEditContactDetailsToolStripMenuItem.Enabled = false;
            this.toolsEditContactDetailsToolStripMenuItem.Name = "toolsEditContactDetailsToolStripMenuItem";
            this.toolsEditContactDetailsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsEditContactDetailsToolStripMenuItem.Text = "Edit &Contact Details";
            this.toolsEditContactDetailsToolStripMenuItem.Click += new System.EventHandler(this.toolsEditContactDetailsToolStripMenuItem_Click);
            // 
            // toolsFieldsToolStripMenuItem
            // 
            this.toolsFieldsToolStripMenuItem.Enabled = false;
            this.toolsFieldsToolStripMenuItem.Name = "toolsFieldsToolStripMenuItem";
            this.toolsFieldsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsFieldsToolStripMenuItem.Text = "Fields";
            this.toolsFieldsToolStripMenuItem.Click += new System.EventHandler(this.toolsFieldsToolStripMenuItem_Click);
            // 
            // toolsNewContactToolStripMenuItem
            // 
            this.toolsNewContactToolStripMenuItem.Enabled = false;
            this.toolsNewContactToolStripMenuItem.Name = "toolsNewContactToolStripMenuItem";
            this.toolsNewContactToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.toolsNewContactToolStripMenuItem.Text = "&New Contact";
            this.toolsNewContactToolStripMenuItem.Click += new System.EventHandler(this.toolsNewContactToolStripMenuItem_Click);
            // 
            // cleanNotesFieldsToolStripMenuItem
            // 
            this.cleanNotesFieldsToolStripMenuItem.Name = "cleanNotesFieldsToolStripMenuItem";
            this.cleanNotesFieldsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.cleanNotesFieldsToolStripMenuItem.Text = "Clean Notes fields";
            this.cleanNotesFieldsToolStripMenuItem.Click += new System.EventHandler(this.cleanNotesFieldsToolStripMenuItem_Click);
            // 
            // markDuplicatesForDeletionToolStripMenuItem
            // 
            this.markDuplicatesForDeletionToolStripMenuItem.Name = "markDuplicatesForDeletionToolStripMenuItem";
            this.markDuplicatesForDeletionToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.markDuplicatesForDeletionToolStripMenuItem.Text = "Mark Duplicates for deletion";
            this.markDuplicatesForDeletionToolStripMenuItem.Click += new System.EventHandler(this.markDuplicatesForDeletionToolStripMenuItem_Click);
            // 
            // fixPhoneNumbersToolStripMenuItem
            // 
            this.fixPhoneNumbersToolStripMenuItem.Name = "fixPhoneNumbersToolStripMenuItem";
            this.fixPhoneNumbersToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.fixPhoneNumbersToolStripMenuItem.Text = "Fix Phone Numbers";
            this.fixPhoneNumbersToolStripMenuItem.Click += new System.EventHandler(this.fixPhoneNumbersToolStripMenuItem_Click);
            // 
            // mergeDuplicateContactsToolStripMenuItem
            // 
            this.mergeDuplicateContactsToolStripMenuItem.Name = "mergeDuplicateContactsToolStripMenuItem";
            this.mergeDuplicateContactsToolStripMenuItem.Size = new System.Drawing.Size(253, 22);
            this.mergeDuplicateContactsToolStripMenuItem.Text = "Merge SameNameAndSuffix Contacts";
            this.mergeDuplicateContactsToolStripMenuItem.Click += new System.EventHandler(this.mergeDuplicateContactsToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpContentsToolStripMenuItem,
            this.helpIndexToolStripMenuItem,
            this.helpSearchToolStripMenuItem,
            this.toolStripSeparator8,
            this.helpAboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(40, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // helpContentsToolStripMenuItem
            // 
            this.helpContentsToolStripMenuItem.Name = "helpContentsToolStripMenuItem";
            this.helpContentsToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.helpContentsToolStripMenuItem.Text = "&Contents";
            // 
            // helpIndexToolStripMenuItem
            // 
            this.helpIndexToolStripMenuItem.Name = "helpIndexToolStripMenuItem";
            this.helpIndexToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.helpIndexToolStripMenuItem.Text = "&Index";
            // 
            // helpSearchToolStripMenuItem
            // 
            this.helpSearchToolStripMenuItem.Name = "helpSearchToolStripMenuItem";
            this.helpSearchToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.helpSearchToolStripMenuItem.Text = "&Search";
            // 
            // toolStripSeparator8
            // 
            this.toolStripSeparator8.Name = "toolStripSeparator8";
            this.toolStripSeparator8.Size = new System.Drawing.Size(115, 6);
            // 
            // helpAboutToolStripMenuItem
            // 
            this.helpAboutToolStripMenuItem.Name = "helpAboutToolStripMenuItem";
            this.helpAboutToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.helpAboutToolStripMenuItem.Text = "&About...";
            // 
            // ndxDataGridViewTextBoxColumn
            // 
            this.ndxDataGridViewTextBoxColumn.DataPropertyName = "Ndx";
            this.ndxDataGridViewTextBoxColumn.Frozen = true;
            this.ndxDataGridViewTextBoxColumn.HeaderText = "Ndx";
            this.ndxDataGridViewTextBoxColumn.Name = "ndxDataGridViewTextBoxColumn";
            this.ndxDataGridViewTextBoxColumn.Width = 54;
            // 
            // fullNameAndCompanyDataGridViewTextBoxColumn
            // 
            this.fullNameAndCompanyDataGridViewTextBoxColumn.DataPropertyName = "FullNameAndCompany";
            this.fullNameAndCompanyDataGridViewTextBoxColumn.Frozen = true;
            this.fullNameAndCompanyDataGridViewTextBoxColumn.HeaderText = "FullNameAndCompany";
            this.fullNameAndCompanyDataGridViewTextBoxColumn.Name = "fullNameAndCompanyDataGridViewTextBoxColumn";
            this.fullNameAndCompanyDataGridViewTextBoxColumn.Width = 157;
            // 
            // titleDataGridViewTextBoxColumn
            // 
            this.titleDataGridViewTextBoxColumn.DataPropertyName = "Title";
            this.titleDataGridViewTextBoxColumn.HeaderText = "Title";
            this.titleDataGridViewTextBoxColumn.Name = "titleDataGridViewTextBoxColumn";
            this.titleDataGridViewTextBoxColumn.Width = 57;
            // 
            // lastNameDataGridViewTextBoxColumn
            // 
            this.lastNameDataGridViewTextBoxColumn.DataPropertyName = "LastName";
            this.lastNameDataGridViewTextBoxColumn.HeaderText = "LastName";
            this.lastNameDataGridViewTextBoxColumn.Name = "lastNameDataGridViewTextBoxColumn";
            this.lastNameDataGridViewTextBoxColumn.Width = 88;
            // 
            // middleNameDataGridViewTextBoxColumn
            // 
            this.middleNameDataGridViewTextBoxColumn.DataPropertyName = "MiddleName";
            this.middleNameDataGridViewTextBoxColumn.HeaderText = "MiddleName";
            this.middleNameDataGridViewTextBoxColumn.Name = "middleNameDataGridViewTextBoxColumn";
            this.middleNameDataGridViewTextBoxColumn.Width = 101;
            // 
            // firstNameDataGridViewTextBoxColumn
            // 
            this.firstNameDataGridViewTextBoxColumn.DataPropertyName = "FirstName";
            this.firstNameDataGridViewTextBoxColumn.HeaderText = "FirstName";
            this.firstNameDataGridViewTextBoxColumn.Name = "firstNameDataGridViewTextBoxColumn";
            this.firstNameDataGridViewTextBoxColumn.Width = 88;
            // 
            // suffixDataGridViewTextBoxColumn
            // 
            this.suffixDataGridViewTextBoxColumn.DataPropertyName = "Suffix";
            this.suffixDataGridViewTextBoxColumn.HeaderText = "Suffix";
            this.suffixDataGridViewTextBoxColumn.Name = "suffixDataGridViewTextBoxColumn";
            this.suffixDataGridViewTextBoxColumn.Width = 64;
            // 
            // nickNameDataGridViewTextBoxColumn
            // 
            this.nickNameDataGridViewTextBoxColumn.DataPropertyName = "NickName";
            this.nickNameDataGridViewTextBoxColumn.HeaderText = "NickName";
            this.nickNameDataGridViewTextBoxColumn.Name = "nickNameDataGridViewTextBoxColumn";
            this.nickNameDataGridViewTextBoxColumn.Width = 90;
            // 
            // companyNameDataGridViewTextBoxColumn
            // 
            this.companyNameDataGridViewTextBoxColumn.DataPropertyName = "CompanyName";
            this.companyNameDataGridViewTextBoxColumn.HeaderText = "CompanyName";
            this.companyNameDataGridViewTextBoxColumn.Name = "companyNameDataGridViewTextBoxColumn";
            this.companyNameDataGridViewTextBoxColumn.Width = 115;
            // 
            // fullNameDataGridViewTextBoxColumn
            // 
            this.fullNameDataGridViewTextBoxColumn.DataPropertyName = "FullName";
            this.fullNameDataGridViewTextBoxColumn.HeaderText = "FullName";
            this.fullNameDataGridViewTextBoxColumn.Name = "fullNameDataGridViewTextBoxColumn";
            this.fullNameDataGridViewTextBoxColumn.Width = 84;
            // 
            // lastNameAndFirstNameDataGridViewTextBoxColumn
            // 
            this.lastNameAndFirstNameDataGridViewTextBoxColumn.DataPropertyName = "LastNameAndFirstName";
            this.lastNameAndFirstNameDataGridViewTextBoxColumn.HeaderText = "LastNameAndFirstName";
            this.lastNameAndFirstNameDataGridViewTextBoxColumn.Name = "lastNameAndFirstNameDataGridViewTextBoxColumn";
            this.lastNameAndFirstNameDataGridViewTextBoxColumn.Width = 166;
            // 
            // companyAndFullNameDataGridViewTextBoxColumn
            // 
            this.companyAndFullNameDataGridViewTextBoxColumn.DataPropertyName = "CompanyAndFullName";
            this.companyAndFullNameDataGridViewTextBoxColumn.HeaderText = "CompanyAndFullName";
            this.companyAndFullNameDataGridViewTextBoxColumn.Name = "companyAndFullNameDataGridViewTextBoxColumn";
            this.companyAndFullNameDataGridViewTextBoxColumn.Width = 157;
            // 
            // primaryTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.primaryTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "PrimaryTelephoneNumber";
            this.primaryTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "PrimaryTelephoneNumber";
            this.primaryTelephoneNumberDataGridViewTextBoxColumn.Name = "primaryTelephoneNumberDataGridViewTextBoxColumn";
            this.primaryTelephoneNumberDataGridViewTextBoxColumn.Width = 176;
            // 
            // mobileTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.mobileTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "MobileTelephoneNumber";
            this.mobileTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "MobileTelephoneNumber";
            this.mobileTelephoneNumberDataGridViewTextBoxColumn.Name = "mobileTelephoneNumberDataGridViewTextBoxColumn";
            this.mobileTelephoneNumberDataGridViewTextBoxColumn.Width = 172;
            // 
            // homeTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.homeTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "HomeTelephoneNumber";
            this.homeTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "HomeTelephoneNumber";
            this.homeTelephoneNumberDataGridViewTextBoxColumn.Name = "homeTelephoneNumberDataGridViewTextBoxColumn";
            this.homeTelephoneNumberDataGridViewTextBoxColumn.Width = 167;
            // 
            // home2TelephoneNumberDataGridViewTextBoxColumn
            // 
            this.home2TelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "Home2TelephoneNumber";
            this.home2TelephoneNumberDataGridViewTextBoxColumn.HeaderText = "Home2TelephoneNumber";
            this.home2TelephoneNumberDataGridViewTextBoxColumn.Name = "home2TelephoneNumberDataGridViewTextBoxColumn";
            this.home2TelephoneNumberDataGridViewTextBoxColumn.Width = 174;
            // 
            // companyMainTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "CompanyMainTelephoneNumber";
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "CompanyMainTelephoneNumber";
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn.Name = "companyMainTelephoneNumberDataGridViewTextBoxColumn";
            this.companyMainTelephoneNumberDataGridViewTextBoxColumn.Width = 213;
            // 
            // businessTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.businessTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "BusinessTelephoneNumber";
            this.businessTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "BusinessTelephoneNumber";
            this.businessTelephoneNumberDataGridViewTextBoxColumn.Name = "businessTelephoneNumberDataGridViewTextBoxColumn";
            this.businessTelephoneNumberDataGridViewTextBoxColumn.Width = 185;
            // 
            // business2TelephoneNumberDataGridViewTextBoxColumn
            // 
            this.business2TelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "Business2TelephoneNumber";
            this.business2TelephoneNumberDataGridViewTextBoxColumn.HeaderText = "Business2TelephoneNumber";
            this.business2TelephoneNumberDataGridViewTextBoxColumn.Name = "business2TelephoneNumberDataGridViewTextBoxColumn";
            this.business2TelephoneNumberDataGridViewTextBoxColumn.Width = 192;
            // 
            // otherTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.otherTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "OtherTelephoneNumber";
            this.otherTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "OtherTelephoneNumber";
            this.otherTelephoneNumberDataGridViewTextBoxColumn.Name = "otherTelephoneNumberDataGridViewTextBoxColumn";
            this.otherTelephoneNumberDataGridViewTextBoxColumn.Width = 166;
            // 
            // homeFaxNumberDataGridViewTextBoxColumn
            // 
            this.homeFaxNumberDataGridViewTextBoxColumn.DataPropertyName = "HomeFaxNumber";
            this.homeFaxNumberDataGridViewTextBoxColumn.HeaderText = "HomeFaxNumber";
            this.homeFaxNumberDataGridViewTextBoxColumn.Name = "homeFaxNumberDataGridViewTextBoxColumn";
            this.homeFaxNumberDataGridViewTextBoxColumn.Width = 127;
            // 
            // businessFaxNumberDataGridViewTextBoxColumn
            // 
            this.businessFaxNumberDataGridViewTextBoxColumn.DataPropertyName = "BusinessFaxNumber";
            this.businessFaxNumberDataGridViewTextBoxColumn.HeaderText = "BusinessFaxNumber";
            this.businessFaxNumberDataGridViewTextBoxColumn.Name = "businessFaxNumberDataGridViewTextBoxColumn";
            this.businessFaxNumberDataGridViewTextBoxColumn.Width = 145;
            // 
            // otherFaxNumberDataGridViewTextBoxColumn
            // 
            this.otherFaxNumberDataGridViewTextBoxColumn.DataPropertyName = "OtherFaxNumber";
            this.otherFaxNumberDataGridViewTextBoxColumn.HeaderText = "OtherFaxNumber";
            this.otherFaxNumberDataGridViewTextBoxColumn.Name = "otherFaxNumberDataGridViewTextBoxColumn";
            this.otherFaxNumberDataGridViewTextBoxColumn.Width = 126;
            // 
            // callbackTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.callbackTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "CallbackTelephoneNumber";
            this.callbackTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "CallbackTelephoneNumber";
            this.callbackTelephoneNumberDataGridViewTextBoxColumn.Name = "callbackTelephoneNumberDataGridViewTextBoxColumn";
            this.callbackTelephoneNumberDataGridViewTextBoxColumn.Width = 184;
            // 
            // carTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.carTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "CarTelephoneNumber";
            this.carTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "CarTelephoneNumber";
            this.carTelephoneNumberDataGridViewTextBoxColumn.Name = "carTelephoneNumberDataGridViewTextBoxColumn";
            this.carTelephoneNumberDataGridViewTextBoxColumn.Width = 154;
            // 
            // pagerNumberDataGridViewTextBoxColumn
            // 
            this.pagerNumberDataGridViewTextBoxColumn.DataPropertyName = "PagerNumber";
            this.pagerNumberDataGridViewTextBoxColumn.HeaderText = "PagerNumber";
            this.pagerNumberDataGridViewTextBoxColumn.Name = "pagerNumberDataGridViewTextBoxColumn";
            this.pagerNumberDataGridViewTextBoxColumn.Width = 108;
            // 
            // iSDNNumberDataGridViewTextBoxColumn
            // 
            this.iSDNNumberDataGridViewTextBoxColumn.DataPropertyName = "ISDNNumber";
            this.iSDNNumberDataGridViewTextBoxColumn.HeaderText = "ISDNNumber";
            this.iSDNNumberDataGridViewTextBoxColumn.Name = "iSDNNumberDataGridViewTextBoxColumn";
            this.iSDNNumberDataGridViewTextBoxColumn.Width = 105;
            // 
            // telexNumberDataGridViewTextBoxColumn
            // 
            this.telexNumberDataGridViewTextBoxColumn.DataPropertyName = "TelexNumber";
            this.telexNumberDataGridViewTextBoxColumn.HeaderText = "TelexNumber";
            this.telexNumberDataGridViewTextBoxColumn.Name = "telexNumberDataGridViewTextBoxColumn";
            this.telexNumberDataGridViewTextBoxColumn.Width = 106;
            // 
            // tTYTDDTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "TTYTDDTelephoneNumber";
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "TTYTDDTelephoneNumber";
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn.Name = "tTYTDDTelephoneNumberDataGridViewTextBoxColumn";
            this.tTYTDDTelephoneNumberDataGridViewTextBoxColumn.Width = 185;
            // 
            // accountDataGridViewTextBoxColumn
            // 
            this.accountDataGridViewTextBoxColumn.DataPropertyName = "Account";
            this.accountDataGridViewTextBoxColumn.HeaderText = "Account";
            this.accountDataGridViewTextBoxColumn.Name = "accountDataGridViewTextBoxColumn";
            this.accountDataGridViewTextBoxColumn.Width = 79;
            // 
            // anniversaryDataGridViewTextBoxColumn
            // 
            this.anniversaryDataGridViewTextBoxColumn.DataPropertyName = "Anniversary";
            this.anniversaryDataGridViewTextBoxColumn.HeaderText = "Anniversary";
            this.anniversaryDataGridViewTextBoxColumn.Name = "anniversaryDataGridViewTextBoxColumn";
            this.anniversaryDataGridViewTextBoxColumn.Width = 98;
            // 
            // assistantNameDataGridViewTextBoxColumn
            // 
            this.assistantNameDataGridViewTextBoxColumn.DataPropertyName = "AssistantName";
            this.assistantNameDataGridViewTextBoxColumn.HeaderText = "AssistantName";
            this.assistantNameDataGridViewTextBoxColumn.Name = "assistantNameDataGridViewTextBoxColumn";
            this.assistantNameDataGridViewTextBoxColumn.Width = 115;
            // 
            // assistantTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.assistantTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "AssistantTelephoneNumber";
            this.assistantTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "AssistantTelephoneNumber";
            this.assistantTelephoneNumberDataGridViewTextBoxColumn.Name = "assistantTelephoneNumberDataGridViewTextBoxColumn";
            this.assistantTelephoneNumberDataGridViewTextBoxColumn.Width = 186;
            // 
            // billingInformationDataGridViewTextBoxColumn
            // 
            this.billingInformationDataGridViewTextBoxColumn.DataPropertyName = "BillingInformation";
            this.billingInformationDataGridViewTextBoxColumn.HeaderText = "BillingInformation";
            this.billingInformationDataGridViewTextBoxColumn.Name = "billingInformationDataGridViewTextBoxColumn";
            this.billingInformationDataGridViewTextBoxColumn.Width = 129;
            // 
            // birthdayDataGridViewTextBoxColumn
            // 
            this.birthdayDataGridViewTextBoxColumn.DataPropertyName = "Birthday";
            this.birthdayDataGridViewTextBoxColumn.HeaderText = "Birthday";
            this.birthdayDataGridViewTextBoxColumn.Name = "birthdayDataGridViewTextBoxColumn";
            this.birthdayDataGridViewTextBoxColumn.Width = 78;
            // 
            // businessAddressDataGridViewTextBoxColumn
            // 
            this.businessAddressDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddress";
            this.businessAddressDataGridViewTextBoxColumn.HeaderText = "BusinessAddress";
            this.businessAddressDataGridViewTextBoxColumn.Name = "businessAddressDataGridViewTextBoxColumn";
            this.businessAddressDataGridViewTextBoxColumn.Width = 127;
            // 
            // businessAddressCityDataGridViewTextBoxColumn
            // 
            this.businessAddressCityDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressCity";
            this.businessAddressCityDataGridViewTextBoxColumn.HeaderText = "BusinessAddressCity";
            this.businessAddressCityDataGridViewTextBoxColumn.Name = "businessAddressCityDataGridViewTextBoxColumn";
            this.businessAddressCityDataGridViewTextBoxColumn.Width = 148;
            // 
            // businessAddressCountryDataGridViewTextBoxColumn
            // 
            this.businessAddressCountryDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressCountry";
            this.businessAddressCountryDataGridViewTextBoxColumn.HeaderText = "BusinessAddressCountry";
            this.businessAddressCountryDataGridViewTextBoxColumn.Name = "businessAddressCountryDataGridViewTextBoxColumn";
            this.businessAddressCountryDataGridViewTextBoxColumn.Width = 170;
            // 
            // businessAddressPostalCodeDataGridViewTextBoxColumn
            // 
            this.businessAddressPostalCodeDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressPostalCode";
            this.businessAddressPostalCodeDataGridViewTextBoxColumn.HeaderText = "BusinessAddressPostalCode";
            this.businessAddressPostalCodeDataGridViewTextBoxColumn.Name = "businessAddressPostalCodeDataGridViewTextBoxColumn";
            this.businessAddressPostalCodeDataGridViewTextBoxColumn.Width = 191;
            // 
            // businessAddressPostOfficeBoxDataGridViewTextBoxColumn
            // 
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressPostOfficeBox";
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn.HeaderText = "BusinessAddressPostOfficeBox";
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn.Name = "businessAddressPostOfficeBoxDataGridViewTextBoxColumn";
            this.businessAddressPostOfficeBoxDataGridViewTextBoxColumn.Width = 207;
            // 
            // businessAddressStateDataGridViewTextBoxColumn
            // 
            this.businessAddressStateDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressState";
            this.businessAddressStateDataGridViewTextBoxColumn.HeaderText = "BusinessAddressState";
            this.businessAddressStateDataGridViewTextBoxColumn.Name = "businessAddressStateDataGridViewTextBoxColumn";
            this.businessAddressStateDataGridViewTextBoxColumn.Width = 157;
            // 
            // businessAddressStreetDataGridViewTextBoxColumn
            // 
            this.businessAddressStreetDataGridViewTextBoxColumn.DataPropertyName = "BusinessAddressStreet";
            this.businessAddressStreetDataGridViewTextBoxColumn.HeaderText = "BusinessAddressStreet";
            this.businessAddressStreetDataGridViewTextBoxColumn.Name = "businessAddressStreetDataGridViewTextBoxColumn";
            this.businessAddressStreetDataGridViewTextBoxColumn.Width = 161;
            // 
            // businessHomePageDataGridViewTextBoxColumn
            // 
            this.businessHomePageDataGridViewTextBoxColumn.DataPropertyName = "BusinessHomePage";
            this.businessHomePageDataGridViewTextBoxColumn.HeaderText = "BusinessHomePage";
            this.businessHomePageDataGridViewTextBoxColumn.Name = "businessHomePageDataGridViewTextBoxColumn";
            this.businessHomePageDataGridViewTextBoxColumn.Width = 143;
            // 
            // categoriesDataGridViewTextBoxColumn
            // 
            this.categoriesDataGridViewTextBoxColumn.DataPropertyName = "Categories";
            this.categoriesDataGridViewTextBoxColumn.HeaderText = "Categories";
            this.categoriesDataGridViewTextBoxColumn.Name = "categoriesDataGridViewTextBoxColumn";
            this.categoriesDataGridViewTextBoxColumn.Width = 92;
            // 
            // childrenDataGridViewTextBoxColumn
            // 
            this.childrenDataGridViewTextBoxColumn.DataPropertyName = "Children";
            this.childrenDataGridViewTextBoxColumn.HeaderText = "Children";
            this.childrenDataGridViewTextBoxColumn.Name = "childrenDataGridViewTextBoxColumn";
            this.childrenDataGridViewTextBoxColumn.Width = 78;
            // 
            // companiesDataGridViewTextBoxColumn
            // 
            this.companiesDataGridViewTextBoxColumn.DataPropertyName = "Companies";
            this.companiesDataGridViewTextBoxColumn.HeaderText = "Companies";
            this.companiesDataGridViewTextBoxColumn.Name = "companiesDataGridViewTextBoxColumn";
            this.companiesDataGridViewTextBoxColumn.Width = 93;
            // 
            // computerNetworkNameDataGridViewTextBoxColumn
            // 
            this.computerNetworkNameDataGridViewTextBoxColumn.DataPropertyName = "ComputerNetworkName";
            this.computerNetworkNameDataGridViewTextBoxColumn.HeaderText = "ComputerNetworkName";
            this.computerNetworkNameDataGridViewTextBoxColumn.Name = "computerNetworkNameDataGridViewTextBoxColumn";
            this.computerNetworkNameDataGridViewTextBoxColumn.Width = 164;
            // 
            // creationTimeDataGridViewTextBoxColumn
            // 
            this.creationTimeDataGridViewTextBoxColumn.DataPropertyName = "CreationTime";
            this.creationTimeDataGridViewTextBoxColumn.HeaderText = "CreationTime";
            this.creationTimeDataGridViewTextBoxColumn.Name = "creationTimeDataGridViewTextBoxColumn";
            this.creationTimeDataGridViewTextBoxColumn.Width = 106;
            // 
            // customerIDDataGridViewTextBoxColumn
            // 
            this.customerIDDataGridViewTextBoxColumn.DataPropertyName = "CustomerID";
            this.customerIDDataGridViewTextBoxColumn.HeaderText = "CustomerID";
            this.customerIDDataGridViewTextBoxColumn.Name = "customerIDDataGridViewTextBoxColumn";
            this.customerIDDataGridViewTextBoxColumn.Width = 97;
            // 
            // departmentDataGridViewTextBoxColumn
            // 
            this.departmentDataGridViewTextBoxColumn.DataPropertyName = "Department";
            this.departmentDataGridViewTextBoxColumn.HeaderText = "Department";
            this.departmentDataGridViewTextBoxColumn.Name = "departmentDataGridViewTextBoxColumn";
            this.departmentDataGridViewTextBoxColumn.Width = 97;
            // 
            // email1AddressDataGridViewTextBoxColumn
            // 
            this.email1AddressDataGridViewTextBoxColumn.DataPropertyName = "Email1Address";
            this.email1AddressDataGridViewTextBoxColumn.HeaderText = "Email1Address";
            this.email1AddressDataGridViewTextBoxColumn.Name = "email1AddressDataGridViewTextBoxColumn";
            this.email1AddressDataGridViewTextBoxColumn.Width = 114;
            // 
            // email1AddressTypeDataGridViewTextBoxColumn
            // 
            this.email1AddressTypeDataGridViewTextBoxColumn.DataPropertyName = "Email1AddressType";
            this.email1AddressTypeDataGridViewTextBoxColumn.HeaderText = "Email1AddressType";
            this.email1AddressTypeDataGridViewTextBoxColumn.Name = "email1AddressTypeDataGridViewTextBoxColumn";
            this.email1AddressTypeDataGridViewTextBoxColumn.Width = 142;
            // 
            // email1DisplayNameDataGridViewTextBoxColumn
            // 
            this.email1DisplayNameDataGridViewTextBoxColumn.DataPropertyName = "Email1DisplayName";
            this.email1DisplayNameDataGridViewTextBoxColumn.HeaderText = "Email1DisplayName";
            this.email1DisplayNameDataGridViewTextBoxColumn.Name = "email1DisplayNameDataGridViewTextBoxColumn";
            this.email1DisplayNameDataGridViewTextBoxColumn.Width = 142;
            // 
            // email2AddressDataGridViewTextBoxColumn
            // 
            this.email2AddressDataGridViewTextBoxColumn.DataPropertyName = "Email2Address";
            this.email2AddressDataGridViewTextBoxColumn.HeaderText = "Email2Address";
            this.email2AddressDataGridViewTextBoxColumn.Name = "email2AddressDataGridViewTextBoxColumn";
            this.email2AddressDataGridViewTextBoxColumn.Width = 114;
            // 
            // email2AddressTypeDataGridViewTextBoxColumn
            // 
            this.email2AddressTypeDataGridViewTextBoxColumn.DataPropertyName = "Email2AddressType";
            this.email2AddressTypeDataGridViewTextBoxColumn.HeaderText = "Email2AddressType";
            this.email2AddressTypeDataGridViewTextBoxColumn.Name = "email2AddressTypeDataGridViewTextBoxColumn";
            this.email2AddressTypeDataGridViewTextBoxColumn.Width = 142;
            // 
            // email2DisplayNameDataGridViewTextBoxColumn
            // 
            this.email2DisplayNameDataGridViewTextBoxColumn.DataPropertyName = "Email2DisplayName";
            this.email2DisplayNameDataGridViewTextBoxColumn.HeaderText = "Email2DisplayName";
            this.email2DisplayNameDataGridViewTextBoxColumn.Name = "email2DisplayNameDataGridViewTextBoxColumn";
            this.email2DisplayNameDataGridViewTextBoxColumn.Width = 142;
            // 
            // email3AddressDataGridViewTextBoxColumn
            // 
            this.email3AddressDataGridViewTextBoxColumn.DataPropertyName = "Email3Address";
            this.email3AddressDataGridViewTextBoxColumn.HeaderText = "Email3Address";
            this.email3AddressDataGridViewTextBoxColumn.Name = "email3AddressDataGridViewTextBoxColumn";
            this.email3AddressDataGridViewTextBoxColumn.Width = 114;
            // 
            // email3AddressTypeDataGridViewTextBoxColumn
            // 
            this.email3AddressTypeDataGridViewTextBoxColumn.DataPropertyName = "Email3AddressType";
            this.email3AddressTypeDataGridViewTextBoxColumn.HeaderText = "Email3AddressType";
            this.email3AddressTypeDataGridViewTextBoxColumn.Name = "email3AddressTypeDataGridViewTextBoxColumn";
            this.email3AddressTypeDataGridViewTextBoxColumn.Width = 142;
            // 
            // email3DisplayNameDataGridViewTextBoxColumn
            // 
            this.email3DisplayNameDataGridViewTextBoxColumn.DataPropertyName = "Email3DisplayName";
            this.email3DisplayNameDataGridViewTextBoxColumn.HeaderText = "Email3DisplayName";
            this.email3DisplayNameDataGridViewTextBoxColumn.Name = "email3DisplayNameDataGridViewTextBoxColumn";
            this.email3DisplayNameDataGridViewTextBoxColumn.Width = 142;
            // 
            // entryIDDataGridViewTextBoxColumn
            // 
            this.entryIDDataGridViewTextBoxColumn.DataPropertyName = "EntryID";
            this.entryIDDataGridViewTextBoxColumn.HeaderText = "EntryID";
            this.entryIDDataGridViewTextBoxColumn.Name = "entryIDDataGridViewTextBoxColumn";
            this.entryIDDataGridViewTextBoxColumn.Width = 74;
            // 
            // fileAsDataGridViewTextBoxColumn
            // 
            this.fileAsDataGridViewTextBoxColumn.DataPropertyName = "FileAs";
            this.fileAsDataGridViewTextBoxColumn.HeaderText = "FileAs";
            this.fileAsDataGridViewTextBoxColumn.Name = "fileAsDataGridViewTextBoxColumn";
            this.fileAsDataGridViewTextBoxColumn.Width = 66;
            // 
            // fTPSiteDataGridViewTextBoxColumn
            // 
            this.fTPSiteDataGridViewTextBoxColumn.DataPropertyName = "FTPSite";
            this.fTPSiteDataGridViewTextBoxColumn.HeaderText = "FTPSite";
            this.fTPSiteDataGridViewTextBoxColumn.Name = "fTPSiteDataGridViewTextBoxColumn";
            this.fTPSiteDataGridViewTextBoxColumn.Width = 77;
            // 
            // governmentIDNumberDataGridViewTextBoxColumn
            // 
            this.governmentIDNumberDataGridViewTextBoxColumn.DataPropertyName = "GovernmentIDNumber";
            this.governmentIDNumberDataGridViewTextBoxColumn.HeaderText = "GovernmentIDNumber";
            this.governmentIDNumberDataGridViewTextBoxColumn.Name = "governmentIDNumberDataGridViewTextBoxColumn";
            this.governmentIDNumberDataGridViewTextBoxColumn.Width = 156;
            // 
            // hobbyDataGridViewTextBoxColumn
            // 
            this.hobbyDataGridViewTextBoxColumn.DataPropertyName = "Hobby";
            this.hobbyDataGridViewTextBoxColumn.HeaderText = "Hobby";
            this.hobbyDataGridViewTextBoxColumn.Name = "hobbyDataGridViewTextBoxColumn";
            this.hobbyDataGridViewTextBoxColumn.Width = 68;
            // 
            // homeAddressDataGridViewTextBoxColumn
            // 
            this.homeAddressDataGridViewTextBoxColumn.DataPropertyName = "HomeAddress";
            this.homeAddressDataGridViewTextBoxColumn.HeaderText = "HomeAddress";
            this.homeAddressDataGridViewTextBoxColumn.Name = "homeAddressDataGridViewTextBoxColumn";
            this.homeAddressDataGridViewTextBoxColumn.Width = 109;
            // 
            // homeAddressCityDataGridViewTextBoxColumn
            // 
            this.homeAddressCityDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressCity";
            this.homeAddressCityDataGridViewTextBoxColumn.HeaderText = "HomeAddressCity";
            this.homeAddressCityDataGridViewTextBoxColumn.Name = "homeAddressCityDataGridViewTextBoxColumn";
            this.homeAddressCityDataGridViewTextBoxColumn.Width = 130;
            // 
            // homeAddressCountryDataGridViewTextBoxColumn
            // 
            this.homeAddressCountryDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressCountry";
            this.homeAddressCountryDataGridViewTextBoxColumn.HeaderText = "HomeAddressCountry";
            this.homeAddressCountryDataGridViewTextBoxColumn.Name = "homeAddressCountryDataGridViewTextBoxColumn";
            this.homeAddressCountryDataGridViewTextBoxColumn.Width = 152;
            // 
            // homeAddressPostalCodeDataGridViewTextBoxColumn
            // 
            this.homeAddressPostalCodeDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressPostalCode";
            this.homeAddressPostalCodeDataGridViewTextBoxColumn.HeaderText = "HomeAddressPostalCode";
            this.homeAddressPostalCodeDataGridViewTextBoxColumn.Name = "homeAddressPostalCodeDataGridViewTextBoxColumn";
            this.homeAddressPostalCodeDataGridViewTextBoxColumn.Width = 173;
            // 
            // homeAddressPostOfficeBoxDataGridViewTextBoxColumn
            // 
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressPostOfficeBox";
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn.HeaderText = "HomeAddressPostOfficeBox";
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn.Name = "homeAddressPostOfficeBoxDataGridViewTextBoxColumn";
            this.homeAddressPostOfficeBoxDataGridViewTextBoxColumn.Width = 189;
            // 
            // homeAddressStateDataGridViewTextBoxColumn
            // 
            this.homeAddressStateDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressState";
            this.homeAddressStateDataGridViewTextBoxColumn.HeaderText = "HomeAddressState";
            this.homeAddressStateDataGridViewTextBoxColumn.Name = "homeAddressStateDataGridViewTextBoxColumn";
            this.homeAddressStateDataGridViewTextBoxColumn.Width = 139;
            // 
            // homeAddressStreetDataGridViewTextBoxColumn
            // 
            this.homeAddressStreetDataGridViewTextBoxColumn.DataPropertyName = "HomeAddressStreet";
            this.homeAddressStreetDataGridViewTextBoxColumn.HeaderText = "HomeAddressStreet";
            this.homeAddressStreetDataGridViewTextBoxColumn.Name = "homeAddressStreetDataGridViewTextBoxColumn";
            this.homeAddressStreetDataGridViewTextBoxColumn.Width = 143;
            // 
            // initialsDataGridViewTextBoxColumn
            // 
            this.initialsDataGridViewTextBoxColumn.DataPropertyName = "Initials";
            this.initialsDataGridViewTextBoxColumn.HeaderText = "Initials";
            this.initialsDataGridViewTextBoxColumn.Name = "initialsDataGridViewTextBoxColumn";
            this.initialsDataGridViewTextBoxColumn.Width = 69;
            // 
            // jobTitleDataGridViewTextBoxColumn
            // 
            this.jobTitleDataGridViewTextBoxColumn.DataPropertyName = "JobTitle";
            this.jobTitleDataGridViewTextBoxColumn.HeaderText = "JobTitle";
            this.jobTitleDataGridViewTextBoxColumn.Name = "jobTitleDataGridViewTextBoxColumn";
            this.jobTitleDataGridViewTextBoxColumn.Width = 77;
            // 
            // journalDataGridViewCheckBoxColumn
            // 
            this.journalDataGridViewCheckBoxColumn.DataPropertyName = "Journal";
            this.journalDataGridViewCheckBoxColumn.HeaderText = "Journal";
            this.journalDataGridViewCheckBoxColumn.Name = "journalDataGridViewCheckBoxColumn";
            this.journalDataGridViewCheckBoxColumn.Width = 54;
            // 
            // languageDataGridViewTextBoxColumn
            // 
            this.languageDataGridViewTextBoxColumn.DataPropertyName = "Language";
            this.languageDataGridViewTextBoxColumn.HeaderText = "Language";
            this.languageDataGridViewTextBoxColumn.Name = "languageDataGridViewTextBoxColumn";
            this.languageDataGridViewTextBoxColumn.Width = 88;
            // 
            // lastModificationTimeDataGridViewTextBoxColumn
            // 
            this.lastModificationTimeDataGridViewTextBoxColumn.DataPropertyName = "LastModificationTime";
            this.lastModificationTimeDataGridViewTextBoxColumn.HeaderText = "LastModificationTime";
            this.lastModificationTimeDataGridViewTextBoxColumn.Name = "lastModificationTimeDataGridViewTextBoxColumn";
            this.lastModificationTimeDataGridViewTextBoxColumn.Width = 152;
            // 
            // mailingAddressDataGridViewTextBoxColumn
            // 
            this.mailingAddressDataGridViewTextBoxColumn.DataPropertyName = "MailingAddress";
            this.mailingAddressDataGridViewTextBoxColumn.HeaderText = "MailingAddress";
            this.mailingAddressDataGridViewTextBoxColumn.Name = "mailingAddressDataGridViewTextBoxColumn";
            this.mailingAddressDataGridViewTextBoxColumn.Width = 117;
            // 
            // mailingAddressCityDataGridViewTextBoxColumn
            // 
            this.mailingAddressCityDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressCity";
            this.mailingAddressCityDataGridViewTextBoxColumn.HeaderText = "MailingAddressCity";
            this.mailingAddressCityDataGridViewTextBoxColumn.Name = "mailingAddressCityDataGridViewTextBoxColumn";
            this.mailingAddressCityDataGridViewTextBoxColumn.Width = 138;
            // 
            // mailingAddressCountryDataGridViewTextBoxColumn
            // 
            this.mailingAddressCountryDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressCountry";
            this.mailingAddressCountryDataGridViewTextBoxColumn.HeaderText = "MailingAddressCountry";
            this.mailingAddressCountryDataGridViewTextBoxColumn.Name = "mailingAddressCountryDataGridViewTextBoxColumn";
            this.mailingAddressCountryDataGridViewTextBoxColumn.Width = 160;
            // 
            // mailingAddressPostalCodeDataGridViewTextBoxColumn
            // 
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressPostalCode";
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn.HeaderText = "MailingAddressPostalCode";
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn.Name = "mailingAddressPostalCodeDataGridViewTextBoxColumn";
            this.mailingAddressPostalCodeDataGridViewTextBoxColumn.Width = 181;
            // 
            // mailingAddressPostOfficeBoxDataGridViewTextBoxColumn
            // 
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressPostOfficeBox";
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn.HeaderText = "MailingAddressPostOfficeBox";
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn.Name = "mailingAddressPostOfficeBoxDataGridViewTextBoxColumn";
            this.mailingAddressPostOfficeBoxDataGridViewTextBoxColumn.Width = 197;
            // 
            // mailingAddressStateDataGridViewTextBoxColumn
            // 
            this.mailingAddressStateDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressState";
            this.mailingAddressStateDataGridViewTextBoxColumn.HeaderText = "MailingAddressState";
            this.mailingAddressStateDataGridViewTextBoxColumn.Name = "mailingAddressStateDataGridViewTextBoxColumn";
            this.mailingAddressStateDataGridViewTextBoxColumn.Width = 147;
            // 
            // mailingAddressStreetDataGridViewTextBoxColumn
            // 
            this.mailingAddressStreetDataGridViewTextBoxColumn.DataPropertyName = "MailingAddressStreet";
            this.mailingAddressStreetDataGridViewTextBoxColumn.HeaderText = "MailingAddressStreet";
            this.mailingAddressStreetDataGridViewTextBoxColumn.Name = "mailingAddressStreetDataGridViewTextBoxColumn";
            this.mailingAddressStreetDataGridViewTextBoxColumn.Width = 151;
            // 
            // managerNameDataGridViewTextBoxColumn
            // 
            this.managerNameDataGridViewTextBoxColumn.DataPropertyName = "ManagerName";
            this.managerNameDataGridViewTextBoxColumn.HeaderText = "ManagerName";
            this.managerNameDataGridViewTextBoxColumn.Name = "managerNameDataGridViewTextBoxColumn";
            this.managerNameDataGridViewTextBoxColumn.Width = 113;
            // 
            // messageClassDataGridViewTextBoxColumn
            // 
            this.messageClassDataGridViewTextBoxColumn.DataPropertyName = "MessageClass";
            this.messageClassDataGridViewTextBoxColumn.HeaderText = "MessageClass";
            this.messageClassDataGridViewTextBoxColumn.Name = "messageClassDataGridViewTextBoxColumn";
            this.messageClassDataGridViewTextBoxColumn.Width = 112;
            // 
            // mileageDataGridViewTextBoxColumn
            // 
            this.mileageDataGridViewTextBoxColumn.DataPropertyName = "Mileage";
            this.mileageDataGridViewTextBoxColumn.HeaderText = "Mileage";
            this.mileageDataGridViewTextBoxColumn.Name = "mileageDataGridViewTextBoxColumn";
            this.mileageDataGridViewTextBoxColumn.Width = 76;
            // 
            // noAgingDataGridViewCheckBoxColumn
            // 
            this.noAgingDataGridViewCheckBoxColumn.DataPropertyName = "NoAging";
            this.noAgingDataGridViewCheckBoxColumn.HeaderText = "NoAging";
            this.noAgingDataGridViewCheckBoxColumn.Name = "noAgingDataGridViewCheckBoxColumn";
            this.noAgingDataGridViewCheckBoxColumn.Width = 61;
            // 
            // officeLocationDataGridViewTextBoxColumn
            // 
            this.officeLocationDataGridViewTextBoxColumn.DataPropertyName = "OfficeLocation";
            this.officeLocationDataGridViewTextBoxColumn.HeaderText = "OfficeLocation";
            this.officeLocationDataGridViewTextBoxColumn.Name = "officeLocationDataGridViewTextBoxColumn";
            this.officeLocationDataGridViewTextBoxColumn.Width = 115;
            // 
            // organizationalIDNumberDataGridViewTextBoxColumn
            // 
            this.organizationalIDNumberDataGridViewTextBoxColumn.DataPropertyName = "OrganizationalIDNumber";
            this.organizationalIDNumberDataGridViewTextBoxColumn.HeaderText = "OrganizationalIDNumber";
            this.organizationalIDNumberDataGridViewTextBoxColumn.Name = "organizationalIDNumberDataGridViewTextBoxColumn";
            this.organizationalIDNumberDataGridViewTextBoxColumn.Width = 169;
            // 
            // otherAddressDataGridViewTextBoxColumn
            // 
            this.otherAddressDataGridViewTextBoxColumn.DataPropertyName = "OtherAddress";
            this.otherAddressDataGridViewTextBoxColumn.HeaderText = "OtherAddress";
            this.otherAddressDataGridViewTextBoxColumn.Name = "otherAddressDataGridViewTextBoxColumn";
            this.otherAddressDataGridViewTextBoxColumn.Width = 108;
            // 
            // otherAddressCityDataGridViewTextBoxColumn
            // 
            this.otherAddressCityDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressCity";
            this.otherAddressCityDataGridViewTextBoxColumn.HeaderText = "OtherAddressCity";
            this.otherAddressCityDataGridViewTextBoxColumn.Name = "otherAddressCityDataGridViewTextBoxColumn";
            this.otherAddressCityDataGridViewTextBoxColumn.Width = 129;
            // 
            // otherAddressCountryDataGridViewTextBoxColumn
            // 
            this.otherAddressCountryDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressCountry";
            this.otherAddressCountryDataGridViewTextBoxColumn.HeaderText = "OtherAddressCountry";
            this.otherAddressCountryDataGridViewTextBoxColumn.Name = "otherAddressCountryDataGridViewTextBoxColumn";
            this.otherAddressCountryDataGridViewTextBoxColumn.Width = 151;
            // 
            // otherAddressPostalCodeDataGridViewTextBoxColumn
            // 
            this.otherAddressPostalCodeDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressPostalCode";
            this.otherAddressPostalCodeDataGridViewTextBoxColumn.HeaderText = "OtherAddressPostalCode";
            this.otherAddressPostalCodeDataGridViewTextBoxColumn.Name = "otherAddressPostalCodeDataGridViewTextBoxColumn";
            this.otherAddressPostalCodeDataGridViewTextBoxColumn.Width = 172;
            // 
            // otherAddressPostOfficeBoxDataGridViewTextBoxColumn
            // 
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressPostOfficeBox";
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn.HeaderText = "OtherAddressPostOfficeBox";
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn.Name = "otherAddressPostOfficeBoxDataGridViewTextBoxColumn";
            this.otherAddressPostOfficeBoxDataGridViewTextBoxColumn.Width = 188;
            // 
            // otherAddressStateDataGridViewTextBoxColumn
            // 
            this.otherAddressStateDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressState";
            this.otherAddressStateDataGridViewTextBoxColumn.HeaderText = "OtherAddressState";
            this.otherAddressStateDataGridViewTextBoxColumn.Name = "otherAddressStateDataGridViewTextBoxColumn";
            this.otherAddressStateDataGridViewTextBoxColumn.Width = 138;
            // 
            // otherAddressStreetDataGridViewTextBoxColumn
            // 
            this.otherAddressStreetDataGridViewTextBoxColumn.DataPropertyName = "OtherAddressStreet";
            this.otherAddressStreetDataGridViewTextBoxColumn.HeaderText = "OtherAddressStreet";
            this.otherAddressStreetDataGridViewTextBoxColumn.Name = "otherAddressStreetDataGridViewTextBoxColumn";
            this.otherAddressStreetDataGridViewTextBoxColumn.Width = 142;
            // 
            // outlookInternalVersionDataGridViewTextBoxColumn
            // 
            this.outlookInternalVersionDataGridViewTextBoxColumn.DataPropertyName = "OutlookInternalVersion";
            this.outlookInternalVersionDataGridViewTextBoxColumn.HeaderText = "OutlookInternalVersion";
            this.outlookInternalVersionDataGridViewTextBoxColumn.Name = "outlookInternalVersionDataGridViewTextBoxColumn";
            this.outlookInternalVersionDataGridViewTextBoxColumn.Width = 161;
            // 
            // outlookVersionDataGridViewTextBoxColumn
            // 
            this.outlookVersionDataGridViewTextBoxColumn.DataPropertyName = "OutlookVersion";
            this.outlookVersionDataGridViewTextBoxColumn.HeaderText = "OutlookVersion";
            this.outlookVersionDataGridViewTextBoxColumn.Name = "outlookVersionDataGridViewTextBoxColumn";
            this.outlookVersionDataGridViewTextBoxColumn.Width = 118;
            // 
            // personalHomePageDataGridViewTextBoxColumn
            // 
            this.personalHomePageDataGridViewTextBoxColumn.DataPropertyName = "PersonalHomePage";
            this.personalHomePageDataGridViewTextBoxColumn.HeaderText = "PersonalHomePage";
            this.personalHomePageDataGridViewTextBoxColumn.Name = "personalHomePageDataGridViewTextBoxColumn";
            this.personalHomePageDataGridViewTextBoxColumn.Width = 142;
            // 
            // professionDataGridViewTextBoxColumn
            // 
            this.professionDataGridViewTextBoxColumn.DataPropertyName = "Profession";
            this.professionDataGridViewTextBoxColumn.HeaderText = "Profession";
            this.professionDataGridViewTextBoxColumn.Name = "professionDataGridViewTextBoxColumn";
            this.professionDataGridViewTextBoxColumn.Width = 91;
            // 
            // radioTelephoneNumberDataGridViewTextBoxColumn
            // 
            this.radioTelephoneNumberDataGridViewTextBoxColumn.DataPropertyName = "RadioTelephoneNumber";
            this.radioTelephoneNumberDataGridViewTextBoxColumn.HeaderText = "RadioTelephoneNumber";
            this.radioTelephoneNumberDataGridViewTextBoxColumn.Name = "radioTelephoneNumberDataGridViewTextBoxColumn";
            this.radioTelephoneNumberDataGridViewTextBoxColumn.Width = 168;
            // 
            // referredByDataGridViewTextBoxColumn
            // 
            this.referredByDataGridViewTextBoxColumn.DataPropertyName = "ReferredBy";
            this.referredByDataGridViewTextBoxColumn.HeaderText = "ReferredBy";
            this.referredByDataGridViewTextBoxColumn.Name = "referredByDataGridViewTextBoxColumn";
            this.referredByDataGridViewTextBoxColumn.Width = 95;
            // 
            // savedDataGridViewCheckBoxColumn
            // 
            this.savedDataGridViewCheckBoxColumn.DataPropertyName = "Saved";
            this.savedDataGridViewCheckBoxColumn.HeaderText = "Saved";
            this.savedDataGridViewCheckBoxColumn.Name = "savedDataGridViewCheckBoxColumn";
            this.savedDataGridViewCheckBoxColumn.Width = 49;
            // 
            // sizeDataGridViewTextBoxColumn
            // 
            this.sizeDataGridViewTextBoxColumn.DataPropertyName = "Size";
            this.sizeDataGridViewTextBoxColumn.HeaderText = "Size";
            this.sizeDataGridViewTextBoxColumn.Name = "sizeDataGridViewTextBoxColumn";
            this.sizeDataGridViewTextBoxColumn.Width = 56;
            // 
            // spouseDataGridViewTextBoxColumn
            // 
            this.spouseDataGridViewTextBoxColumn.DataPropertyName = "Spouse";
            this.spouseDataGridViewTextBoxColumn.HeaderText = "Spouse";
            this.spouseDataGridViewTextBoxColumn.Name = "spouseDataGridViewTextBoxColumn";
            this.spouseDataGridViewTextBoxColumn.Width = 74;
            // 
            // subjectDataGridViewTextBoxColumn
            // 
            this.subjectDataGridViewTextBoxColumn.DataPropertyName = "Subject";
            this.subjectDataGridViewTextBoxColumn.HeaderText = "Subject";
            this.subjectDataGridViewTextBoxColumn.Name = "subjectDataGridViewTextBoxColumn";
            this.subjectDataGridViewTextBoxColumn.Width = 75;
            // 
            // unReadDataGridViewCheckBoxColumn
            // 
            this.unReadDataGridViewCheckBoxColumn.DataPropertyName = "UnRead";
            this.unReadDataGridViewCheckBoxColumn.HeaderText = "UnRead";
            this.unReadDataGridViewCheckBoxColumn.Name = "unReadDataGridViewCheckBoxColumn";
            this.unReadDataGridViewCheckBoxColumn.Width = 59;
            // 
            // user1DataGridViewTextBoxColumn
            // 
            this.user1DataGridViewTextBoxColumn.DataPropertyName = "User1";
            this.user1DataGridViewTextBoxColumn.HeaderText = "User1";
            this.user1DataGridViewTextBoxColumn.Name = "user1DataGridViewTextBoxColumn";
            this.user1DataGridViewTextBoxColumn.Width = 65;
            // 
            // user2DataGridViewTextBoxColumn
            // 
            this.user2DataGridViewTextBoxColumn.DataPropertyName = "User2";
            this.user2DataGridViewTextBoxColumn.HeaderText = "User2";
            this.user2DataGridViewTextBoxColumn.Name = "user2DataGridViewTextBoxColumn";
            this.user2DataGridViewTextBoxColumn.Width = 65;
            // 
            // user3DataGridViewTextBoxColumn
            // 
            this.user3DataGridViewTextBoxColumn.DataPropertyName = "User3";
            this.user3DataGridViewTextBoxColumn.HeaderText = "User3";
            this.user3DataGridViewTextBoxColumn.Name = "user3DataGridViewTextBoxColumn";
            this.user3DataGridViewTextBoxColumn.Width = 65;
            // 
            // user4DataGridViewTextBoxColumn
            // 
            this.user4DataGridViewTextBoxColumn.DataPropertyName = "User4";
            this.user4DataGridViewTextBoxColumn.HeaderText = "User4";
            this.user4DataGridViewTextBoxColumn.Name = "user4DataGridViewTextBoxColumn";
            this.user4DataGridViewTextBoxColumn.Width = 65;
            // 
            // userCertificateDataGridViewTextBoxColumn
            // 
            this.userCertificateDataGridViewTextBoxColumn.DataPropertyName = "UserCertificate";
            this.userCertificateDataGridViewTextBoxColumn.HeaderText = "UserCertificate";
            this.userCertificateDataGridViewTextBoxColumn.Name = "userCertificateDataGridViewTextBoxColumn";
            this.userCertificateDataGridViewTextBoxColumn.Width = 116;
            // 
            // webPageDataGridViewTextBoxColumn
            // 
            this.webPageDataGridViewTextBoxColumn.DataPropertyName = "WebPage";
            this.webPageDataGridViewTextBoxColumn.HeaderText = "WebPage";
            this.webPageDataGridViewTextBoxColumn.Name = "webPageDataGridViewTextBoxColumn";
            this.webPageDataGridViewTextBoxColumn.Width = 87;
            // 
            // yomiCompanyNameDataGridViewTextBoxColumn
            // 
            this.yomiCompanyNameDataGridViewTextBoxColumn.DataPropertyName = "YomiCompanyName";
            this.yomiCompanyNameDataGridViewTextBoxColumn.HeaderText = "YomiCompanyName";
            this.yomiCompanyNameDataGridViewTextBoxColumn.Name = "yomiCompanyNameDataGridViewTextBoxColumn";
            this.yomiCompanyNameDataGridViewTextBoxColumn.Width = 142;
            // 
            // yomiFirstNameDataGridViewTextBoxColumn
            // 
            this.yomiFirstNameDataGridViewTextBoxColumn.DataPropertyName = "YomiFirstName";
            this.yomiFirstNameDataGridViewTextBoxColumn.HeaderText = "YomiFirstName";
            this.yomiFirstNameDataGridViewTextBoxColumn.Name = "yomiFirstNameDataGridViewTextBoxColumn";
            this.yomiFirstNameDataGridViewTextBoxColumn.Width = 115;
            // 
            // yomiLastNameDataGridViewTextBoxColumn
            // 
            this.yomiLastNameDataGridViewTextBoxColumn.DataPropertyName = "YomiLastName";
            this.yomiLastNameDataGridViewTextBoxColumn.HeaderText = "YomiLastName";
            this.yomiLastNameDataGridViewTextBoxColumn.Name = "yomiLastNameDataGridViewTextBoxColumn";
            this.yomiLastNameDataGridViewTextBoxColumn.Width = 115;
            // 
            // genderDataGridViewTextBoxColumn
            // 
            this.genderDataGridViewTextBoxColumn.DataPropertyName = "Gender";
            this.genderDataGridViewTextBoxColumn.HeaderText = "Gender";
            this.genderDataGridViewTextBoxColumn.Name = "genderDataGridViewTextBoxColumn";
            this.genderDataGridViewTextBoxColumn.Width = 73;
            // 
            // importanceDataGridViewTextBoxColumn
            // 
            this.importanceDataGridViewTextBoxColumn.DataPropertyName = "Importance";
            this.importanceDataGridViewTextBoxColumn.HeaderText = "Importance";
            this.importanceDataGridViewTextBoxColumn.Name = "importanceDataGridViewTextBoxColumn";
            this.importanceDataGridViewTextBoxColumn.Width = 95;
            // 
            // selectedMailingAddressDataGridViewTextBoxColumn
            // 
            this.selectedMailingAddressDataGridViewTextBoxColumn.DataPropertyName = "SelectedMailingAddress";
            this.selectedMailingAddressDataGridViewTextBoxColumn.HeaderText = "SelectedMailingAddress";
            this.selectedMailingAddressDataGridViewTextBoxColumn.Name = "selectedMailingAddressDataGridViewTextBoxColumn";
            this.selectedMailingAddressDataGridViewTextBoxColumn.Width = 167;
            // 
            // sensitivityDataGridViewTextBoxColumn
            // 
            this.sensitivityDataGridViewTextBoxColumn.DataPropertyName = "Sensitivity";
            this.sensitivityDataGridViewTextBoxColumn.HeaderText = "Sensitivity";
            this.sensitivityDataGridViewTextBoxColumn.Name = "sensitivityDataGridViewTextBoxColumn";
            this.sensitivityDataGridViewTextBoxColumn.Width = 90;
            // 
            // outlookNdxDataGridViewTextBoxColumn
            // 
            this.outlookNdxDataGridViewTextBoxColumn.DataPropertyName = "OutlookNdx";
            this.outlookNdxDataGridViewTextBoxColumn.HeaderText = "OutlookNdx";
            this.outlookNdxDataGridViewTextBoxColumn.Name = "outlookNdxDataGridViewTextBoxColumn";
            this.outlookNdxDataGridViewTextBoxColumn.Width = 98;
            // 
            // bodyDataGridViewTextBoxColumn
            // 
            this.bodyDataGridViewTextBoxColumn.DataPropertyName = "Body";
            this.bodyDataGridViewTextBoxColumn.HeaderText = "Body";
            this.bodyDataGridViewTextBoxColumn.Name = "bodyDataGridViewTextBoxColumn";
            this.bodyDataGridViewTextBoxColumn.Width = 60;
            // 
            // myContactBindingSource
            // 
            this.myContactBindingSource.DataSource = typeof(OutlookContactsAnalyser.MyContact);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(881, 638);
            this.Controls.Add(this.tlpMain);
            this.Controls.Add(this.menuStrip1);
            this.HelpButton = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "OutlookContactsAnalyser";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tlpMain.ResumeLayout(false);
            this.tlpMain.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvContacts)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.groupBoxViewRadioButtons.ResumeLayout(false);
            this.groupBoxViewRadioButtons.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.myContactBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tlpMain;
        private System.Windows.Forms.DataGridView dgvContacts;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.BindingSource myContactBindingSource;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewHideUnusedColumnsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowAllContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowDuplicateContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowSimilarContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowVaguelySimilarContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowWithTheSameNameToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowWithTruncatedNotesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowWithDislikedContentsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileNewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileOpenToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem fileSaveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileSaveAsToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripMenuItem filePrintToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem filePrintPreviewToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripMenuItem fileExitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editUndoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editRedoToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripMenuItem editCutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editCopyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem editPasteToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
        private System.Windows.Forms.ToolStripMenuItem editSelectAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsCustomizeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsOptionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpContentsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpIndexToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpSearchToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
        private System.Windows.Forms.ToolStripMenuItem helpAboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileRefreshFoldersToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileLoadContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileSaveChangedContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fileStopGetContactsThreadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsAnalyseContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsEditContactDetailsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsFieldsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsNewContactToolStripMenuItem;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnRefreshFolders;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbFolder;
        private System.Windows.Forms.Button btnGetContacts;
        private System.Windows.Forms.TextBox txtMaxContactstoRead;
        private System.Windows.Forms.Label lblMaxContactstoRead;
        private System.Windows.Forms.Button btnContactDetails;
        private System.Windows.Forms.Button btnAnalyse;
        private System.Windows.Forms.Button btnViewFields;
        private System.Windows.Forms.ToolStripMenuItem cleanNotesFieldsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem markDuplicatesForDeletionToolStripMenuItem;
        //private System.Windows.Forms.DataGridViewTextBoxColumn indiciesIdenticalDataGridViewTextBoxColumn;
        //private System.Windows.Forms.DataGridViewTextBoxColumn indiciesSimilarDataGridViewTextBoxColumn;
        //private System.Windows.Forms.DataGridViewTextBoxColumn indiciesVaguelySimilarDataGridViewTextBoxColumn;
        //private System.Windows.Forms.DataGridViewTextBoxColumn indiciesSameNameDataGridViewTextBoxColumn;
        public System.Windows.Forms.TextBox txtStatusLine;
        private System.Windows.Forms.ToolStripMenuItem fixPhoneNumbersToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowWithDuplicateAddressesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem mergeDuplicateContactsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowWithTheSameNameAndSuffixToolStripMenuItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn ndxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fullNameAndCompanyDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn RowSelect;
        private System.Windows.Forms.DataGridViewTextBoxColumn titleDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn lastNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn middleNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn firstNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn suffixDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn nickNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn companyNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fullNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn lastNameAndFirstNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn companyAndFullNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn primaryTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mobileTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn home2TelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn companyMainTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn business2TelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeFaxNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessFaxNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherFaxNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn callbackTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn carTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn pagerNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn iSDNNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn telexNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn tTYTDDTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn accountDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn anniversaryDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn assistantNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn assistantTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn billingInformationDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn birthdayDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressCityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressCountryDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressPostalCodeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressPostOfficeBoxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressStateDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessAddressStreetDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn businessHomePageDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn categoriesDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn childrenDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn companiesDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn computerNetworkNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn creationTimeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn customerIDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn departmentDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email1AddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email1AddressTypeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email1DisplayNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email2AddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email2AddressTypeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email2DisplayNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email3AddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email3AddressTypeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn email3DisplayNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn entryIDDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fileAsDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fTPSiteDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn governmentIDNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn hobbyDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressCityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressCountryDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressPostalCodeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressPostOfficeBoxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressStateDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn homeAddressStreetDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn initialsDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn jobTitleDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn journalDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn languageDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn lastModificationTimeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressCityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressCountryDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressPostalCodeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressPostOfficeBoxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressStateDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mailingAddressStreetDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn managerNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn messageClassDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn mileageDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn noAgingDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn officeLocationDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn organizationalIDNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressCityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressCountryDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressPostalCodeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressPostOfficeBoxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressStateDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn otherAddressStreetDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn outlookInternalVersionDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn outlookVersionDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn personalHomePageDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn professionDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn radioTelephoneNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn referredByDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn savedDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn sizeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn spouseDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn subjectDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewCheckBoxColumn unReadDataGridViewCheckBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn user1DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn user2DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn user3DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn user4DataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn userCertificateDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn webPageDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yomiCompanyNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yomiFirstNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yomiLastNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn genderDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn importanceDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn selectedMailingAddressDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn sensitivityDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn outlookNdxDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn bodyDataGridViewTextBoxColumn;
        private System.Windows.Forms.ToolStripMenuItem mailingAddressesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowUnlinkedMailingAddressesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowLinkedMailingAddressesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowEmptyMailingAddressesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewShowMismatchedAddressComponentsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showWithDuplicatePhoneNumbersToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem showWithDuplicateEmailAddressesToolStripMenuItem;
        private System.Windows.Forms.RadioButton radioViewAll;
        private System.Windows.Forms.GroupBox groupBoxViewRadioButtons;
        private System.Windows.Forms.RadioButton radioViewDuplicates;
        private System.Windows.Forms.RadioButton radioViewDislikedContents;
        private System.Windows.Forms.RadioButton radioViewMismatchedAddrAndComp;
        private System.Windows.Forms.RadioButton radioViewEmptyMailingAddresses;
        private System.Windows.Forms.RadioButton radioViewLinkedMailingAddresses;
        private System.Windows.Forms.RadioButton radioViewUnlinkedMailingAddresses;
        private System.Windows.Forms.RadioButton radioViewTruncatedNotes;
        private System.Windows.Forms.RadioButton radioViewDuplicateEmails;
        private System.Windows.Forms.RadioButton radioViewDuplicatePhoneNumbers;
        private System.Windows.Forms.RadioButton radioViewDuplicateAddresses;
        private System.Windows.Forms.RadioButton radioViewSameNameAndSuffix;
        private System.Windows.Forms.RadioButton radioViewSameName;
        private System.Windows.Forms.RadioButton radioViewVaguelySimilar;
        private System.Windows.Forms.RadioButton radioViewSimilar;
        private System.Windows.Forms.RadioButton radioViewAllByName;
        private System.Windows.Forms.RadioButton radioViewMinimalInfo;
    }
}

