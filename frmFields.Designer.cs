namespace OutlookContactsAnalyser
{
  partial class frmFields
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
        this.tlpMain = new System.Windows.Forms.TableLayoutPanel();
        this.dgvFields = new System.Windows.Forms.DataGridView();
        this.ndxDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
        this.fieldNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
        this.occurancesDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
        this.myFieldBindingSource = new System.Windows.Forms.BindingSource(this.components);
        this.txtLog = new System.Windows.Forms.TextBox();
        this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
        this.btnViewFields = new System.Windows.Forms.Button();
        this.cboFieldViewSelector = new System.Windows.Forms.ComboBox();
        this.label1 = new System.Windows.Forms.Label();
        this.txtNumberOfFields = new System.Windows.Forms.TextBox();
        this.tlpMain.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dgvFields)).BeginInit();
        ((System.ComponentModel.ISupportInitialize)(this.myFieldBindingSource)).BeginInit();
        this.tableLayoutPanel1.SuspendLayout();
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
        this.tlpMain.Controls.Add(this.dgvFields, 0, 0);
        this.tlpMain.Controls.Add(this.txtLog, 0, 4);
        this.tlpMain.Controls.Add(this.tableLayoutPanel1, 0, 2);
        this.tlpMain.Location = new System.Drawing.Point(19, 4);
        this.tlpMain.Name = "tlpMain";
        this.tlpMain.RowCount = 5;
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 43F));
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle());
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 70F));
        this.tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
        this.tlpMain.Size = new System.Drawing.Size(503, 478);
        this.tlpMain.TabIndex = 1;
        // 
        // dgvFields
        // 
        this.dgvFields.AllowUserToAddRows = false;
        this.dgvFields.AllowUserToDeleteRows = false;
        this.dgvFields.AllowUserToOrderColumns = true;
        this.dgvFields.AutoGenerateColumns = false;
        this.dgvFields.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
        this.dgvFields.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
        dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle1.BackColor = System.Drawing.Color.NavajoWhite;
        dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        dataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black;
        dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
        dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
        dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
        this.dgvFields.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
        this.dgvFields.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dgvFields.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ndxDataGridViewTextBoxColumn,
            this.fieldNameDataGridViewTextBoxColumn,
            this.occurancesDataGridViewTextBoxColumn});
        this.dgvFields.DataSource = this.myFieldBindingSource;
        dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Info;
        dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
        dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
        dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
        dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
        this.dgvFields.DefaultCellStyle = dataGridViewCellStyle2;
        this.dgvFields.Dock = System.Windows.Forms.DockStyle.Fill;
        this.dgvFields.Location = new System.Drawing.Point(3, 3);
        this.dgvFields.Name = "dgvFields";
        dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
        dataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control;
        dataGridViewCellStyle3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        dataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText;
        dataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight;
        dataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
        dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
        this.dgvFields.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
        this.dgvFields.RowHeadersVisible = false;
        this.dgvFields.Size = new System.Drawing.Size(497, 359);
        this.dgvFields.TabIndex = 1;
        // 
        // ndxDataGridViewTextBoxColumn
        // 
        this.ndxDataGridViewTextBoxColumn.DataPropertyName = "Ndx";
        this.ndxDataGridViewTextBoxColumn.HeaderText = "Ndx";
        this.ndxDataGridViewTextBoxColumn.Name = "ndxDataGridViewTextBoxColumn";
        this.ndxDataGridViewTextBoxColumn.Width = 54;
        // 
        // fieldNameDataGridViewTextBoxColumn
        // 
        this.fieldNameDataGridViewTextBoxColumn.DataPropertyName = "FieldName";
        this.fieldNameDataGridViewTextBoxColumn.HeaderText = "FieldName";
        this.fieldNameDataGridViewTextBoxColumn.Name = "fieldNameDataGridViewTextBoxColumn";
        this.fieldNameDataGridViewTextBoxColumn.Width = 91;
        // 
        // occurancesDataGridViewTextBoxColumn
        // 
        this.occurancesDataGridViewTextBoxColumn.DataPropertyName = "Occurances";
        this.occurancesDataGridViewTextBoxColumn.HeaderText = "Occurances";
        this.occurancesDataGridViewTextBoxColumn.Name = "occurancesDataGridViewTextBoxColumn";
        // 
        // myFieldBindingSource
        // 
        this.myFieldBindingSource.DataSource = typeof(OutlookContactsAnalyser.MyField);
        // 
        // txtLog
        // 
        this.txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                    | System.Windows.Forms.AnchorStyles.Left)
                    | System.Windows.Forms.AnchorStyles.Right)));
        this.txtLog.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
        this.txtLog.Location = new System.Drawing.Point(3, 411);
        this.txtLog.Multiline = true;
        this.txtLog.Name = "txtLog";
        this.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
        this.txtLog.Size = new System.Drawing.Size(497, 64);
        this.txtLog.TabIndex = 6;
        // 
        // tableLayoutPanel1
        // 
        this.tableLayoutPanel1.ColumnCount = 5;
        this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
        this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 287F));
        this.tableLayoutPanel1.Controls.Add(this.btnViewFields, 2, 0);
        this.tableLayoutPanel1.Controls.Add(this.cboFieldViewSelector, 4, 0);
        this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
        this.tableLayoutPanel1.Controls.Add(this.txtNumberOfFields, 1, 0);
        this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 368);
        this.tableLayoutPanel1.Name = "tableLayoutPanel1";
        this.tableLayoutPanel1.RowCount = 1;
        this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
        this.tableLayoutPanel1.Size = new System.Drawing.Size(394, 30);
        this.tableLayoutPanel1.TabIndex = 7;
        // 
        // btnViewFields
        // 
        this.btnViewFields.Location = new System.Drawing.Point(101, 3);
        this.btnViewFields.Name = "btnViewFields";
        this.btnViewFields.Size = new System.Drawing.Size(75, 22);
        this.btnViewFields.TabIndex = 13;
        this.btnViewFields.Text = "View";
        this.btnViewFields.UseVisualStyleBackColor = true;
        this.btnViewFields.Click += new System.EventHandler(this.btnViewFields_Click);
        // 
        // cboFieldViewSelector
        // 
        this.cboFieldViewSelector.FormattingEnabled = true;
        this.cboFieldViewSelector.Items.AddRange(new object[] {
            "All Fields",
            "Never used",
            "Seldom used (<10%)",
            "Sometimes used (not never, not always)",
            "Always used"});
        this.cboFieldViewSelector.Location = new System.Drawing.Point(182, 3);
        this.cboFieldViewSelector.Name = "cboFieldViewSelector";
        this.cboFieldViewSelector.Size = new System.Drawing.Size(200, 21);
        this.cboFieldViewSelector.TabIndex = 19;
        this.cboFieldViewSelector.SelectedIndexChanged += new System.EventHandler(this.cboFieldViewSelector_SelectedIndexChanged);
        // 
        // label1
        // 
        this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
        this.label1.AutoSize = true;
        this.label1.Location = new System.Drawing.Point(3, 8);
        this.label1.Name = "label1";
        this.label1.Size = new System.Drawing.Size(34, 13);
        this.label1.TabIndex = 17;
        this.label1.Text = "Fields";
        // 
        // txtNumberOfFields
        // 
        this.txtNumberOfFields.Location = new System.Drawing.Point(43, 3);
        this.txtNumberOfFields.Name = "txtNumberOfFields";
        this.txtNumberOfFields.ReadOnly = true;
        this.txtNumberOfFields.Size = new System.Drawing.Size(52, 20);
        this.txtNumberOfFields.TabIndex = 2;
        this.txtNumberOfFields.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
        // 
        // frmFields
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(542, 494);
        this.Controls.Add(this.tlpMain);
        this.Name = "frmFields";
        this.Text = "Field Usage";
        this.Load += new System.EventHandler(this.FieldsForm_Load);
        this.tlpMain.ResumeLayout(false);
        this.tlpMain.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)(this.dgvFields)).EndInit();
        ((System.ComponentModel.ISupportInitialize)(this.myFieldBindingSource)).EndInit();
        this.tableLayoutPanel1.ResumeLayout(false);
        this.tableLayoutPanel1.PerformLayout();
        this.ResumeLayout(false);

    }

    #endregion

    private System.Windows.Forms.TableLayoutPanel tlpMain;
    private System.Windows.Forms.DataGridView dgvFields;
    private System.Windows.Forms.TextBox txtLog;
    private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
    private System.Windows.Forms.TextBox txtNumberOfFields;
    private System.Windows.Forms.Button btnViewFields;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.ComboBox cboFieldViewSelector;
    private System.Windows.Forms.DataGridViewTextBoxColumn ndxDataGridViewTextBoxColumn;
    private System.Windows.Forms.DataGridViewTextBoxColumn fieldNameDataGridViewTextBoxColumn;
    private System.Windows.Forms.DataGridViewTextBoxColumn occurancesDataGridViewTextBoxColumn;
    private System.Windows.Forms.BindingSource myFieldBindingSource;
  }
}