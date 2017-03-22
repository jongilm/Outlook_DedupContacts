using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
//using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookContactsAnalyser
{
  public partial class frmFields : Form
  {
    public List<MyField> m_FieldList = null;
    private List<MyField> m_FieldListView = null;
    public int m_ContactList_Count = 0;

    public frmFields()
    {
      InitializeComponent();
      m_FieldList = new List<MyField>();
      m_FieldListView = new List<MyField>();
      txtLog.AppendText("frmFields Construct\r\n");
    }
    //////////////////////////////////////
    // When the form is loaded.
    //////////////////////////////////////
    private void FieldsForm_Load(object sender, EventArgs e)
    {
      //m_FieldList = new List<MyField>();
      //m_FieldListView = new List<MyField>();
      txtLog.AppendText("frmFields Load\r\n");

      this.Cursor = Cursors.WaitCursor;
      //dgvFields.DataSource = m_ContactList;
      cboFieldViewSelector.SelectedIndex = 0;
      //btnViewFields_Click(sender, e);
      this.Cursor = Cursors.Arrow;
    }

    private void btnViewFields_Click(object sender, EventArgs e)
    {
      txtLog.AppendText("View Fields [" + cboFieldViewSelector.SelectedIndex + "]\r\n");
      this.Cursor = Cursors.WaitCursor;
      m_FieldListView.Clear();
      switch(cboFieldViewSelector.SelectedIndex)
      {
        case 0: // All Fields
          txtLog.AppendText("View All Fields...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            m_FieldListView.Add(fld);
            txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
          }
          break;
        case 1: // Never used
          txtLog.AppendText("View Never used...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            if (fld.Occurances==0)
            {
              m_FieldListView.Add(fld);
              txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
            }
          }
          break;
        case 2: // Seldom used (<10%)
          txtLog.AppendText("View Seldom used (<10%)...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            if (fld.Occurances>0 && fld.Occurances<(m_ContactList_Count/20)) // < 5%
            {
              m_FieldListView.Add(fld);
              txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
            }
          }
          break;
        case 3: // Sometimes used (not never, not always)
          txtLog.AppendText("View Sometimes used (not never, not always)...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            if (fld.Occurances!=0 && fld.Occurances!=m_ContactList_Count)
            {
              m_FieldListView.Add(fld);
              txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
            }
          }
          break;
        case 4: // Always used
          txtLog.AppendText("View Always used...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            if (fld.Occurances==m_ContactList_Count)
            {
              m_FieldListView.Add(fld);
              txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
            }
          }
          break;

        default:
          txtLog.AppendText("View Default (All Fields)...\r\n");
          foreach (MyField fld in m_FieldList)
          {
            m_FieldListView.Add(fld);
            txtLog.AppendText(fld.Ndx + ": " + fld.FieldName + " = " + fld.Occurances + "\r\n");
          }
          break;
      }
      txtLog.AppendText("View Complete\r\n");
      // DataGridView
      dgvFields.DataSource = null;
      dgvFields.Update();
      dgvFields.DataSource = m_FieldListView;
      dgvFields.Update();
      txtNumberOfFields.Text = m_FieldListView.Count.ToString();
      this.Cursor = Cursors.Arrow;
    }

    private void cboFieldViewSelector_SelectedIndexChanged(object sender, EventArgs e)
    {
      btnViewFields_Click(sender, e);
    }
  }
}
