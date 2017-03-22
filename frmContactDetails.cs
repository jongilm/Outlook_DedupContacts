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
  public partial class frmContactDetails : Form
  {
    public MainForm m_MainForm = null;
    public frmContactDetails()
    {
      InitializeComponent();
    }

    private void frmContactDetails_Load(object sender, EventArgs e)
    {
      //m_MainForm = sender;
      myContactBindingSource.DataSource = m_MainForm.m_ContactList;
      //myContactBindingSource = OutlookContactsAnalyser.MyContact;

    }

    private void myContactBindingNavigatorSaveItem_Click(object sender, EventArgs e)
    {
      //myContactBindingSource.EndEdit();
      String s = myContactBindingSource.Current.ToString();
      int n = myContactBindingSource.List.Count;
    }
  }
}
