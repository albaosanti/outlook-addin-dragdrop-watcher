using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DragDrapWatcher_AddIn
{
    public partial class frmManager : Form
    {
        public Outlook.Rules rules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();
        public List<myWatchEmail> watch_list = new List<myWatchEmail>();
      
        #region Classes
        public class myWatchEmail
        {
            public string destination_folder;
            public string email_add;
            public string email_name;
            public string rule_name;
            public string folder_name;


            public myWatchEmail() { }
            public myWatchEmail(string _dest, string _eadd, string _ename, string _rule, string _foldername)
            {
                this.destination_folder = _dest;
                this.email_add = _eadd;
                this.rule_name = _rule;
                this.email_name = _ename;
                this.folder_name = _foldername;
            }
        }

        private void UpdateWatchList()
        {
            string rule_prefix = Properties.Settings.Default.RuleName_Prefix.ToLower().Trim();
            watch_list = new List<myWatchEmail>();

            lblStatus.Text= "Loading Rules.. Please wait.";
            this.Refresh();

            //Outlook.Rules rules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();
            //rules["xaddin_"].Conditions.rec
          try
          {
              foreach (Outlook.Rule rule in rules)
              {
                  if (rule.Name.ToLower().StartsWith(rule_prefix))
                  {
                      {
                      if (rule.RuleType == Outlook.OlRuleType.olRuleReceive)
                          foreach (Outlook.Recipient rp in rule.Conditions.From.Recipients )
                          {
                              myWatchEmail tmp_ew = new myWatchEmail();
                              tmp_ew.destination_folder = rule.Actions.MoveToFolder.Folder.FolderPath;
                              tmp_ew.folder_name = rule.Actions.MoveToFolder.Folder.Name;
                              tmp_ew.email_add = rp.Address;
                              tmp_ew.email_name = rp.Name;
                              tmp_ew.rule_name = rule.Name;
                              watch_list.Add(tmp_ew);
                          }
                      }
                  }
              }
          }
          catch (Exception ex)
          {
              MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In");
          }
           
        }

        public bool DeleteWatchItem(string email_add, string rule_name)
        {
            bool rem = false;
            for (int i=0;i < watch_list.Count;i++)
            {
                if (watch_list[i].rule_name.ToLower() == rule_name.ToLower() &&
                    watch_list[i].email_add.ToLower() == email_add.ToLower())
                {
                    watch_list.RemoveAt(i);
                    rem = true;
                    break;
                }
            }
            return rem;
        }

        public Outlook.Rule fnGetRule(string ruleName)
        {
            Outlook.Rule rule = null;
            if (rules != null)
            {
                foreach (Outlook.Rule itm in rules)
                {
                    if (itm.Name.ToLower() == ruleName.ToLower())
                    {
                        rule = itm;
                        break;
                    }
                }
                             
            }
            return rule;
        }
#endregion

        public frmManager()
        {
            InitializeComponent();
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string key_word = textBox1.Text.Trim().ToLower();
            bool match = false;

            dgvList.Rows.Clear();
            lblStatus.Text = "Searching... Please wait.";
            this.Refresh();
            
            if(!string.IsNullOrWhiteSpace(key_word)){
                try
                {
                    foreach (myWatchEmail em in watch_list)
                    {
                        match = (checkedListBox1.GetItemChecked(0) &&
                                em.email_name.ToLower().Contains(key_word));

                        if (!match)
                        {
                            match = (checkedListBox1.GetItemChecked(1) &&
                              em.email_add.ToLower().Contains(key_word));
                        }
                        if (!match)
                        {
                            match = (checkedListBox1.GetItemChecked(2) &&
                              em.folder_name.ToLower().Contains(key_word));
                        }
                       

                        if (match)
                        {
                             dgvList.Rows.Add(new object[]{ em.email_name ,
                                em.email_add,
                                em.folder_name, 
                                em.destination_folder,
                                em.rule_name});
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message + ex.StackTrace, "FarCap Outlook Add-In");
                }
            }
            else
            {
                foreach (myWatchEmail em in watch_list)
                {
                    dgvList.Rows.Add(new object[]{ em.email_name ,
                                em.email_add,
                                em.folder_name, 
                                em.destination_folder,
                                em.rule_name});
                }
            }
            lblStatus.Text = "[" + dgvList.RowCount + "] account match found.";
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            UpdateWatchList();
            dgvList.Rows.Clear();
            foreach (myWatchEmail em in watch_list)
            {
                dgvList.Rows.Add(new object[]{ em.email_name ,
                        em.email_add,
                        em.folder_name, 
                        em.destination_folder,
                        em.rule_name});
            }
            lblStatus.Text = "[" + watch_list.Count + "] email account/s on watch list.";
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return) btnSearch.PerformClick();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
                checkedListBox1.SetItemChecked(i, true);

        }

        private void frmManager_Load(object sender, EventArgs e)
        {
            linkLabel1_LinkClicked(sender, null);
            txtRuleName.Text = Properties.Settings.Default.RuleName_Prefix;
            txtFolder.Text = Properties.Settings.Default.WatchFolder_Prefix;
            txtRecipient.Text = Properties.Settings.Default.Recipient;
            btnRefresh.PerformClick();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvList.SelectedRows.Count > 0)
            {
                if (MessageBox.Show("Are you sure to DELETE the selected account [" + dgvList.SelectedRows.Count + "] on watch list?", "Confirm Delete - FarCap Outlook Add-In", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        int remove_count = 0;

                        while (dgvList.SelectedRows.Count > 0)
                        {
                            DataGridViewRow itm = dgvList.SelectedRows[0];
                            Outlook.Rule rule = fnGetRule(itm.Cells[4].Value.ToString());//PASS RULE_NAME
                            bool changed = false;

                            if (rule != null)
                            {
                                foreach (Outlook.Recipient rc in rule.Conditions.From.Recipients)
                                {
                                    if (rc.Address.ToLower() == itm.Cells[1].Value.ToString().ToLower())
                                    {
                                        rc.Delete();
                                        rc.Resolve();
                                        changed = true;
                                        remove_count += 1;
                                    }
                                }

                                if (changed)
                                {
                                    if (rule.Conditions.From.Recipients.Count == 0) 
                                        rules.Remove(itm.Cells[4].Value.ToString());
                                    else
                                        rule.Conditions.From.Recipients.ResolveAll();
                                    
                                    DeleteWatchItem(itm.Cells[1].Value.ToString(), itm.Cells[4].Value.ToString());
                                    dgvList.Rows.Remove(itm);
                                }
                            }
                        }
                        if (remove_count > 0)
                        {
                            rules.Save(true);
                            MessageBox.Show("Deleted Email/s [" + remove_count + "] !", "FarCap Outlook Add-In");
                        }
                    }
                    catch ( Exception ex){
                        MessageBox.Show(ex.Message + ex.StackTrace, "Error @ Delete Email - FarCap Outlook Add-In");
                    }
                    
                    lblStatus.Text = "[" + dgvList.RowCount + "] email account/s on watch list.";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string rule_prefix = txtRuleName.Text.Trim();
            string folder_prefix = txtFolder.Text.Trim();
            string err_recipients = txtRecipient.Text.Trim();

            if (rule_prefix != "" && folder_prefix != "" &&  err_recipients!= "")
            {
                if (MessageBox.Show("Confirm to UPDATE the configuration?", "Confirm Update - FarCap Outlook Add-In", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    Properties.Settings.Default.WatchFolder_Prefix = folder_prefix;
                    Properties.Settings.Default.RuleName_Prefix = rule_prefix;
                    Properties.Settings.Default.Recipient = err_recipients;

                    Properties.Settings.Default.Save();
                    btnRefresh.PerformClick();
                }
            }
            else
                MessageBox.Show("All fields require!", "FarCap Outlook Add-In");

        }

        private void btnAdd_Click(object sender, EventArgs e)
        {

        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (dgvList.SelectedRows.Count > 0)
            {
                frmEditTarget f_edit = new frmEditTarget();
                f_edit.selected_emails = new List<DataGridViewRow>();
                f_edit.rules = this.rules;

                foreach (DataGridViewRow itm in dgvList.SelectedRows)
                    f_edit.selected_emails.Add(itm);

                if (f_edit.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
                    btnRefresh.PerformClick();
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            clsSendNotif err_notif = new clsSendNotif();
            if (err_notif.SendTestNotification("This is a test message.", txtRecipient.Text))
                MessageBox.Show("Sent!");
            else
                MessageBox.Show("Failed to send!");

        }
    }
}
