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
    public partial class frmEditTarget : Form
    {
        public Outlook.Rules rules;
        public List<DataGridViewRow> selected_emails = null;
        public List<string[]> ValidFolders = null;

        public frmEditTarget()
        {
            InitializeComponent();
        }
        private void frmEditTarget_Load(object sender, EventArgs e)
        {
            initList();
            LoadFolders();
        }

        #region Functions & Procedures
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

       // Returns Folder object based on folder path
        private Outlook.Folder fnGetFolder(string folderPath)
        {
            Outlook.Folder folder;
            string backslash = @"\";
            try
            {
                if (folderPath.StartsWith(@"\\"))
                    folderPath = folderPath.Remove(0, 2);

                String[] folders = folderPath.Split(backslash.ToCharArray());
                folder = Globals.ThisAddIn.Application.Session.Folders[folders[0]] as Outlook.Folder;
                if (folder != null)
                {
                    for (int i = 1; i <= folders.GetUpperBound(0); i++)
                    {
                        Outlook.Folders subFolders = folder.Folders;
                        folder = subFolders[folders[i]] as Outlook.Folder;
                        if (folder == null) return null;
                    }
                }
                return folder;
            }
            catch { return null; }
        }

        private void initList()
        {
            lblCount.Text = selected_emails.Count.ToString();
        }
        private void LoadFolders()
        {
            try
            {
                Outlook._NameSpace outNS;
                Outlook.Application application = Globals.ThisAddIn.Application;
                //Get the MAPI namespace
                outNS = application.GetNamespace("MAPI");
                //Get UserName
                string profileName = outNS.CurrentUser.Name;

                Outlook.Folders folders = outNS.Folders;

                ValidFolders = new List<string[]>();
                cmbTarget.Items.Clear();

                if (folders.Count > 0)
                {
                    IterateFolder(folders);
                    //foreach (Outlook.Folder sub_fldr in folders)
                    //{
                    //    if (sub_fldr.Name.ToLower() == "deleted items") continue;

                    //    if (sub_fldr.Name.ToLower().StartsWith(Properties.Settings.Default.WatchFolder_Prefix.ToLower()))
                    //    {
                    //        ValidFolders.Add(new string[] { sub_fldr.Name, sub_fldr.FolderPath});
                    //        cmbTarget.Items.Add(sub_fldr.Name);
                    //    }

                    //    if (sub_fldr.Folders.Count > 0)
                    //    {
                    //        IterateFolder(sub_fldr.Folders, "");
                    //    }
                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + ex.StackTrace, "Error Loading Drag & Drop AddIn");
            }
        }

        private void IterateFolder(Outlook.Folders parent_folder)
        {
            foreach(Outlook.Folder sub_fldr in parent_folder)
            {
                if (sub_fldr.Name.ToLower() == "deleted items") continue;
                //if (path_ != "") path_ += "\\";
                // path_ += sub_fldr.Name;

                if (sub_fldr.Name.ToLower().StartsWith(Properties.Settings.Default.WatchFolder_Prefix.ToLower()))
                {
                    ValidFolders.Add(new string[] { sub_fldr.Name, sub_fldr.FolderPath });
                    cmbTarget.Items.Add(sub_fldr.Name);
                }

                if (sub_fldr.Folders.Count > 0)
                    IterateFolder(sub_fldr.Folders);
            }
        }
        #endregion

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            string folder_path = "" ;
            string tar_rulename = "";
            string src_rulename = "";

            string sender_address = "";
            bool eadd_exist = false;
            bool has_changed = false;

            if (cmbTarget.SelectedIndex > -1)
            {
                folder_path = ValidFolders[cmbTarget.SelectedIndex][1];
                
                if (MessageBox.Show("Are you to change the target folder to " + cmbTarget.Text + "?", 
                    "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    try
                    {
                        Outlook.Folder tar_folder = fnGetFolder(folder_path);
                        tar_rulename = Properties.Settings.Default.RuleName_Prefix + tar_folder.Name;
                        Outlook.Rule tar_rule = fnGetRule(Properties.Settings.Default.RuleName_Prefix + tar_folder.Name);
                        Outlook.Rule src_rule = null;

                        //CREATE RULE 
                        if (tar_rule == null)
                        {
                            //CREATE NEW RULE
                            tar_rule = rules.Create(tar_rulename, Outlook.OlRuleType.olRuleReceive);
                            //SET TARGET FOLDER
                            tar_rule.Actions.MoveToFolder.Folder = (tar_folder);
                            tar_rule.Actions.MoveToFolder.Enabled = true;
                        }

                        //CHECK EACH SENDER_ADDRESS
                        foreach (DataGridViewRow row in selected_emails)
                        {
                            sender_address = row.Cells[1].Value.ToString().Trim();
                            src_rulename = row.Cells[4].Value.ToString();
                            eadd_exist = false;

                            if (sender_address != "" &&
                                    row.Cells[3].Value.ToString().ToLower() !=
                                        ValidFolders[cmbTarget.SelectedIndex][1].ToLower())
                            {
                                //DELETE THE EMAIL FROM THE PREVIOUS RULE
                                src_rule = fnGetRule(src_rulename);
                                if (src_rule != null)
                                {
                                    foreach (Outlook.Recipient rc in src_rule.Conditions.From.Recipients)
                                    {
                                        if (rc.Address.ToLower() == sender_address.ToLower())
                                        {
                                            rc.Delete();
                                            rc.Resolve();
                                            has_changed = true;
                                            break;
                                        }
                                    }
                                    if (src_rule.Conditions.From.Recipients.Count == 0)
                                        rules.Remove(src_rulename);
                                }


                                //ADD THE EMAIL TO THE NEW RULE
                                if (tar_rule.Conditions.From.Recipients.Count > 0)
                                {
                                    foreach (Outlook.Recipient rec in tar_rule.Conditions.From.Recipients)
                                    {
                                        eadd_exist = (rec.Address.ToLower() == row.Cells[1].Value.ToString().ToLower());
                                        if (eadd_exist) break;
                                    }
                                }

                                //ADD FROM EMAIL
                                if (!eadd_exist)
                                {
                                    tar_rule.Conditions.From.Recipients.Add(sender_address);
                                    tar_rule.Conditions.From.Recipients.ResolveAll();
                                    tar_rule.Conditions.From.Enabled = true;
                                    has_changed = true;
                                }
                            }
                        }
                        if (has_changed && rules != null)
                        {
                            rules.Save(true);
                        }

                        this.DialogResult = System.Windows.Forms.DialogResult.OK;
                        this.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message + ex.StackTrace);
                    }                    
                }                
            }
        }
             
    }
}
