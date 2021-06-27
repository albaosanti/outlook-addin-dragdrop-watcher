using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Text.RegularExpressions;

using Outlook = Microsoft.Office.Interop.Outlook;



namespace DragDrapWatcher_AddIn
{
    class clsSendNotif
    {
        private const string local_log_path = "C:\\FarCap_Outlook_AddIn\\Error.log";
        private const string Subject = "FarCap Outlook Add-In";
                       
        public bool SendNotification(string str_message)
        {
            bool ok_sent = false;

            Outlook.MailItem mail = null;
            Outlook.Recipients mailRecipients = null;
            Outlook.Recipient mailRecipient = null;

            List<string> recipients = Split_Recipients(Properties.Settings.Default.Recipient);
            string ex_msg = "";
            try
            {
                if (recipients.Count > 0)
                {
                    mail =  (Outlook.MailItem) Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
                    mail.Subject = Subject + " Error";
                    mail.Body = "An exception has occured in the code of add-in.";
                    mail.Body += "\n\n" + str_message;
                    mailRecipients = mail.Recipients;
                    
                    foreach (string eadd in recipients)
                    {
                        mailRecipient = mailRecipients.Add(eadd);
                        mailRecipient.Resolve();
                    }
                    if (mailRecipient.Resolved)
                    {
                        ((Outlook._MailItem)mail).Send();
                        ok_sent=true;
                    }
                    else
                    {
                        ex_msg = "Unable to send the error notification";
                    }
                }
                else
                {
                    ex_msg = "No recipient.";
                }            
            }
            catch (Exception ex)
            {
                ex_msg = ex.Message + ex.StackTrace;
            }
            finally
            {
                if (mailRecipient!=null)
                    Marshal.ReleaseComObject(mailRecipient);
                if (mailRecipients!=null)
                    Marshal.ReleaseComObject(mailRecipients);
                if (mail!=null)
                    Marshal.ReleaseComObject(mail);
            }

            if (!ok_sent) 
                WriteLog(ex_msg, str_message);

            return ok_sent;

        }

        public bool SendTestNotification(string str_message,string str_recipients)
        {
            bool ok_sent = false;

            Outlook.MailItem mail = null;
            Outlook.Recipients mailRecipients = null;
            Outlook.Recipient mailRecipient = null;

            List<string> recipients = Split_Recipients(str_recipients);
            string ex_msg = "";
            try
            {
                if (recipients.Count > 0)
                {
                    mail = (Outlook.MailItem)Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem);
                    mail.Subject = Subject + " Error";
                    mail.Body = "An exception has occured in the code of add-in.";
                    mail.Body += "\n\n" + str_message;
                    mailRecipients = mail.Recipients;
                    foreach (string eadd in recipients)
                    {
                        mailRecipient = mailRecipients.Add(eadd);
                        mailRecipient.Resolve();
                    }
                    if (mailRecipient.Resolved)
                    {
                        ((Outlook._MailItem)mail).Send();
                        ok_sent = true;
                    }
                    else
                    {
                        ex_msg = "Unable to send the error notification";
                    }
                }
                else
                {
                    ex_msg = "No recipient.";
                }
            }
            catch (Exception ex)
            {
                ex_msg = ex.Message + ex.StackTrace;
            }
            finally
            {
                if (mailRecipient != null)
                    Marshal.ReleaseComObject(mailRecipient);
                if (mailRecipients != null)
                    Marshal.ReleaseComObject(mailRecipients);
                if (mail != null)
                    Marshal.ReleaseComObject(mail);
            }

            if (!ok_sent)
                WriteLog(ex_msg, str_message);

            return ok_sent;
        }
    
        private void WriteLog(string ex_msg,string str_message)
        {

            StreamWriter writer = null;
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(local_log_path)))
                    Directory.CreateDirectory(Path.GetDirectoryName(local_log_path));

                writer = new StreamWriter(local_log_path, true);
                if(!string.IsNullOrWhiteSpace(ex_msg))
                {
                    writer.WriteLine("Unsend Notification Error: " +  ex_msg);
                }
                writer.WriteLine("Timestamp: " + DateTime.Now.ToString());
                writer.WriteLine("Error: " + str_message);
                writer.WriteLine();
                writer.Close();
                writer.Dispose();
                writer = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to write log!\nException: " + ex.Message +
                    "\n\nMessage:" + str_message,"FarCap Outlook Add-in");
            }
            finally
            {
                if (writer != null)
                {
                    writer.Close();
                    writer.Dispose();
                }
            }
        }

        private bool IsValidEAdd(string email_add)
        {
            Regex regex = new Regex(@"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$",
             RegexOptions.CultureInvariant | RegexOptions.Singleline);

            if (string.IsNullOrWhiteSpace(email_add))
                return false;

            return  regex.IsMatch(email_add.Trim());
        }

        private List<string> Split_Recipients(string str_recipients)
        {
            string[] sp = str_recipients.Split(new char[] { ';' });
            List<string> rec = new List<string>();

            foreach (string s1 in sp)
            {
                if (IsValidEAdd(s1))
                {
                    rec.Add(s1.Trim());
                }
            }
            return rec;
        }
    }
}
