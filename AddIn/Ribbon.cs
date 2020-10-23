using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace KeepAttachmentsOnReply
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {

        private static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private void GetAttachmentsFromConversation(MailItem mailItem)
        {
            if (mailItem == null)
            {
                return;
            }

            if (mailItem.Attachments.CountNonEmbeddedAttachments() > 0)
            {
                return;
            }

            System.Collections.Generic.Stack<MailItem> st = new System.Collections.Generic.Stack<MailItem>();

            // Determine the store of the mail item. 
            Outlook.Folder folder = mailItem.Parent as Outlook.Folder;
            Outlook.Store store = folder.Store;
            if (store.IsConversationEnabled == true)
            {
                // Obtain a Conversation object. 
                Outlook.Conversation conv = mailItem.GetConversation();
                // Check for null Conversation. 
                if (conv != null)
                {
                    // Obtain Table that contains rows 
                    // for each item in the conversation. 
                    Outlook.Table table = conv.GetTable();
                    _logger.Debug("Conversation Items Count: " + table.GetRowCount().ToString());
                    _logger.Debug("Conversation Items from Root:");

                    // Obtain root items and enumerate the conversation. 
                    Outlook.SimpleItems simpleItems = conv.GetRootItems();
                    foreach (object item in simpleItems)
                    {
                        // In this example, enumerate only MailItem type. 
                        // Other types such as PostItem or MeetingItem 
                        // can appear in the conversation. 
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mail = item as Outlook.MailItem;
                            Outlook.Folder inFolder = mail.Parent as Outlook.Folder;
                            string msg = mail.Subject + " in folder [" + inFolder.Name + "] EntryId [" + (mail.EntryID.ToString() ?? "NONE") + "]";
                            _logger.Debug(msg);
                            _logger.Debug(mail.Sender);
                            _logger.Debug(mail.ReceivedByEntryID);

                            if (mail.EntryID != null && (mail.Sender != null || mail.ReceivedByEntryID != null))
                            {
                                st.Push(mail);
                            }
                        }
                        // Call EnumerateConversation 
                        // to access child nodes of root items. 
                        EnumerateConversation(st, item, conv);
                    }
                }
            }

            while (st.Count > 0)
            {
                MailItem it = st.Pop();

                if (it.Attachments.CountNonEmbeddedAttachments() > 0)
                {
                    //_logger.Debug(it.Attachments.CountNonEmbeddedAttachments());

                    try
                    {
                        if (mailItem.IsMailItemSignedOrEncrypted())
                            if (MessageBox.Show(null, "Es handelt sich um eine signierte Nachricht. Soll diese für die Anhänge ohne Zertifikat dupliziert werden?", "Nachricht duplizieren?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                mailItem.Close(OlInspectorClose.olDiscard);
                                mailItem = mailItem.Copy();
                                mailItem.Unsign();
                                mailItem.Save();
                                mailItem.Close(OlInspectorClose.olDiscard);
                                mailItem.Save();
                            }
                            else
                            {
                                st.Clear();
                                break;
                            }

                        mailItem.CopyAttachmentsFrom(it);
                        mailItem.Save();
                    }
                    catch (System.Exception ex)
                    {
                        //mailItem.Close(OlInspectorClose.olDiscard);
                        MessageBox.Show(ex.Message);
                    }

                    st.Clear();
                }
            }
            st.Clear();

            Marshal.ReleaseComObject(mailItem);
        }

        private void EnumerateConversation(System.Collections.Generic.Stack<MailItem> st, object item, Outlook.Conversation conversation)
        {
            Outlook.SimpleItems items =
            conversation.GetChildren(item);
            if (items.Count > 0)
            {
                foreach (object myItem in items)
                {
                    // In this example, enumerate only MailItem type. 
                    // Other types such as PostItem or MeetingItem 
                    // can appear in the conversation. 
                    if (myItem is Outlook.MailItem)
                    {
                        Outlook.MailItem mailItem = myItem as Outlook.MailItem;
                        Outlook.Folder inFolder = mailItem.Parent as Outlook.Folder;

                        string msg = mailItem.Subject + " in folder [" + inFolder.Name + "] EntryId [" + (mailItem.EntryID.ToString() ?? "NONE") + "]";
                        _logger.Debug(msg);
                        _logger.Debug(mailItem.Sender);
                        _logger.Debug(mailItem.ReceivedByEntryID);

                        if (mailItem.EntryID != null && (mailItem.Sender != null || mailItem.ReceivedByEntryID != null))
                        {
                            st.Push(mailItem);
                        }
                    }
                    // Continue recursion. 
                    EnumerateConversation(st, myItem, conversation);
                }
            }
        }

        public Ribbon()
        {
        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("KeepAttachmentsOnReply.Ribbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. 
        //Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
        }

        public void OnInfo(Office.IRibbonControl e)
        {
            Form frm = new AboutForm();
            frm.ShowDialog();
            frm.Close();
            frm.Dispose();
        }

        public void OnFindMissingAttachments(Office.IRibbonControl e)
        {
            List<MailItem> mailItems = new List<MailItem>();

            // single email opend 
            if (e.Context is Inspector m)
            {
                mailItems.Add(m.CurrentItem as MailItem);
            }
            else
            {
                foreach (object selectedItem in Globals.ThisAddIn.Application.ActiveExplorer().Selection)
                {
                    if (selectedItem is Outlook.MailItem)
                    {
                        mailItems.Add(selectedItem as Outlook.MailItem);
                    }
                }
            }


            if (mailItems.Count > 0)
            {
                foreach (MailItem mailItem in mailItems)
                {
                    if (mailItem.EntryID == null)
                    {
                        MessageBox.Show("noch nicht implementiert...");
                        break;
                    }

                    // if (mailitem.Attachments.Count > 0) return;
                    GetAttachmentsFromConversation(mailItem);

                    // Abbruch-Bedingung
                    if (mailItems.Count > 1 && mailItems.First().Equals(mailItem))
                    {
                        if (MessageBox.Show(null, "Es wurden mehrere Nachrichten ausgewählt. " + Environment.NewLine + "Sollen die Anhänge für alle " +
                                                    mailItems.Count.ToString() +
                                                    " Nachrichten gesucht werden?", "Alle suchen?", MessageBoxButtons.YesNo) == DialogResult.No)
                        {
                            break;
                        }
                    }
                }
            }

            mailItems.Clear();
        }

        #endregion

        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
