using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.IO;
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

        private Office.IRibbonUI ribbon;

        private void getAttachmentsFromConversation(MailItem mailItem)
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
                    Debug.WriteLine("Conversation Items Count: " + table.GetRowCount().ToString());

                    Debug.WriteLine("Conversation Items from Root:");
                    // Obtain root items and enumerate the conversation. 
                    Outlook.SimpleItems simpleItems = conv.GetRootItems();
                    foreach (object item in simpleItems)
                    {
                        // In this example, enumerate only MailItem type. 
                        // Other types such as PostItem or MeetingItem 
                        // can appear in the conversation. 
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem mail = item
                            as Outlook.MailItem;
                            Outlook.Folder inFolder = mail.Parent as Outlook.Folder;
                            string msg = mail.Subject + " in folder [" + inFolder.Name + "] EntryId [" + (mail.EntryID.ToString() ?? "NONE") + "]";
                            Debug.WriteLine(msg);
                            Debug.WriteLine(mail.Sender);
                            Debug.WriteLine(mail.ReceivedByEntryID);

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
                    //Debug.WriteLine(it.Attachments.CountNonEmbeddedAttachments());

                    try
                    {
                        if (mailItem.IsMailItemSignedOrEncrypted())
                            if (MessageBox.Show(null, "Es handelt sich um eine signierte Nachricht. Soll diese für die Anhänge ohne Zertifikat dupliziert werden?", "Nachricht duplizieren?", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                mailItem = mailItem.Copy();
                                mailItem.Unsign();
                            }
                            else
                            {
                                st.Clear();
                                break;
                            }

                        ThisAddIn.addParentAttachments(mailItem, it);
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
            mailItem = null;
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
                        Debug.WriteLine(msg);
                        Debug.WriteLine(mailItem.Sender);
                        Debug.WriteLine(mailItem.ReceivedByEntryID);

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

        public void OnFindMissingAttachments(Office.IRibbonControl e)
        {
            MailItem mailItem = null;
            Inspector m = e.Context as Inspector;

            // single email opend 
            if (m != null)
            {
                mailItem = m.CurrentItem as MailItem;
            }
            else
            {
                object selectedItem = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
                if (selectedItem is Outlook.MailItem)
                {
                    mailItem = selectedItem as Outlook.MailItem;
                }
            }


            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    MessageBox.Show("noch nicht implementiert...");
                    return;
                }

                //if (mailitem.Attachments.Count > 0) return;
                getAttachmentsFromConversation(mailItem);
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
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
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
