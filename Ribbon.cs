using Microsoft.Office.Interop.Outlook;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace KeepAttachmentsOnReply
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        private void getAttachmentsFromConversation(MailItem dest, MailItem search, int lvl)
        {
            if (dest == null)
            {
                return;
            }

            if (countAttachements(dest) > 0)
            {
                return;
            }

            System.Collections.Generic.Stack<MailItem> st = new System.Collections.Generic.Stack<MailItem>();

            // Cast selectedItem to MailItem. 

            // Determine the store of the mail item. 
            Outlook.Folder folder = search.Parent as Outlook.Folder;
            Outlook.Store store = folder.Store;
            if (store.IsConversationEnabled == true)
            {
                // Obtain a Conversation object. 
                Outlook.Conversation conv = search.GetConversation();
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
                if (countAttachements(it) > 0)
                {
                    Debug.WriteLine(countAttachements(it));
                    ThisAddIn.addParentAttachments(dest, it);
                    st.Clear();
                }
            }

        }

        private int countAttachements(MailItem mailItem)
        {
            int cnt = 0;
            if (mailItem != null)
            {
                // any attachments?
                if (mailItem.Attachments.Count > 0)
                {
                    // iterate attachments
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        // flags to check whether the attachment is embedded (inline) or not
                        dynamic flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");

                        // ignore embedded attachment...
                        if (flags != 4)
                        {
                            // rtf mail attachment comes here - and the embeded image is treated as attachment then Type value is 6 and ignore it
                            if ((int)attachment.Type != 6)
                            {
                                cnt++;
                            }
                        }
                    }
                }
            }
            return cnt;
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

        public void OnButton1(Office.IRibbonControl e)
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
                getAttachmentsFromConversation(mailItem, mailItem, 0);
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
