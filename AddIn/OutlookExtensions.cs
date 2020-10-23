using Microsoft.Office.Interop.Outlook;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace KeepAttachmentsOnReply
{
    public static class OutlookExtensions
    {
        private static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // signed / crypted
        private static readonly string PR_SECURITY_FLAGS = @"http://schemas.microsoft.com/mapi/proptag/0x6E010003";
        private static readonly string AttachmentFlags = @"http://schemas.microsoft.com/mapi/proptag/0x37140003";

        public static void Unsign(this MailItem mailItem)
        {
            mailItem.PropertyAccessor.SetProperty(PR_SECURITY_FLAGS, 0);
        }
        public static bool IsMailItemSignedOrEncrypted(this MailItem dest)
        {
            dynamic str = dest.PropertyAccessor.GetProperty(PR_SECURITY_FLAGS);
            return (str % 32 > 0);
        }
        public static bool IsEmbedded(this Attachment attachment)
        {
            // flags to check whether the attachment is embedded (inline) or not
            dynamic flags = attachment.PropertyAccessor.GetProperty(AttachmentFlags);

            // ignore embedded attachment...
            // https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxcmsg/af8700bc-9d2a-47e4-b107-5ebf4467a418
            /*
             * attInvisibleInHtml = 0x00000001
             * attInvisibleInRtf  = 0x00000002
             * attRenderedInBody  = 0x00000004
             * 
             * none AND combinations are possible! 0,1,2,3,4,5,6,7
             */
            if (flags < 4)
            {
                // rtf mail attachment comes here - and the embeded image is treated as attachment then Type value is 6 and ignore it
                if ((int)attachment.Type != 6)
                {
                    return false;
                }
            }

            return true;
        }
        public static int CountNonEmbeddedAttachments(this Attachments attachments)
        {
            int cnt = 0;
            if (attachments != null)
            {
                // any attachments?
                if (attachments.Count > 0)
                {
                    // iterate attachments
                    foreach (Attachment attachment in attachments)
                    {
                        if (!attachment.IsEmbedded()) cnt++;
                    }
                }
            }
            return cnt;
        }

        /// <summary>
        /// Parse an opend item if it is a mail
        /// </summary>  
        public static void ParseItemForAttachments(object item)
        {
            if (item == null) return;

            // is it a mail? 
            if (!(item is Outlook.MailItem)) return;

            // cast to MailItem
            if (!(item is MailItem mailItem)) return;

            Application app = Globals.ThisAddIn.Application;

            // it is a new one
            if (mailItem.EntryID == null && mailItem.Sent == false)
            {
                // is it a reply?
                bool isReply = !string.IsNullOrEmpty(mailItem.To) || mailItem.Recipients.Count > 0;
                // it is!
                if (isReply)
                {
                    // keep attachments from original mail
                    // get selected item (reply is open, so this should be the original mail) 
                    MailItem src = null;
                    if (app.ActiveWindow() is Outlook.Explorer)
                    {
                        //Debug.WriteLine("Explorer");
                        if (app.ActiveExplorer().Selection.Count == 1)
                        {
                            object selectedItem = app.ActiveExplorer().Selection[1];
                            if (selectedItem is Outlook.MailItem)
                            {
                                src = selectedItem as Outlook.MailItem;
                            }
                        }
                    }
                    else
                    if (app.ActiveWindow() is Outlook.Inspector)
                    {
                        //Debug.WriteLine("Inspector");
                        object selectedItem = app.ActiveInspector().CurrentItem;
                        if (selectedItem is Outlook.MailItem)
                        {
                            src = selectedItem as Outlook.MailItem;
                        }
                    }

                    if (src != null)
                        mailItem.CopyAttachmentsFrom(src);

                }
            }
        }

        /// <summary>
        /// Copy attachments from original mail to the reply
        /// </summary>  
        public static void CopyAttachmentsFrom(this MailItem newMail, MailItem attFrom)
        {
            // is it a mail? 
            if (attFrom is Outlook.MailItem)
            {
                // cast to MailItem. 
                MailItem mailItem = attFrom;

                // any attachments?
                if (mailItem.Attachments.CountNonEmbeddedAttachments() > 0)
                {
                    // iterate attachments
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        if (!attachment.IsEmbedded())
                        {
                            // does our temp dir exists? if not -> create
                            string tmpDir = Path.GetTempPath() + "OutlookAddIn_KeepAttachmentsOnReply" + Path.DirectorySeparatorChar.ToString();
                            if (!Directory.Exists(tmpDir)) Directory.CreateDirectory(tmpDir);

                            // generate a temp path
                            string tmp = tmpDir + attachment.FileName;

                            // save file to tmp
                            attachment.SaveAsFile(tmp);

                            // add tmp file as new attachment to the reply mail
                            newMail.Attachments.Add(tmp, Outlook.OlAttachmentType.olByValue, 1, attachment.DisplayName);

                            // save replay mail
                            //newMail.Save();

                            // delete tmp file
                            File.Delete(tmp);
                        }
                    }
                }
            }
        }

    }
}
