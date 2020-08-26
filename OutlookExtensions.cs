using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KeepAttachmentsOnReply
{
    public static class OutlookExtensions
    {
        // signed / crypted
        private static string PR_SECURITY_FLAGS = @"http://schemas.microsoft.com/mapi/proptag/0x6E010003";
        private static string AttachmentFlags = @"http://schemas.microsoft.com/mapi/proptag/0x37140003";

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

    }
}
