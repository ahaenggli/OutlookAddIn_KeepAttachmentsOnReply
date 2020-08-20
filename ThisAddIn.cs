using KeepAttachmentsOnReply.Properties;
using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Windows.Forms;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Deployment.Application;
using System.Reflection;
using System.Text.RegularExpressions;

namespace KeepAttachmentsOnReply
{
    public partial class ThisAddIn
    {

        private static string tmpDir = Path.GetTempPath() + "OutlookAddIn_KeepAttachmentsOnReply" + Path.DirectorySeparatorChar.ToString();

        /// <summary>
        /// add menu items to ribbon bars
        /// </summary>
        /// <returns>ribbonbar</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }


        /// <summary>
        /// register addIn in outlook
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Background for better performance
            Thread worker = new Thread(Init);
            worker.IsBackground = true;
            worker.Start();
        }

        /// <summary>
        /// initialize AddIn, called in background from Startup
        /// </summary>  
        private void Init()
        {
            this.Application.Inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ActiveExplorer().InlineResponse += new Outlook.ExplorerEvents_10_InlineResponseEventHandler(parseItem);

            //update vsto/clickonce directory in background
            Thread worker = new Thread(Update);
            worker.IsBackground = true;
            worker.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        /// <summary>
        /// handle for new Inspector
        /// </summary>  
        private void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            if (Inspector != null)
                if (Inspector is Microsoft.Office.Interop.Outlook.Inspector)
                    parseItem(Inspector.CurrentItem);
        }

        /// <summary>
        /// Parse an opend item if it is a mail
        /// </summary>  
        private void parseItem(object item)
        {
            // is it a mail? 
            if (item is Outlook.MailItem)
            {
                // cast to MailItem
                MailItem mailItem = item as MailItem;

                // item not NULL
                if (mailItem != null)
                {
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
                            if (Application.ActiveExplorer().Selection.Count == 1)
                            {
                                object selectedItem = Application.ActiveExplorer().Selection[1];
                                if (selectedItem is Outlook.MailItem)
                                    addParentAttachments(mailItem, selectedItem as Outlook.MailItem);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Copy attachments from original mail to the reply
        /// </summary>  
        public static void addParentAttachments(MailItem newMail, MailItem attFrom)
        {
            // is it a mail? 
            if (attFrom is Outlook.MailItem)
            {
                // cast to MailItem. 
                MailItem mailItem = attFrom as Outlook.MailItem;

                // any attachments?
                if (mailItem.Attachments.Count > 0)
                {
                    // iterate attachments
                    foreach (Attachment attachment in mailItem.Attachments)
                    {
                        // flags to check whether the attachment is embedded (inline) or not
                        var flags = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37140003");

                        // ignore embedded attachment...
                        if (flags != 4)
                        {
                            // rtf mail attachment comes here - and the embeded image is treated as attachment then Type value is 6 and ignore it
                            if ((int)attachment.Type != 6)
                            {
                                // does our temp dir exists? if not -> create
                                if (!Directory.Exists(tmpDir)) Directory.CreateDirectory(tmpDir);
                                // generate a temp path
                                var tmp = tmpDir + attachment.FileName;
                                // save file to tmp
                                attachment.SaveAsFile(tmp);
                                // add tmp file as new attachment to the reply mail
                                newMail.Attachments.Add(tmp, Outlook.OlAttachmentType.olByValue, 1, attachment.DisplayName);
                                // save replay mail
                                // newMail.Save();
                                // delete tmp file
                                File.Delete(tmp);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Check for updates and download them
        /// </summary>
        private void Update()
        {
            try
            {
                if (Settings.Default.LastUpdateCheck == null)
                {
                    Settings.Default.LastUpdateCheck = DateTime.Now;
                    Properties.Settings.Default.Save();
                }

                Version a = (ApplicationDeployment.IsNetworkDeployed) ? ApplicationDeployment.CurrentDeployment.CurrentVersion : Assembly.GetExecutingAssembly().GetName().Version;
                Version b = a;


                // once a day should be enougth....
                if (Settings.Default.LastUpdateCheck.AddHours(1) <= DateTime.Now)
                {
                    System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
                    HttpWebRequest wr = (HttpWebRequest)WebRequest.Create(Settings.Default.VersionUrl);
                    wr.UserAgent = "ahaenggli/OutlookAddIn_KeepAttachmentsOnReply";
                    var x = wr.GetResponse(); ;

                    using (var reader = new System.IO.StreamReader(x.GetResponseStream()))
                    {
                        string json = reader.ReadToEnd();
                        if (json.Contains("tag_name"))
                        {
                            Regex pattern = new Regex("\"tag_name\":\"v\\d+(\\.\\d+){2,}\",");
                            Match m = pattern.Match(json);
                            b = new Version(m.Value.Replace("\"", "").Replace("tag_name:v", "").Replace(",", ""));
                        }
                    }

                    if (b > a)
                    {

                        string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\haenggli.NET\";
                        string AddInData = ProgramData + @"OutlookAddIn_KeepAttachmentsOnReply\";
                        string StartFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.vsto";
                        string localFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.zip";
                        string DownloadUrl = Environment.GetEnvironmentVariable("KeepAttachmentsOnReply_DownloadUrl", EnvironmentVariableTarget.Machine) ?? Settings.Default.UpdateUrl;

                        if (DownloadUrl.Equals("---"))
                        {
                            Settings.Default.LastUpdateCheck = DateTime.Now;
                            Properties.Settings.Default.Save();
                            return;
                        }

                        if (!Directory.Exists(AddInData)) Directory.CreateDirectory(AddInData);
                        foreach (System.IO.FileInfo file in new DirectoryInfo(AddInData).GetFiles()) file.Delete();
                        foreach (System.IO.DirectoryInfo subDirectory in new DirectoryInfo(AddInData).GetDirectories()) subDirectory.Delete(true);

                        if (DownloadUrl.StartsWith("http"))
                        {

                            WebClient webClient = new WebClient();
                            webClient.DownloadFile(DownloadUrl, localFile);
                            webClient.Dispose();
                            DownloadUrl = localFile;
                        }

                        ZipFile.ExtractToDirectory(DownloadUrl, AddInData);
                    }
                    Settings.Default.LastUpdateCheck = DateTime.Now;
                    Properties.Settings.Default.Save();
                }
            }
            catch (System.Exception Ex)
            {
                if (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("KeepAttachmentsOnReply_DownloadUrl", EnvironmentVariableTarget.Machine)))
                    MessageBox.Show(Ex.Message);
            }
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
