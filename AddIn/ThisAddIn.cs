using KeepAttachmentsOnReply.Properties;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Deployment.Application;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace KeepAttachmentsOnReply
{
    public partial class ThisAddIn
    {

        private static readonly log4net.ILog _logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static readonly System.Threading.Timer _timerUpdateCheck = new System.Threading.Timer(OnTimerUpdateCheck, null, 60000, 0);

        private static void OnTimerUpdateCheck(object state)
        {
            _logger.Debug("OnTimerUpdateCheck");

            try
            {
                if (Settings.Default.IsUpgradeRequired)
                {
                    Settings.Default.Upgrade();
                    Settings.Default.Reload();
                    Settings.Default.IsUpgradeRequired = false;
                    Settings.Default.Save();
                }

                // AutoUpdate disabled
                if (!Settings.Default.IsAutoUpdate) return;

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
                    WebResponse x = wr.GetResponse(); ;

                    using (StreamReader reader = new System.IO.StreamReader(x.GetResponseStream()))
                    {
                        string json = reader.ReadToEnd();
                        if (json.Contains("tag_name"))
                        {
                            Regex pattern = new Regex("\"tag_name\":\"v\\d+(\\.\\d+){2,}\",");
                            Match m = pattern.Match(json);
                            b = new Version(m.Value.Replace("\"", "").Replace("tag_name:v", "").Replace("tag_name:", "").Replace(",", ""));
                        }
                    }

                    if (b > a)
                    {
                        string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\haenggli.NET\";
                        string AddInData = ProgramData + @"OutlookAddIn_KeepAttachmentsOnReply\";
                        string StartFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.vsto";
                        string localFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.zip";
                        string DownloadUrl = Settings.Default.UpdateUrl;

                        if (String.IsNullOrWhiteSpace(DownloadUrl))
                        {
                            Settings.Default.LastUpdateCheck = DateTime.Now;
                            Settings.Default.Save();
                            return;
                        }

                        if (!Directory.Exists(AddInData))
                        {
                            Directory.CreateDirectory(AddInData);
                        }

                        foreach (System.IO.FileInfo file in new DirectoryInfo(AddInData).GetFiles())
                        {
                            file.Delete();
                        }

                        foreach (System.IO.DirectoryInfo subDirectory in new DirectoryInfo(AddInData).GetDirectories())
                        {
                            subDirectory.Delete(true);
                        }

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
                _logger.Error(Ex);
            }

        }

        /// <summary>
        /// add menu items to ribbon bars
        /// </summary>
        /// <returns>ribbonbar</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _logger.Debug("CreateRibbonExtensibilityObject");
            return new Ribbon();
        }

        /// <summary>
        /// register addIn in outlook
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _logger.Debug("Startup");
            this.Application.Inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ActiveExplorer().InlineResponse += new Outlook.ExplorerEvents_10_InlineResponseEventHandler(OutlookExtensions.ParseItemForAttachments);
            OnTimerUpdateCheck(null);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            // muss ausgeführt werden, wenn Outlook heruntergefahren wird. 
            // Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
            _logger.Debug("Shutdown");
        }

        /// <summary>
        /// handle for new Inspector
        /// </summary>  
        private void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            if (Inspector != null)
            {
                if (Inspector is Microsoft.Office.Interop.Outlook.Inspector)
                {
                    OutlookExtensions.ParseItemForAttachments(Inspector.CurrentItem);
                }
            }
        }

        #region Von VSTO generierter Code
        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
