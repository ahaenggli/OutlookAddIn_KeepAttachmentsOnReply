using OutlookAddIn_KeepAttachmentsOnReply_Installer.Properties;
using System;
using System.IO;
using System.IO.Compression;
using System.Net;

namespace Setup
{
    internal class Program
    {
        private static readonly string ProgramData = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"\haenggli.NET\";
        private static readonly string AddInData = ProgramData + @"OutlookAddIn_KeepAttachmentsOnReply\";
        private static readonly string StartFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.vsto";
        private static readonly string localFile = AddInData + @"OutlookAddIn_KeepAttachmentsOnReply.zip";

        private static void Main(string[] args)
        {
            string DownloadUrl = Environment.GetEnvironmentVariable("KeepAttachmentsOnReply_DownloadUrl", EnvironmentVariableTarget.Machine) ?? Settings.Default.UpdateUrl;

            bool runVSTO = true;

            if (args.Length > 0)
            {
                DownloadUrl = args[0];

                foreach (string a in args)
                {
                    if (a.ToLower().Equals("/silent"))
                    {
                        runVSTO = false;
                    }
                }
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
            //System.Diagnostics.Process.Start("VSTOInstaller.exe /silent /uninstall \""+ StartFile + "\"");
            if (runVSTO)
            {
                System.Diagnostics.Process.Start(StartFile);
            }
        }
    }
}
