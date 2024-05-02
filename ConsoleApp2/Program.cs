// To customize application configuration such as set high DPI settings or default font,
// see https://aka.ms/applicationconfiguration.

using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Exception = System.Exception;

internal class Program
{
    private static void Main(string[] args)
    {
        // Create an instance of the Outlook application
        Application outlookApp = new Application();
        
        // Get the MAPI namespace
        NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

        // Get the Inbox folder
        MAPIFolder inbox = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        Items items = inbox.Items;
        items.Sort("[ReceivedTime]", true);
        // Loop through the messages in the Inbox
        for (int i = 1; i < 30; i++)
        {
            MailItem mailItem = items[i] as MailItem;
            if (mailItem != null)
            {

                string taskSubject = Regex.Replace(mailItem.TaskSubject, @"[ $@"" /:\\]+", "_");
                string receiveDate = Regex.Replace(mailItem.ReceivedTime.ToString(), @"[ $@"" /:\\]+", "-");
                string path = $@"C:\Users\tornike.gumberidze\downloads\Outlook_Files\{taskSubject}\{receiveDate}\";
                foreach (Attachment attachment in mailItem.Attachments)
                {
                    //if (Regex.IsMatch(attachment.FileName, @"\.(png|jpg|jpeg|gif|bmp|tif|tiff|ico)$", RegexOptions.IgnoreCase))
                    //{
                    //    continue;
                    //}
                    if(!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }
                    // Save attachments to a specific directory
                    string attachmentPath = Path.Combine(path, attachment.FileName);
                    attachment.SaveAsFile(attachmentPath);
                    Console.WriteLine("Attachment saved: {0}", attachmentPath);
                }
            }

        }
        // Release COM objects
        System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookNamespace);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);

        // Optionally, you may want to explicitly terminate Outlook application process
        // This is necessary because Outlook might not close properly after you use it through Interop
        // Be cautious while using this, as it will terminate all instances of Outlook running on the machine.
        //System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
        //foreach (System.Diagnostics.Process process in processes)
        //{
        //    process.Kill();
        //}
    }
}