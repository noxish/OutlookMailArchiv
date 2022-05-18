using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace ERPlusArchiv
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MyRibbon()
        {
        }

        #region IRibbonExtensibility-Member

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ERPlusArchiv.MyRibbon.xml");
        }

        #endregion

        #region Menübandrückrufe
        //Erstellen Sie hier Rückrufmethoden. Weitere Informationen zum Hinzufügen von Rückrufmethoden finden Sie unter https://go.microsoft.com/fwlink/?LinkID=271226.

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
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

        private List<MailItem> EmailItems = new List<MailItem>();
        private List<Attachment> AttachmentItems = new List<Attachment>();
        public void OnEmailExportButton(Office.IRibbonControl control)
        {
            List<string> FileList = new List<string>();
            EmailItems = new List<MailItem>();
            foreach (MailItem x in new Application().ActiveExplorer().Selection)
            {
                EmailItems.Add(x);
            }

            foreach(var y in EmailItems)
            {
                y.SaveAs(Path.GetTempPath() + CleanFileName(y.Subject) + ".msg", OlSaveAsType.olMSGUnicode);
                FileList.Add(Path.GetTempPath() + CleanFileName(y.Subject) + ".msg");
            }

            SaveFiles(FileList);
        }

        public void OnAttachmentExportButton(Office.IRibbonControl control)
        {
            List<string> FileList = new List<string>();
            AttachmentItems = new List<Attachment>();

            var window = new Application().ActiveWindow();
            var attachsel = window.AttachmentSelection();



            int? index = null;
            if (attachsel.count > 0)
            {
                var attachment = attachsel[1];
                index = attachment.Index;
            }

            var explorer = new Application().ActiveExplorer();
            var selection = explorer.Selection;

            if ((selection.Count > 0) && (index != null) && (attachsel.count == 1))
            {
                object selectedItem = selection[1];
                var mailItem = selectedItem as MailItem;
                foreach (Attachment attach in mailItem.Attachments)
                {
                    if (attach.Index == index)
                    {
                        attach.SaveAsFile(Path.Combine(Path.GetTempPath(), CleanFileName(attach.FileName)));
                        FileList.Add(Path.Combine(Path.GetTempPath() + CleanFileName(attach.FileName)));
                    }
                }

            }

            SaveFiles(FileList);
        }

        public void OnAttachmentExportAllButton(Office.IRibbonControl control)
        {
            List<string> FileList = new List<string>();
            AttachmentItems = new List<Attachment>();

            var window = new Application().ActiveWindow();
            var attachsel = window.AttachmentSelection();



            int? index = null;
            if (attachsel.count > 0)
            {
                var attachment = attachsel[1];
                index = attachment.Index;
            }

            var explorer = new Application().ActiveExplorer();
            var selection = explorer.Selection;

            if ((selection.Count > 0) && (index != null) && (attachsel.count > 1))
            {
                object selectedItem = selection[1];
                var mailItem = selectedItem as MailItem;
                foreach (Attachment attach in mailItem.Attachments)
                {
                    attach.SaveAsFile(Path.Combine(Path.GetTempPath(), CleanFileName(attach.FileName)));
                    FileList.Add(Path.Combine(Path.GetTempPath() + CleanFileName(attach.FileName)));
                }
            }

            SaveFiles(FileList);
        }

        private void SaveFiles(List<string> FileList)
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            startInfo.FileName = "C:\\Program Files\\ERPlus\\bin\\erplus.exe";
            startInfo.Arguments = $"/AddDoc {String.Join(" ", FileList)}";
            process.StartInfo = startInfo;
            process.Start();
        }
        private static string CleanFileName(string fileName)
        {
            return Path.GetInvalidFileNameChars().Aggregate(fileName, (current, c) => current.Replace(c.ToString(), "_"));
        }
    }
}
