using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Office.CommandBar newToolBar;
        private Office.CommandBarButton newToolBarButton;
        private System.Speech.Synthesis.SpeechSynthesizer syn = new System.Speech.Synthesis.SpeechSynthesizer();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Add new Toolbar, if not exist
            if (newToolBar == null)
            {
                newToolBar = Application.ActiveExplorer().CommandBars.Add("MyToolBar", Office.MsoBarPosition.msoBarTop, false, true);
                // Add Button to Toolbar
                newToolBarButton = (Office.CommandBarButton)newToolBar.Controls.Add(Office.MsoControlType.msoControlButton, missing, missing, missing, true);
                newToolBarButton.Caption = "Stop Voice";
                newToolBarButton.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(newToolBarButton_Click);
                newToolBar.Visible = true;
            }
            Application.NewMailEx += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailExEventHandler(Application_NewMailEx);
        }
        // New Mail Event Handler
        void Application_NewMailEx(string EntryIDCollection)
        {
            try
            {
                var mailItem = ((Outlook.MailItem)this.Application.Session.GetItemFromID(EntryIDCollection, missing));
                if (mailItem.UnRead)
                {
                    var body = mailItem.Body;
                    var subject = mailItem.Subject;
                    var sender = mailItem.SenderName;
                    if (subject.Contains("Hello"))
                        syn.SpeakAsync("Mail From " + sender + "Subject is " + subject + body);
                }
            }
            catch
            {
            }
        }

        void newToolBarButton_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            syn.SpeakAsyncCancelAll();
        }

      

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
