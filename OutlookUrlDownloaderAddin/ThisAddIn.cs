using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using OutlookUrlDownloaderAddin.Helpers;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookUrlDownloaderAddin
{
    public partial class ThisAddIn
    {
        private IUriProcessor uriProcessor;
        private Explorer currentExplorer = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            currentExplorer = Application.ActiveExplorer();
            currentExplorer.SelectionChange += OnSelectionChange;

            uriProcessor = new UriProcessor();
        }

        private void OnSelectionChange()
        {
            MAPIFolder selectedFolder = this.Application.ActiveExplorer().CurrentFolder;

            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    var activeExplorer = this.Application.ActiveExplorer();

                    MailItem mailItem = activeExplorer.Selection[1] as MailItem;

                    if (mailItem != null && processUrlRibbon != null)
                    {
                        processUrlRibbon.CurrentItem = mailItem;
                        processUrlRibbon.IsEnabled = false;

                        var task = uriProcessor.ParseAsync(mailItem);

                        task.ContinueWith(x =>
                        {
                            if (!object.ReferenceEquals(processUrlRibbon.CurrentItem, mailItem) || (x.Result == null))
                                return;

                            processUrlRibbon.IsEnabled = x.Result.HasUrls;
                            processUrlRibbon.ParsedObject = x.Result;

                        }, TaskContinuationOptions.ExecuteSynchronously);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (currentExplorer != null)
                currentExplorer.SelectionChange -= OnSelectionChange;
        }

        private ProcessUrlRibbon processUrlRibbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            processUrlRibbon = new ProcessUrlRibbon();
            return processUrlRibbon;
        }

        #region Infrastructure


        #endregion


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
