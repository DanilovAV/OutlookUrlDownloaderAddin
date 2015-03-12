using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using OutlookUrlDownloaderAddin.Helpers;
using OutlookUrlDownloaderAddin.Properties;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OutlookUrlDownloaderAddin
{
    [ComVisible(true)]
    public class ProcessUrlRibbon : Office.IRibbonExtensibility
    {
        private const string btDownloadOpen = "btDownloadOpen";
        private object locker = new object();

        private Office.IRibbonUI ribbon;     

        public ProcessUrlRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("CloudUriRewriter.ProcessUrlRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnButtonAction(Office.IRibbonControl ribbonControl)
        {
            if (ParsedObject == null)
                return;

            StringBuilder sb = new StringBuilder();

            foreach(var url in ParsedObject.UrlObjects)
                sb.AppendLine(url.Url);

            MessageBox.Show(sb.ToString(), "Demo: urls in email");
        }

        public Bitmap GetImage(Office.IRibbonControl ribbonControl)
        {
            return Resources.cloud;
        }

        public bool GetEnabled(Office.IRibbonControl ribbonControl)
        {
            return isEnabled;
        }

        #endregion

        #region Ribbon methods

        private bool isEnabled;
        public bool IsEnabled
        {
            get
            {
                lock (locker)
                {
                    return isEnabled;
                }
            }
            set
            {
                if (isEnabled != value)
                {
                    lock (locker)
                    {
                        if (isEnabled != value)
                        {
                            isEnabled = value;
                            InternalRefresh();
                        }
                    }
                }
            }
        }

        private MailItem currentItem;
        public MailItem CurrentItem
        {
            get
            {
                lock (locker)
                {
                    return currentItem;
                }
            }
            set
            {
                if (!ReferenceEquals(currentItem, value))
                    lock (locker)
                        if (!ReferenceEquals(currentItem, value))
                            currentItem = value;
            }
        }

        private ParsedObject parsedObject;
        public ParsedObject ParsedObject
        {
            get
            {
                lock (locker)
                {
                    return parsedObject;
                }
            }
            set
            {
                if (!ReferenceEquals(parsedObject, value))
                    lock (locker)
                        if (!ReferenceEquals(parsedObject, value))
                            parsedObject = value;
            }
        }

        private void InternalRefresh()
        {
            if (ribbon != null)
                ribbon.InvalidateControl(btDownloadOpen);
        }

        public void Refresh()
        {
            lock (locker)
                ribbon.InvalidateControl(btDownloadOpen);
        }

        #endregion

        #region Helpers

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
