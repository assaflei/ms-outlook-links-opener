using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Office = Microsoft.Office.Core;

namespace LinksOpener
{
    public class LOBrowserDetails
    {
        public string CommandPath { get; set; }
        public string Label { get; set; }
        public int BrowserID { get; set; }
    }

    public class AsyncTaskResult
    {
        public List<LOBrowserDetails> Browsers { get; set; }
        public bool IsMenuVisible { get; set; }
        public string CustomUIXml { get; set; }
    }

    [ComVisible(true)]
    public class LinksOpenerRibbon : Office.IRibbonExtensibility
    {
        private const string IDPrefix = "LOBrowser";
        private const string IDEditPrefix = "LOEditBrowser";
        private Office.IRibbonUI ribbon;
        Task<AsyncTaskResult> searchBrowsersTask;

        TraceSwitch debugLevel = new TraceSwitch("DebugLevel", "The Output level of tracing");

        public LinksOpenerRibbon()
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: LinksOpener Ctor start.", startPoint));
#endif
            searchBrowsersTask = Task<AsyncTaskResult>.Factory.StartNew(() => SearchBrowsersInRegistry());
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: LinksOpener Ctor end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
        }

        AsyncTaskResult SearchBrowsersInRegistry()
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Async Search start ", startPoint));
#endif
            StringBuilder sb = new StringBuilder();
            bool isMenuVisible;
            AsyncTaskResult result = new AsyncTaskResult()
            {
                Browsers = new List<LOBrowserDetails>(),
                IsMenuVisible = true,
                CustomUIXml = string.Empty
            };

            try
            {
                Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Searching registry " + @"SOFTWARE\WOW6432Node\Clients\StartMenuInternet", DateTime.Now));
                RegistryKey browserKeys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Clients\StartMenuInternet");
                if (browserKeys == null)
                {
                    Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Searching registry " + @"SOFTWARE\Clients\StartMenuInternet", DateTime.Now));
                    browserKeys = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Clients\StartMenuInternet");
                }

                string[] browserNames = browserKeys.GetSubKeyNames();
                isMenuVisible = browserNames.Length > 1;

                if (!isMenuVisible)
                {
                    result.IsMenuVisible = false;
                    return result;
                }

                Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Found {1} browsers", DateTime.Now, browserNames.Length));
                for (int i = 0; i < browserNames.Length; i++)
                {
                    LOBrowserDetails browser = new LOBrowserDetails();
                    RegistryKey browserKey = browserKeys.OpenSubKey(browserNames[i]);
                    browser.BrowserID = result.Browsers.Count + 1;
                    browser.Label = (string)browserKey.GetValue(null);

                    // if browser with the same name already exists in the browsers list, skip it
                    if (result.Browsers.Any(b => b.Label.CompareTo(browser.Label) == 0))
                        continue;
                    RegistryKey browserKeyPath = browserKey.OpenSubKey(@"shell\open\command");
                    browser.CommandPath = (string)browserKeyPath.GetValue(null);
                    result.Browsers.Add(browser);
                    sb.Append(string.Format(ConfigurationManager.AppSettings["contextMenuString"], browser.BrowserID, browser.Label));
                }
                Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Created browsers list", DateTime.Now));
            }
            catch (System.Exception ex)
            {
                // log the error always
                Trace.TraceError("{0}: [class]:{1} [method]:{2}{3}[message]:{4}{5}[Stack]:{6}",
                    DateTime.Now, // when was the error happened
                    MethodInfo.GetCurrentMethod().DeclaringType.Name, // the class name
                    MethodInfo.GetCurrentMethod().Name,  // the method name
                    Environment.NewLine,
                    ex.Message, // the error message
                    Environment.NewLine,
                    ex.StackTrace // the stack trace information
                );
                result.Browsers.Clear();
                result.IsMenuVisible = false;
            }

            result.CustomUIXml = sb.ToString();
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Async Search end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
            return result;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetCustomUI {1} start.", startPoint, ribbonID));
#endif
            string customUI = string.Empty;
            // get ribbon xml
            customUI = GetResourceText("LinksOpener.LinksOpenerRibbon.xml");

            // adjust email read context menu xml
            customUI = customUI.Replace(ConfigurationManager.AppSettings["customUIReplaceString"], searchBrowsersTask.Result.CustomUIXml).Replace("Item", IDPrefix);

            // adjust email compose context menu xml
            customUI = customUI.Replace(ConfigurationManager.AppSettings["customUIReplaceStringEdit"], searchBrowsersTask.Result.CustomUIXml).Replace("Item", IDEditPrefix);
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetCustomUI end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
            return customUI;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {

#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Ribbon_Load start.", startPoint));
#endif
            this.ribbon = ribbonUI;
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: Ribbon_Load end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
        }
        
        /// <summary>
        /// Scan registry location where all browsers register their details and if there is only one key in there, 
        ///    the eis no need in displaying the menu since only one browser is used to open links
        /// </summary>
        /// <param name="control">Links opener context menu</param>
        /// <returns>display or not</returns>
        public bool IsLinksOpenerVisible(IRibbonControl control)
        {
            return searchBrowsersTask.Result.IsMenuVisible;
        }

        /// <summary>
        /// Assign a link open button for each installed browser, set visibility
        /// </summary>
        public bool IsLOButtonVisible(IRibbonControl control)
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: IsLOButtonVisible start.", startPoint));
#endif
            int numId = GetButtonID(control.Id);
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: IsLOButtonVisible end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
            return numId <= searchBrowsersTask.Result.Browsers.Count;
        }

        /// <summary>
        /// Assign a link open button for each installed browser, set label 
        /// </summary>
        public string GetLOButtonLabel(IRibbonControl control)
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetLOButtonLabel start.", startPoint));
#endif
            int numId = GetButtonID(control.Id);
            string label = searchBrowsersTask.Result.Browsers.Where(browser => browser.BrowserID == numId).Single().Label;
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetLOButtonLabel end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
            return label;
        }

        private int GetButtonID(string control)
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetButtonId start {1}", startPoint, control));
#endif
            string prefixToUse = (control.IndexOf(IDPrefix) > -1 ? IDPrefix : IDEditPrefix);
            int index = control.IndexOf(prefixToUse);
            string cleanId = control.Remove(index, prefixToUse.Length);

            int numId = -1;
            int.TryParse(cleanId, out numId);
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceVerbose, String.Format("{0}: GetButtonId end - {1} milliseconds.", endPoint, endPoint.Subtract(startPoint).Milliseconds));
#endif
            return numId;
        }

        public void OpenLink(Office.IRibbonControl control)
        {
            Inspector currentInspector = Globals.ThisAddIn.Application.ActiveInspector();
            if (currentInspector == null)
            {
                Explorer currentExplorer = Globals.ThisAddIn.Application.ActiveExplorer();
                Selection currentSelection = currentExplorer.Selection;
                object o = currentSelection[1];
                MailItem oMail = (MailItem)o;
                currentInspector = oMail.GetInspector;
            }
            Microsoft.Office.Interop.Word.Document document = 
                (Microsoft.Office.Interop.Word.Document)currentInspector.WordEditor;
            string address = document.Application.Selection.Hyperlinks[1].Address;
            int numId = GetButtonID(control.Id);
            string path = searchBrowsersTask.Result.Browsers.Where(browser => browser.BrowserID == numId).Single().CommandPath;
            Process.Start(path, address);
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
