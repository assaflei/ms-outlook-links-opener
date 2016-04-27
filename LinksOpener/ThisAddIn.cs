using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Threading;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace LinksOpener
{
    public partial class ThisAddIn
    {
        // Read Tracelevel from app.config and remember the value
        TraceSwitch debugLevel = new TraceSwitch("DebugLevel", "The Output level of tracing");

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Trace.WriteLineIf(debugLevel.TraceInfo, String.Format("{0}: Application startup.", DateTime.Now));
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Trace.WriteLineIf(debugLevel.TraceInfo, String.Format("{0}: Application shutdown.", DateTime.Now));
            Trace.Close();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
#endif
            InitializeLog();
            // Since this event is triggered before addin startup, set the localization here
            Outlook.Application app = this.GetHostItem<Outlook.Application>(typeof(Outlook.Application), "Application");
            int lcid = app.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDUI);
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(lcid);

#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceInfo, String.Format("CreateRibbonExtensibilityObject end - {0} milliseconds.", endPoint.Subtract(startPoint).Milliseconds));
#endif
            return new LinksOpenerRibbon();
        }

        private void InitializeLog()
        {
#if Debug
            DateTime startPoint = DateTime.Now;
            DateTime endPoint;
#endif           // define the datastore for application specific files
            string applicationPath =
                Path.Combine(System.Environment.GetEnvironmentVariable("LOCALAPPDATA"),
                "LinksOpener");

            // Create Directory if it doesn't exists
            Directory.CreateDirectory(applicationPath);
            Directory.CreateDirectory(Path.Combine(applicationPath, "logs"));

            // define the logging destination
            string logFile = Path.Combine(applicationPath, "logs\\tracelog.txt");
            // Configure TraceListener
            Trace.AutoFlush = true;
            Trace.Listeners.Add(new TextWriterTraceListener(logFile, "LinksOpenerListener"));
#if Debug
            endPoint = DateTime.Now;
            Trace.WriteLineIf(debugLevel.TraceInfo, String.Format("InitializeLog end - {0} milliseconds.", endPoint.Subtract(startPoint).Milliseconds));
#endif
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
