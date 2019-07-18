using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Administration;

namespace Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler
{
    public class LoggingService : SPDiagnosticsServiceBase
    {
        public static string DiagnosticAreaName = "Docx2HtmlSPHandler";
        private static LoggingService _Current;
        public static LoggingService Current
        {
            get
            {
                if (_Current == null)
                {
                    _Current = new LoggingService();
                }

                return _Current;
            }
        }

        private LoggingService()
            : base("Logging Service", SPFarm.Local)
        {

        }

        protected override IEnumerable<SPDiagnosticsArea> ProvideAreas()
        {
            List<SPDiagnosticsArea> areas = new List<SPDiagnosticsArea>
            {
                new SPDiagnosticsArea(DiagnosticAreaName, new List<SPDiagnosticsCategory>
                {
                    new SPDiagnosticsCategory("Docx2HtmlSPHandlerLog", TraceSeverity.Unexpected, EventSeverity.Error)
                })
            };

            return areas;
        }

        public static void LogError(string categoryName, string errorMessage)
        {
            SPDiagnosticsCategory category = LoggingService.Current.Areas[DiagnosticAreaName].Categories[categoryName];
            LoggingService.Current.WriteTrace(0, category, TraceSeverity.Unexpected, errorMessage);
        }
    }
}
