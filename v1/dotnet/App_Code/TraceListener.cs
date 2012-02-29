using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using System.Xml;

namespace EWSMAPI
{
    class TraceListener : ITraceListener
    {
        private String _traceDir = "";

        public TraceListener(String traceDirectory)
        {
            _traceDir = traceDirectory;
        }

        #region ITraceListener Members

        public void Trace(string traceType, string traceMessage)
        {
            CreateXMLTextFile(traceType, traceMessage.ToString());
        }

        #endregion

        private void CreateXMLTextFile(string fileName, string traceContent)
        {
            // Create a new XML file for the trace information.
            try
            {
                // If the trace data is valid XML, create an XmlDocument object and save.
                System.IO.File.AppendAllText(_traceDir + DateTime.Now.ToString("ddMMyyyyhhmmss") + DateTime.Now.Ticks + "-" + fileName + ".txt", traceContent);
            }
            catch
            {
                // If the trace data is not valid XML, save it as a text document.
                System.IO.File.WriteAllText(fileName + ".txt", traceContent);
            }
        }
    }
}