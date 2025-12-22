using System;
using System.IO;
using Word;

namespace WpsToPdf
{
    class WpsWord2Pdf : IDisposable
    {
        dynamic wps;

        public WpsWord2Pdf()
        {
            Type type = Type.GetTypeFromProgID("KWps.Application");//KWps.Application KET.Application KWPP.Application
            wps = Activator.CreateInstance(type);
        }
        public void ToPdf(string wpsFilename, string pdfFilename = null)
        {
            if (wpsFilename == null) 
            { 
                throw new ArgumentNullException("wpsFilename"); 
            }

            if (pdfFilename == null)
            {
                pdfFilename = Path.ChangeExtension(wpsFilename, "pdf");
            }

            dynamic doc = wps.Documents.Open(wpsFilename, Visible: false);
            doc.ExportAsFixedFormat(pdfFilename, WdExportFormat.wdExportFormatPDF);
            doc.Close();
        }

        public void Dispose()
        {
            if (wps != null) { wps.Quit(); }
        }
    }
}
