using Excel;
using System;
using System.IO;

namespace WpsToPdf
{
    class WpsExcel2Pdf : IDisposable
    {
        dynamic wps = null;

        public WpsExcel2Pdf()
        {
            Type type = Type.GetTypeFromProgID("KET.Application");//KWps.Application KET.Application KWPP.Application
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

            dynamic doc = wps.Workbooks.Open(wpsFilename, ReadOnly: false);
            doc.ExportAsFixedFormat(
                Type: XlFixedFormatType.xlTypePDF, // 导出为PDF
                Filename: pdfFilename
                // 其他可选参数，如指定导出范围、忽略打印区域等
                // IgnorePrintAreas: false,
                // From: 1,
                // To: 5
            );
            doc.Close();
        }

        public void Dispose()
        {
            if (wps != null) { wps.Quit(); }
        }
    }
}
