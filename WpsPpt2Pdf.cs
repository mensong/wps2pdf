using System;
using System.IO;
using PowerPoint;

namespace WpsToPdf
{
    class WpsPpt2Pdf : IDisposable
    {
        dynamic wps;

        public WpsPpt2Pdf()
        {
            Type type = Type.GetTypeFromProgID("KWPP.Application");//KWps.Application KET.Application KWPP.Application
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

            //wps.Visible = MsoTriState.msoFalse;
            wps.DisplayAlerts = MsoTriState.msoFalse;

            //dynamic presentation = wps.Presentations;
            dynamic doc = wps.Presentations.Open(wpsFilename, ReadOnly: MsoTriState.msoTrue);

            doc.SaveAs(pdfFilename, 32);

            //doc.ExportAsFixedFormat2(
            //    Path: pdfFilename,
            //    FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF
            //// 其他可选参数（参数名可能与Office不完全一致，请参考WPS对象模型）：
            //// PpFixedFormatIntent.ppFixedFormatIntentScreen, // 意图：屏幕显示
            //// MsoTriState.msoFalse, // 不包含文档属性
            //// PpPrintRangeType.ppPrintAll, // 打印全部幻灯片
            //// PrintRange: null, // 打印范围
            //// FrameSlides: MsoTriState.msoFalse, // 是否为每张幻灯片加框
            //// HandoutOrder: PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
            //// OutputType: PpPrintOutputType.ppPrintOutputSlides
            //);

            doc.Close();
        }

        public void Dispose()
        {
            if (wps != null) { wps.Quit(); }
        }
    }
}
