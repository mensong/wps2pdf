using System;
using System.IO;

namespace WpsToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // 显示Logo
            Version();

            // 如果不带参数，输出帮助信息
            if (args.Length == 0)
            {
                Help();
                Environment.Exit(9);
                return;
            }

            // 判断第1个参数是否-v或-h，如果是，输出相应的信息
            switch (args[0].ToLower().Substring(0, 2))
            {
                case "-v":
                    Environment.Exit(0);
                    return;
                case "-h":
                    Help();
                    Environment.Exit(0);
                    return;
            }

            // 解析文件名
            string wpsFilename = null;
            string pdfFilename = null;
            try
            {
                wpsFilename = Path.GetFullPath(args[0]);
                if (args.Length > 1) { pdfFilename = Path.GetFullPath(args[1]); }
            }
            catch (Exception ex)
            {
                Console.WriteLine("参数中包含不正确的文件名");
                Environment.Exit(2);
                return;
            }

            // 判断输入文件是否存在
            if (!File.Exists(wpsFilename))
            {
                Console.WriteLine("错误：指定文件不存在");
                Environment.Exit(1);
                return;
            }

            // 转换
            int exitCode = 0;
            try
            {
                string ext = Path.GetExtension(wpsFilename);
                if (ext.StartsWith(".doc", StringComparison.OrdinalIgnoreCase))
                {
                    using (WpsWord2Pdf word = new WpsWord2Pdf())
                    {
                        word.ToPdf(wpsFilename, pdfFilename);
                    }
                }
                else if (ext.StartsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    using (WpsExcel2Pdf excel = new WpsExcel2Pdf())
                    {
                        excel.ToPdf(wpsFilename, pdfFilename);
                    }
                }
                else if (ext.StartsWith(".ppt", StringComparison.OrdinalIgnoreCase))
                {
                    using (WpsPpt2Pdf ppt = new WpsPpt2Pdf())
                    {
                        ppt.ToPdf(wpsFilename, pdfFilename);
                    }
                }
                else
                {
                    exitCode = -2;
                }
            }
            catch (Exception ex)
            {
                exitCode = -1;
            }
            
            Environment.Exit(exitCode);
        }

        static void Version()
        {
            Console.WriteLine(
@"wps2pdf - 将WPS文档(含DOC/DOCX/XLS/XLSX/PPT/PPTX)转换为PDF
Copyright (c) 2025 Mensong
版本：1.0
");
        }

        static void Help()
        {
            Console.WriteLine(
@"命令：wps2pdf WPS文件 [PDF文件]
      将指定的WPS文件转换为PDF文件，若未指定PDF文件，
      生成的PDF文件与WPS文件同名，且扩展名改为PDF。
");
        }
    }
}
