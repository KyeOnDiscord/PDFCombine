using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PDFCombine
{
    class Program
    {
        static void Main()
        {
            //Dlls inside the exe cool method https://stackoverflow.com/a/62929101/12897035
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;


            string[] pdfs = Directory.GetFiles(AppDomain.CurrentDomain.BaseDirectory, "*pdf");

            if (pdfs.Length == 0)
            {
                Console.WriteLine("Please put the pdfs you want to combine the same folder as this");
                Console.ReadKey();
                return;
            }

            Array.Sort(pdfs, new WindowsFileNames());
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("====================");
            Console.WriteLine("PDF Combiner by Kye | github.com/kyeondiscord");
            Console.WriteLine("====================");
            Console.WriteLine("");
            Console.WriteLine($"Found {pdfs.Length} PDF's to combine");

            Console.WriteLine("The pdf's will be written in this order:");
            foreach (string pdf in pdfs)
            {
                Console.WriteLine(Path.GetFileName(pdf));
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
            Console.Clear();
            Console.WriteLine("Merging pdfs....");
            string exportfilepath = $"._{Path.GetFileName(pdfs[0])}_Combined.pdf";
            MergeMultiplePDFIntoSinglePDF(exportfilepath, pdfs);
            Console.Clear();
            Console.WriteLine($"Exported PDF to {exportfilepath}");
            Console.ReadKey();
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            AssemblyName assemblyName = new AssemblyName(args.Name);

            var path = assemblyName.Name + ".dll";
            if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false) path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null) return null;

                var assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }


        //https://stackoverflow.com/a/17640332/12897035
        private static void MergeMultiplePDFIntoSinglePDF(string outputFilePath, string[] pdfFiles)
        {
            PdfDocument outputPDFDocument = new PdfDocument();
            foreach (string pdfFile in pdfFiles)
            {
                PdfDocument inputPDFDocument = PdfReader.Open(pdfFile, PdfDocumentOpenMode.Import);
                outputPDFDocument.Version = inputPDFDocument.Version;
                foreach (PdfPage page in inputPDFDocument.Pages)
                {
                    outputPDFDocument.AddPage(page);
                }
            }
            outputPDFDocument.Save(outputFilePath);
        }

        public class WindowsFileNames : IComparer<string>
        {

            [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
            static extern int StrCmpLogicalW(String x, String y);

            public int Compare(string x, string y)
            {
                return StrCmpLogicalW(x, y);
            }
        }
    }
}
