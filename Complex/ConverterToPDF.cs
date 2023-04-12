using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Tasks = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace Complex
{
    internal class ConverterToPDF
    {

        public async Tasks Converter(IEnumerable<string> files)
        {
            await Tasks.Delay(100);
            Word.Application word = new Word.Application();

            foreach (string file in files)
            {
                Word.Document doc = word.Documents.Open(file);

                string pdfName = Path.ChangeExtension(file, ".pdf");
                doc.ExportAsFixedFormat(pdfName, Word.WdExportFormat.wdExportFormatPDF);
                doc.Close(false);
            }

            word.Quit();
            MessageBox.Show("Conversion procedure ended");
        }

    }
}
