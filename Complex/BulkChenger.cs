using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace Complex
{
    internal class BulkChenger
    {

        public async Task WordChanger(List<string> wordList, List<string[]> chnges)
        {
            await Task.Delay(1000);
            foreach (var word in wordList) 
            {
                Word.Application wordApplication = new Word.Application();
                Word.Document wordDocument = wordApplication.Documents.Open(word);

                // Заменяем строку один на строку два
                wordDocument.Content.Find.Execute("Один", ReplaceWith: "Два");

                // Сохраняем и закрываем документ
                wordDocument.Save();
                wordDocument.Close();
                wordApplication.Quit();
            }
        }

        public async Task ExcelChanger(List<string> excelList, List<string[]> chnges)
        {
            await Task.Delay(1000);
            foreach(var excel in excelList)
            {
                Excel.Application excelApplication = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApplication.Workbooks.Open(excel);
                Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;

                // Заменяем строку один на строку два
                Excel.Range range = excelWorksheet.UsedRange;
                object[,] values = range.Value;
                for (int i = 1; i <= values.GetLength(0); i++)
                {
                    for (int j = 1; j <= values.GetLength(1); j++)
                    {
                        if (values[i, j].ToString() == "Один")
                        {
                            values[i, j] = "Два";
                        }
                    }
                }
                range.Value = values;

                // Сохраняем и закрываем документ
                excelWorkbook.Save();
                excelWorkbook.Close();
                excelApplication.Quit();
            }

        }

    }
}
