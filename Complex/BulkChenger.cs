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

        public async Task WordChanger(IEnumerable<string> wordList, List<string?[]> changes)
        {
            await Task.Delay(1000);
            foreach (var word in wordList) 
            {
                Word.Application wordApplication = new Word.Application();
                Word.Document wordDocument = wordApplication.Documents.Open(word);
                foreach (var chn in changes)
                {
                    // Заменяем строку один на строку два
                    wordDocument.Content.Find.Execute(chn[0], ReplaceWith: chn[1]);
                }                
                // Сохраняем и закрываем документ
                wordDocument.Save();
                wordDocument.Close();
                wordApplication.Quit();
                
            }
            MessageBox.Show("Word changes completed");
        }

        public async Task ExcelChanger(IEnumerable<string> excelList, List<string?[]> changes)
        {
            await Task.Delay(1000);
            foreach(var excel in excelList)
            {
                Excel.Application excelApplication = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApplication.Workbooks.Open(excel);
                Excel.Worksheet excelWorksheet = excelWorkbook.ActiveSheet;
                try
                {
                    Excel.Range range = excelWorksheet.UsedRange;
                    object[,] values = range.Value;
                    
                    // Заменяем строку один на строку два
                    foreach (var chn in changes)
                    {
                        for (int i = 1; i <= values.GetLength(0); i++)
                        {
                            for (int j = 1; j <= values.GetLength(1); j++)
                            {
                                if ((values[i, j] ?? "").ToString().Contains(chn[0]))
                                {
                                    values[i, j] = values[i, j].ToString().Replace(chn[0], chn[1]);
                                }
                            }
                        }
                    }

                    range.Value = values;

                    // Сохраняем и закрываем документ
                    excelWorkbook.Save();
                    excelWorkbook.Close();
                    excelApplication.Quit();
                }

                catch
                {
                    excelWorkbook.Save();
                    excelWorkbook.Close();
                    excelApplication.Quit();
                }
                
            }
            MessageBox.Show("Excel changes completed");
        }

    }
}
