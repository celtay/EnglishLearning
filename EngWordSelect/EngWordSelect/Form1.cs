using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EngWordSelect
{
    public partial class Form1 : Form
    {
        private List<int> usedNumbers = new List<int>();
        private int upperLimit;
        private Random random = new Random();
        private Dictionary<int, WordData> wordDictionary = new Dictionary<int, WordData>();

        public Form1()
        {
            InitializeComponent();
            LoadDataFromExcel();
            //Tayfun Start Code 01.07.205
        }

        private class WordData
        {
            public string Word { get; set; }
            public string Translation { get; set; }
            public string Sentence1 { get; set; }
            public string Sentence2 { get; set; }
            public string Sentence3 { get; set; }
        }

        private void LoadDataFromExcel()
        {
            Excel.Application excel = null;
            Excel.Workbook workbook = null;
            Excel._Worksheet worksheet = null;

            try
            {
                excel = new Excel.Application();
                // Excel dosyanızın tam yolunu buraya yazın
                workbook = excel.Workbooks.Open(@"C:\YourPath\YourFile.xlsx");
                worksheet = workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

                // Upper limit'i Excel'den oku (B3 hücresinden)
                upperLimit = (int)worksheet.Cells[3, 2].Value2;

                for (int i = 4; i <= upperLimit + 3; i++)
                {
                    int number = (int)range.Cells[i, 5].Value2;    // E kolonu
                    string word = range.Cells[i, 6].Value2?.ToString();    // F kolonu
                    string translation = range.Cells[i, 7].Value2?.ToString();    // G kolonu
                    string sentence1 = range.Cells[i, 8].Value2?.ToString();    // H kolonu
                    string sentence2 = range.Cells[i, 9].Value2?.ToString();    // I kolonu
                    string sentence3 = range.Cells[i, 10].Value2?.ToString();   // J kolonu

                    wordDictionary.Add(number, new WordData
                    {
                        Word = word,
                        Translation = translation,
                        Sentence1 = sentence1,
                        Sentence2 = sentence2,
                        Sentence3 = sentence3
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata oluştu: " + ex.Message);
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Close();
                }
                if (excel != null)
                {
                    excel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
                }
            }
        }

    }
}
