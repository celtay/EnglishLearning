using System;
using System.Collections.Generic;
using System.IO;
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
                string excelPath = Path.Combine(Application.StartupPath, "words.xlsx");

                excel = new Excel.Application();
                workbook = excel.Workbooks.Open(excelPath);
                worksheet = workbook.Sheets[1];

                // B3'teki upper limit değerini oku
                var upperLimitCell = worksheet.Cells[3, 2].Value2?.ToString(); // B sütunu = 2
                MessageBox.Show($"B3'teki değer: {upperLimitCell}");
                upperLimit = int.Parse(upperLimitCell ?? "0");

                // Test amaçlı ilk birkaç hücrenin değerini görelim
                string testMessage = "";
                for (int col = 1; col <= 10; col++)
                {
                    var value = worksheet.Cells[4, col].Value2?.ToString() ?? "boş";
                    testMessage += $"Sütun {col}: {value}\n";
                }
                MessageBox.Show(testMessage);

                // Şimdi verileri doğru sütunlardan okuyalım
                for (int i = 4; i <= upperLimit + 3; i++)
                {
                    try
                    {
                        // Sütun numaralarını Excel'deki gerçek konumlara göre ayarlayın
                        string numberStr = (worksheet.Cells[i, 5].Value2?.ToString() ?? "").Trim();  // E sütunu
                        string word = (worksheet.Cells[i, 6].Value2?.ToString() ?? "").Trim();      // F sütunu
                        string translation = (worksheet.Cells[i, 7].Value2?.ToString() ?? "").Trim();// G sütunu
                        string sentence1 = (worksheet.Cells[i, 8].Value2?.ToString() ?? "").Trim();  // H sütunu
                        string sentence2 = (worksheet.Cells[i, 9].Value2?.ToString() ?? "").Trim();  // I sütunu
                        string sentence3 = (worksheet.Cells[i, 10].Value2?.ToString() ?? "").Trim(); // J sütunu

                        MessageBox.Show($"Satır {i}:\nNumara: {numberStr}\nKelime: {word}\nÇeviri: {translation}");

                        if (!string.IsNullOrEmpty(word))
                        {
                            int number = int.Parse(numberStr);
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
                        MessageBox.Show($"Satır {i} işlenirken hata: {ex.Message}");
                        continue;
                    }
                }

                MessageBox.Show($"Toplam {wordDictionary.Count} kelime yüklendi.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Genel hata: " + ex.Message);
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

        // Next butonu için event handler
        private void btnNext_Click(object sender, EventArgs e)
        {
            if (usedNumbers.Count >= upperLimit)
            {
                MessageBox.Show("Tüm kelimeler kullanıldı! Reset'e basın.");
                return;
            }

            int randomNumber;
            do
            {
                randomNumber = random.Next(1, upperLimit + 1);
            } while (usedNumbers.Contains(randomNumber));

            usedNumbers.Add(randomNumber);
            DisplayWord(randomNumber);
        }
        // Kelimeyi gösterme metodu
        private void DisplayWord(int number)
        {
            if (wordDictionary.TryGetValue(number, out WordData word))
            {
                txtWord.Text = word.Word;
                txtTranslation.Text = word.Translation;
                txtSentence1.Text = word.Sentence1;
                txtSentence2.Text = word.Sentence2;
                txtSentence3.Text = word.Sentence3;
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            usedNumbers.Clear();
            txtWord.Clear();
            txtTranslation.Clear();
            txtSentence1.Clear();
            txtSentence2.Clear();
            txtSentence3.Clear();
        }
    }
}
