using CsvHelper;
using System;
using System.Collections.Generic;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelData
{
    class QnA
    {
        public string Question { get; set; }
        public string AnswerDescription { get; set; }    
        public string State { get; set; }
        public string Answer { get; set; }
    }
    class Program
    {
        const string PSV_PATH = @"C:\Dev\Github\RMC\Data.csv";
        const string EXCEL_FILE_PATH = @"C:\Dev\Github\RMC\Data-New.xlsx";
        excel.Application app = null;
        excel.Workbook workbook = null;

        static void Main(string[] args)
        {
            var prog = new Program();
            prog.ProcessFiles();
        }

        private void WriteDataToText(IEnumerable<QnA> qas)
        {
            using (var textWriter = File.CreateText(PSV_PATH))
            {
                var csv = new CsvWriter(textWriter);
                csv.Configuration.Delimiter = "|";
                csv.WriteRecords(qas);
                textWriter.Close();
            }
        }

        private void ProcessFiles()
        {
            var qnas = new List<QnA>();

            try
            {
                app = new excel.Application();
                workbook = app.Workbooks.Open(EXCEL_FILE_PATH);
                excel.Range previous = null;                
                
                // Reading the first worksheet at this time
                var sheet = workbook.Sheets[1];

                var columns = sheet.UsedRange.Cells.Columns.Count + 1;
                var rows = sheet.UsedRange.Cells.Rows.Count + 1;
                var max = columns * rows;

                qnas = new List<QnA>(max);

                var qna = new QnA();
                int qCnt = 1;
                var states = new string[columns-3];
                
                // First row has the state names
                for (int column = 1; column <= columns-1; column++)
                {                    
                    if(!string.IsNullOrWhiteSpace(sheet.Cells[1, column].Text.Trim()))
                    {
                        states[column-3] = sheet.Cells[1, column].Text.Trim();
                    }                    
                }

                //Iterating from row 2 because first row contains states  
                for (int row = 2; row < sheet.UsedRange.Cells.Rows.Count; row++)
                {
                    previous = sheet.Cells[row, 1];

                    if (Convert.ToString(previous.Text).Trim().StartsWith("Q:"))
                    {
                        for (int cnt = 0; cnt < states.Length; cnt++)
                        {
                            var state = states[cnt];

                            qna = new QnA
                            {
                                Question = Convert.ToString(sheet.Cells[row, 2].Text).Trim(),
                                AnswerDescription = Convert.ToString(sheet.Cells[row + 1, 2].Text).Trim(),
                                State = state,
                                Answer = Convert.ToString(sheet.Cells[row + 1, 3 + cnt].Text).Trim(),

                            };
                            qnas.Add(qna);
                        }
                        Console.WriteLine("Done reading {0} answers for question: {1}", states.Length, qCnt);
                        qCnt++;
                    }
                    else
                    {
                        continue;
                    }
                }

                Console.WriteLine("We have {0} questions & answers", qnas.Count);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Release the Excel objects     
                workbook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                app.Workbooks.Close();
                app.Quit();
                app = null;
                workbook = null;

                GC.GetTotalMemory(false);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.GetTotalMemory(true);

                //Write data to CSV files
                WriteDataToText(qnas);
            }
        }
    }
}
