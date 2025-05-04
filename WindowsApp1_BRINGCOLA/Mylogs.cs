using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;

namespace WindowsApp1_BRINGCOLA
{
    class Mylogs
    {
        Workbook book = new Workbook();
        
        public void insertLogs(string user, string message)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet2 = book.Worksheets[1];
            int row = sheet2.Rows.Length;
            sheet2.Range[row, 1].Value = user;
            sheet2.Range[row, 2].Value = message;
            sheet2.Range[row, 3].Value = DateTime.Now.ToString("mm/dd/yyyy");
            sheet2.Range[row, 4].Value = DateTime.Now.ToString("hh:mm:ss:tt");

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
        }
        public DataTable LoadLogs()
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet2 = book.Worksheets[1];
            DataTable dataTable = new DataTable();

            // Define columns
            dataTable.Columns.Add("User");
            dataTable.Columns.Add("Message");
            dataTable.Columns.Add("Date");
            dataTable.Columns.Add("Time");

            // Fill DataTable with data from Excel
            int rowCount = sheet2.LastRow;
            for (int row = 2; row <= rowCount; row++) // Assuming first row is header
            {
                DataRow dataRow = dataTable.NewRow();
                dataRow["User"] = sheet2.Range[row, 1].Value;
                dataRow["Message"] = sheet2.Range[row, 2].Value;
                dataRow["Date"] = sheet2.Range[row, 3].Value;
                dataRow["Time"] = sheet2.Range[row, 4].Value;
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        public void UpdateLog(int logId, string newLogValue)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet2 = book.Worksheets[1];

            // Assuming logId corresponds to the row index, adjust as necessary
            for (int row = 2; row <= sheet2.LastRow; row++) // Assuming first row is the header
            {
                if (sheet2.Range[row, 1].Value == logId.ToString()) // Assuming ID is in the first column
                {
                    // Update the log message in the second column (adjust as necessary)
                    sheet2.Range[row, 2].Value = newLogValue; // Assuming the message is in the second column
                    break; // Exit loop after updating the log
                }
            }

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
        }

        public void DeleteLog(int logId)
        {
            book.LoadFromFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx");
            Worksheet sheet2 = book.Worksheets[1];

            // Assuming logId corresponds to the row index, adjust as necessary
            for (int row = 2; row <= sheet2.LastRow; row++) // Assuming first row is the header
            {
                if (sheet2.Range[row, 1].Value == logId.ToString()) // Assuming ID is in the first column
                {
                    // Remove the row using DeleteRow
                    sheet2.DeleteRow(row); // Correct method to delete a row
                    break; // Exit loop after deleting the log
                }
            }

            book.SaveToFile(@"C:\Users\ACT-STUDENT\Downloads\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGAnew\WindowsApp1_BORINAGA\Book1(1).xlsx", ExcelVersion.Version2016);
        }

    }
}
