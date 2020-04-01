using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;


namespace HeaderMarkup.Markup
{
    class MarkBookHolder
    {
        public Dictionary<string, MarkBook> markBooks = new Dictionary<string, MarkBook>();

        public MarkBook GetMarkBook(Excel.Workbook workbook)
        {
            if (!markBooks.ContainsKey(workbook.Name))
                markBooks.Add(workbook.Name, new MarkBook());
            return markBooks[workbook.Name];
        }

        public MarkBook GetMarkBook() => GetMarkBook(Utils.GetActiveWorkbook());

        public MarkSheet GetMarkSheet()
        {
            Excel.Workbook workbook = Utils.GetActiveWorkbook();
            if (!markBooks.ContainsKey(workbook.Name))
                markBooks.Add(workbook.Name, new MarkBook());
            MarkBook book = markBooks[workbook.Name];
            Excel.Worksheet worksheet = Utils.GetActiveWorksheet(workbook);
            if (!book.markSheets.ContainsKey(worksheet.Name))
                book.markSheets.Add(worksheet.Name, new MarkSheet());
            return book.markSheets[worksheet.Name];
        }

        public void Remove(Excel.Workbook workbook) => markBooks.Remove(workbook.Name);

        public void Remove(string bookName) => markBooks.Remove(bookName);
    }
}
