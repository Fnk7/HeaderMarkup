using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;


namespace HeaderMarkup.Markup
{
    class BookHolder
    {
        public Dictionary<string, Book> books = new Dictionary<string, Book>();

        public Book GetBook(Excel.Workbook workbook)
        {
            if (!books.ContainsKey(workbook.Name))
                books.Add(workbook.Name, new Book());
            return books[workbook.Name];
        }

        public Book GetBook() => GetBook(Utils.GetActiveWorkbook());

        public Sheet GetSheet()
        {
            Excel.Workbook workbook = Utils.GetActiveWorkbook();
            if (!books.ContainsKey(workbook.Name))
                books.Add(workbook.Name, new Book());
            Book book = books[workbook.Name];
            Excel.Worksheet worksheet = Utils.GetActiveWorksheet(workbook);
            if (!book.sheets.ContainsKey(worksheet.Name))
                book.sheets.Add(worksheet.Name, new Sheet());
            return book.sheets[worksheet.Name];
        }

        public void Remove(Excel.Workbook workbook) => books.Remove(workbook.Name);

        public void Remove(string bookName) => books.Remove(bookName);
    }
}
