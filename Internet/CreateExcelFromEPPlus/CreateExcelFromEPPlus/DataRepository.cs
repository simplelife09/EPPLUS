using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateExcelFromEPPlus
{
    public class DataRepository
    {
        private IList<Book> _bookList = new List<Book> {
            new Book { Id = 1, Name = "Dong Chu Liet Quoc" },
            new Book { Id = 2, Name = "Su ky Tu Ma Thien" },
            new Book { Id = 3, Name = "Viet Nam su luoc" }           
        };

        public List<Book> GetBookList()
        {
            return _bookList.ToList();
        }

        public DataTable GetBookDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("Name");
            foreach (var book in _bookList)
            {
                var dr = dt.NewRow();
                dr["ID"] = book.Id;
                dr["Name"] = book.Name.ToString();
                dt.Rows.Add(dr);
            }
            return dt;

        }
    }
}
