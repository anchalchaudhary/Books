using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Books.Models
{
    public class BooksModel
    {
        public int BookID { get; set; }
        public string barcode { get; set; }
        public string itemcallnumber { get; set; }
        public string author { get; set; }
        public string title { get; set; }
        public string publishercode { get; set; }
    }
}