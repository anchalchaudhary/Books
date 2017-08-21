using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Books.Models;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Books.Controllers
{
    public class UploadFileController : Controller
    {

        DBBooksEntities db = new DBBooksEntities();
        // GET: UploadFile
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase excelBooks, HttpPostedFileBase excelBarcode)
        {
            if (excelBooks.FileName.EndsWith("xls") || excelBooks.FileName.EndsWith("xlsx") && excelBarcode.FileName.EndsWith("xls") || excelBarcode.FileName.EndsWith("xlsx"))
            {
                string pathBooks = Server.MapPath("~/UploadedExcel/" + excelBooks.FileName);
                string pathBarcode = Server.MapPath("~/UploadedExcel/" + excelBarcode.FileName);

                if (System.IO.File.Exists(pathBooks))
                {
                    System.IO.File.Delete(pathBooks);
                }
                excelBooks.SaveAs(pathBooks);

                if (System.IO.File.Exists(pathBarcode))
                {
                    System.IO.File.Delete(pathBarcode);
                }
                excelBarcode.SaveAs(pathBarcode);
                Excel.Application application = new Excel.Application();
                Excel.Workbook workbookBooks = application.Workbooks.Open(pathBooks);
                Excel.Worksheet worksheetBooks = workbookBooks.ActiveSheet;
                Excel.Range rangeBooks = worksheetBooks.UsedRange;

                Excel.Workbook workbookBarcode = application.Workbooks.Open(pathBarcode);
                Excel.Worksheet worksheetBooksBarcode = workbookBarcode.ActiveSheet;
                Excel.Range rangeBarcode = worksheetBooksBarcode.UsedRange;

                List<tblBook> datatblBooks = new List<tblBook>();
                List<tblBarcode> datatblBarcode = new List<tblBarcode>();

                for (int row = 1; row <= rangeBooks.Rows.Count; row++)
                {
                    if (true)
                    {
                        datatblBooks.Add(new tblBook
                        {
                            barcode = ((Excel.Range)rangeBooks.Cells[row, 1]).Text,
                            itemcallnumber = ((Excel.Range)rangeBooks.Cells[row, 2]).Text,
                            author = ((Excel.Range)rangeBooks.Cells[row, 3]).Text,
                            title = ((Excel.Range)rangeBooks.Cells[row, 4]).Text,
                            publishercode = ((Excel.Range)rangeBooks.Cells[row, 5]).Text,
                        });
                    }
                }

                for (int row = 1; row <= rangeBarcode.Rows.Count; row++)
                {
                    if (true)
                    {
                        datatblBarcode.Add(new tblBarcode
                        {
                            Barcode = ((Excel.Range)rangeBarcode.Cells[row, 1]).Text
                        });
                    }
                }
                //using (DBBooksEntities db = new DBBooksEntities())
                //{
                //    foreach (var item1 in datatblBooks)
                //    {
                //        tblBook objtblBook = new tblBook();
                //        objtblBook.barcode = item1.barcode;
                //        objtblBook.itemcallnumber = item1.itemcallnumber;
                //        objtblBook.author = item1.author;
                //        objtblBook.title = item1.title;
                //        objtblBook.publishercode = item1.publishercode;

                //        db.tblBooks.Add(objtblBook);
                //        db.SaveChanges();
                //    }
                //    foreach (var item2 in datatblBarcode)
                //    {
                //        tblBarcode objtblBarcode = new tblBarcode();
                //        objtblBarcode.Barcode = item2.Barcode;

                //        db.tblBarcodes.Add(objtblBarcode);
                //        db.SaveChanges();
                //    }
                //}

                foreach (var item2 in datatblBarcode)
                {
                    foreach (var item1 in datatblBooks)
                    {
                        if (item1.barcode != item2.Barcode)
                        {
                            tblBook objtblBook = new tblBook();
                            objtblBook.barcode = item1.barcode;
                            objtblBook.title = item1.title;
                            objtblBook.itemcallnumber = item1.itemcallnumber;
                            objtblBook.publishercode = item1.publishercode;


                        }
                        else
                        {
                            break;
                        }
                    }
                }
                return View("Index");
            }
            return View("Index");
        }
    }
}