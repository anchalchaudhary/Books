using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Books.Models;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

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
            bool isPresent = false;
            if (excelBooks.FileName.EndsWith("xls") || excelBooks.FileName.EndsWith("xlsx") && excelBarcode.FileName.EndsWith("xls") || excelBarcode.FileName.EndsWith("xlsx"))
            {
                //path for Books excel
                string pathBooks = Server.MapPath("~/UploadedExcel/" + excelBooks.FileName);

                //path for Barcode excel
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

                //Books
                Excel.Workbook workbookBooks = application.Workbooks.Open(pathBooks);
                Excel.Worksheet worksheetBooks = workbookBooks.ActiveSheet;
                Excel.Range rangeBooks = worksheetBooks.UsedRange;

                //Barcode
                Excel.Workbook workbookBarcode = application.Workbooks.Open(pathBarcode);
                Excel.Worksheet worksheetBooksBarcode = workbookBarcode.ActiveSheet;
                Excel.Range rangeBarcode = worksheetBooksBarcode.UsedRange;

                //List for books
                List<tblBook> datatblBooks = new List<tblBook>();

                //List for barcodes
                List<tblBarcode> datatblBarcode = new List<tblBarcode>();

                //importing books excel data into list for books
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

                //importing barcode excel data into list for barcode
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

                

                //list for books with no barcode in datatblBarcode 
                List<tblBook> tblBooksList = new List<tblBook>();
                
                List<tblBook> newList = new List<tblBook>();
                newList.Add(new tblBook
                {
                    barcode = "",
                    itemcallnumber = "",
                    author = "",
                    title = "",
                    publishercode = "",
                });

                //check if every book has it's barcode in list of barcodes
                foreach (var itemBook in datatblBooks)
                {
                    isPresent = false;
                    foreach (var itemBarcode in datatblBarcode)
                    {
                        if (itemBook.barcode == itemBarcode.Barcode)
                        {
                            isPresent = true;
                            break;
                        }
                    }
                    if (isPresent == false)
                    {
                        if (itemBook.barcode != "barcode")
                        {
                            tblBook objtblBook = new tblBook();
                            objtblBook.barcode = itemBook.barcode;
                            objtblBook.itemcallnumber = itemBook.itemcallnumber;
                            objtblBook.author = itemBook.author;
                            objtblBook.title = itemBook.title;
                            objtblBook.publishercode = itemBook.publishercode;

                            tblBooksList.Add(objtblBook);
                        }
                    }
                }

                var gridView = new GridView();
                gridView.DataSource = tblBooksList;
                gridView.DataBind();
                Response.ClearContent();
                Response.Buffer = true;
                Response.AddHeader("content-disposition", "attachment; filename=ListOfBooks.xls");
                Response.ContentType = "application/ms-excel";
                Response.Charset = "";
                StringWriter objStringWriter = new StringWriter();
                HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
                gridView.RenderControl(objHtmlTextWriter);
                Response.Output.Write(objStringWriter.ToString());
                Response.Flush();
                Response.End();
                return View("Index");
            }
            return View("Index");
        }
    }
}