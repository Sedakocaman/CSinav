using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WebApplication4.Controllers
{
    public class HomeController : Controller
    {

        public ActionResult Index()
        {

            return View();
        }

        public class Kisi //nesne oluşturdum.
        {
                public string ad { get; set; }
                public string soyad { get; set; }
                public string adres { get; set; }
                public string email { get; set; }

         }


    public ActionResult Excel(string isim, string soyisim,  string adres, string email)
        {
            //kisi nesnesinden instance oluşturdum.
            var kisi = new Kisi();
            kisi.ad = isim;
            kisi.soyad = soyisim;
            kisi.adres = adres;
            kisi.email = email;
            var kisiler = new List<Kisi>(); //dizi, liste elemanı oluşturdum.

            if (Session["kisiler"] == null) //sessionı kontrol ettim, yoksa oluşturdum, varsa session (51) aldım, tekrar session a gönderdim.
            {
                
                kisiler.Add(kisi);
                Session["kisiler"] = kisiler;
                
            }
            else
            {
                kisiler = (List<Kisi>)(Session["kisiler"]);
                kisiler.Add(kisi);
                Session["kisiler"] = kisiler;
            }

           

            using (MemoryStream mem = new MemoryStream())
            {

                var spreadsheetDocument = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook);
               // SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(mem, SpreadsheetDocumentType.Workbook, false);

                var workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

            
                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                SheetData sheetData1 = new SheetData();


                var baslikSatiri = new Row();
                baslikSatiri.Append(CreateCell("İsim"));
                baslikSatiri.Append(CreateCell("Soyisim"));
                baslikSatiri.Append(CreateCell("Adres"));
                baslikSatiri.Append(CreateCell("E-mail"));

                sheetData1.Append(baslikSatiri);

                foreach (var item in kisiler)
                {
                    var tRow = new Row();
                    tRow.Append(CreateCell(item.ad));
                    tRow.Append(CreateCell(item.soyad));
                    tRow.Append(CreateCell(item.adres));
                    tRow.Append(CreateCell(item.email));

                    sheetData1.Append(tRow);
                }
                

                worksheetPart.Worksheet = new Worksheet(sheetData1);

         
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

    
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "sayfa1"
                };

                
                sheets.Append(sheet);
                workbookpart.Workbook.Save();

                spreadsheetDocument.Close();

            

                string handle = Guid.NewGuid().ToString();

                mem.Position = 0;
                TempData[handle] = mem.ToArray();
                      
                return new JsonResult()
                {
                    Data = new { FileGuid = handle, FileName = "dosya.xlsx" }
                };
               
            }

           
        }

        private Cell CreateCell(string text)
        {
            Cell cell = new Cell();
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(text);
            return cell;
        }

        [HttpGet]
        public virtual ActionResult Download(string fileGuid, string fileName)
        {
            if (TempData[fileGuid] != null)
            {
                byte[] data = TempData[fileGuid] as byte[];
                return File(data, "application/vnd.ms-excel", fileName);
            }
            else
            {

                return new EmptyResult();
            }
        }

    }
}