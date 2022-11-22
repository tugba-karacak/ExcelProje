using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProje.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using Microsoft.AspNetCore.Http;

namespace ExcelProje.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public const int Status404NotFound = 404;
        public IActionResult Index()
        {
            string url = "http://www.tcmb.gov.tr/kurlar/";
            string date = DateTime.Now.AddDays(-1).ToShortDateString();
          
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            var excelBlank = excelPackage.Workbook.Worksheets.Add("Ödev2");
            excelBlank.Cells[1, 1].Value = "Tarih";
            excelBlank.Cells[1, 2].Value = "Kur";
            excelBlank.Cells[1, 3].Value = "EURO/USD";

            List<Kur> kurUsdListesi = new List<Kur>();
            List<Kur> kurEuroListesi = new List<Kur>();
            for (int i = -1; i > -31; i--)
            {
               
                try
                {

                    Kur kurUSD = new Kur();
                    Kur kurEURO = new Kur();
                    url = TarihConvertXml(date);
                    var xmldoc = new XmlDocument();
                    xmldoc.Load(url);
                    kurUSD.kur =Convert.ToDouble( xmldoc.SelectSingleNode("Tarih_Date/Currency [@Kod='USD']/BanknoteSelling").InnerXml.Replace('.',','));
                    kurUSD.Kurtarihi = date;
                    kurEURO.kur = Convert.ToDouble(xmldoc.SelectSingleNode("Tarih_Date/Currency [@Kod='EUR']/BanknoteSelling").InnerXml.Replace('.', ','));
                    kurEURO.Kurtarihi = date;
                    kurUsdListesi.Add(kurUSD);
                    kurEuroListesi.Add(kurEURO);
                    date = DateTime.Now.AddDays(i - 1).ToShortDateString();
                }
                catch (Exception)
                {
                    i--;
                    date = DateTime.Now.AddDays(i - 1).ToShortDateString();
                    continue;
                    
                }
                
            }

            List<Kur> kurUsdListesiSirali=kurUsdListesi.OrderByDescending(x=>x.kur).ToList();
            List<Kur> kurEuroListesiSirali = kurEuroListesi.OrderByDescending(x => x.kur).ToList();
            int row = 2;
            
            for (int i = 0; i < 5; i++)
            {
                excelBlank.Cells[row, 1].Value = kurUsdListesiSirali[i].Kurtarihi;
                excelBlank.Columns[1].AutoFit();
                excelBlank.Cells[row, 2].Value = kurUsdListesiSirali[i].kur;
                excelBlank.Columns[1].AutoFit();
                excelBlank.Cells[row, 3].Value = "USD";
                excelBlank.Columns[1].AutoFit();
                row++;
            }
            for (int i = 0; i < 5; i++)
            {
                excelBlank.Cells[row, 1].Value = kurEuroListesiSirali[i].Kurtarihi;
                excelBlank.Columns[1].AutoFit();
                excelBlank.Cells[row, 2].Value = kurEuroListesiSirali[i].kur;
                excelBlank.Columns[1].AutoFit();
                excelBlank.Cells[row, 3].Value = "EURO";
                excelBlank.Columns[1].AutoFit();
                row++;
            }
            var bytes = excelPackage.GetAsByteArray();
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ödev.xlsx");

            


        }

        public string TarihConvertXml(String date) {
            string url = "http://www.tcmb.gov.tr/kurlar/";
            List<string> tarihBilgileri = date.Split('.').ToList();
            if (tarihBilgileri[0].Length<=1)
            {
                tarihBilgileri[0] = "0" + tarihBilgileri[0];
            }
            if (tarihBilgileri[1].Length <= 1)
            {
                tarihBilgileri[1] = "0" + tarihBilgileri[1];
            }
            string yilAy = tarihBilgileri[2] + tarihBilgileri[1];
            string gunAyYil = tarihBilgileri[0] + tarihBilgileri[1]+tarihBilgileri[2];
            url = url + yilAy + "/" + gunAyYil+".xml";
            return url;
        }


     
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
