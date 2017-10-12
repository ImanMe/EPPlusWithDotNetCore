using System.Collections.Generic;
using ExportToExcelDemo.Model;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExportToExcelDemo.Controllers
{
    [Route("api/[controller]")]
    public class ValuesController : Controller
    {          
        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            
        [HttpGet]
        public IActionResult Get(int id)
        {
            var listOfVehicles = new List<Vehicle>
            {
                new Vehicle {Id = 1, Make = "make1",Model = "model1"},
                new Vehicle {Id = 2, Make = "make2",Model = "model2"},
                new Vehicle {Id = 1, Make = "make3",Model = "model3"}
            };

            byte[] reportBytes;

            using (var package = CreateExcelPackage(listOfVehicles,"worksheetName"))
            {
                reportBytes = package.GetAsByteArray();
            }

            return File(reportBytes, XlsxContentType, "reportName.xlsx");
        }        
        
        private ExcelPackage CreateExcelPackage<T>(IEnumerable<T> dataToBeExported, string worksheetName)
        {            
            var package = new ExcelPackage();
            
            var worksheet = package.Workbook.Worksheets.Add(worksheetName);

            var headerCount =1;

            foreach (var header in typeof(T).GetProperties())
            {
                worksheet.Cells[1, headerCount].Value = header.Name;
                headerCount++;
            }
                       
            var rowCounter = 2;

            foreach (var v in dataToBeExported)
            {
                var columnCount = 0;
                foreach (var prop in v.GetType().GetProperties())
                {                    
                    worksheet.Cells[rowCounter, ++columnCount].Value = prop.GetValue(v, null);
                }                             
                rowCounter++;
            }
                                                
            // AutoFitColumns
            //worksheet.Cells[1, 1, 4, 4].AutoFitColumns();
           
            return package;
        }
    }
}
