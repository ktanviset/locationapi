using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using locationapi.Models;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data.SqlClient;
using System.Data;

namespace locationapi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ServiceOrderController : ControllerBase
    {
        private readonly IConfiguration configuration;

        public ServiceOrderController(IConfiguration config) 
        {
            configuration = config;
        }


        [HttpGet("{id}")]
        public ActionResult<TodoItem> GetServiceOrder(long id)
        {
            var todoItem = new TodoItem();
            todoItem.Id = id;
            todoItem.Name = "Test Name";

            if (todoItem == null)
            {
                return NotFound();
            }

            return todoItem;
        }

        [HttpPost]
        [Route("ImportExcelDataSheet")]
        public async Task<ActionResult> ImportExcelDataSheet(List<IFormFile> files)
        {
            // var connectionString = configuration.GetConnectionString("DefaultConnection");

            // var todoItem = new TodoItem();
            // todoItem.Id = 100;
            // todoItem.Name = connectionString +  Directory.GetCurrentDirectory();


            if (files == null || files.Count == 0)
                return Content("file not selected");

            var path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcel");//files[0].FileName

            string fullPath = Path.Combine(path, files[0].FileName);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string sb = "";
            string sFileExtension = Path.GetExtension(files[0].FileName).ToLower();
            ISheet sheet;
            using (var stream = new FileStream(fullPath, FileMode.Create))
            {
                await files[0].CopyToAsync(stream);

                stream.Position = 0;
                if (sFileExtension == ".xls")
                {
                    HSSFWorkbook hssfwb = new HSSFWorkbook(stream); //This will read the Excel 97-2000 formats  
                    sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook  
                }
                else
                {
                    XSSFWorkbook hssfwb = new XSSFWorkbook(stream); //This will read 2007 Excel format  
                    sheet = hssfwb.GetSheetAt(0); //get first sheet from workbook   
                }

                IRow headerRow = sheet.GetRow(0); //Get Header Row
                int cellCount = headerRow.LastCellNum;

                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++) //Read Excel File
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;
                    if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        if (row.GetCell(j) != null)
                            sb += "<td>" + row.GetCell(j).ToString() + "</td>";
                    }
                    sb += "</tr>";
                }
            }

            return new OkResult();
        }

        [HttpPost]
        [Route("ImportData")]
        public async Task<ActionResult> ImportData([FromBody] ImportServiceOrderModel serviceOrderModel)
        {
            try
            {
                await ImportToDb(serviceOrderModel);
            }
            catch (Exception e)
            {

            }
            return new OkResult();

        }

        private async Task ImportToDb(ImportServiceOrderModel serviceOrderModel)
        {
            var connectionString = configuration.GetConnectionString("DefaultConnection");
            using (var conn = new SqlConnection(connectionString))
            using (var command = new SqlCommand("PROC_Import_ServiceOrder", conn) { CommandType = CommandType.StoredProcedure }) {

                if (!string.IsNullOrEmpty(serviceOrderModel.StoreNo))            command.Parameters.Add("@StoreNo",             SqlDbType.NVarChar).Value = serviceOrderModel.StoreNo;
                if (serviceOrderModel.OrderCreationDate.HasValue)                command.Parameters.Add("@OrderCreationDate",   SqlDbType.Date).Value     = serviceOrderModel.OrderCreationDate;
                if (!string.IsNullOrEmpty(serviceOrderModel.DocumentNo))         command.Parameters.Add("@DocumentNo",          SqlDbType.NVarChar).Value = serviceOrderModel.DocumentNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceOrderNo))     command.Parameters.Add("@ServiceOrderNo",      SqlDbType.NVarChar).Value = serviceOrderModel.ServiceOrderNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceItemNo))      command.Parameters.Add("@ServiceItemNo",       SqlDbType.NVarChar).Value = serviceOrderModel.ServiceItemNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceName))        command.Parameters.Add("@ServiceName",         SqlDbType.NVarChar).Value = serviceOrderModel.ServiceName;
                if (serviceOrderModel.ServiceDate.HasValue)                      command.Parameters.Add("@ServiceDate",         SqlDbType.Date).Value     = serviceOrderModel.ServiceDate;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceTimeSlot))    command.Parameters.Add("@ServiceTimeSlot",     SqlDbType.NVarChar).Value = serviceOrderModel.ServiceTimeSlot;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceStatus))      command.Parameters.Add("@ServiceStatus",       SqlDbType.NVarChar).Value = serviceOrderModel.ServiceStatus;
                if (serviceOrderModel.ServiceGoodsValue.HasValue)                command.Parameters.Add("@ServiceGoodsValue",   SqlDbType.Float).Value    = serviceOrderModel.ServiceGoodsValue;
                if (!string.IsNullOrEmpty(serviceOrderModel.CapacityUnit))       command.Parameters.Add("@CapacityUnit",        SqlDbType.NVarChar).Value = serviceOrderModel.CapacityUnit;
                if (serviceOrderModel.CapacityValueWeight.HasValue)              command.Parameters.Add("@CapacityValueWeight", SqlDbType.Float).Value    = serviceOrderModel.CapacityValueWeight;
                if (serviceOrderModel.CapacityValueVolume.HasValue)              command.Parameters.Add("@CapacityValueVolume", SqlDbType.Float).Value    = serviceOrderModel.CapacityValueVolume;
                if (serviceOrderModel.BookedQty.HasValue)                        command.Parameters.Add("@BookedQty",           SqlDbType.Float).Value    = serviceOrderModel.BookedQty;
                if (serviceOrderModel.ServicePriceExclVAT.HasValue)              command.Parameters.Add("@ServicePriceExclVAT", SqlDbType.Float).Value    = serviceOrderModel.ServicePriceExclVAT;
                if (serviceOrderModel.ServicePriceInclVAT.HasValue)              command.Parameters.Add("@ServicePriceInclVAT", SqlDbType.Float).Value    = serviceOrderModel.ServicePriceInclVAT;
                if (!string.IsNullOrEmpty(serviceOrderModel.PriceCalcMethod))    command.Parameters.Add("@PriceCalcMethod",     SqlDbType.NVarChar).Value = serviceOrderModel.PriceCalcMethod;
                if (serviceOrderModel.NoofItems.HasValue)                        command.Parameters.Add("@NoofItems",           SqlDbType.Float).Value    = serviceOrderModel.NoofItems;
                if (serviceOrderModel.NoofPackages.HasValue)                     command.Parameters.Add("@NoofPackages",        SqlDbType.Float).Value    = serviceOrderModel.NoofPackages;
                if (serviceOrderModel.TotalOrderValue.HasValue)                  command.Parameters.Add("@TotalOrderValue",     SqlDbType.Float).Value    = serviceOrderModel.TotalOrderValue;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceProviderName))command.Parameters.Add("@ServiceProviderName", SqlDbType.NVarChar).Value = serviceOrderModel.ServiceProviderName;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceProviderID))  command.Parameters.Add("@ServiceProviderID",   SqlDbType.NVarChar).Value = serviceOrderModel.ServiceProviderID;
                if (!string.IsNullOrEmpty(serviceOrderModel.PaymentStatus))      command.Parameters.Add("@PaymentStatus",       SqlDbType.NVarChar).Value = serviceOrderModel.PaymentStatus;
                if (!string.IsNullOrEmpty(serviceOrderModel.PaymenttoIKEA_SP))   command.Parameters.Add("@PaymenttoIKEA_SP",    SqlDbType.NVarChar).Value = serviceOrderModel.PaymenttoIKEA_SP;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToCustomerName)) command.Parameters.Add("@ShipToCustomerName",  SqlDbType.NVarChar).Value = serviceOrderModel.ShipToCustomerName;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToAddress))      command.Parameters.Add("@ShipToAddress",       SqlDbType.NVarChar).Value = serviceOrderModel.ShipToAddress;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToAddress2))     command.Parameters.Add("@ShipToAddress2",      SqlDbType.NVarChar).Value = serviceOrderModel.ShipToAddress2;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToPostcode))     command.Parameters.Add("@ShipToPostcode",      SqlDbType.NVarChar).Value = serviceOrderModel.ShipToPostcode;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToCity))         command.Parameters.Add("@ShipToCity",          SqlDbType.NVarChar).Value = serviceOrderModel.ShipToCity;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToPhoneNo))      command.Parameters.Add("@ShipToPhoneNo",       SqlDbType.NVarChar).Value = serviceOrderModel.ShipToPhoneNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.ShipToEmail))        command.Parameters.Add("@ShipToEmail",         SqlDbType.NVarChar).Value = serviceOrderModel.ShipToEmail;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToCustomerName)) command.Parameters.Add("@SellToCustomerName",  SqlDbType.NVarChar).Value = serviceOrderModel.SellToCustomerName;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToAddress))      command.Parameters.Add("@SellToAddress",       SqlDbType.NVarChar).Value = serviceOrderModel.SellToAddress;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToAddress2))     command.Parameters.Add("@SellToAddress2",      SqlDbType.NVarChar).Value = serviceOrderModel.SellToAddress2;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToPostcode))     command.Parameters.Add("@SellToPostcode",      SqlDbType.NVarChar).Value = serviceOrderModel.SellToPostcode;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToCity))         command.Parameters.Add("@SellToCity",          SqlDbType.NVarChar).Value = serviceOrderModel.SellToCity;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToPhoneNo))      command.Parameters.Add("@SellToPhoneNo",       SqlDbType.NVarChar).Value = serviceOrderModel.SellToPhoneNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToMobilePhoneNo))command.Parameters.Add("@SellToMobilePhoneNo", SqlDbType.NVarChar).Value = serviceOrderModel.SellToMobilePhoneNo;
                if (!string.IsNullOrEmpty(serviceOrderModel.SellToEmail))        command.Parameters.Add("@SellToEmail",         SqlDbType.NVarChar).Value = serviceOrderModel.SellToEmail;
                if (!string.IsNullOrEmpty(serviceOrderModel.ServiceComment))     command.Parameters.Add("@ServiceComment",      SqlDbType.NVarChar).Value = serviceOrderModel.ServiceComment;
                if (!string.IsNullOrEmpty(serviceOrderModel.OrderComment))       command.Parameters.Add("@OrderComment",        SqlDbType.NVarChar).Value = serviceOrderModel.OrderComment;
                if (!string.IsNullOrEmpty(serviceOrderModel.SalesPerson))        command.Parameters.Add("@SalesPerson",         SqlDbType.NVarChar).Value = serviceOrderModel.SalesPerson;
                if (!string.IsNullOrEmpty(serviceOrderModel.CRMCaseID))          command.Parameters.Add("@CRMCaseID",           SqlDbType.NVarChar).Value = serviceOrderModel.CRMCaseID;
                if (serviceOrderModel.HandoverDate.HasValue)                     command.Parameters.Add("@HandoverDate",        SqlDbType.Date).Value     = serviceOrderModel.HandoverDate;
                if (serviceOrderModel.HandoverTime.HasValue)                     command.Parameters.Add("@HandoverTime",        SqlDbType.Time).Value     = serviceOrderModel.HandoverTime;

                conn.Open();
                await command.ExecuteNonQueryAsync();
            }
        }
    }

    /*
    {
	"StoreNo": "479",
	"OrderCreationDate": "2019-06-01",
	"DocumentNo": "R479-19000322",
	"ServiceOrderNo": "V479-190201910",
	"ServiceItemNo": "80000510",
	"ServiceName": "Return Service",
	"ServiceDate": "2019-06-04",
	"ServiceTimeSlot": "9:00..12:00",
	"ServiceStatus": "Waiting for Payment",
	"ServiceGoodsValue": 600.00,
	"CapacityUnit": "TRANSPORT",
	"CapacityValueWeight": 2.4,
	"CapacityValueVolume": 0.00393,
	"BookedQty": 1,
	"ServicePriceExclVAT": 0,
	"ServicePriceInclVAT": 0,
	"PriceCalcMethod": "PER SERVIC",
	"NoofItems": 1,
	"NoofPackages": 1,
	"TotalOrderValue": 600.00,
	"ServiceProviderName": "M-World Logistics (Thailand)",
	"ServiceProviderID": "SERVPROV1900003",
	"PaymentStatus": "Not Paid",
	"PaymenttoIKEA_SP": "Pay to IKEA",
	"ShipToCustomerName": "ขนิษภา วงษ์พิพัฒน์พันธ์",
	"ShipToAddress": "5/9 หมู่บ้านมอตโต้ ซ.A1 ถ.กาญจนาภิเษก",
	"ShipToAddress2": "แขวงบางบอนใต้ เขตบางบอน",
	"ShipToPostcode": "10150",
	"ShipToCity": "Bangkok",
	"ShipToPhoneNo": "+66896622252",
	"ShipToEmail": "cher_rylee@hotmail.com",
	"SellToCustomerName": "ขนิษภา วงษ์พิพัฒน์พันธ์",
	"SellToAddress": "5/9 หมู่บ้านมอตโต้ ซ.A1 ถ.กาญจนาภิเษก",
	"SellToAddress2": "แขวงบางบอนใต้ เขตบางบอน",
	"SellToPostcode": "10150",
	"SellToCity": "Bangkok",
	"SellToPhoneNo": "+66896622252",
	"SellToMobilePhoneNo": "+66896622252",
	"SellToEmail": "cher_rylee@hotmail.com",
	"ServiceComment": "01/06/19 คุณ ขนิษภา โทร.089-6622252 01/06/19 นำสินค้าตัวใหม่ไปให้สินค้าเก่ากลับมาที่อิเกียเเล้วค่ะ 01/06/19 อิเกียจ่ายค่าจัดส่ง 510 บาท 01/06/19 HD ดูเเลเรื่องค่าประกอบต่อค่ะ 01/06/19 ยืนยันวันนัดเข้าหน้างาน 04/06/2019 เวลา 09.00-12.00 น. ",
	"OrderComment": "",
	"SalesPerson": "RADCHANEEW",
	"CRMCaseID": "",
	"HandoverDate": "2019-07-05",
	"HandoverTime": "02:58:30"
}

    */

    public class TodoItem {
        public long Id { get; set; }
        public string Name { get; set; }
    }
}