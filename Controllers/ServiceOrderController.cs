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
using locationapi.AppExtensions;

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

        [HttpPost]
        [Route("ImportExcelDataSheet")]
        public async Task<ActionResult<ResponseModel>> ImportExcelDataSheet(List<IFormFile> files)
        {
            if (files == null || files.Count == 0)
                return new ResponseModel() { IsSuccess = false, Message = "file not selected" };

            string fileExtension = Path.GetExtension(files[0].FileName).ToLower();
            if (!(fileExtension == ".xls" || fileExtension == ".xlsx"))
                return new ResponseModel() { IsSuccess = false, Message = "file is not excel" };

            int processRow = 0;
            try
            {
                string path = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadExcel");//files[0].FileName
                string fullPath = Path.Combine(path, files[0].FileName);

                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);

                int fileRunning = 1;
                while (System.IO.File.Exists(fullPath))
                {
                    fullPath = Path.Combine(path, $"{Path.GetFileNameWithoutExtension(files[0].FileName)}_{fileRunning}{fileExtension}");
                    fileRunning++;
                }

                ISheet sheet;
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    await files[0].CopyToAsync(stream);

                    stream.Position = 0;
                    if (fileExtension == ".xls")
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
                        processRow = i;
                        IRow row = sheet.GetRow(i);
                        if (row == null) continue;
                        if (row.Cells.All(d => d.CellType == CellType.Blank)) continue;
                        // for (int j = row.FirstCellNum; j < cellCount; j++)
                        // {
                        //     if (row.GetCell(j) != null){
                        //         var value = row.GetCell(j).ToString();
                        //     }
                        // }
                        ImportServiceOrderModel importModel = new ImportServiceOrderModel();

                        importModel.StoreNo                  = $"{row.GetCell(0)}";
                        importModel.OrderCreationDate           = row.GetCell(1).DateCellValue.ToNullableDateTimeValue();
                        importModel.DocumentNo               = $"{row.GetCell(2)}";
                        importModel.ServiceOrderNo           = $"{row.GetCell(3)}";
                        importModel.ServiceItemNo            = $"{row.GetCell(4)}";
                        importModel.ServiceName              = $"{row.GetCell(5)}";
                        importModel.ServiceDate                 = row.GetCell(6).DateCellValue.ToNullableDateTimeValue();
                        importModel.ServiceTimeSlot          = $"{row.GetCell(7)}";
                        importModel.ServiceStatus            = $"{row.GetCell(8)}";
                        importModel.ServiceGoodsValue    = (decimal)row.GetCell(9).NumericCellValue;
                        importModel.CapacityUnit             = $"{row.GetCell(10)}";
                        importModel.CapacityValueWeight  = (decimal)row.GetCell(11).NumericCellValue;
                        importModel.CapacityValueVolume  = (decimal)row.GetCell(12).NumericCellValue;
                        importModel.BookedQty            = (decimal)row.GetCell(13).NumericCellValue;
                        importModel.ServicePriceExclVAT  = (decimal)row.GetCell(14).NumericCellValue;
                        importModel.ServicePriceInclVAT  = (decimal)row.GetCell(15).NumericCellValue;
                        importModel.PriceCalcMethod          = $"{row.GetCell(16)}";
                        importModel.NoofItems            = (decimal)row.GetCell(17).NumericCellValue;
                        importModel.NoofPackages         = (decimal)row.GetCell(18).NumericCellValue;
                        importModel.TotalOrderValue      = (decimal)row.GetCell(19).NumericCellValue;
                        importModel.ServiceProviderName      = $"{row.GetCell(20)}";
                        importModel.ServiceProviderID        = $"{row.GetCell(21)}";
                        importModel.PaymentStatus            = $"{row.GetCell(22)}";
                        importModel.PaymenttoIKEA_SP         = $"{row.GetCell(23)}";
                        importModel.ShipToCustomerName       = $"{row.GetCell(24)}";
                        importModel.ShipToAddress            = $"{row.GetCell(25)}";
                        importModel.ShipToAddress2           = $"{row.GetCell(26)}";
                        importModel.ShipToPostcode           = $"{row.GetCell(27)}";
                        importModel.ShipToCity               = $"{row.GetCell(28)}";
                        importModel.ShipToPhoneNo            = $"{row.GetCell(29)}";
                        importModel.ShipToEmail              = $"{row.GetCell(30)}";
                        importModel.SellToCustomerName       = $"{row.GetCell(31)}";
                        importModel.SellToAddress            = $"{row.GetCell(32)}";
                        importModel.SellToAddress2           = $"{row.GetCell(33)}";
                        importModel.SellToPostcode           = $"{row.GetCell(34)}";
                        importModel.SellToCity               = $"{row.GetCell(35)}";
                        importModel.SellToPhoneNo            = $"{row.GetCell(36)}";
                        importModel.SellToMobilePhoneNo      = $"{row.GetCell(37)}";
                        importModel.SellToEmail              = $"{row.GetCell(38)}";
                        importModel.ServiceComment           = $"{row.GetCell(39)}";
                        importModel.OrderComment             = $"{row.GetCell(40)}";
                        importModel.SalesPerson              = $"{row.GetCell(41)}";
                        importModel.CRMCaseID                = $"{row.GetCell(42)}";
                        importModel.HandoverDate                = row.GetCell(43).DateCellValue.ToNullableDateTimeValue();
                        importModel.HandoverTime                = row.GetCell(44).DateCellValue.ToNullableDateTimeValue();

                        await ImportToDb(importModel);
                    }
                }
            }
            catch (Exception e)
            {
                return new ResponseModel() { IsSuccess = false, Message = $"Row:{processRow}\r\nMessage:{e.Message}\r\nStackTrace:{e.StackTrace}" };
            }

            return new ResponseModel();
        }

        [HttpPost]
        [Route("ImportData")]
        public async Task<ActionResult<ResponseModel>> ImportData([FromBody] ImportServiceOrderModel serviceOrderModel)
        {
            try
            {
                await ImportToDb(serviceOrderModel);
            }
            catch (Exception e)
            {
                return new ResponseModel() { IsSuccess = false, Message = e.Message };
            }
            return new ResponseModel();
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
                if (serviceOrderModel.ServiceGoodsValue.HasValue)                command.Parameters.Add("@ServiceGoodsValue",   SqlDbType.Decimal).Value  = serviceOrderModel.ServiceGoodsValue;
                if (!string.IsNullOrEmpty(serviceOrderModel.CapacityUnit))       command.Parameters.Add("@CapacityUnit",        SqlDbType.NVarChar).Value = serviceOrderModel.CapacityUnit;
                if (serviceOrderModel.CapacityValueWeight.HasValue)              command.Parameters.Add("@CapacityValueWeight", SqlDbType.Decimal).Value  = serviceOrderModel.CapacityValueWeight;
                if (serviceOrderModel.CapacityValueVolume.HasValue)              command.Parameters.Add("@CapacityValueVolume", SqlDbType.Decimal).Value  = serviceOrderModel.CapacityValueVolume;
                if (serviceOrderModel.BookedQty.HasValue)                        command.Parameters.Add("@BookedQty",           SqlDbType.Decimal).Value  = serviceOrderModel.BookedQty;
                if (serviceOrderModel.ServicePriceExclVAT.HasValue)              command.Parameters.Add("@ServicePriceExclVAT", SqlDbType.Decimal).Value  = serviceOrderModel.ServicePriceExclVAT;
                if (serviceOrderModel.ServicePriceInclVAT.HasValue)              command.Parameters.Add("@ServicePriceInclVAT", SqlDbType.Decimal).Value  = serviceOrderModel.ServicePriceInclVAT;
                if (!string.IsNullOrEmpty(serviceOrderModel.PriceCalcMethod))    command.Parameters.Add("@PriceCalcMethod",     SqlDbType.NVarChar).Value = serviceOrderModel.PriceCalcMethod;
                if (serviceOrderModel.NoofItems.HasValue)                        command.Parameters.Add("@NoofItems",           SqlDbType.Decimal).Value  = serviceOrderModel.NoofItems;
                if (serviceOrderModel.NoofPackages.HasValue)                     command.Parameters.Add("@NoofPackages",        SqlDbType.Decimal).Value  = serviceOrderModel.NoofPackages;
                if (serviceOrderModel.TotalOrderValue.HasValue)                  command.Parameters.Add("@TotalOrderValue",     SqlDbType.Decimal).Value  = serviceOrderModel.TotalOrderValue;
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
                if (serviceOrderModel.HandoverTime.HasValue)                     command.Parameters.Add("@HandoverTime",        SqlDbType.Time).Value     = serviceOrderModel.HandoverTime.Value.ToString("HH:mm:ss");

                conn.Open();
                await command.ExecuteNonQueryAsync();
            }
        }
    }
}