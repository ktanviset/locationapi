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

            using (var stream = new FileStream(fullPath, FileMode.Create))
            {
                await files[0].CopyToAsync(stream);
            }

            return new OkResult();
        }
    }

    

    public class TodoItem {
        public long Id { get; set; }
        public string Name { get; set; }
    }
}