using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using locationapi.AppExtensions;
using locationapi.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace locationapi.Controllers
{
    [Route("[controller]")]
    public class ScancodeController : Controller
    {
        private readonly IConfiguration configuration;

        public ScancodeController(IConfiguration config)
        {
            configuration = config;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        [Route("SaveCodeId/{codeid}")]
        public ActionResult<string> SaveCodeId(string codeid)
        {
            string response = codeid;
            var connectionString = configuration.GetConnectionString("DefaultConnection");

            try
            {
                using (var conn = new SqlConnection(connectionString))
                using (var command = new SqlCommand("PROC_Push_Scancode", conn) { CommandType = CommandType.StoredProcedure })
                {
                    if (!string.IsNullOrEmpty(codeid))
                        command.Parameters.Add("@code_id", SqlDbType.NVarChar).Value = codeid;
                    command.Parameters.Add("import_time", SqlDbType.DateTime).Value = DateTime.Now;
                    command.Parameters.Add("export_time", SqlDbType.DateTime).Value = DateTime.Now;

                    conn.Open();
                    int result = command.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                response = e.Message;
            }

            return response;
        }

        [HttpGet]
        [Route("GetList/{datetype}/")]
        public ActionResult<List<ScancodeModel>> GetList(string datetype, [FromQuery] string datestring)
        {
            var connectionString = configuration.GetConnectionString("DefaultConnection");
            List<ScancodeModel> responde = new List<ScancodeModel>();
            try
            {
                string sql = "select * from scancode";
                var datefromto = datestring.Split("between", 2, StringSplitOptions.None);
                if (datefromto != null && datefromto.Length == 2)
                {
                    bool haswhere = false;
                    if (!string.IsNullOrWhiteSpace(datefromto[0]))
                    {
                        haswhere = true;
                        sql += $" where {datetype} >= '{datefromto[0].Replace('T', ' ')}'";
                    }

                    if (!string.IsNullOrWhiteSpace(datefromto[1]))
                    {
                        if (haswhere)
                            sql += " and";
                        else
                            sql += " where";

                        sql += $" {datetype} <= '{datefromto[1].Replace('T', ' ')}'";
                    }
                }

                sql += " order by import_time desc";

                using (var conn = new SqlConnection(connectionString))
                using (var command = new SqlCommand(sql, conn))
                {
                    conn.Open();
                    using (SqlDataReader data = command.ExecuteReader())
                    {
                        while (data.Read())
                        {
                            var scMosel = new ScancodeModel()
                            {
                                CodeId = SqlExtensions.Read<string>(data, "code_id"),
                                ImportTime = SqlExtensions.Read<DateTime?>(data, "import_time"),
                                ExportTime = SqlExtensions.Read<DateTime?>(data, "export_time"),
                            };

                            responde.Add(scMosel);
                        }
                    }
                }

                return responde;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}