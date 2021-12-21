using LogictecaTest.Data;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Dynamic.Core;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Http;
using LogictecaTest.LinqToDataTable;
using OfficeOpenXml;
using System.Text;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using System.Data.OleDb;
using System.Data;
using LogictecaTest.Models;
using LogictecaTest.Utilities;

namespace LogictecaTest.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ItemController : ControllerBase
    {
        private readonly ApplicationDbContext context;
        private readonly IServiceProvider _serviceProvider;
        private readonly IHostingEnvironment _hostingEnvironment;


        public ItemController(IHostingEnvironment hostingEnvironment,ApplicationDbContext context, IServiceProvider serviceProvider)
        {
            this.context = context;
            _serviceProvider = serviceProvider;
            _hostingEnvironment = hostingEnvironment;
        }

        [HttpPost("Import")]
        public async Task<IActionResult> Import([FromForm]IFormFile file)
        {
            string uploads = Path.Combine(_hostingEnvironment.WebRootPath, "import");
            Directory.CreateDirectory(uploads);
            string filePath = Path.Combine(uploads, Guid.NewGuid().ToString()+ file.FileName);
            using (Stream fileStream = new FileStream(filePath, FileMode.OpenOrCreate))
            {
                await file.CopyToAsync(fileStream);
            }
            var scope = _serviceProvider.CreateScope();
            Task.Run(() => ProcessImport(scope, filePath));
            return Ok();
        }

        private void ProcessImport(IServiceScope scope, string filePath)
        {
            using (scope)
            {
               var applicationDbContext = scope.ServiceProvider.GetService<ApplicationDbContext>();
                try
                {
                    var ExcelRows = SelectRowsFromExcel(filePath, $"SELECT * FROM [Cisco PSS Services - Dec 2020$]");
                    var min_save_length = 20000;
                    List<Item> records = new List<Item>();
                    foreach (IDataReader reader in ExcelRows.Skip(2))
                    {
                        if (!reader.IsDBNull(5)&& reader.FieldCount >= 8)
                        {
                            var Part_SKU = reader.IsDBNull(4) ? null : reader.GetString(4);
                            records=records.Where(s => s.Part_SKU != Part_SKU).ToList();
                            records.Add(new Item
                            {
                                Band = reader.IsDBNull(1) ? null : reader.GetString(1),
                                Category_Code = reader.IsDBNull(2) ? null : reader.GetString(2),
                                Manufacturer = reader.IsDBNull(3) ? null : reader.GetString(3),
                                Part_SKU = Part_SKU,
                                Item_Description = reader.IsDBNull(5) ? null : reader.GetString(5),
                                List_Price = reader.IsDBNull(6) ? null : reader.GetString(6),
                                 Minimum_Discount = reader.IsDBNull(7) ? null : reader.GetString(7),
                                 Discounted_Price = reader.IsDBNull(8) ? null : reader.GetString(8)
                            });

                        }
                        SaveToDb(records, applicationDbContext, min_save_length);
                    }
                    SaveToDb(records, applicationDbContext, 0);
                }



                catch (Exception ex)
                {
                }

            }
        }

        private void SaveToDb(List<Item> records,ApplicationDbContext applicationDbContext, int min_save_length)
        {
            if (records.Count >= min_save_length && records.Count>0)
            {

                try
                {
                    var updateItems = applicationDbContext.Items.Where(s => records.Select(q => q.Part_SKU).Any(Part_SKU => Part_SKU == s.Part_SKU)).ToList();
                    var insertItems = records.Where(s => !updateItems.Any(item => item.Part_SKU == s.Part_SKU));
                    foreach (var item in updateItems)
                    {
                        var newdata = records.LastOrDefault(s => s.Part_SKU == item.Part_SKU);
                        item.Band = newdata.Band;
                        item.Category_Code = newdata.Category_Code;
                        item.Manufacturer = newdata.Manufacturer;
                        item.Item_Description = newdata.Item_Description;
                        item.List_Price = newdata.List_Price;
                        item.Minimum_Discount = newdata.Minimum_Discount;
                        item.Discounted_Price = newdata.Discounted_Price;
                    }
                    if (updateItems.Count > 0)
                        applicationDbContext.BulkUpdate(updateItems);
                    if (insertItems.Count() > 0)
                        applicationDbContext.BulkInsert(insertItems);
                    applicationDbContext.SaveChanges();
                }
                catch (Exception ex)
                {

                }
                records.Clear();
            }

        }
        private IEnumerable<IDataRecord> SelectRowsFromExcel(string filePath,string query)
        {

                string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + "; Extended Properties='Excel 8.0;HDR=NO;IMEX=1;'";


                // Create the connection object
                OleDbConnection oledbConn = new OleDbConnection(connString);
                     // Open connection
                    oledbConn.Open();

                    // Create OleDbCommand object and select data from worksheet Sample-spreadsheet-file
                    //here sheet name is Sample-spreadsheet-file, usually it is Sheet1, Sheet2 etc..
                    OleDbCommand cmd = new OleDbCommand(query, oledbConn);


            using (IDataReader rdr = cmd.ExecuteReader())
            {
                while (rdr.Read())
                {
                    yield return (IDataRecord)rdr;
                }

            }
 
 


        }

        [HttpPost]
        public IActionResult GetItems()
        {

            try
            {

                var draw = Request.Form["draw"].FirstOrDefault();
                var start = Request.Form["start"].FirstOrDefault();
                var length = Request.Form["length"].FirstOrDefault();
                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                int pageSize = length != null ? Convert.ToInt32(length) : 20;
                int skip = start != null ? Convert.ToInt32(start) : 0;
                int recordsTotal = 0;
                var itemData = (from tempcustomer in context.Items select tempcustomer);
                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
                {
                    itemData = itemData.OrderBy(sortColumn + " " + sortColumnDirection);
                }

                #region Search

                var Search_Band = Request.Form["columns[0][search][value]"].FirstOrDefault();
                var Search_Category_Code = Request.Form["columns[1][search][value]"].FirstOrDefault();
                var Search_Manufacturer = Request.Form["columns[2][search][value]"].FirstOrDefault();
                var Search_Part_SKU = Request.Form["columns[3][search][value]"].FirstOrDefault();
                var Search_Item_Description = Request.Form["columns[4][search][value]"].FirstOrDefault();
                var Search_List_Price = Request.Form["columns[5][search][value]"].FirstOrDefault();
                var Search_Minimum_Discount = Request.Form["columns[6][search][value]"].FirstOrDefault();
                var Search_Discounted_Price = Request.Form["columns[7][search][value]"].FirstOrDefault();

                if (!string.IsNullOrEmpty(Search_Band))
                    itemData = itemData.Where(m => m.Band.Contains(Search_Band));


                if (!string.IsNullOrEmpty(Search_Category_Code))
                    itemData = itemData.Where(m => m.Category_Code.Contains(Search_Category_Code));


                if (!string.IsNullOrEmpty(Search_Manufacturer))
                    itemData = itemData.Where(m => m.Manufacturer.Contains(Search_Manufacturer));
               
                if (!string.IsNullOrEmpty(Search_Part_SKU))
                    itemData = itemData.Where(m => m.Part_SKU.Contains(Search_Part_SKU));

                if (!string.IsNullOrEmpty(Search_Item_Description))
                    itemData = itemData.Where(m => m.Item_Description.Contains(Search_Item_Description));


                if (!string.IsNullOrEmpty(Search_List_Price))
                    itemData = itemData.Where(m => m.List_Price.Contains(Search_List_Price));


                if (!string.IsNullOrEmpty(Search_Minimum_Discount))
                    itemData = itemData.Where(m => m.Minimum_Discount.Contains(Search_Minimum_Discount));

                if (!string.IsNullOrEmpty(Search_Discounted_Price))
                    itemData = itemData.Where(m => m.Discounted_Price.Contains(Search_Discounted_Price));

                #endregion

                recordsTotal = itemData.Count();
                var data = itemData.Skip(skip).Take(pageSize).ToList();
                var jsonData = new { draw = draw, recordsFiltered = recordsTotal, recordsTotal = recordsTotal, data = data };
                return Ok(jsonData);

            }
            catch (Exception ex)
            {
                throw;
            }
        }

        [HttpPost("Export")]
        public IActionResult Export()
        {

            try
            {

                var sortColumn = Request.Form["columns[" + Request.Form["order[0][column]"].FirstOrDefault() + "][data]"].FirstOrDefault();
                var sortColumnDirection = Request.Form["order[0][dir]"].FirstOrDefault();
                var itemData = (from tempcustomer in context.Items select tempcustomer);
                if (!(string.IsNullOrEmpty(sortColumn) && string.IsNullOrEmpty(sortColumnDirection)))
                {
                    itemData = itemData.OrderBy(sortColumn + " " + sortColumnDirection);
                }

                #region Search

                var Search_Band = Request.Form["columns[0][search][value]"].FirstOrDefault();
                var Search_Category_Code = Request.Form["columns[1][search][value]"].FirstOrDefault();
                var Search_Manufacturer = Request.Form["columns[2][search][value]"].FirstOrDefault();
                var Search_Part_SKU = Request.Form["columns[3][search][value]"].FirstOrDefault();
                var Search_Item_Description = Request.Form["columns[4][search][value]"].FirstOrDefault();
                var Search_List_Price = Request.Form["columns[5][search][value]"].FirstOrDefault();
                var Search_Minimum_Discount = Request.Form["columns[6][search][value]"].FirstOrDefault();
                var Search_Discounted_Price = Request.Form["columns[7][search][value]"].FirstOrDefault();

                if (!string.IsNullOrEmpty(Search_Band))
                    itemData = itemData.Where(m => m.Band.Contains(Search_Band));


                if (!string.IsNullOrEmpty(Search_Category_Code))
                    itemData = itemData.Where(m => m.Category_Code.Contains(Search_Category_Code));


                if (!string.IsNullOrEmpty(Search_Manufacturer))
                    itemData = itemData.Where(m => m.Manufacturer.Contains(Search_Manufacturer));

                if (!string.IsNullOrEmpty(Search_Part_SKU))
                    itemData = itemData.Where(m => m.Part_SKU.Contains(Search_Part_SKU));

                if (!string.IsNullOrEmpty(Search_Item_Description))
                    itemData = itemData.Where(m => m.Item_Description.Contains(Search_Item_Description));


                if (!string.IsNullOrEmpty(Search_List_Price))
                    itemData = itemData.Where(m => m.List_Price.Contains(Search_List_Price));


                if (!string.IsNullOrEmpty(Search_Minimum_Discount))
                    itemData = itemData.Where(m => m.Minimum_Discount.Contains(Search_Minimum_Discount));

                if (!string.IsNullOrEmpty(Search_Discounted_Price))
                    itemData = itemData.Where(m => m.Discounted_Price.Contains(Search_Discounted_Price));

                #endregion

                DataSet ds = new DataSet();
                ds.Tables.Add(itemData.Select(s=>new { 
                    s.Band,
                    s.Category_Code,
                    s.Manufacturer,
                    s.Part_SKU,
                    s.Item_Description,
                    s.List_Price,
                    s.Minimum_Discount,
                    s.Discounted_Price,

                }).CopyToDataTable());

                
               
                string uploads = Path.Combine(_hostingEnvironment.WebRootPath, "export");
                Directory.CreateDirectory(uploads);
                string filename = $"Items-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";
                string filePath = Path.Combine(uploads, filename);
                ds.Tables[0].ExportToExcel(filePath, "Cisco PSS Services - Dec 2020");



                string url = "/export/"+ filename;
                return Ok(new { url });

            }
            catch (Exception ex)
            {
                throw;
            }
        }
    }
    public static class My_DataTable_Extensions
    {

        // Export DataTable into an excel file with field names in the header line
        // - Save excel file without ever making it visible if filepath is given
        // - Don't save excel file, just make it visible if no filepath is given
        public static void ExportToExcel(this DataTable tbl, string excelFilePath,string sheetName)
        {
            try
            {
                if (tbl == null || tbl.Columns.Count == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Add();

                // single worksheet
                Microsoft.Office.Interop.Excel._Worksheet workSheet = excelApp.ActiveSheet;
                workSheet.Name = sheetName;
                // column headings
                workSheet.Cells[1, 1] = "Date : " + DateTime.Now.ToString("dd/MM/yyyy");
                for (var i = 0; i < tbl.Columns.Count; i++)
                {
                    workSheet.Cells[2, i + 1] = tbl.Columns[i].ColumnName;
                }


                // Get dimensions of the 2-d array
                int rowCount = tbl.Rows.Count;
                int columnCount = tbl.Columns.Count;
                Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)workSheet.Cells[3, 1];
                range = range.get_Resize(rowCount, columnCount);
                var matrix =tbl.Rows.OfType<DataRow>().Select(s => new string[] { 
                    s.ItemArray[0].ToString(),
                    s.ItemArray[1].ToString(),
                    s.ItemArray[2].ToString(),
                    s.ItemArray[3].ToString(),
                    s.ItemArray[4].ToString(),
                    s.ItemArray[5].ToString(),
                    s.ItemArray[6].ToString(),
                    s.ItemArray[7].ToString(),
                });
                range.Value2 = CreateRectangularArray<string>(matrix.ToArray());

                // check file path
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        workSheet.SaveAs(excelFilePath);
                        excelApp.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                                            + ex.Message);
                    }
                }
                else
                { // no file path is given
                    excelApp.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }
        static T[,] CreateRectangularArray<T>(T[][] arrays)
        {
            // TODO: Validation and special-casing for arrays.Count == 0
            int minorLength = arrays[0].Length;
            T[,] ret = new T[arrays.Length, minorLength];
            for (int i = 0; i < arrays.Length; i++)
            {
                var array = arrays[i];
                if (array.Length != minorLength)
                {
                    throw new ArgumentException
                        ("All arrays must be the same length");
                }
                for (int j = 0; j < minorLength; j++)
                {
                    ret[i, j] = array[j];
                }
            }
            return ret;
        }

    }
}
