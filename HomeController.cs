using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;
using UploadExcel.Models;

namespace UploadExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult ExcelUpload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ReadExcel()
        {
            List<Student> studentList = new List<Student>();
            if (ModelState.IsValid)
            {

                string filePath = string.Empty;
                if (Request != null)
                {
                    HttpPostedFileBase file = Request.Files["file"];
                    if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                    {

                        string fileName = file.FileName;
                        string fileContentType = file.ContentType;
                        string path = Server.MapPath("~/Uploads/");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                        filePath = path + Path.GetFileName(file.FileName);
                        string extension = Path.GetExtension(file.FileName);
                        file.SaveAs(filePath);
                        Stream stream = file.InputStream;
                       
                        
                        IExcelDataReader reader = null;
                        if (file.FileName.EndsWith(".xls"))
                        {
                            reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        else if (file.FileName.EndsWith(".xlsx"))
                        {
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        else
                        {
                            ModelState.AddModelError("File", "This file format is not supported");
                            return RedirectToAction("ExcelUpload");
                        }
                      


                        DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });


                        reader.Close();
                        
                        
                        string filedetails = path + fileName;
                        FileInfo fileinfo = new FileInfo(filedetails);
                        if (fileinfo.Exists)
                        {
                            fileinfo.Delete();
                        }
                        DataTable dt = result.Tables[0];
                        studentList = ConvertDataTable<Student>(dt); 
                        TempData["Excelstudent"] = studentList;
                    }
                }

            }
            // var files = Request.Files;  

            return new JsonResult { Data = studentList, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }
        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }
        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name.Trim().ToUpper() == column.ColumnName.Trim().ToUpper())
                        pro.SetValue(obj, Convert.ChangeType(dr[column.ColumnName], pro.PropertyType) ?? null, null);
                    else
                        continue;
                }
            }
            return obj;
        }
        public ActionResult UpdateCustomerTable()
        {
            int length = 0;
            try
            {
                if (TempData["Excelstudent"] != null)
                {
                    List<Student> lstStudent = (List<Student>)TempData["Excelstudent"];
                    using (ApplicationDbContext db = new ApplicationDbContext())
                    {
                        foreach (var s in lstStudent)
                        {
                            db.Students.Add(s);
                        }
                        db.SaveChanges();
                        length = lstStudent.Count();
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
            }
            return new JsonResult { Data = length, JsonRequestBehavior = JsonRequestBehavior.AllowGet };
        }
    }
}