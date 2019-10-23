using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace WebApplication7.Controllers
{
    public class HomeController : Controller
    {

        Models.MVCtest2Entities db = new Models.MVCtest2Entities();
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult About(HttpPostedFileBase file)
        {
            if (file != null)
            {
                string extension = Path.GetExtension(file.FileName);
                string fileLocation = Server.MapPath("~/Content/") + file.FileName;
                if (extension == ".xls" || extension == ".xlsx")
                {

                    if (System.IO.File.Exists(fileLocation)) // 驗證檔案是否存在
                    {
                        System.IO.File.Delete(fileLocation);
                    }

                    file.SaveAs(fileLocation); // 存放檔案到伺服器上
                }
                if (extension == ".xls")
                {
                    HSSFWorkbook excel;

                    var members = new List<Models.test1>();
                    using (FileStream files = new FileStream(fileLocation, FileMode.Open, FileAccess.Read))
                    {
                        excel = new HSSFWorkbook(files);
                    }
                    ISheet sheet = excel.GetSheetAt(0);
                    for (int row = 1; row <= sheet.LastRowNum; row++) // 使用For 走訪所有的資料列
                    {
                        if (sheet.GetRow(row) != null) // 驗證是不是空白列
                        {
                            HSSFCell cell = (HSSFCell)sheet.GetRow(row).GetCell(0);
                            cell.SetCellType(CellType.String);
                            string cellValue = cell.StringCellValue;

                            HSSFCell cell1 = (HSSFCell)sheet.GetRow(row).GetCell(1);
                            cell1.SetCellType(CellType.Numeric);
                            string cellValue1 = cell1.DateCellValue.ToString("yyyy/MM/dd");

                            HSSFCell cell2 = (HSSFCell)sheet.GetRow(row).GetCell(2);
                            cell2.SetCellType(CellType.String);
                            string cellValue2 = cell2.StringCellValue;
                            if (db.test1.Find(cellValue) == null)
                            {
                                var member = new Models.test1()
                                {
                                    account = cellValue,
                                    birth = cellValue1,
                                    sex = cellValue2
                                };
                                members.Add(member);
                            }


                        }

                        //for (int c = 0; c <= sheet.GetRow(row).LastCellNum; c++) // 使用For 走訪資料欄
                        //{
                        //    // 資料取得，等等說明
                        //}

                    }
                    if (members.Count > 0)
                    {
                        db.test1.AddRange(members);
                        db.SaveChanges();
                    }
                    return View(members);
                }
                else if (extension == ".xlsx")
                {
                    XSSFWorkbook excel;
                    var members = new List<Models.test1>();
                    // 檔案讀取
                    using (FileStream files = new FileStream(fileLocation, FileMode.Open, FileAccess.Read))
                    {
                        excel = new XSSFWorkbook(files); // 將剛剛的Excel 讀取進入到工作簿中
                    }
                    ISheet sheet = excel.GetSheetAt(0);
                    for (int row = 1; row <= sheet.LastRowNum; row++) // 使用For 走訪所有的資料列
                    {
                        if (sheet.GetRow(row) != null) // 驗證是不是空白列
                        {
                            XSSFCell cell = (XSSFCell)sheet.GetRow(row).GetCell(0);
                            cell.SetCellType(CellType.String);
                            string cellValue = cell.StringCellValue;

                            XSSFCell cell1 = (XSSFCell)sheet.GetRow(row).GetCell(1);
                            cell1.SetCellType(CellType.Numeric);
                            string cellValue1 = cell1.DateCellValue.ToString("yyyy/MM/dd");

                            XSSFCell cell2 = (XSSFCell)sheet.GetRow(row).GetCell(2);
                            cell2.SetCellType(CellType.String);
                            string cellValue2 = cell2.StringCellValue;

                            if (db.test1.Find(cellValue) == null)
                            {
                                var member = new Models.test1()
                                {
                                    account = cellValue,
                                    birth = cellValue1,
                                    sex = cellValue2
                                };
                                members.Add(member);
                            }

                            //for (int c = 0; c <= sheet.GetRow(row).LastCellNum; c++) // 使用For 走訪資料欄
                            //{
                            //    // 資料取得，等等說明
                            //}

                        }
                    }
                    if (members.Count > 0)
                    {
                        db.test1.AddRange(members);
                        db.SaveChanges();
                    }
                    return View(members);

                }
            }
            return View();
        }

        public ActionResult Download()
        {
            string[] titleList = { "帳號", "生日", "性別" };
            HSSFWorkbook book = new HSSFWorkbook();
            //新增一個sheet
            ISheet sheet1 = book.CreateSheet("Sheet1");
            //給sheet1新增第一行的頭部標題
            IRow headerrow = sheet1.CreateRow(0);

            HSSFCellStyle headStyle = (HSSFCellStyle)book.CreateCellStyle();
            HSSFFont font = (HSSFFont)book.CreateFont();
            font.FontHeightInPoints = 10;
            font.Boldweight = 700;
            headStyle.SetFont(font);
            for (int i = 0; i < 3; i++)
            {
                ICell cell = headerrow.CreateCell(i);
                cell.CellStyle = headStyle;
                cell.SetCellValue(titleList[i]);
            }
            var members = db.test1.ToList();
            //將資料逐步寫入sheet1各個行
            for(int i=0,c=members.Count;i<c;i++)
            {
                IRow rowtemp = sheet1.CreateRow(i + 1);
                rowtemp.CreateCell(0).SetCellValue(members[i].account);
                rowtemp.CreateCell(1).SetCellValue(members[i].birth);
                rowtemp.CreateCell(2).SetCellValue(members[i].sex);
            }
            // 寫入到客戶端 
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            book.Write(ms);
            ms.Seek(0, SeekOrigin.Begin);
            return File(ms, "application/vnd.ms-excel", "學員報名詳情.xls");
        }
        public ActionResult About()
        {

            return View();
        }

        public ActionResult Contact()
        {


            return View();
        }
    }
}