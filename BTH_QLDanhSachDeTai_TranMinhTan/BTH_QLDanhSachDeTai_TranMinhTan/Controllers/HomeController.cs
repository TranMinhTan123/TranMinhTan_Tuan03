using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BTH_QLDanhSachDeTai_TranMinhTan.Models;
using System.Data;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

namespace BTH_QLDanhSachDeTai_TranMinhTan.Controllers
{
    public class HomeController : Controller
    {

        //BÀI TẬP 1

        ThucTapDataContext data = new ThucTapDataContext();
        public ActionResult Index()
        {
            List<TBLDeTai> dsdt = data.TBLDeTais.ToList();
            return View(dsdt);
        }
        // Thêm đề tài
        public ActionResult ThemDeTai()
        {
            return View();
        }
        [HttpPost]
        public ActionResult XuLyThemDeTai(TBLDeTai dt,FormCollection col)
        {
            string madt = col["Madt"];
            var tendt = col["Tendt"];
            if(madt != "" && tendt != "")
            {
                try
                {
                    data.TBLDeTais.InsertOnSubmit(dt);
                    data.SubmitChanges();
                    return RedirectToAction("Index", "Home");
                }
                catch
                {
                    ViewBag.thongbao = "Mã đề tài đã tồn tại";
                }
            }
            else
            {
                ViewBag.thongbao = "Thông tin chưa đầy đủ !";
            }
            return View();
        }

        //Xóa đề tài
        public ActionResult XoaDeTai(string madt)
        {
            var ktTrungMa = data.TBLHuongDans.Where(t => t.Madt == madt);
            if (ktTrungMa==null)
            {
                TBLDeTai del = data.TBLDeTais.SingleOrDefault(t => t.Madt == madt);
                if (del != null)
                {
                    data.TBLDeTais.DeleteOnSubmit(del);
                    data.SubmitChanges();
                    return RedirectToAction("Index", "Home");
                }
            }
            else 
            {
                ViewBag.thongbao = "Không thể xóa dữ liệu đã sử dụng ở table hướng dẫn";
            }
            return View();
        }
        //Sửa đề tài
        public ActionResult SuaDeTai(string madt)
        {
            TBLDeTai dt = data.TBLDeTais.SingleOrDefault(t => t.Madt == madt);
            return View(dt);
        }
         [HttpPost]
        public ActionResult XyLySuaDeTai(FormCollection col)
        {
            string madt = col["Madt"];
            var tendt = col["Tendt"];
            int kinhphi = int.Parse(col["Kinhphi"]);
            var noithuctap = col["Noithuctap"];
            if (tendt != null)
            {
                TBLDeTai dt = data.TBLDeTais.SingleOrDefault(t => t.Madt == madt);
                dt.Tendt = tendt;
                dt.Kinhphi = kinhphi;
                dt.Noithuctap = noithuctap;

                data.SubmitChanges();
            }
            List<TBLDeTai> sdt = data.TBLDeTais.Where(t => t.Madt != null).ToList();
            return RedirectToAction("Index", "Home", sdt);
        }
        //tìm kiếm
         public ActionResult TimKiem()
         {        
             return View();
         }
        [HttpPost]
         public ActionResult XuLyTimKiem(FormCollection col)
         {
             try
             {
                 string tendt = col["txtTenDT"];
                 string Noithuctap = col["txtNoiTT"];
                 int bd = int.Parse(col["txtBatDau"]);
                 int kt = int.Parse(col["txtKetThuc"]);
                 List<TBLDeTai> dsdt = data.TBLDeTais.Where(t => t.Tendt.Contains(tendt) == true && t.Kinhphi >= bd && t.Kinhphi <= kt && t.Noithuctap.Contains(Noithuctap) == true).ToList();
                 return View("Index", dsdt);
             }
             catch
             {
                 ViewBag.thongbao = "Chưa nhập đủ thông tin hoặc sai định dạng";
             }
             return View();
         }
        //Clear
        public ActionResult Clear()
        {
            return RedirectToAction("Index", "Home");
        }
        //Xuất excel
        public ActionResult XuatExcel()
        {

            var detai = this.data.TBLDeTais.ToList();

            var grid = new GridView();
            grid.DataSource = from p in detai
                              select new
                              {
                                  madt = p.Madt,
                                  tendt = p.Tendt,
                                  kinhphi = p.Kinhphi,
                                  noithuctap = p.Noithuctap
                              };
            grid.DataBind();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=MyExcelFile.xls");
            Response.ContentType = "application/ms-excel";

            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);

            grid.RenderControl(htw);

            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();

            return View(); 
        }

    }
}
