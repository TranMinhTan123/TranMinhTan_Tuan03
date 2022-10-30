using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using BTH_QLDanhSachDeTai_TranMinhTan.Models;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;
using Microsoft.Build.Tasks;
using Microsoft.Reporting.WebForms;
using CrystalDecisions.CrystalReports.Engine;
using System.Data;
using BTH_QLDanhSachDeTai_TranMinhTan.Report;

namespace BTH_QLDanhSachDeTai_TranMinhTan.Controllers
{
    public class Home2Controller : Controller
    {
        //BÀI TẬP 2
        ThucTapDataContext data = new ThucTapDataContext();
        public ActionResult Index()
        {
            List<TBLSinhVien> sv = data.TBLSinhViens.ToList();
            List<TBLGiangVien> gv = data.TBLGiangViens.ToList();
            List<TBLDeTai> dt = data.TBLDeTais.ToList();
            List<TBLHuongDan> hd = data.TBLHuongDans.ToList();
            var obj = from h in hd
                      join s in sv on h.Masv equals s.Masv
                      join g in gv on h.Magv equals g.Magv
                      join d in dt on h.Madt equals d.Madt
                      select new Class1
                      {
                          sinhvien = s,
                          giangvien = g,
                          detai = d,
                          huongdan = h
                      };
            return View(obj);
        }
        // Thêm mới danh sách hướng dẫn đề tài
        public ActionResult ThemHuongDanDeTai()
        {
            ViewBag.MADT = new SelectList(data.TBLDeTais.ToList(), "MADT", "Tendt");
            ViewBag.MASV = new SelectList(data.TBLSinhViens.ToList(), "MASV", "Hotensv");
            ViewBag.MAGV = new SelectList(data.TBLGiangViens.ToList(), "MAGV", "Hotengv");
            return View();
        }
        [HttpPost]
        public ActionResult XuLyThemHuongDanDeTai(TBLHuongDan hd, FormCollection col)
        {
            int masv = int.Parse(col["MASV"]);
            var ktTrungMa = data.TBLHuongDans.Where(t => t.Masv == masv);
            if (ktTrungMa == null)
            {
                data.TBLHuongDans.InsertOnSubmit(hd);
                data.SubmitChanges();
                return RedirectToAction("Index", "Home2");
            }
            else
            {
                ViewBag.thongbao = "Sinh viên đã tồn tại";
            }
            return View();
        }
        //Clear
        public ActionResult Clear()
        {
            return RedirectToAction("Index", "Home2");
        }
        //xóa hướng dẫn đề tài
        public ActionResult XoaHuongDanDeTai(int masv)
        {
            TBLHuongDan del = data.TBLHuongDans.SingleOrDefault(t => t.Masv == masv);
            data.TBLHuongDans.DeleteOnSubmit(del);
            data.SubmitChanges();
            return RedirectToAction("Index", "Home2");
        }

        //Nhập kết quả hướng dẫn đề tài
        public ActionResult SuaHuongDanDeTai(int masv)
        {
            TBLHuongDan hd = data.TBLHuongDans.SingleOrDefault(t => t.Masv == masv);
            return View(hd);
        }
        [HttpPost]
        public ActionResult XyLySuaHuongDanDeTai(FormCollection col)
        {
            int masv = int.Parse(col["Masv"]);
            decimal ketqua = decimal.Parse(col["KetQua"]);
            if (ketqua != null)
            {
                TBLHuongDan hd = data.TBLHuongDans.SingleOrDefault(t => t.Masv == masv);
                hd.KetQua = ketqua;
                data.SubmitChanges();
            }
            //List<TBLDeTai> sdt = data.TBLDeTais.Where(t => t.Madt != null).ToList();
            return RedirectToAction("Index", "Home2");
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
                string detai = col["txtDeTai"];
                string giangvien = col["txtGiangVien"];
                string sinhvien = col["txtSinhVien"];
                string kd = col["cbxKetQua"];
                if (col["cbxKetQua"] == "2")
                {
                    List<TBLSinhVien> sv = data.TBLSinhViens.ToList();
                    List<TBLGiangVien> gv = data.TBLGiangViens.ToList();
                    List<TBLDeTai> dt = data.TBLDeTais.ToList();
                    List<TBLHuongDan> hd = data.TBLHuongDans.ToList();
                    var obj = from h in hd
                              join s in sv on h.Masv equals s.Masv
                              join g in gv on h.Magv equals g.Magv
                              join d in dt on h.Madt equals d.Madt
                              where d.Tendt == detai && s.Hotensv == sinhvien && g.Hotengv == giangvien && h.KetQua == null
                              select new Class1
                              {
                                  sinhvien = s,
                                  giangvien = g,
                                  detai = d,
                                  huongdan = h
                              };
                    return View("Index", obj);
                }
                else
                {
                    List<TBLSinhVien> sv = data.TBLSinhViens.ToList();
                    List<TBLGiangVien> gv = data.TBLGiangViens.ToList();
                    List<TBLDeTai> dt = data.TBLDeTais.ToList();
                    List<TBLHuongDan> hd = data.TBLHuongDans.ToList();
                    var obj = from h in hd
                              join s in sv on h.Masv equals s.Masv
                              join g in gv on h.Magv equals g.Magv
                              join d in dt on h.Madt equals d.Madt
                              where d.Tendt == detai && s.Hotensv == sinhvien && g.Hotengv == giangvien && h.KetQua != null
                              select new Class1
                              {
                                  sinhvien = s,
                                  giangvien = g,
                                  detai = d,
                                  huongdan = h
                              };
                    return View("Index", obj);
                }

            }
            catch
            {
                ViewBag.thongbao = "Chưa nhập đủ thông tin hoặc sai định dạng";
            }
            return View();
        }
        //Xuất excel 
        public ActionResult XuatExcel()
        {

            var sv = this.data.TBLSinhViens.ToList();
            var gv = this.data.TBLGiangViens.ToList();
            var dt = this.data.TBLDeTais.ToList();
            var hd = this.data.TBLHuongDans.ToList();

            var grid = new GridView();
            grid.DataSource = from h in hd
                              join s in sv on h.Masv equals s.Masv
                              join g in gv on h.Magv equals g.Magv
                              join d in dt on h.Madt equals d.Madt
                              select new
                              {
                                  makhoa = s.Makhoa,
                                  masv = s.Masv,
                                  hotensv = s.Hotensv,
                                  magv = g.Magv,
                                  hotengv = g.Magv,
                                  madt = d.Madt,
                                  tendt = d.Tendt,
                                  ketqua = h.KetQua
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

        public ActionResult Report()
        {

            var sv = this.data.TBLSinhViens.ToList();
            var gv = this.data.TBLGiangViens.ToList();
            var dt = this.data.TBLDeTais.ToList();
            var hd = this.data.TBLHuongDans.ToList();

            ReportDocument rd = new ReportDocument();
            rd.Load(Path.Combine(Server.MapPath("~/Report"), "CrystalReport1.rpt"));
            //rd.SetDataSource(data.TBLHuongDans.Select(t => new
            //{
            //    Masv = t.Masv,
            //    Madt = t.Madt,
            //    Magv = t.Magv,
            //    KetQua = t.KetQua
            //}).ToList());

            DataTable dts = new DataTable();
            dts.Clear();
            dts.Columns.Add("Masv");
            dts.Columns.Add("Madt");
            dts.Columns.Add("Magv");
            dts.Columns.Add("KetQua");

            dts.Rows.Add(new object[] { "ms1", "madt", 6, 6 });

            rd.SetDataSource(dts);

            Response.Buffer = false;
            Response.ClearContent();
            Response.ClearHeaders();


            rd.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape;
            rd.PrintOptions.ApplyPageMargins(new CrystalDecisions.Shared.PageMargins(5, 5, 5, 5));
            rd.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.PaperA5;

            Stream stream = rd.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
            stream.Seek(0, SeekOrigin.Begin);

            return File(stream, "application/pdf", "CustomerList.pdf");
        }
    }
}
