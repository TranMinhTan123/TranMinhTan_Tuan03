using System.Web;
using System.Web.Mvc;

namespace BTH_QLDanhSachDeTai_TranMinhTan
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}