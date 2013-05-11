using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Collections.Generic;

namespace FlexService
{
    public class OperConst
    {
        //public static Solution sln;
        //public static SAPbobsCOM.Company oCompany1;
        public static int totalUser = 5;
        public static List<SAPbobsCOM.Company> oCompanys = new List<SAPbobsCOM.Company>(totalUser);
    }
}
