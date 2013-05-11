using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

using SAPbobsCOM;
using System.Xml;

namespace FlexService
{
    public class Solution
    {
        string SetupPath = ConfigurationSettings.AppSettings["SetupPath"];
        public Solution()
        {
            try
            {
                DEFAULT_LOG_FILE_NAME = "BCSS.log";
                // PATH_SEPRATOR = @"\";
                //LogPath = System.Windows.Forms.Application.UserAppDataPath + @"\" + DateTime.Now.ToString("yyyy.MM.dd.") + Solution.DEFAULT_LOG_FILE_NAME;
                LogPath = SetupPath + @"\log\" + DateTime.Now.ToString("yyyy.MM.dd.") + Solution.DEFAULT_LOG_FILE_NAME;
            }
            catch (Exception ex)
            {
                Log(ex.Message);
            }
        }

        public Solution(string B1UserName, string B1PassWord):this()
        {
            try
            {
                GetDataAccess();
                //  OperConst.oCompany1=ConnectionToDB("(local)", SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005, true, "", "", "manager", "manager", "SBODemoCN");
                B1UserName1 = B1UserName;
                B1PassWord1 = B1PassWord;
                //OperConst.oCompany1 = ConnectionToDB(Server1, DBType, UserTrusted1, DBUserName1, DBPassWord1, B1UserName, B1PassWord, DBName1);
            }
            catch (Exception ex)
            {
                Log(B1UserName + "-" + ex.Message);
            }
        }

        private static string sErrMsg = null;
        private static int lErrCode = 0;
        private static int lRetCode;
        private static string DEFAULT_LOG_FILE_NAME;
        private static string LogPath;


        private string Server1 = "";
        private bool UserTrusted1;
        private string DBType = "";
        private string DBName1 = "";
        private string B1UserName1 = "";
        private string B1PassWord1 = "";
        private string DBUserName1 = "";
        private string DBPassWord1 = "";


        public SAPbobsCOM.Company ConnectionToDB()
        {
            SAPbobsCOM.Company oCompany;
            oCompany = new SAPbobsCOM.Company();
            oCompany.Server = Server1;
            //Language
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Chinese;
            if (DBType == "2005")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
            }
            else
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            }
            // Use Windows authentication for database server.
            // True for NT server authentication,
            // False for database server authentication.
            oCompany.UseTrusted = UserTrusted1;
            if (!UserTrusted1)
            {
                oCompany.DbUserName = DBUserName1;
                oCompany.DbPassword = DBPassWord1;
            }
            oCompany.UserName = B1UserName1;
            oCompany.Password = B1PassWord1;
            oCompany.CompanyDB = DBName1;
            lRetCode = oCompany.Connect();
            if (lRetCode == 0)//连接成功
            {
                Log("公司连接成功!" + oCompany.CompanyName + "-" + oCompany.UserName);//仅在调试时使用
                return oCompany;
            }
            else
            {
                oCompany.GetLastError(out lErrCode, out sErrMsg);
                Log("公司连接失败，原因为：" + sErrMsg);//仅在调试时使用
                return null;
            }
        }
/**
        public SAPbobsCOM.Company ConnectionToDB(String ServerName, String DBType, bool UseTrusted, String DBUserName, String DBPassword, String UserName, String Password, String CompanyDB)
        {
            SAPbobsCOM.Company oCompany;
            oCompany = new SAPbobsCOM.Company();
            oCompany.Server = ServerName;
            //Language
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Chinese;
            if (DBType == "2005")
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
            }
            else
            {
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
            }
            // Use Windows authentication for database server.
            // True for NT server authentication,
            // False for database server authentication.
            oCompany.UseTrusted = UseTrusted;
            if (!UseTrusted)
            {
                oCompany.DbUserName = DBUserName;
                oCompany.DbPassword = DBPassword;
            }
            oCompany.UserName = UserName;
            oCompany.Password = Password;
            oCompany.CompanyDB = CompanyDB;
            lRetCode = oCompany.Connect();
            if (lRetCode == 0)//连接成功
            {
                Log("公司连接成功!" + oCompany.CompanyName + "-" + oCompany.UserName);//仅在调试时使用
                return oCompany;
            }
            else
            {
                oCompany.GetLastError(out lErrCode, out sErrMsg);
                Log("公司连接失败，原因为：" + sErrMsg);//仅在调试时使用
                return null;
            }
        }
 * */

        public void Log(string sMessage)
        {

            try
            {
                System.IO.StreamWriter sw = System.IO.File.AppendText(LogPath);
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss: ") + sMessage);
                sw.Close();
            }
            catch
            { }

        }


        public void GetDataAccess()
        {
            XmlDocument Doc = new XmlDocument();
            Doc.Load(SetupPath + @"\DataAccess.xml");
            //Doc.Load(@"E:\DataAccess.xml");
            XmlNode CompNode;

            CompNode = Doc.SelectSingleNode("//Company1");
            foreach (XmlNode Node in CompNode.ChildNodes)
            {

                switch (Node.LocalName)
                {
                    case "Server":
                        Server1 = Node.InnerText;
                        break;
                    case "DBName":
                        DBName1 = Node.InnerText;
                        break;
                    //case "B1UserName":
                    //    B1UserName1 = Node.InnerText;
                    //    break;
                    //case "B1PassWord":
                    //    B1PassWord1 = Node.InnerText;
                    //    break;
                    case "UserTrusted":
                        UserTrusted1 = bool.Parse(Node.InnerText);
                        break;
                    case "DBUserName":
                        DBUserName1 = Node.InnerText;
                        break;
                    case "DBPassWord":
                        DBPassWord1 = Node.InnerText;
                        break;
                    case "DBType":
                        DBType = Node.InnerText;
                        break;
                }

            }
        }
    }
}
