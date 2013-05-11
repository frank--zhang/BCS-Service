using System;
using System.Data;
using System.Web;
using System.Collections;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.ComponentModel;

using SAPbobsCOM;
using System.Collections.Generic;
using FlexService.model;

namespace FlexService
{
    /// <summary>
    /// BCSService 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [ToolboxItem(false)]

    public class BCSService : System.Web.Services.WebService
    {
        public Solution sln;

        public Recordset CreateRecordSet(SAPbobsCOM.Company oCompany)
        {
            try
            {
                return (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            }
            catch
            {
            }
            return null;
        }

        public int getIndex(string username)
        {
            for (int i = 0; i < OperConst.oCompanys.Count; i++)
            {
                SAPbobsCOM.Company thisCompany = OperConst.oCompanys[i];
                if (thisCompany.UserName == username)
                {
                    return i;
                }
            }
            return -1;
        }

        [WebMethod]
        public string HelloWorld()
        {
            /**
            List<string> list = new List<string>(5);
            list.Add("i0");
            list.Add("i1");
            list.Add("i2");
            list.Add("i3");
            list.RemoveAt(1);
            list.Add("new I1");

            string bbb = "103400030111201/13/01";
            bbb = bbb.Substring(9);
             * */

            return "Hello World";
        }

        //登陆
        [WebMethod]
        public string[] login(string username, string password)
        {
            string ifUser = "0";
            int index = this.getIndex(username);
            if (index == -1)
            {//不是已登录用户
                if (OperConst.oCompanys.Count >= OperConst.totalUser)
                {//达到最大用户数
                    ifUser = "2";

                    string[] result0 = new string[2] { "2", username };
                    return result0;
                }
            }
            //创建连接
            sln = new Solution(username, password);
            SAPbobsCOM.Company thisCompany = sln.ConnectionToDB();
            if (thisCompany != null)
            {//登陆成功
                if (thisCompany.Connected)
                {       
                    if (index == -1)
                    {
                        OperConst.oCompanys.Add(thisCompany);
                    }
                    ifUser = "1";
                }
            }

            string[] result = new string[2] { ifUser,username };
            return result;
        }

        //注销
        [WebMethod]
        public int logoutN(string username)
        {
            int index = this.getIndex(username);
            if (index != -1)
            {
                SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
                thisCompany.Disconnect();
                OperConst.oCompanys.RemoveAt(index);
            }
            return 1;
        }

        //所有颜色
        [WebMethod]
        public List<KV> getAllColors(string username)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"SELECT Code  FROM [@color] order by Code";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    string code = oRes.Fields.Item("Code").Value.ToString().Trim();
                    k.Str0 = code;
                    k.Str1 = code;
                    list.Add(k);

                    oRes.MoveNext();
                }
            }

            return list;
        }

        //供应商
        [WebMethod]
        public List<KV> getAllOCRD(string username, string cCode)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"select CardCode,CardName from OCRD where CardType='C'";
            if (!string.IsNullOrEmpty(cCode))
            {
                sql += @" and CardCode like '%" + cCode + "%'";
            }
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    k.Str0 = oRes.Fields.Item("CardCode").Value.ToString().Trim();
                    k.Str1 = oRes.Fields.Item("CardName").Value.ToString().Trim();
                    list.Add(k);

                    oRes.MoveNext();
                }
            }

            return list;
        }

        //仓库
        [WebMethod]
        public List<KV> getAllWH(string username, string whCode)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"select WhsCode,WhsName from OWHS";
            if (!string.IsNullOrEmpty(whCode))
            {
                sql += @" where WhsCode like '%" + whCode + "%'";
            }
            sql += @" order by WhsCode";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    k.Str0 = oRes.Fields.Item("WhsCode").Value.ToString().Trim();
                    k.Str1 = oRes.Fields.Item("WhsName").Value.ToString().Trim();
                    list.Add(k);

                    oRes.MoveNext();
                }
            }

            return list;
        }

        //扫描
        [WebMethod]
        public List<KV> scanItem(string username, string bcs, string combox1, string combox2, string WHT, string tPHH)
        {
            List<KV> list = new List<KV>();
            KV k = new KV();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_ColorNo='" + combox2 + "'"
                                + " and right(t1.ItemCode,1)='" + combox1 + "'"
                                + " and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and $[OWTR.U_Asi]='998' and '" + combox1 + "'='B'"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and T1.U_ColorNo='" + combox2 + "' and $[OWTR.U_Asi]<>'998' and '" + combox1 + "'='B'"
                                + " union all"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and T1.U_ColorNo='" + combox2 + "'and ('" + combox1 + "'<>'Z' or '" + combox1 + "'<>'A' or '" + combox1 + "'<>'B')"
                                + " UNION ALL"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo='" + bcs + "' and T1.U_ColorNo='" + combox2 + "'"
                                + " and right(t1.ItemCode,1)='" + combox1 + "'"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and $[OWTR.U_Asi]='998'"
                                + " and (substring('" + bcs + "',9,1) NOT IN ('L','R') )and '" + combox1 + "'='B'"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6)"
                                + " and T1.U_SizeNo=substring('" + bcs + "',7,2)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "'"
                                + " and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and $[OWTR.U_Asi]='998'"
                                + " and (substring('" + bcs + "',9,1)  IN ('L','R') )"
                                + " and '" + combox1 + "'='B'";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();

                k.Str0 = bcs;
                k.Str1 = oRes.Fields.Item("ItemCode").Value.ToString().Trim();
                k.Str2 = oRes.Fields.Item("ItemName").Value.ToString().Trim();
                k.Str3 = "0";
                k.Str4 = WHT;
                k.Str6 = bcs.Substring(9);
                k.Str8 = tPHH;
            }
            list.Add(k);
            return list ;
        }

        //库存转储草稿
        [WebMethod]
        public string saveDraft(string username, List<KV> list, string cardCode, string date01, string date02, string WHF, string combox1, string combox2, string tTSBJ, string tPHH,string tDesc,string tNote)
        {
            string result = "";
            string success = "0";

            sln = new Solution();
            sln.Log("User: " + username + "-调用保存草稿.");

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            

            try
            {
                string ErrMsg;
                int ErrCode;
                int RetVal;
                string tempStr = null;

                SAPbobsCOM.Documents oStockTransfer1 = (SAPbobsCOM.Documents)thisCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oStockTransfer1.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                oStockTransfer1.CardCode = cardCode;
                string DocDate = date01.Substring(0, 4) + "-" + date01.Substring(4, 2) + "-" + date01.Substring(6, 2);
                oStockTransfer1.DocDate = DateTime.Parse(DocDate);
                string TaxDate = date02.Substring(0, 4) + "-" + date02.Substring(4, 2) + "-" + date02.Substring(6, 2);
                oStockTransfer1.TaxDate = DateTime.Parse(TaxDate);
                oStockTransfer1.UserFields.Fields.Item("U_FWH").Value = WHF;
                oStockTransfer1.UserFields.Fields.Item("U_Bsi").Value = combox1;
                oStockTransfer1.UserFields.Fields.Item("U_Asi").Value = combox2;
                oStockTransfer1.UserFields.Fields.Item("U_Csi").Value = tTSBJ;
                //tPHH
                oStockTransfer1.JournalMemo = tDesc;
                oStockTransfer1.Comments = tNote;

                for (int i = 0; i < list.Count; i++)
                {
                    if (i != 0)
                    {
                        oStockTransfer1.Lines.Add();
                    }
                    oStockTransfer1.Lines.SetCurrentLine(i);

                    KV k = list[i];

                    oStockTransfer1.Lines.UserFields.Fields.Item("U_SubCode").Value = k.Str0;
                    oStockTransfer1.Lines.ItemCode = k.Str1;
                    oStockTransfer1.Lines.ItemDescription = k.Str2;
                    oStockTransfer1.Lines.Quantity = double.Parse(k.Str3);
                    oStockTransfer1.Lines.WarehouseCode = k.Str4;

                    oStockTransfer1.Lines.UserFields.Fields.Item("U_CPNo1").Value = k.Str5;
                    oStockTransfer1.Lines.UserFields.Fields.Item("U_CPNo2").Value = k.Str6;
                    oStockTransfer1.Lines.UserFields.Fields.Item("U_PHNo1").Value = k.Str7;
                    oStockTransfer1.Lines.UserFields.Fields.Item("U_PHNo1").Value = k.Str8;
                }

                RetVal = oStockTransfer1.Add();
                thisCompany.GetNewObjectCode(out tempStr);
                if (RetVal != 0)
                {
                    thisCompany.GetLastError(out ErrCode, out ErrMsg);
                    //sln.Log("User: " + username + ":  [添加草稿]错误!原因：" + ErrCode + "---" + ErrMsg);
                    throw new Exception("添加错误" + ErrCode + "---" + ErrMsg);
                }
                else
                {
                    sln.Log("User: " + username + "The draft(OWTR) was added successfully. And DocEntry is:" + tempStr);
                    result = "操作成功!所生成的草稿编号(DocEntry)为: " + tempStr;
                }
            }
            catch (Exception ex)
            {
                success = "0";
                sln.Log("User: " + username + "-[添加草稿]错误!原因：" + ex.ToString());
                result = "操作失败!原因：" + ex.Message.ToString();
            }

            return success + "&" + result;
        }
    }
}
