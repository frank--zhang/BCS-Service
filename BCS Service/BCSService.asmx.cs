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

            return "Hello World,I am Frank。Frank is here";
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
            string sql = @"SELECT Code,Name FROM [@color] order by Code";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    string code = oRes.Fields.Item("Code").Value.ToString().Trim();
                    string name = oRes.Fields.Item("Name").Value.ToString().Trim();
                    k.Str0 = code;
                    k.Str1 = code;
                    k.Str2 = name;
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

        [WebMethod]
        public List<KV> getAllBZBJ(string username)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"SELECT Code,U_Name FROM [@BZBJ]";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    string code = oRes.Fields.Item("Code").Value.ToString().Trim();
                    string name = oRes.Fields.Item("U_Name").Value.ToString().Trim();
                    k.Str0 = code;
                    k.Str1 = name;
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
            string bcsName = bcs;
            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_ColorNo='" + combox2 + "'"
                                + " and right(t1.ItemCode,1)='" + combox1 + "'"
                                + " and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and '" + combox2 + "'='998' and '" + combox1 + "'='B'"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6) and T1.U_SizeNo=substring('" + bcs + "',7,2) and T1.U_BZBJ=substring('" + bcs + "',9,1)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and T1.U_ColorNo='" + combox2 + "' and '" + combox2 + "'<>'998' and '" + combox1 + "'='B'"
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
                                + " and right(t1.ItemCode,1)='" + combox1 + "' and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and '" + combox2 + "'='998'"
                                + " and (substring('" + bcs + "',9,1) NOT IN ('L','R') )and '" + combox1 + "'='B'"
                                + " UNION"
                                + " SELECT T1.ItemCode,T1.ItemName from OITM T1"
                                + " where t1.U_HHNo=left('" + bcs + "',6)"
                                + " and T1.U_SizeNo=substring('" + bcs + "',7,2)"
                                + " and right(t1.ItemCode,1)='" + combox1 + "'"
                                + " and (T1.U_ColorNo='998' OR T1.U_ColorNo='Z00' )and '" + combox2 + "'='998'"
                                + " and (substring('" + bcs + "',9,1)  IN ('L','R') )"
                                + " and '" + combox1 + "'='B'";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();

                k.Str0 = bcsName;
                k.Str1 = oRes.Fields.Item("ItemCode").Value.ToString().Trim();
                k.Str2 = oRes.Fields.Item("ItemName").Value.ToString().Trim();
                k.Str3 = "0";
                k.Str4 = WHT;
                k.Str6 = bcs.Substring(9);
                k.Str7 = tPHH;
                list.Add(k);
            }

            return list;
        }

        //库存转储草稿
        [WebMethod]
        public string saveDraft(string username, List<KV> list, string cardCode, string date01, string date02, string WHF, string combox1, string combox2, string tTSBJ, string tPHH, string tDesc, string tNote)
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
                //SAPbobsCOM.StockTransfer oStockTransfer = (SAPbobsCOM.StockTransfer)thisCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                SAPbobsCOM.Documents oStockTransfer1 = (SAPbobsCOM.Documents)thisCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                oStockTransfer1.DocObjectCode = SAPbobsCOM.BoObjectTypes.oStockTransfer;
                oStockTransfer1.CardCode = cardCode;
                string DocDate = date01.Substring(0, 4) + "-" + date01.Substring(5, 2) + "-" + date01.Substring(8, 2);
                oStockTransfer1.DocDate = DateTime.Parse(DocDate);
                string TaxDate = date02.Substring(0, 4) + "-" + date02.Substring(5, 2) + "-" + date02.Substring(8, 2);
                oStockTransfer1.TaxDate = DateTime.Parse(TaxDate);
                oStockTransfer1.UserFields.Fields.Item("U_FWH").Value = WHF;
                oStockTransfer1.UserFields.Fields.Item("U_Bsi").Value = combox1;
                oStockTransfer1.UserFields.Fields.Item("U_Asi").Value = combox2;
                oStockTransfer1.UserFields.Fields.Item("U_Csi").Value = tTSBJ;
                //tPHH
                oStockTransfer1.JournalMemo = tDesc;
                oStockTransfer1.Comments = tNote;

                //for (int i = 0; i < list.Count; i++)
                int j = 0;
                for (int i = list.Count - 1; i >= 0; i--)
                {
                    if (j != 0)
                    {
                        oStockTransfer1.Lines.Add();
                    }
                    oStockTransfer1.Lines.SetCurrentLine(j);

                    KV k = list[i];

                    oStockTransfer1.Lines.UserFields.Fields.Item("U_CB").Value = k.Str0;
                    oStockTransfer1.Lines.ItemCode = k.Str1;
                    oStockTransfer1.Lines.ItemDescription = k.Str2;
                    oStockTransfer1.Lines.Quantity = double.Parse(k.Str3);
                    oStockTransfer1.Lines.WarehouseCode = k.Str4;
                    if (k.Str5 != null)
                    {
                        oStockTransfer1.Lines.UserFields.Fields.Item("U_CPNo1").Value = k.Str5;
                    }
                    if (k.Str6 != null)
                    {
                        oStockTransfer1.Lines.UserFields.Fields.Item("U_CPNo2").Value = k.Str6;
                    }
                    if (k.Str7 != null)
                    {
                        oStockTransfer1.Lines.UserFields.Fields.Item("U_PHNo1").Value = k.Str7;
                    }
                    if (k.Str8 != null)
                    {
                        oStockTransfer1.Lines.UserFields.Fields.Item("U_PHNo1").Value = k.Str8;
                    }
                    j++;
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
                    success = "1";
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

        //检索销售订单行
        [WebMethod]
        public AutoDLN searchRDR(string userName, KV sea,int lineIndex)
        {
            AutoDLN ad = new AutoDLN();

            int index = this.getIndex(userName);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];

            string sql = @"select t1.OpenCreQty,t1.DocEntry,t1.LineNum,t3.ItemName,t0.DocDueDate,t1.Factor1,t1.Factor2,t1.Factor3,t1.Factor4,t1.Quantity,t1.WhsCode from RDR1 t1 left join ORDR t0"
                        + @" on t1.DocEntry=t0.DocEntry"
                        + @" left join OITM t3"
                        + @" on t1.ItemCode=t3.ItemCode"
                        + @" where t1.LineStatus='O'";
            if (!string.IsNullOrEmpty(sea.Str1))
            {
                sql += @" and t0.CardCode='" + sea.Str1 + "'";
            }
            if (!string.IsNullOrEmpty(sea.Str2))
            {
                sql += @" and t1.ItemCode='" + sea.Str2 + "'";
            }
            if (!string.IsNullOrEmpty(sea.Str4))
            {
                sql += @" and t0.DocDueDate>='" + sea.Str4 + "'";
            }
            if (!string.IsNullOrEmpty(sea.Str5))
            {
                sql += @" and t0.DocDueDate<='" + sea.Str5 + "'";
            }
            if (!string.IsNullOrEmpty(sea.Str6))
            {
                sql += @" and t0.Comments like '%" + sea.Str6 + "%'";
            }
            sql+= @" order by t0.DocDueDate";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();

                ad.LineIndex = lineIndex.ToString();
                ad.SignA = "A";
                ad.SignB = "B";
                ad.CardCode = "";
                ad.ItemCode = sea.Str2;
                ad.ItemName = "";
                ad.Sum = sea.Str3;
                //总数
                int totNum = int.Parse(sea.Str3);
                //bool needBreak = false;
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    AutoDLN1 adL = new AutoDLN1();
                    adL.LineType = "订单行";
                    adL.OrdrEntry = oRes.Fields.Item("DocEntry").Value.ToString().Trim();
                    adL.OrdrLine = oRes.Fields.Item("LineNum").Value.ToString().Trim();
                    adL.ItemName = oRes.Fields.Item("ItemName").Value.ToString().Trim();
                    adL.DDD = DateTime.Parse(oRes.Fields.Item("DocDueDate").Value.ToString().Trim()).ToString("yyyy.MM.dd");
                    adL.Fa1 = oRes.Fields.Item("Factor1").Value.ToString().Trim();
                    adL.Fa2 = oRes.Fields.Item("Factor2").Value.ToString().Trim();
                    adL.Fa3 = oRes.Fields.Item("Factor3").Value.ToString().Trim();
                    adL.Fa4 = oRes.Fields.Item("Factor4").Value.ToString().Trim();
                    adL.UnSum = oRes.Fields.Item("OpenCreQty").Value.ToString().Trim();
                    //数量处理
                    int sumInt = int.Parse(oRes.Fields.Item("Quantity").Value.ToString().Trim());
                    if (totNum >= sumInt)
                    {
                        adL.Sum = sumInt.ToString();
                        totNum -= sumInt;//减 求出需要数量
                    }
                    else
                    {
                        adL.Sum = totNum.ToString();
                        totNum = 0;
                        //needBreak = true;
                    }
                    
                    adL.WhsCode = oRes.Fields.Item("WhsCode").Value.ToString().Trim();

                    ad.Lines.Add(adL);

                    oRes.MoveNext();

                    /**
                    if (needBreak)
                    {
                        break;
                    }
                     * */
                }
                if (totNum > 0)
                {
                    AutoDLN1 adL = new AutoDLN1();
                    adL.LineType = "自动补数";
                    adL.Fa1 = "1";
                    adL.Fa2 = "1";
                    adL.Fa3 = "1";
                    adL.Fa4 = "1";
                    adL.Sum = totNum.ToString();
                    adL.UnSum = "";
                    adL.WhsCode = "";
                    ad.Lines.Add(adL);
                }

            }
            else
            {
                ad = null;
            }
            return ad;
        }

        //获取客户
        [WebMethod]
        public List<KV> getAllOCRDC(string username, KV obj)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            string sql = @"select CardCode,CardName from OCRD where CardType='C'";
            if (!string.IsNullOrEmpty(obj.Str1))
            {
                //sql += @" and CardCode like '%" + obj.Str1 + "%'";
                sql += @" and CardName like N'%" + obj.Str1 + "%'";
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


        //获取物料
        [WebMethod]
        public List<KV> getItemList(string username, KV obj)
        {
            List<KV> list = new List<KV>();

            int index = this.getIndex(username);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];
            //string sql = @"select ItemCode,ItemName,SalFactor1,SalFactor2,SalFactor3,SalFactor4,DfltWH,OnHand from OITM where 1=1";
            string sql = @"select OITM.ItemCode,ItemName,SalFactor1,SalFactor2,SalFactor3,SalFactor4,DfltWH,OITW.OnHand from OITM left join OITW on OITM.ItemCode=OITW.ItemCode and OITM.DfltWH=OITW.WhsCode where 1=1";
            if (!string.IsNullOrEmpty(obj.Str1))
            {
                //sql += @" and ItemCode like '%" + obj.Str1 + "%'";
                sql += @" and ItemName like N'%" + obj.Str1 + "%'";
            }
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                for (int i = 0; i < oRes.RecordCount; i++)
                {
                    KV k = new KV();
                    k.Str0 = oRes.Fields.Item("ItemCode").Value.ToString().Trim();
                    k.Str1 = oRes.Fields.Item("ItemName").Value.ToString().Trim();
                    k.Str6 = oRes.Fields.Item("SalFactor1").Value.ToString().Trim();
                    k.Str7 = oRes.Fields.Item("SalFactor2").Value.ToString().Trim();
                    k.Str8 = oRes.Fields.Item("SalFactor3").Value.ToString().Trim();
                    k.Str9 = oRes.Fields.Item("SalFactor4").Value.ToString().Trim();
                    k.WH = oRes.Fields.Item("DfltWH").Value.ToString().Trim();
                    k.IH = oRes.Fields.Item("OnHand").Value.ToString().Trim();
                    list.Add(k);

                    oRes.MoveNext();
                }
            }

            return list;
        }


        //创建交货单
        [WebMethod]
        public string saveODLN(string thisUser, List<AutoDLN> dlnList,string cardCode)
        {
            string result = "";
            string success = "0";


            int index = this.getIndex(thisUser);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];

            try
            {
                string ErrMsg;
                int ErrCode;
                int RetVal;
                string tempStr = null;

                //SAPbobsCOM.Documents oDeliveryNotes = (SAPbobsCOM.Documents)thisCompany.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                SAPbobsCOM.Documents oDeliveryNotes = (SAPbobsCOM.Documents)thisCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                oDeliveryNotes.DocObjectCode = SAPbobsCOM.BoObjectTypes.oDeliveryNotes;
                oDeliveryNotes.CardCode = cardCode;                   

                int totleLine = 0;
                for (int i = 0; i < dlnList.Count; i++)
                {
                    AutoDLN dln = dlnList[i];
                    for (int j = 0; j < dln.Lines.Count; j++)
                    {
                        AutoDLN1 dln1 = dln.Lines[j];
                        
                        double dlnSum = double.Parse(dln1.Sum);
                        if (dlnSum == 0)
                        {
                            continue;
                        }
                        //添加行
                        if (totleLine != 0)
                        {
                            oDeliveryNotes.Lines.Add();
                        }
                        oDeliveryNotes.Lines.SetCurrentLine(totleLine);

                        if (dln1.LineType == "订单行")
                        {
                            oDeliveryNotes.Lines.BaseType = 17;
                            oDeliveryNotes.Lines.BaseEntry = int.Parse(dln1.OrdrEntry);
                            oDeliveryNotes.Lines.BaseLine = int.Parse(dln1.OrdrLine);
                            oDeliveryNotes.Lines.UserFields.Fields.Item("U_typeMa").Value = "订单行";
                        }
                        else
                        {
                            oDeliveryNotes.Lines.ItemCode = dln.ItemCode;
                            oDeliveryNotes.Lines.UserFields.Fields.Item("U_typeMa").Value = "补足";
                        }


                        oDeliveryNotes.Lines.Factor1 = double.Parse(dln1.Fa1);
                        oDeliveryNotes.Lines.Factor2 = double.Parse(dln1.Fa2);
                        oDeliveryNotes.Lines.Factor3 = double.Parse(dln1.Fa3);
                        oDeliveryNotes.Lines.Factor4 = double.Parse(dln1.Fa4);                        
                        oDeliveryNotes.Lines.Quantity = dlnSum;
                        oDeliveryNotes.Lines.WarehouseCode = dln.WH;
                        /**
                        string tempWhs = dln1.WhsCode;
                        if (string.IsNullOrEmpty(tempWhs))
                        {
                            throw new Exception("第" + (i + 1) + "行的数据中有为空的仓库!");
                        }
                        else
                        {
                            oDeliveryNotes.Lines.WarehouseCode = tempWhs;
                        }
                         * */




                        totleLine++;
                    }
                }
                RetVal = oDeliveryNotes.Add();
                thisCompany.GetNewObjectCode(out tempStr);

                sln = new Solution();
                if (RetVal != 0)
                {
                    thisCompany.GetLastError(out ErrCode, out ErrMsg);
                    sln.Log("User: " + thisUser + ":  [保存交货单]错误!原因：" + ErrCode + "---" + ErrMsg);
                    throw new Exception(ErrCode + "---" + ErrMsg);
                    //result = "操作错误!原因：" + ErrCode + "---" + ErrMsg;
                }
                else
                {
                    success = "1";
                    sln.Log("User: " + thisUser + "The ODLN was added successfully. And DocEntry is:" + tempStr);
                    result = "操作成功!所生成的草稿编号(DocEntry)为: " + tempStr;
                }
            }
            catch (Exception ex)
            {
                success = "0";
                //sln.Log("User: " + username + "-[添加草稿]错误!原因：" + ex.ToString());
                result = "操作失败!原因：" + ex.Message.ToString();
            }

            return success + "&" + result;
        }

        //查库存
        [WebMethod]
        public string getWHStock(string thisUser, string itemCode, string whCode)
        {
            string stock = "0";

            int index = this.getIndex(thisUser);
            SAPbobsCOM.Company thisCompany = OperConst.oCompanys[index];

            string sql = @"select OnHand from OITW where ItemCode='" + itemCode + "' and WhsCode='" + whCode + "'";
            Recordset oRes = this.CreateRecordSet(thisCompany);
            oRes.DoQuery(sql);
            if (oRes.RecordCount > 0)
            {
                oRes.MoveFirst();
                stock = oRes.Fields.Item("OnHand").Value.ToString().Trim();
            }

            return stock;
        }
    }
}
