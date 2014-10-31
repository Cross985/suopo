﻿using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.ApplicationBlocks.Data;
using System.Text.RegularExpressions;

namespace WebApplication1
{
    public partial class _Default : System.Web.UI.Page
    {
        string strConn = ConfigurationManager.ConnectionStrings["connectionString"].ToString(); //链接SQL数据库
        string importtype = string.Empty;
        

        public DataSet ExecleDs(string filenameurl, string table)
        {
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + filenameurl + ";Extended Properties='Excel 12.0 Xml;HDR=YES'";
            //string strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + filenameurl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter odda = new OleDbDataAdapter("select * from [Sheet1$]", conn);
            odda.Fill(ds, table);
            return ds;
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile == false)//HasFile用来检查FileUpload是否有指定文件
            {
                Response.Write("<script>alert('Please select Excel files!')</script> ");
                return;//当无文件时,返回
            }
            string IsXls = System.IO.Path.GetExtension(FileUpload1.FileName).ToString().ToLower();//System.IO.Path.GetExtension获得文件的扩展名
            if (IsXls != ".xls" && IsXls != ".xlsx")
            {
                Response.Write("<script>alert('只可以选择Excel文件')</script>");
                return;//当选择的不是Excel文件时,返回
            }
            try
            {
                SqlConnection cn = new SqlConnection(strConn);
                cn.Open();
                string filename = DateTime.Now.ToString("yyyymmddhhMMss") + FileUpload1.FileName;              //获取Execle文件名  DateTime日期函数
                string savePath = Server.MapPath(("~\\upfiles\\") + filename);//Server.MapPath 获得虚拟服务器相对路径
                FileUpload1.SaveAs(savePath);                        //SaveAs 将上传的文件内容保存在服务器上
                DataSet ds = ExecleDs(savePath, filename);           //调用自定义方法
                DataRow[] rowArray = ds.Tables[0].Select();            //定义一个DataRow数组
                int rowsnum = ds.Tables[0].Rows.Count;
                if (rowsnum == 0)
                {
                    Response.Write("<script>alert('这个Excel文档是空的!')</script>");   //当Excel表为空时,对用户进行提示
                }
                else
                {
                    int insertcount = 0, updatecount = 0;
                    int errorcount = 0; string errormessage = string.Empty;
                    importtype = Request.QueryString.Get("importtype");
                    string serialnum = DateTime.Now.ToShortDateString() + DateTime.Now.ToLongTimeString();
                    if (importtype == "1")
                    {
                        #region 导入客户
                        for (int i = 0; i < rowArray.Length; i++)
                        {
                            string ownername = string.Empty;
                            string ownergender = string.Empty;
                            string ownerphone = string.Empty;
                            string ownerqq = string.Empty;
                            string owneremail = string.Empty;
                            string owneraddress = string.Empty;
                            string contacterlname = string.Empty;
                            string contacterlphone = string.Empty;
                            string contacterlqq = string.Empty;
                            string contacterlemail = string.Empty;
                            string contacter2name = string.Empty;
                            string contacter2phone = string.Empty;
                            string contacter2qq = string.Empty;
                            string contacter2email = string.Empty;
                            string nextfollowdate = string.Empty;
                            string followstatus = string.Empty;
                            string priority = string.Empty;
                            string admit = string.Empty;
                            string signupid = string.Empty;
                            string ownerattitude = string.Empty;
                            string ifaddqq = string.Empty;
                            string wangwangid = string.Empty;
                            string comment = string.Empty;
                            string storename = string.Empty;
                            string followsaler = string.Empty;
                            //string temp = string.Empty;
                            string taobaourl = string.Empty;
                            string loanpurpose = string.Empty;
                            string customerstatus = string.Empty;
                            string companyname = string.Empty;
                            string corporation = string.Empty;
                            string basiclimit = string.Empty;
                            string templimit = string.Empty;
                            string quicklimit = string.Empty;
                            string jhslimit = string.Empty;
                            string dealcount = string.Empty;
                            string dealamount = string.Empty;
                            string lastdealdate = string.Empty;
                            string lastloanlimit = string.Empty;
                            string scale = string.Empty;
                            string followcount = string.Empty;
                            string level = string.Empty;
                            string comment1 = string.Empty;

                            string updatestr = string.Empty;

                            if (ds.Tables[0].Columns.Contains("规模"))
                            {
                                scale = rowArray[i]["规模"].ToString();
                                switch (scale.Trim())
                                {
                                    case "大": scale = "1"; break;
                                    case "中": scale = "2"; break;
                                    case "小": scale = "3"; break;
                                    default: scale = ""; break;
                                }
                                if (!string.IsNullOrEmpty(scale))
                                {
                                    updatestr = updatestr + " comp_scale='" + scale + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("备注1"))
                            {
                                comment1 = rowArray[i]["备注1"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("客户状态"))
                            {
                                customerstatus = rowArray[i]["客户状态"].ToString();
                                if (!string.IsNullOrEmpty(customerstatus))
                                {
                                    updatestr = updatestr + " comp_customerstatus='" + customerstatus + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("客户级别"))
                            {
                                level = rowArray[i]["客户级别"].ToString();
                                if (!string.IsNullOrEmpty(level))
                                {
                                    switch (level)
                                    {
                                        case "普通客户":
                                            level = "normal";
                                            break;
                                        case "重要客户":
                                            level = "important";
                                            break;
                                        case "X客户":
                                            level = "unimportant";
                                            break;
                                        default:
                                            level = "nomarl";
                                            break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(level))
                                {
                                    updatestr = updatestr + " comp_customerpriority='" + level + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("常规额度"))
                            {
                                basiclimit = rowArray[i]["常规额度"].ToString();
                                if (!string.IsNullOrEmpty(basiclimit))
                                {
                                    updatestr = updatestr + " comp_basiclimit='" + basiclimit + "',";
                                }
                                else
                                    basiclimit = "0";

                            }
                            if (ds.Tables[0].Columns.Contains("临时额度"))
                            {
                                templimit = rowArray[i]["临时额度"].ToString();
                                if (!string.IsNullOrEmpty(templimit))
                                {
                                    updatestr = updatestr + " comp_templimit='" + templimit + "',";
                                }
                                else
                                    templimit = "0";

                            }
                            if (ds.Tables[0].Columns.Contains("聚划算临时额度"))
                            {
                                jhslimit = rowArray[i]["聚划算临时额度"].ToString();
                                if (!string.IsNullOrEmpty(jhslimit))
                                {
                                    updatestr = updatestr + " comp_jhstemplimit='" + jhslimit + "',";
                                }
                                else
                                    jhslimit = "0";
                            }
                            if (ds.Tables[0].Columns.Contains("快审临时额度"))
                            {
                                quicklimit = rowArray[i]["快审临时额度"].ToString();
                                if (!string.IsNullOrEmpty(quicklimit))
                                {
                                    updatestr = updatestr + " comp_quicktemplimit='" + quicklimit + "',";
                                }
                                else
                                    quicklimit = "0";
                            }
                            if (ds.Tables[0].Columns.Contains("成交次数"))
                            {
                                dealcount = rowArray[i]["成交次数"].ToString();
                                if (!string.IsNullOrEmpty(dealcount))
                                {
                                    updatestr = updatestr + " comp_dealcount='" + dealcount + "',";
                                }
                                else
                                    dealcount = "0";
                            }
                            if (ds.Tables[0].Columns.Contains("累计成交金额"))
                            {
                                dealamount = rowArray[i]["累计成交金额"].ToString();
                                if (!string.IsNullOrEmpty(dealamount))
                                {
                                    updatestr = updatestr + " comp_totaldealamount='" + dealamount + "',";
                                }
                                else
                                    dealamount = "0";
                            }
                            if (ds.Tables[0].Columns.Contains("上次成交日期"))
                            {
                                lastdealdate = rowArray[i]["上次成交日期"].ToString();
                                if (!string.IsNullOrEmpty(lastdealdate))
                                {
                                    updatestr = updatestr + " comp_lastdealdate='" + lastdealdate + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("上次借款期限"))
                            {
                                lastloanlimit = rowArray[i]["上次借款期限"].ToString();
                                if (!string.IsNullOrEmpty(lastloanlimit))
                                {
                                    updatestr = updatestr + " comp_lastloanlimit='" + lastloanlimit + "',";
                                }
                                else
                                {
                                    lastloanlimit = "0";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("企业名称"))
                            {
                                companyname = rowArray[i]["企业名称"].ToString();
                                if (!string.IsNullOrEmpty(companyname))
                                {
                                    updatestr = updatestr + " comp_compname='" + companyname + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("企业法人"))
                            {
                                corporation = rowArray[i]["企业法人"].ToString();
                                if (!string.IsNullOrEmpty(corporation))
                                {
                                    updatestr = updatestr + " comp_corporation='" + corporation + "',";
                                }
                            }

                            if (ds.Tables[0].Columns.Contains("贷款意向"))
                            {
                                loanpurpose = rowArray[i]["贷款意向"].ToString();
                                if (loanpurpose.ToLower() == "y" || loanpurpose.Trim() == "是")
                                    loanpurpose = "Y";
                                else loanpurpose = null;
                                if (!string.IsNullOrEmpty(loanpurpose))
                                {
                                    updatestr = updatestr + " comp_loanpurpose='" + loanpurpose + "',";
                                }
                            }//comp_storetype
                            if (ds.Tables[0].Columns.Contains("店主姓名"))
                            {
                                ownername = rowArray[i]["店主姓名"].ToString();
                                if (!string.IsNullOrEmpty(ownername))
                                {
                                    updatestr = updatestr + " comp_name='" + ownername + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("店主性别"))
                            {
                                ownergender = rowArray[i]["店主性别"].ToString();
                                if (!string.IsNullOrEmpty(ownergender))
                                {
                                    updatestr = updatestr + " comp_ownergender='" + ownergender + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系电话"))
                            {
                                ownerphone = rowArray[i]["联系电话"].ToString();
                                if (!string.IsNullOrEmpty(ownerphone))
                                {
                                    updatestr = updatestr + " comp_ownerphone='" + ownerphone + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("店主QQ"))
                            {
                                ownerqq = rowArray[i]["店主QQ"].ToString();
                                if (!string.IsNullOrEmpty(ownerqq))
                                {
                                    updatestr = updatestr + " comp_ownerqq='" + ownerqq + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("邮箱地址"))
                            {
                                owneremail = rowArray[i]["邮箱地址"].ToString();
                                if (!string.IsNullOrEmpty(owneremail))
                                {
                                    updatestr = updatestr + " comp_owneremail='" + owneremail + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("店主联系地址"))
                            {
                                owneraddress = rowArray[i]["店主联系地址"].ToString();
                                if (!string.IsNullOrEmpty(owneraddress))
                                {
                                    updatestr = updatestr + " comp_owneraddress='" + owneraddress + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1"))
                            {
                                contacterlname = rowArray[i]["联系人1"].ToString();
                                if (!string.IsNullOrEmpty(contacterlname))
                                {
                                    updatestr = updatestr + " comp_contacterlname='" + contacterlname + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1手机号"))
                            {
                                contacterlphone = rowArray[i]["联系人1手机号"].ToString();
                                if (!string.IsNullOrEmpty(contacterlphone))
                                {
                                    updatestr = updatestr + " comp_contacterlphone='" + contacterlphone + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1QQ"))
                            {
                                contacterlqq = rowArray[i]["联系人1QQ"].ToString();
                                if (!string.IsNullOrEmpty(contacterlqq))
                                {
                                    updatestr = updatestr + " comp_contacterlqq='" + contacterlqq + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1Email"))
                            {
                                contacterlemail = rowArray[i]["联系人1Email"].ToString();
                                if (!string.IsNullOrEmpty(contacterlemail))
                                {
                                    updatestr = updatestr + " comp_contacterlemail='" + contacterlemail + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2"))
                            {
                                contacter2name = rowArray[i]["联系人2"].ToString();
                                if (!string.IsNullOrEmpty(contacter2name))
                                {
                                    updatestr = updatestr + " comp_contacter2name='" + contacter2name + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2手机号"))
                            {
                                contacter2phone = rowArray[i]["联系人2手机号"].ToString();
                                if (!string.IsNullOrEmpty(contacter2phone))
                                {
                                    updatestr = updatestr + " comp_contacter2phone='" + contacter2phone + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2QQ"))
                            {
                                contacter2qq = rowArray[i]["联系人2QQ"].ToString();
                                if (!string.IsNullOrEmpty(contacter2qq))
                                {
                                    updatestr = updatestr + " comp_contacter2qq='" + contacter2qq + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2Email"))
                            {
                                contacter2email = rowArray[i]["联系人2Email"].ToString();
                                if (!string.IsNullOrEmpty(contacter2email))
                                {
                                    updatestr = updatestr + " comp_contacter2email='" + contacter2email + "',";
                                }
                            }

                            if (ds.Tables[0].Columns.Contains("当前联系阶段"))
                            {
                                followstatus = rowArray[i]["当前联系阶段"].ToString();
                                if (!string.IsNullOrEmpty(followstatus))
                                {
                                    updatestr = updatestr + " comp_followstatus='" + followstatus + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("优先级"))
                            {
                                priority = rowArray[i]["优先级"].ToString();
                                if (!string.IsNullOrEmpty(priority))
                                {
                                    updatestr = updatestr + " comp_priority='" + priority + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("准入"))
                            {
                                admit = rowArray[i]["准入"].ToString();
                                if (!string.IsNullOrEmpty(admit))
                                {
                                    updatestr = updatestr + " comp_admit='" + admit + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("已注册ID号"))
                            {
                                signupid = rowArray[i]["已注册ID号"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("掌柜态度"))
                            {
                                ownerattitude = rowArray[i]["掌柜态度"].ToString();
                                if (!string.IsNullOrEmpty(ownerattitude))
                                {
                                    updatestr = updatestr + " comp_ownerattitude='" + ownerattitude + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("是否添加企业QQ"))
                            {
                                ifaddqq = rowArray[i]["是否添加企业QQ"].ToString();
                                if (!string.IsNullOrEmpty(ifaddqq))
                                {
                                    updatestr = updatestr + " comp_addenterpriseqq='" + ifaddqq + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("淘宝旺旺"))
                            {
                                wangwangid = rowArray[i]["淘宝旺旺"].ToString();
                                if (!string.IsNullOrEmpty(wangwangid))
                                {
                                    updatestr = updatestr + " comp_wangwangid='" + wangwangid + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("备注"))
                            {
                                comment = rowArray[i]["备注"].ToString();
                            }
                            string usersecterr = string.Empty;
                            if (ds.Tables[0].Columns.Contains("跟进销售"))
                            {
                                followsaler = rowArray[i]["跟进销售"].ToString();
                                if (!string.IsNullOrEmpty(followsaler))
                                {
                                    string GetSalesSql = "select user_userid,User_PrimaryTerritory from users where User_deleted is null and User_LastName='" + followsaler + "'";
                                    SqlDataReader RDUser = SqlHelper.ExecuteReader(cn, CommandType.Text, GetSalesSql);
                                    if (RDUser.Read())
                                    {
                                        followsaler = RDUser["user_userid"].ToString();
                                        usersecterr = RDUser["User_PrimaryTerritory"].ToString();
                                    }
                                    RDUser.Close();
                                    updatestr = updatestr + " comp_primaryuserid='" + followsaler + "',comp_createdby='" + followsaler + "',comp_secterr='" + usersecterr + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("店铺名称"))
                            {
                                storename = rowArray[i]["店铺名称"].ToString();
                                if (!string.IsNullOrEmpty(storename))
                                {
                                    updatestr = updatestr + " comp_storename='" + storename + "',";
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("淘宝URL"))
                            {
                                taobaourl = rowArray[i]["淘宝URL"].ToString();
                                if (!string.IsNullOrEmpty(taobaourl) && taobaourl.Contains("http"))
                                {
                                    if (!taobaourl.Contains("rate."))
                                    {
                                        string temp = taobaourl.Substring(0, taobaourl.IndexOf(".com") + 4);
                                        taobaourl = temp;
                                    }
                                    else
                                    {
                                        string temp = taobaourl.Substring(0, taobaourl.IndexOf(".htm") + 4);
                                        taobaourl = temp;
                                    }
                                }
                                if (!string.IsNullOrEmpty(taobaourl))
                                {
                                    updatestr = updatestr + " comp_storeurl='" + taobaourl + "',";
                                }
                            }//comp_storeurl
                            comment = comment + comment1;
                            if (!string.IsNullOrEmpty(comment))
                            {
                                updatestr = updatestr + " comp_comment='" + comment + "',";
                            }

                            string checksql = "select comp_companyid from company where comp_deleted is null and comp_signupid='{0}'";
                            string countsql = "select comp_followcount from company where comp_deleted is null and comp_signupid='{0}'";
                            SqlDataReader flcount = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(countsql, signupid));
                            if (flcount.Read())
                            {
                                followcount = flcount["comp_followcount"].ToString();

                            }
                            flcount.Close();
                            string insertsql = @"insert into company (comp_name,comp_ownergender,comp_ownerqq,comp_ownerphone,comp_owneremail 
                         ,comp_owneraddress,comp_contacter1name,comp_contacter1phone,comp_contacter1qq,comp_contacter1email,comp_contacter2name
                          ,comp_contacter2phone,comp_contacter2qq,comp_contacter2email,comp_storeid,comp_followstatus
                         ,comp_priority,comp_admit,comp_signupid,comp_ownerattitude,comp_addenterpriseqq,comp_wangwangid,comp_comment,comp_primaryuserid,comp_storename,comp_customertype,comp_companyid,comp_secterr,comp_createdby,comp_assignstatus,comp_serialnum
                         ,comp_storeurl,comp_loanpurpose,comp_customerstatus,comp_scale,comp_compname,comp_corporation,comp_basiclimit,comp_templimit,comp_quicktemplimit,comp_jhstemplimit,comp_dealcount,comp_totaldealamount,comp_lastdealdate,comp_lastloanlimit,comp_flag,comp_customerpriority) values ('{0}',
                         '{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','2','{25}','{26}','{27}','2','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}',
                         '{36}','{37}','{38}','{39}','{40}','{41}','{42}','1','{43}')";

                            SqlDataReader reader = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(checksql, signupid));

                            if (reader.Read())
                            {
                                string companyid = reader[0].ToString();
                                string companyupdatesql = string.Empty;
                                reader.Close();
                                if ((string.IsNullOrEmpty(followcount) || followcount == "0") && !string.IsNullOrEmpty(updatestr))
                                {
                                    companyupdatesql = "update company set " + updatestr.Remove(updatestr.Length - 1, 1) + " where comp_companyid=" + companyid;

                                    try
                                    {
                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, companyupdatesql);
                                        updatecount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorcount++;
                                        errormessage += ex.Message + ";";
                                    }
                                }
                            }
                            else
                            {
                                reader.Close();
                                int id = autogenerateid();

                                string companyinsertsql = string.Format(insertsql, ownername, ownergender, ownerqq, ownerphone, owneremail,
                             owneraddress, contacterlname, contacterlphone, contacterlqq, contacterlemail, contacter2name, contacter2phone, contacter2qq, contacter2email, nextfollowdate, followstatus,
                             priority, admit, signupid, ownerattitude, ifaddqq, wangwangid, comment, followsaler, storename, id, usersecterr, followsaler, serialnum, taobaourl, loanpurpose, customerstatus, scale, companyname, corporation,
                             basiclimit, templimit, quicklimit, jhslimit, dealcount, dealamount, lastdealdate, lastloanlimit, level);
                                try
                                {
                                    if (!string.IsNullOrEmpty(signupid))
                                    {
                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, companyinsertsql);
                                        insertcount++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorcount++;
                                    errormessage += ex.Message + ";";
                                }
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        #region 导入店铺
                        #region 导入字段
                        for (int i = 0; i < rowArray.Length; i++)
                        {
                            string taobaourl = string.Empty;
                            string taobaostorecode = string.Empty;
                            string storetype = string.Empty;
                            string storerange = string.Empty;
                            string storearea = string.Empty;
                            string wangwang = string.Empty;
                            string goodrate = string.Empty;
                            string solddegree = string.Empty;
                            string descriptionrate = string.Empty;
                            string descriptionlevel = string.Empty;
                            string autitude = string.Empty;
                            string autitudelevel = string.Empty;
                            string shipspeed = string.Empty;
                            string shiplevel = string.Empty;
                            string mongoodrate = string.Empty;
                            string monmiddlerate = string.Empty;
                            string monbadrate = string.Empty;
                            string yeargoodrate = string.Empty;
                            string yearmiddlerate = string.Empty;
                            string yearbadrate = string.Empty;
                            string storename = string.Empty;//店铺名称
                            string ownername = string.Empty;
                            string ownergender = string.Empty;
                            string ownerphone = string.Empty;
                            string ownerqq = string.Empty;
                            string owneremail = string.Empty;
                            string owneraddress = string.Empty;
                            string contacterlname = string.Empty;
                            string contacterlphone = string.Empty;
                            string contacterlqq = string.Empty;
                            string contacterlemail = string.Empty;
                            string contacter2name = string.Empty;
                            string contacter2phone = string.Empty;
                            string contacter2qq = string.Empty;
                            string contacter2email = string.Empty;
                            string scale = string.Empty;
                            string corporation = string.Empty;
                            string companyname = string.Empty;
                            string ifaddqq = string.Empty;
                            string comment = string.Empty;
                            string followsaler = string.Empty;
                            string usersecterr = string.Empty;
                            string loanpurpose = string.Empty;
                            string assignstatus = string.Empty;
                            string priority = string.Empty;
                            string storeproper = string.Empty;
                            if (ds.Tables[0].Columns.Contains("店铺名称"))
                            {
                                storename = rowArray[i]["店铺名称"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("优先级"))
                            {
                                priority = rowArray[i]["优先级"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("淘宝URL"))
                            {
                                taobaourl = rowArray[i]["淘宝URL"].ToString();
                                if (!string.IsNullOrEmpty(taobaourl))
                                {
                                    if (!taobaourl.Contains("rate.") && taobaourl.Contains("http"))
                                    {
                                        string temp = taobaourl.Substring(0, taobaourl.IndexOf(".com") + 4);
                                        taobaourl = temp;
                                    }
                                    else
                                    {
                                        string temp = taobaourl.Substring(0, taobaourl.IndexOf(".htm") + 4);
                                        taobaourl = temp;
                                    }

                                }
                            }//comp_storeurl
                            if (ds.Tables[0].Columns.Contains("淘宝店铺编号"))
                            {
                                taobaostorecode = rowArray[i]["淘宝店铺编号"].ToString();
                            }//comp_storeid
                            if (ds.Tables[0].Columns.Contains("店铺类型"))
                            {
                                storetype = rowArray[i]["店铺类型"].ToString();
                                switch (storetype)
                                {
                                    case "淘宝店铺":
                                        storetype = "1"; break;
                                    case "天猫店铺":
                                        storetype = "2"; break;
                                    default:
                                        storetype = ""; break;
                                }

                            }//comp_storetype
                            if (ds.Tables[0].Columns.Contains("经营范围"))
                            {
                                storerange = rowArray[i]["经营范围"].ToString();
                            }//comp_businessarea
                            if (ds.Tables[0].Columns.Contains("地区"))
                            {
                                storearea = rowArray[i]["地区"].ToString();
                            } //comp_region
                            if (ds.Tables[0].Columns.Contains("淘宝旺旺"))
                            {
                                wangwang = rowArray[i]["淘宝旺旺"].ToString();
                            } //comp_wangwangid
                            if (ds.Tables[0].Columns.Contains("好评率"))
                            {
                                goodrate = rowArray[i]["好评率"].ToString();
                                if (string.IsNullOrEmpty(goodrate))
                                {
                                    goodrate = "0";
                                }

                            } //comp_reputablerate
                            else
                            {
                                goodrate = "0";
                            }
                            if (ds.Tables[0].Columns.Contains("店铺等级"))
                            {
                                solddegree = rowArray[i]["店铺等级"].ToString();
                                if (string.IsNullOrEmpty(solddegree))
                                {
                                    solddegree = "0";
                                }
                            }//comp_storelevel
                            if (ds.Tables[0].Columns.Contains("商品与描述相符度"))
                            {
                                descriptionrate = rowArray[i]["商品与描述相符度"].ToString();
                            } //comp_discribecorrespondrate
                            if (ds.Tables[0].Columns.Contains("商品与描述相符度同业水平"))
                            {
                                descriptionlevel = rowArray[i]["商品与描述相符度同业水平"].ToString();
                            } //comp_discribecorrespondlevel
                            if (ds.Tables[0].Columns.Contains("服务态度"))
                            {
                                autitude = rowArray[i]["服务态度"].ToString();
                            } //comp_serviceattitude
                            if (ds.Tables[0].Columns.Contains("服务态度同业水平"))
                            {
                                autitudelevel = rowArray[i]["服务态度同业水平"].ToString();
                            }//comp_serviceattitudelevel
                            if (ds.Tables[0].Columns.Contains("发货速度"))
                            {
                                shipspeed = rowArray[i]["发货速度"].ToString();
                            }//comp_deliverspeed
                            if (ds.Tables[0].Columns.Contains("发货速度同业水平"))
                            {
                                shiplevel = rowArray[i]["发货速度同业水平"].ToString();
                            } //comp_deliverspeedlevel
                            if (ds.Tables[0].Columns.Contains("近一个月好评"))
                            {
                                mongoodrate = rowArray[i]["近一个月好评"].ToString();

                            } //comp_amonthgood
                            if (ds.Tables[0].Columns.Contains("近一个月中评"))
                            {
                                monmiddlerate = rowArray[i]["近一个月中评"].ToString();
                            } //comp_amonthmedium
                            if (ds.Tables[0].Columns.Contains("近一个月差评"))
                            {
                                monbadrate = rowArray[i]["近一个月差评"].ToString();
                            } //comp_amonthbad
                            if (ds.Tables[0].Columns.Contains("近半年好评"))
                            {
                                yeargoodrate = rowArray[i]["近半年好评"].ToString();
                            } //comp_halfyeargood
                            if (ds.Tables[0].Columns.Contains("近半年中评"))
                            {
                                yearmiddlerate = rowArray[i]["近半年中评"].ToString();
                            }//comp_halfyearmedium
                            if (ds.Tables[0].Columns.Contains("近半年差评"))
                            {
                                yearbadrate = rowArray[i]["近半年差评"].ToString();
                            }//comp_halfyearbad
                            //*******
                            if (ds.Tables[0].Columns.Contains("店主姓名"))
                            {
                                ownername = rowArray[i]["店主姓名"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("店主性别"))
                            {
                                ownergender = rowArray[i]["店主性别"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("店主手机号"))
                            {
                                ownerphone = rowArray[i]["店主手机号"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("店主QQ"))
                            {
                                ownerqq = rowArray[i]["店主QQ"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("店主邮箱地址"))
                            {
                                owneremail = rowArray[i]["店主邮箱地址"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("店主联系地址"))
                            {
                                owneraddress = rowArray[i]["店主联系地址"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1"))
                            {
                                contacterlname = rowArray[i]["联系人1"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1手机号"))
                            {
                                contacterlphone = rowArray[i]["联系人1手机号"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1QQ"))
                            {
                                contacterlqq = rowArray[i]["联系人1QQ"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人1Email"))
                            {
                                contacterlemail = rowArray[i]["联系人1Email"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2"))
                            {
                                contacter2name = rowArray[i]["联系人2"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2手机号"))
                            {
                                contacter2phone = rowArray[i]["联系人2手机号"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2QQ"))
                            {
                                contacter2qq = rowArray[i]["联系人2QQ"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("联系人2Email"))
                            {
                                contacter2email = rowArray[i]["联系人2Email"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("规模"))
                            {
                                scale = rowArray[i]["规模"].ToString();
                                switch (scale.Trim())
                                {
                                    case "大": scale = "1"; break;
                                    case "中": scale = "2"; break;
                                    case "小": scale = "3"; break;
                                    default: scale = ""; break;
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("企业名称"))
                            {
                                companyname = rowArray[i]["企业名称"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("企业法人"))
                            {
                                corporation = rowArray[i]["企业法人"].ToString();
                            }
                            if (ds.Tables[0].Columns.Contains("是否添加企业QQ"))
                            {
                                ifaddqq = rowArray[i]["是否添加企业QQ"].ToString();
                                switch (ifaddqq.Trim())
                                {
                                    case "是": ifaddqq = "Y"; break;
                                    default: ifaddqq = ""; break;

                                }
                            }
                            if (ds.Tables[0].Columns.Contains("备注"))
                            {
                                comment = rowArray[i]["备注"].ToString();
                            }
                            //**********
                            if (ds.Tables[0].Columns.Contains("跟进销售"))
                            {
                                followsaler = rowArray[i]["跟进销售"].ToString();
                                if (!string.IsNullOrEmpty(followsaler))
                                {
                                    string GetSalesSql = "select user_userid,User_PrimaryTerritory from users where User_deleted is null and User_LastName='" + followsaler + "'";
                                    SqlDataReader RDUser = SqlHelper.ExecuteReader(cn, CommandType.Text, GetSalesSql);
                                    if (RDUser.Read())
                                    {
                                        followsaler = RDUser["user_userid"].ToString();
                                        usersecterr = RDUser["User_PrimaryTerritory"].ToString();
                                    }
                                    RDUser.Close();
                                }
                            }
                            if (ds.Tables[0].Columns.Contains("贷款意向"))
                            {
                                loanpurpose = rowArray[i]["贷款意向"].ToString();
                                if (loanpurpose.ToLower() == "y" || loanpurpose.Trim() == "是")
                                    loanpurpose = "Y";
                                else loanpurpose = "";
                            }//comp_storetype

                            if (ds.Tables[0].Columns.Contains("店铺属性"))
                            {
                                storeproper = rowArray[i]["店铺属性"].ToString();
                            }
                            string selsql = "select comp_assignstatus,comp_priority from Company where comp_deleted is null and comp_customertype<>'2' and comp_storeurl='{0}'";
                            SqlDataReader selReader = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(selsql, taobaourl));
                            if (selReader.Read())
                            {
                                assignstatus = selReader[0].ToString();

                                if (assignstatus == "2" || string.IsNullOrEmpty(priority))
                                {
                                    priority = selReader[1].ToString();

                                }
                            }

                            selReader.Close();
                        #endregion
                            string checksql = "select comp_companyid from company where comp_deleted is null and comp_customertype<>'2' and comp_storeurl='{0}'";
                            string updatesql = string.Empty;
                            if (followsaler != "0" && !string.IsNullOrEmpty(followsaler))
                                updatesql = @"update company 
                        set comp_storeurl='{0}',comp_wangwangid='{1}',comp_storetype='{2}',comp_businessarea='{3}'
                        ,comp_region='{4}',comp_reputablerate='{5}',comp_storelevel='{6}',comp_discribecorrespondrate='{7}'
                        ,comp_discribecorrespondlevel='{8}',comp_serviceattitude='{9}',comp_serviceattitudelevel='{10}'
                        ,comp_deliverspeed='{11}',comp_deliverspeedlevel='{12}',comp_amonthgood='{13}'
                        ,comp_amonthmedium='{14}',comp_amonthbad='{15}',comp_halfyeargood='{16}',comp_halfyearmedium='{17}'
                       ,comp_halfyearbad='{18}' ,comp_storename='{19}',comp_comment='{20}',comp_primaryuserid = '{21}',comp_loanpurpose = '{22}',comp_secterr='{23}',comp_assignstatus = '2',comp_storeproperty='{24}' where comp_companyid='{25}'";
                            else
                                updatesql = @"update company 
                        set comp_storeurl='{0}',comp_wangwangid='{1}',comp_storetype='{2}',comp_businessarea='{3}'
                        ,comp_region='{4}',comp_reputablerate='{5}',comp_storelevel='{6}',comp_discribecorrespondrate='{7}'
                        ,comp_discribecorrespondlevel='{8}',comp_serviceattitude='{9}',comp_serviceattitudelevel='{10}'
                        ,comp_deliverspeed='{11}',comp_deliverspeedlevel='{12}',comp_amonthgood='{13}'
                        ,comp_amonthmedium='{14}',comp_amonthbad='{15}',comp_halfyeargood='{16}',comp_halfyearmedium='{17}'
                       ,comp_halfyearbad='{18}' ,comp_storename='{19}',comp_storeproperty='{20}' where comp_companyid='{21}'";

                            string insertsql = @"insert into company ( comp_companyid,comp_storeurl,comp_storeid,comp_storetype,comp_businessarea 
                         ,comp_region,comp_wangwangid,comp_reputablerate,comp_storelevel,comp_discribecorrespondrate,comp_discribecorrespondlevel
                          ,comp_serviceattitude,comp_serviceattitudelevel,comp_deliverspeed,comp_deliverspeedlevel,comp_amonthgood
						,comp_amonthmedium,comp_amonthbad,comp_halfyeargood,comp_halfyearmedium,comp_halfyearbad,comp_storename
						,comp_name,comp_ownergender,comp_ownerqq,comp_ownerphone,comp_owneremail 
                         ,comp_owneraddress,comp_contacter1name,comp_contacter1phone,comp_contacter1qq,comp_contacter1email,comp_contacter2name
                          ,comp_contacter2phone,comp_contacter2qq,comp_contacter2email,comp_customertype,comp_serialnum,comp_compname,comp_corporation,comp_scale,comp_ifaddqq,comp_comment,comp_primaryuserid,comp_loanpurpose,comp_assignstatus,comp_secterr,comp_priority,comp_storeproperty) values ('{0}',
                         '{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}'
,'{23}','{24}','{25}','{26}','{27}','{28}','{29}','{30}','{31}','{32}','{33}','{34}','{35}','1','{36}','{37}','{38}','{39}','{40}','{41}','{42}','{43}','{44}','{45}','{46}','{47}')";
                            if (!string.IsNullOrEmpty(taobaourl))
                            {
                                SqlDataReader reader = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(checksql, taobaourl));

                                if (reader.Read())
                                {
                                    string companyid = reader[0].ToString();
                                    string companyupdatesql = string.Empty;
                                    reader.Close();
                                    if (followsaler != "0" && !string.IsNullOrEmpty(followsaler))
                                        companyupdatesql = string.Format(updatesql, taobaourl, wangwang, storetype, storerange, storearea,
            goodrate, solddegree, descriptionrate, descriptionlevel, autitude, autitudelevel, shipspeed, shiplevel, mongoodrate, monmiddlerate, monbadrate,
            yeargoodrate, yearmiddlerate, yearbadrate, storename, comment, followsaler, loanpurpose, usersecterr, storeproper, companyid);
                                    else
                                        companyupdatesql = string.Format(updatesql, taobaourl, wangwang, storetype, storerange, storearea,
         goodrate, solddegree, descriptionrate, descriptionlevel, autitude, autitudelevel, shipspeed, shiplevel, mongoodrate, monmiddlerate, monbadrate,
         yeargoodrate, yearmiddlerate, yearbadrate, storename, storeproper, companyid);

                                    try
                                    {
                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, companyupdatesql);
                                        updatecount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorcount++;
                                        errormessage += ex.Message + ";";
                                    }

                                }
                                else
                                {
                                    reader.Close();
                                    //插入的数据有销售字段变为分配状态
                                    if (followsaler != "0" && !string.IsNullOrEmpty(followsaler))
                                    {
                                        assignstatus = "2";
                                    }
                                    int id = autogenerateid();
                                    string companyinsertsql = string.Format(insertsql, id, taobaourl, taobaostorecode, storetype, storerange, storearea, wangwang,
         goodrate, solddegree, descriptionrate, descriptionlevel, autitude, autitudelevel, shipspeed, shiplevel, mongoodrate, monmiddlerate, monbadrate,
         yeargoodrate, yearmiddlerate, yearbadrate, storename, ownername, ownergender, ownerqq, ownerphone, owneremail,
                             owneraddress, contacterlname, contacterlphone, contacterlqq, contacterlemail, contacter2name, contacter2phone, contacter2qq
                             , contacter2email, serialnum, companyname, corporation, scale, ifaddqq, comment, followsaler, loanpurpose, assignstatus, usersecterr, priority, storeproper);
                                    try
                                    {
                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, companyinsertsql);
                                        insertcount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        errorcount++;
                                        errormessage += ex.Message + ";";
                                    }
                                }
                            }
                        }
                        # endregion
                    }
                    cn.Close();
                    //lblmessage.Text = "Excel表导入成功!成功导入" + (insertcount + updatecount).ToString() + "条记录，其中新增" + insertcount.ToString() + "条记录,更新" + updatecount.ToString() + "条记录，表中一共" + rowArray.Length + "条记录，其中错误数量为" + errorcount.ToString() + ",错误信息为:" + errormessage;
                }
                cn.Close();
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }

        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            string sid = Request.QueryString["SID"].ToString();
            string path = Request.Url.AbsoluteUri;
            string root = path.Split(new string[] { "CustomPages" }, StringSplitOptions.RemoveEmptyEntries)[0];
            string url = string.Empty;
            if (importtype != "1")
            {
                url = root + "eware.dll/Do?SID=" + sid + "&Act=432&Mode=1&CLk=T&dotnetdll=AssignJob&dotnetfunc=RunCheckBoxListPage&J=WaittedAssignStoerList&T=CustomerManagement";
            }
            else
            {
                url = root + "eware.dll/Do?SID=" + sid + "&Act=432&Mode=1&CLk=T&dotnetdll=company&dotnetfunc=RunCustomerListPage&J=CustomerList&T=CustomerManagement";
            }
            Response.Redirect(url);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            importtype = Request.QueryString.Get("importtype");
            
        }

        private int autogenerateid()
        {
            string CommandText = string.Empty;
            CommandText = "select Id_nextid as id from SQL_Identity where Id_TableId= 5";
            int maxid = 0;
            maxid = int.Parse(SqlHelper.ExecuteScalar(strConn, CommandType.Text, CommandText).ToString());
            maxid = maxid + 1;
            CommandText = "update SQL_Identity set Id_nextid=" + maxid + " where Id_TableId= 5";
            //执行sqlindentity中下一行的id更新


            //protected void Grid_SelectedIndexChanged(object sender, EventArgs e)
            //{

            //}   
        SqlHelper.ExecuteNonQuery(strConn, CommandType.Text, CommandText);
                return maxid;
            }

            //protected void Grid_SelectedIndexChanged()
            //{

            //}
        }
    }
