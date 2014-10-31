using System;
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
        string sheetname = string.Empty;

        public void ExecleDs(string filenameurl, string table)
        {
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + filenameurl + ";Extended Properties='Excel 12.0 Xml;HDR=YES'";
            //string strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + filenameurl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            string namestr = string.Empty;
            DataTable ds1 = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,null);
            for (int i = 0; i < ds1.Rows.Count; i++)
            {
                string name = ds1.Rows[i]["Table_Name"].ToString();
                namestr = namestr + name.Replace("$",",");
                
                string sql = "select * from [" + name + "]";
                DataSet ds = new DataSet();
                OleDbDataAdapter odda = new OleDbDataAdapter(sql, conn);
                odda.Fill(ds, table);
                InsertUpdateTable(ds);
            }
            //Response.Write("已插入表：" + namestr + " 中数据");
            //return ds;
        }
        
        //判断单引号
        public string SingleQuoteCheck(string str)
        {
            string rnt = string.Empty;
            if (str.Contains("'"))
            {
                string[] split = str.Split('\'');
                for (int i = 0; i < split.Length; i++)
                {
                    split[i].Insert(split[i].Length,"'");
                    rnt = rnt + split[i];
                }               
                return rnt;
            }
            else
                return str;
        }

        public void InsertUpdateTable(DataSet ds)
        {
            try
            {
                SqlConnection cn = new SqlConnection(strConn);
                cn.Open();
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

                    #region 导入产品
                    for (int i = 0; i < rowArray.Length; i++)
                    {
                        #region param
                        string modelname = string.Empty;//Model
                        string modelid = string.Empty;
                        string sizename = string.Empty;//Size
                        string sizeid = string.Empty;
                        string prodname = string.Empty;//Product
                        string prodid = string.Empty;
                        string language = string.Empty;//Language
                        string area = string.Empty;//Area
                        string curr = string.Empty;
                        string currid = string.Empty;
                        string price = string.Empty;
                        string priceid = string.Empty;
                        //feature
                        string f1 = string.Empty;
                        string f2 = string.Empty;
                        string f3 = string.Empty;
                        string f4 = string.Empty;
                        string f5 = string.Empty;
                        string f6 = string.Empty;
                        string f7 = string.Empty;
                        string f8 = string.Empty;

                        #endregion

                        #region GetValue
                        if (ds.Tables[0].Columns.Contains("Model"))
                        {
                            modelname = rowArray[i]["Model"].ToString();
                            if (!string.IsNullOrEmpty(modelname))
                            {
                                string modelidsql = "select mode_Modelid from Model where mode_deleted is null and mode_Name='" + modelname + "'";
                                SqlDataReader ModeRD = SqlHelper.ExecuteReader(cn, CommandType.Text, modelidsql);
                                if (ModeRD.Read())
                                {
                                    modelid = ModeRD["mode_Modelid"].ToString();
                                }
                                ModeRD.Close();
                            }
                            else
                            {
                                //errormessage = errormessage + "Model 列不可为空!\n";
                            }
                        }
                        if (ds.Tables[0].Columns.Contains("Size"))
                        {
                            sizename = rowArray[i]["Size"].ToString();
                            if (!string.IsNullOrEmpty(sizename))
                            {
                                string sizeidsql = "select size_SizeID from Size where size_modelid = '"+modelid+"' and  size_Name ='" + sizename + "'";
                                SqlDataReader SizeRD = SqlHelper.ExecuteReader(cn, CommandType.Text, sizeidsql);
                                if (SizeRD.Read())
                                {
                                    sizeid = SizeRD["size_SizeID"].ToString();
                                }
                                SizeRD.Close();
                            }
                            else
                            {
                                //errormessage = errormessage + "Size 列不可为空!\n";
                            }
                        }
                        if (ds.Tables[0].Columns.Contains("Product"))
                        {
                            prodname = rowArray[i]["Product"].ToString();
                        }
                        if (ds.Tables[0].Columns.Contains("Area"))
                        {
                            area = rowArray[i]["Area"].ToString();
                        }
                        if (ds.Tables[0].Columns.Contains("Language"))
                        {
                            language = rowArray[i]["Language"].ToString();
                            switch (language)
                            {
                                case "EN": if (area == "usa") { language = "US"; }
                                    else { language = "English";  } break;
                                case "ES": language = "Spanish";  break;
                                case "FR": language = "French"; break;
                                case "DE": language = "German";  break;
                                case "CN": language = "Chinese";break;
                                default: break;
                            }
                        }
                        if (ds.Tables[0].Columns.Contains("Currency"))
                        {
                            curr = rowArray[i]["Currency"].ToString();
                            switch (curr)
                            {
                                case "euro":  { curr = "EUR"; } break;
                                case "usd": curr = "USD"; break;
                                case "rmb": language = "CNY"; break;                                
                                default: break;
                            }
                        }
                        if (ds.Tables[0].Columns.Contains("Price"))
                        {
                            price = rowArray[i]["Price"].ToString();
                            if (string.IsNullOrEmpty(price))
                                price = "0";
                        }

                        //feature


                        if (ds.Tables[0].Columns.Contains("Feature 1"))
                        {
                            f1 = rowArray[i]["Feature 1"].ToString();
                            f1 = SingleQuoteCheck(f1);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 2"))
                        {
                            f2 = rowArray[i]["Feature 2"].ToString();
                            f2 = SingleQuoteCheck(f2);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 3"))
                        {
                            f3 = rowArray[i]["Feature 3"].ToString();
                            f3 = SingleQuoteCheck(f3);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 4"))
                        {
                            f4 = rowArray[i]["Feature 4"].ToString();
                            f4 = SingleQuoteCheck(f4);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 5"))
                        {
                            f5 = rowArray[i]["Feature 5"].ToString(); 
                            f5 = SingleQuoteCheck(f5);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 6"))
                        {
                            f6 = rowArray[i]["Feature 6"].ToString();
                            f6 = SingleQuoteCheck(f6);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 7"))
                        {
                            f7 = rowArray[i]["Feature 7"].ToString();
                            f7 = SingleQuoteCheck(f7);
                        }
                        if (ds.Tables[0].Columns.Contains("Feature 8"))
                        {
                            f8 = rowArray[i]["Feature 8"].ToString();
                            f8 = SingleQuoteCheck(f8);
                        }
                        #endregion

                        if (!string.IsNullOrEmpty(modelid) && !string.IsNullOrEmpty(sizeid))
                        {
                            //get product id
                            string prodidsql = "select sopr_SophProductID from SophProduct where sopr_deleted is null and sopr_modelid = " + modelid + " and sopr_sizeid = " + sizeid;
                            SqlDataReader ProdRD = SqlHelper.ExecuteReader(cn, CommandType.Text, prodidsql);
                            if (ProdRD.Read())
                            {
                                prodid = ProdRD["sopr_SophProductID"].ToString();
                            }
                            ProdRD.Close();

                            if (string.IsNullOrEmpty(prodid))
                            {
                                //insert product
                                prodid = autogenerateid(10230).ToString();
                                if (!string.IsNullOrEmpty(prodid))
                                {
                                    string prodinsert = string.Format("insert  SophProduct (sopr_SophProductID,sopr_Name,sopr_modelid,sopr_sizeid) values ({0},'{1}',{2},{3})", prodid, prodname, modelid, sizeid);
                                    SqlHelper.ExecuteNonQuery(cn, CommandType.Text, prodinsert);
                                    insertcount++;
                                }
                                else
                                {
                                    //errormessage = errormessage + "产品id为空无法插入数据\n";
                                }
                            }

                            else
                            {
                                //update product
                                string produpdatesql = "Update SophProduct set sopr_Name = '" + prodname + "' where sopr_SophProductId=" + prodid;
                                SqlHelper.ExecuteNonQuery(cn, CommandType.Text, produpdatesql);
                            }
                            if (!string.IsNullOrEmpty(prodid))
                            {
                                //feature
                                string checksql = "select nafe_NameFeatureID from NameFeature  where nafe_language = '" + language + "'  and  nafe_sophproductid =" + prodid;
                                SqlDataReader reader = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(checksql, checksql));
                                if (reader.Read())
                                {   //update feature

                                    string featureid = reader[0].ToString();
                                    reader.Close();
                                    string updatesql = @"Update NameFeature set nafe_featureone = '" + f1 + @"',nafe_featuretwo = '" + f2 + @"',nafe_featurethree = '" + f3 + @"',nafe_featurefour = '" + f4 + @"', 
                                    nafe_featurefive = '" + f5 + "',nafe_featuresix = '" + f6 + "',nafe_featureseven ='" + f7 + "',nafe_featureeight = '" + f8 + "',nafe_Name = '" + prodname + "' where nafe_NameFeatureID =" + featureid;
                                    try
                                    {
                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, updatesql);
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
                                    //insert feature
                                    reader.Close();
                                    string featureid = autogenerateid(10237).ToString();
                                    string insertsql = string.Format(@"insert NameFeature (nafe_NameFeatureID,nafe_language,nafe_featureone,nafe_featuretwo,nafe_featurethree,nafe_featurefour,nafe_featurefive,nafe_featuresix,nafe_featureseven,nafe_featureeight,nafe_CreatedBy,nafe_Name) values"
                                        + @"({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}')", featureid, language, f1, f2, f3, f4, f5, f6, f7, f8, 1, prodname);
                                    SqlHelper.ExecuteNonQuery(cn, CommandType.Text, insertsql);
                                    insertcount++;
                                }

                                if (!string.IsNullOrEmpty(curr) && !string.IsNullOrEmpty(price) && !string.IsNullOrEmpty(prodid))
                                {
                                    string currsql = "select Curr_CurrencyID from Currency where curr_deleted is null and Curr_Symbol ='" + curr + "'";
                                    SqlDataReader curRD = SqlHelper.ExecuteReader(cn, CommandType.Text, currsql);
                                    if (curRD.Read())
                                    {
                                        currid = curRD["Curr_CurrencyID"].ToString();
                                    }
                                    curRD.Close();
                                    if (!string.IsNullOrEmpty(currid))
                                    {
                                        string pricesql = "select pric_PriceID from Price where pric_Deleted is null and pric_sophproductid = '" + prodid + "' and pric_currency = '" + currid + "' ";
                                        SqlDataReader check = SqlHelper.ExecuteReader(cn, CommandType.Text, pricesql);
                                        if (check.Read())
                                        {
                                            priceid = check["pric_PriceID"].ToString();
                                            check.Close();
                                            string updatesql = "update Price set pric_prices='" + price + "' where pric_PriceID=" + priceid;
                                            SqlHelper.ExecuteNonQuery(cn, CommandType.Text, updatesql);
                                        }
                                        else
                                        {
                                            check.Close();
                                            priceid = autogenerateid(10233).ToString();
                                            string insert = string.Format("insert Price (pric_PriceID,pric_sophproductid,pric_currency,pric_prices) values ({0},{1},{2},{3})", priceid, prodid, currid, price);
                                            SqlHelper.ExecuteNonQuery(cn, CommandType.Text, insert);
                                        }
                                    }
                                    else
                                    {
                                        //errormessage = errormessage + "无货币类型\n";
                                    }
                                }

                            }
                            else
                            {
                                errormessage = errormessage + "产品id不能为空\n";
                            }
                        }
                        else { //errormessage = errormessage + "无法查到产品信息。"; 
                        }

                    }
                    #endregion

                    cn.Close();
                    lblmessage.Text = "Excel表导入成功!成功导入" + (insertcount + updatecount).ToString() + "条记录，其中新增" + insertcount.ToString() + "条记录,更新" + updatecount.ToString() + "条记录，表中一共" + rowArray.Length + "条记录，其中错误数量为" + errorcount.ToString() + ",错误信息为:" + errormessage;
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message);
            }
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
                //SqlConnection cn = new SqlConnection(strConn);
                //cn.Open();
                string filename = DateTime.Now.ToString("yyyymmddhhMMss") + FileUpload1.FileName;              //获取Execle文件名  DateTime日期函数
                string savePath = Server.MapPath(("~\\upfiles\\") + filename);//Server.MapPath 获得虚拟服务器相对路径
                FileUpload1.SaveAs(savePath);                        //SaveAs 将上传的文件内容保存在服务器上
                //DataSet ds = 
                ExecleDs(savePath, filename);           //调用自定义方法
                #region old
//                DataRow[] rowArray = ds.Tables[0].Select();            //定义一个DataRow数组
//                int rowsnum = ds.Tables[0].Rows.Count;
//                if (rowsnum == 0)
//                {
//                    Response.Write("<script>alert('这个Excel文档是空的!')</script>");   //当Excel表为空时,对用户进行提示
//                }
//                else
//                {
//                    int insertcount = 0, updatecount = 0;
//                    int errorcount = 0; string errormessage = string.Empty;
                   
//                    #region 导入产品
//                    for (int i = 0; i < rowArray.Length; i++)
//                    {
//                        #region param
//                        string modelname = string.Empty;//Model
//                        string modelid = string.Empty;
//                        string sizename = string.Empty;//Size
//                        string sizeid = string.Empty;
//                        string prodname = string.Empty;//Product
//                        string prodid = string.Empty;
//                        string language = string.Empty;//Language
//                        string area = string.Empty;//Area
//                        string curr = string.Empty;
//                        string currid = string.Empty;
//                        string price = string.Empty;
//                        string priceid = string.Empty;
//                        //feature
//                        string f1 = string.Empty;
//                        string f2 = string.Empty;
//                        string f3 = string.Empty;
//                        string f4 = string.Empty;
//                        string f5 = string.Empty;
//                        string f6 = string.Empty;
//                        string f7 = string.Empty;
//                        string f8 = string.Empty;

//                        #endregion

//                        #region GetValue
//                        if (ds.Tables[0].Columns.Contains("Model"))
//                        {
//                            modelname = rowArray[i]["Model"].ToString();
//                            if (!string.IsNullOrEmpty(modelname))
//                            {
//                                string modelidsql = "select mode_Modelid from Model where mode_deleted is null and mode_Name='" + modelname + "'";
//                                SqlDataReader ModeRD = SqlHelper.ExecuteReader(cn, CommandType.Text, modelidsql);
//                                if (ModeRD.Read())
//                                {
//                                    modelid = ModeRD["mode_Modelid"].ToString();
//                                }
//                                ModeRD.Close();
//                            }
//                            else {
//                                errormessage = errormessage + "Model 列不可为空!\n";
//                            } 
//                        }
//                        if (ds.Tables[0].Columns.Contains("Size"))
//                        {
//                            sizename = rowArray[i]["Size"].ToString();
//                            if (!string.IsNullOrEmpty(sizename))
//                            {
//                                string sizeidsql = "select size_SizeID from Size where size_Name ='"+sizename+"'";
//                                SqlDataReader SizeRD = SqlHelper.ExecuteReader(cn, CommandType.Text, sizeidsql);
//                                if (SizeRD.Read())
//                                {
//                                    sizeid = SizeRD["size_SizeID"].ToString();
//                                }
//                                SizeRD.Close();
//                            }
//                            else
//                            {
//                                errormessage = errormessage + "Size 列不可为空!\n";
//                            } 
//                        }
//                        if (ds.Tables[0].Columns.Contains("Product"))
//                        {
//                            prodname = rowArray[i]["Product"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Area"))
//                        {
//                            area = rowArray[i]["Area"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Language"))
//                        {
//                            language = rowArray[i]["Language"].ToString();
//                            switch (language)
//                            {
//                                case "EN": if (area == "usa") { language = "US"; curr = "USD"; }
//                                    else { language = "English"; curr = "GBP"; } break;
//                                case "ES": language = "Spanish"; curr = "EUR"; break;
//                                case "FR": language = "French"; curr = "EUR"; break;
//                                case "DE": language = "German"; curr = "EUR"; break;
//                                case "CN": language = "Chinese"; curr = "CNY"; break;
//                                default: break;
//                            }
//                        }
//                        //feature
                       

//                        if (ds.Tables[0].Columns.Contains("Feature 1"))
//                        {
//                            f1 = rowArray[i]["Feature 1"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 2"))
//                        {
//                            f2 = rowArray[i]["Feature 2"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 3"))
//                        {
//                            f3 = rowArray[i]["Feature 3"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 4"))
//                        {
//                            f4 = rowArray[i]["Feature 4"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 5"))
//                        {
//                            f5 = rowArray[i]["Feature 5"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 6"))
//                        {
//                            f6 = rowArray[i]["Feature 6"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 7"))
//                        {
//                            f7 = rowArray[i]["Feature 7"].ToString();
//                        }
//                        if (ds.Tables[0].Columns.Contains("Feature 8"))
//                        {
//                            f8 = rowArray[i]["Feature 8"].ToString();
//                        }
//                        #endregion

//                        if (!string.IsNullOrEmpty(modelid) && !string.IsNullOrEmpty(sizeid))
//                        {
//                            //get product id
//                            string prodidsql = "select sopr_SophProductID from SophProduct where sopr_deleted is null and sopr_modelid = " + modelid + " and sopr_sizeid = " + sizeid;
//                            SqlDataReader ProdRD = SqlHelper.ExecuteReader(cn, CommandType.Text, prodidsql);
//                            if (ProdRD.Read())
//                            {
//                                prodid = ProdRD["sopr_SophProductID"].ToString();
//                            }
//                            ProdRD.Close();

//                            if (string.IsNullOrEmpty(prodid))
//                            {
//                                //insert product
//                                prodid = autogenerateid(10230).ToString();
//                                if (!string.IsNullOrEmpty(prodid))
//                                {
//                                    string prodinsert = string.Format("insert  SophProduct (sopr_SophProductID,sopr_Name,sopr_modelid,sopr_sizeid) values ({0},'{1}',{2},{3})", prodid, prodname, modelid, sizeid);
//                                    SqlHelper.ExecuteNonQuery(cn, CommandType.Text, prodinsert);
//                                    insertcount++;
//                                }
//                                else
//                                {
//                                    errormessage = errormessage + "产品id为空无法插入数据\n";
//                                }
//                            }

//                            else { 
//                                //update product

//                            }
//                            if (!string.IsNullOrEmpty(prodid))
//                            { 
//                                //feature
//                                string checksql = "select nafe_NameFeatureID from NameFeature  where nafe_language = '"+language+"'  and  nafe_sophproductid ="+prodid;
//                                SqlDataReader reader = SqlHelper.ExecuteReader(cn, CommandType.Text, string.Format(checksql, checksql));
//                                if (reader.Read())
//                                {   //update feature
                                    
//                                    string featureid = reader[0].ToString();
//                                    reader.Close();
//                                    string updatesql = @"Update NameFeature set nafe_featureone = '"+f1+@"',nafe_featuretwo = '"+f2+@"',nafe_featurethree = '"+f3+@"',nafe_featurefour = '"+f4+@"', 
//                                    nafe_featurefive = '"+f5+"',nafe_featuresix = '"+f6+"',nafe_featureseven ='"+f7+"',nafe_featureeight = '"+f8+"',nafe_Name = '"+prodname+"' where nafe_NameFeatureID =" + featureid;
//                                    try
//                                    {
//                                        SqlHelper.ExecuteNonQuery(cn, CommandType.Text, updatesql);
//                                        updatecount++;
//                                    }
//                                    catch (Exception ex)
//                                    {
//                                        errorcount++;
//                                        errormessage += ex.Message + ";";
//                                    }
//                                } 
//                                else { 
//                                    //insert feature
//                                    reader.Close();
//                                    string featureid = autogenerateid(10237).ToString();
//                                    string insertsql = string.Format(@"insert NameFeature (nafe_NameFeatureID,nafe_language,nafe_featureone,nafe_featuretwo,nafe_featurefour,nafe_featurefive,nafe_featuresix,nafe_featureseven,nafe_featureeight,nafe_CreatedBy,nafe_Name) values"
//                                        + @"({0},'{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}',{10},'{11}')", featureid, language, f1, f2, f3, f4, f5, f6, f7, f8, 1,prodname);
//                                    insertcount++;
//                                }

//                                if (!string.IsNullOrEmpty(curr) && !string.IsNullOrEmpty(price))
//                                {
//                                    string currsql = "select Curr_CurrencyID from Currency where curr_deleted is null and Curr_Symbol ='" + curr + "'";
//                                    SqlDataReader curRD = SqlHelper.ExecuteReader(cn, CommandType.Text, currsql);
//                                    if (curRD.Read())
//                                    {
//                                        currid = curRD["Curr_CurrencyID"].ToString();
//                                    }
//                                    curRD.Close();
//                                    if (!string.IsNullOrEmpty(currid))
//                                    {
//                                        string pricesql = "select pric_PriceID from Price where pric_Deleted is null and pric_sophproductid = '" + prodid + "' and pric_currency = '" + currid + "' ";
//                                        SqlDataReader check = SqlHelper.ExecuteReader(cn, CommandType.Text, pricesql);
//                                        if (check.Read())
//                                        {
//                                            priceid = check["pric_PriceID"].ToString();
//                                            check.Close();
//                                            string updatesql = "update Price set pric_prices='" + price + "' where pric_PriceID=" + priceid;
//                                            SqlHelper.ExecuteNonQuery(cn, CommandType.Text, updatesql);
//                                        }
//                                        else {
//                                            check.Close();
//                                            priceid = autogenerateid(10233).ToString();
//                                            string insert = string.Format("insert Price (pric_PriceID,pric_sophproductid,pric_currency,pric_prices) values ({0},{1},{2},{3})", priceid, prodid, currid, price);
//                                            SqlHelper.ExecuteNonQuery(cn,CommandType.Text,insert);
//                                        }
//                                    }
//                                    else
//                                    {
//                                        errormessage = errormessage + "无货币类型\n";
//                                    }
//                                }

//                            }
//                            else
//                            {
//                                errormessage = errormessage + "产品id不能为空\n";
//                            }
//                        }
//                        else { errormessage = errormessage + "无法查到产品信息。"; }
                        
//                    }
//                    #endregion
                   
//                    cn.Close();
//                    lblmessage.Text = "Excel表导入成功!成功导入" + (insertcount + updatecount).ToString() + "条记录，其中新增" + insertcount.ToString() + "条记录,更新" + updatecount.ToString() + "条记录，表中一共" + rowArray.Length + "条记录，其中错误数量为" + errorcount.ToString() + ",错误信息为:" + errormessage;
                //                }
                #endregion
                //cn.Close();
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
            {
                url =root + "eware.dll/Do?SID="+sid+"&Act=432&Mode=1&CLk=T&dotnetdll=Product&dotnetfunc=RunSearchPage&J=SophProduct&T=find";
            }
            Response.Redirect(url);
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            sheetname=Request.QueryString.Get("sheetname");
        }

        private int autogenerateid(int tableid )
        {
            string CommandText = string.Empty;
            CommandText = "select Id_nextid as id from SQL_Identity where Id_TableId= "+tableid;
            int maxid = 0;
            maxid = int.Parse(SqlHelper.ExecuteScalar(strConn, CommandType.Text, CommandText).ToString());
            maxid = maxid + 1;
            CommandText = "update SQL_Identity set Id_nextid=" + maxid + " where Id_TableId= "+ tableid ;
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
