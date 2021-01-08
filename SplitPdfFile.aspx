using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;
using org.apache.pdfbox.pdmodel;
using org.apache.pdfbox.util;
using System.IO;
using System.Text.RegularExpressions;
using java.util;
using System.Data;
using iTextSharp.text.html.simpleparser;
using System.Drawing;
using System.Diagnostics;
using System.Net;
using System.Data.SqlClient;
using System.Web.Services;
using RawPrint;
using System.Drawing.Printing;

public partial class SplitPdfFile : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    //lay ca du lieu va vao trong database lay du lieu
    [WebMethod]
    public static List<string> getDatatoTable(string[] empdetails)
    {
        List<string> emp = new List<string>();
        List<string> emp1 = new List<string>();
        ArrayList lstcyuno = new ArrayList();
        ArrayList lstkz = new ArrayList();
        string checkzaiko = empdetails[3].ToString();

        if (empdetails[2].ToString() == "3")
        {
            for (int j = 1; j < 3; j++)
            {
                SqlConnection sqlconn = new SqlConnection();
                string sqlquery = "";
                if (j == 1)
                {
                    sqlconn = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
                    sqlquery = getSqlQuery(empdetails[5].ToString(), empdetails[0].ToString(), empdetails[1].ToString(), 2, empdetails[4].ToString(), j.ToString());
                }
                else if (j == 2)
                {
                    sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
                    sqlquery = getSqlQuery(empdetails[5].ToString(), empdetails[0].ToString(), empdetails[1].ToString(), 1, empdetails[4].ToString(), j.ToString());
                }
                sqlconn.Open();
                SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);
                SqlDataReader sdr = sqlcomn.ExecuteReader();

                while (sdr.Read())
                {
                    lstcyuno.add(sdr["CYUNO"].ToString());
                    lstkz.add(j);
                    string mt = sdr["CYUNO"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/" + sdr["KMNO"].ToString() + "/*/" + sdr["JYUNO"].ToString() + "/*/"
                       + sdr["JYUSU"].ToString() + "/*/" + sdr["ZUBAN"].ToString() + "/*/" + sdr["YDATE"].ToString() + "/*/" + sdr["NOUKI"].ToString() + "/*/"
                       + sdr["TANA"].ToString() + "/*/" + j + "/*/" + sdr["TOKCD"].ToString();
                    emp.Add(mt);
                }
                sqlconn.Close();
            }
            ArrayList zaikou = GetDataPO(lstcyuno, empdetails[2], lstkz);
            for (int i = 0; i < emp.Count(); i++)
            {

                string text_zaikou = zaikou.get(i).ToString().Trim();
                if (checkzaiko == "1")
                {
                    if (text_zaikou == "在庫あり")
                    {
                        emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);
                    }
                }
                else
                {
                    emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);
                }

            }
            return emp1;
        }
        else
        {
            SqlConnection sqlconn = new SqlConnection();
            if (empdetails[2].ToString() == "1")
            {
                sqlconn = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
            }
            else if (empdetails[2].ToString() == "2")
            {
                sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
            }
            string sqlquery = getSqlQuery(empdetails[5].ToString(), empdetails[0].ToString(), empdetails[1].ToString(), 1, empdetails[4].ToString(), empdetails[2].ToString());
            sqlconn.Open();
            SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);
            SqlDataReader sdr = sqlcomn.ExecuteReader();
            while (sdr.Read())
            {
                lstcyuno.add(sdr["CYUNO"].ToString());
                lstkz.add(empdetails[2].ToString());
                string mt = sdr["CYUNO"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/" + sdr["KMNO"].ToString() + "/*/" + sdr["JYUNO"].ToString() + "/*/"
                   + sdr["JYUSU"].ToString() + "/*/" + sdr["ZUBAN"].ToString() + "/*/" + sdr["YDATE"].ToString() + "/*/" + sdr["NOUKI"].ToString() + "/*/"
                   + sdr["TANA"].ToString() + "/*/" + empdetails[2].ToString() + "/*/" + sdr["TOKCD"].ToString();
                emp.Add(mt);
            }
            sqlconn.Close();
            ArrayList zaikou = GetDataPO(lstcyuno, empdetails[2], lstkz);
            for (int i = 0; i < zaikou.size(); i++)
            {
                string text_zaikou = zaikou.get(i).ToString().Trim();
                if (checkzaiko == "1")
                {
                    if (text_zaikou == "在庫あり")
                    {
                        emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);
                    }
                }
                else
                {
                    emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);
                }
            }

            return emp1;
        }
    }

    //ham tao cau truy van sql theo cac dieu kien tu dau vao
    [WebMethod]
    public static String getSqlQuery(string companyTokcd, string startDateSearch, string endDateSearch, int checkBothFactory, string searchWhere, string facid)
    {
        string sqlquery = "";
        //cau truy van chung
        string sqlselect = " WITH tb1 AS( SELECT [D1000].CYUNO,[M0120].KABUH,[D1000].JYUNO " +
   " ,[D1000].JYUSU,[D1000].ZUBAN,[D1000].YDATE,[D1000].NOUKI ,[D1000].TOKCD" +
   " FROM [TESC].[dbo].[D1000] " +
   " left join  [TESC].[dbo].[M0100]" +
   " on  [D1000].SEICD=[M0100].ZAICD " +
   " left join [TESC].[dbo].[M0120] " +
   " on  [M0100].ZAICD=[M0120].ZAICD  ";
        //cau truy van where theo nouki hay ydate
        string sqlwhereNOUKI = "";
        string sqlwhereYDATE = "";
        //cau truy van sap xep theo gi
        string sqlorder = "";

        if (companyTokcd == "11")
        {
            sqlwhereNOUKI = " where [M0120].JUNJ='001' and [D1000].TOKCD='00164' and  [D1000].NOUKI>='" + startDateSearch + "' and [D1000].NOUKI<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlwhereYDATE = " where [M0120].JUNJ='001' and [D1000].TOKCD='00164' and  [D1000].YDATE>='" + startDateSearch + "' and [D1000].YDATE<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlorder = " order by  tb1.TOKCD,tb1.NOUKI,[M0100].TANA ";
        }
        else if (companyTokcd == "21")
        {
            sqlwhereNOUKI = " where [M0120].JUNJ='001' and [D1000].TOKCD='00002' and  [D1000].NOUKI>='" + startDateSearch + "' and [D1000].NOUKI<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlwhereYDATE = " where [M0120].JUNJ='001' and [D1000].TOKCD='00002' and  [D1000].YDATE>='" + startDateSearch + "' and [D1000].YDATE<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlorder = " order by  tb1.TOKCD,tb1.NOUKI,[M0100].TANA ";
        }
        else if (companyTokcd == "22")
        {
            sqlwhereNOUKI = " where [M0120].JUNJ='001' and ( [D1000].TOKCD='00012' or [D1000].TOKCD='00095' or [D1000].TOKCD='00099' or [D1000].TOKCD='00207' ) and  [D1000].NOUKI>='" + startDateSearch + "' and [D1000].NOUKI<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlwhereYDATE = " where [M0120].JUNJ='001' and ( [D1000].TOKCD='00012' or [D1000].TOKCD='00095' or [D1000].TOKCD='00099' or [D1000].TOKCD='00207' ) and  [D1000].YDATE>='" + startDateSearch + "' and [D1000].YDATE<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
            sqlorder = "  order by tb1.TOKCD,tb1.NOUKI,tb1.ZUBAN ";
        }
        else if (companyTokcd == "23")
        {
            if (facid == "1")
            {
                sqlwhereNOUKI = " where [M0120].JUNJ='001' and (  [D1000].TOKCD='00203' ) and  [D1000].NOUKI>='" + startDateSearch + "' and [D1000].NOUKI<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
                sqlwhereYDATE = " where [M0120].JUNJ='001' and (  [D1000].TOKCD='00203' ) and  [D1000].YDATE>='" + startDateSearch + "' and [D1000].YDATE<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
                sqlorder = "  order by tb1.TOKCD,tb1.NOUKI,tb1.ZUBAN ";
            }
            else
            {
                sqlwhereNOUKI = " where [M0120].JUNJ='001' and (  [D1000].TOKCD='00164' ) and  [D1000].NOUKI>='" + startDateSearch + "' and [D1000].NOUKI<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
                sqlwhereYDATE = " where [M0120].JUNJ='001' and (  [D1000].TOKCD='00164' ) and  [D1000].YDATE>='" + startDateSearch + "' and [D1000].YDATE<='" + endDateSearch + "' and [M0100].ZAIKB='A' ) ";
                sqlorder = "  order by tb1.TOKCD,tb1.NOUKI,tb1.ZUBAN ";
            }
        }

        //cau truy van lay du lieu tuy vao tim kiem 1 nha may hay la 2 nha may
        string sqlselectApart = " select tb1.CYUNO,[M0100].TANA,[D5000].KMNO " +
     " ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,[TESCex].[dbo].[D1000exLog].JYUNO,tb1.TOKCD  from tb1 " +
     " left join [TESC].[dbo].[M0100] " +
     " on tb1.KABUH=[M0100].ZAICD " +
     " left join [TESC].[dbo].[D5000] " +
     " on  [D5000].JYUNO=tb1.JYUNO and [D5000].JSKBN='J' " +
     " left join [TESCex].[dbo].[D1000exLog] " +
     " on [TESCex].[dbo].[D1000exLog].JYUNO=tb1.JYUNO " +
     " where  [M0100].ZAIKB='B'  and [TESCex].[dbo].[D1000exLog].JYUNO is null ";

        string sqlselectBoth = " select tb1.CYUNO,[M0100].TANA,[D5000].KMNO " +
            " ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb2.JYUNO,tb1.TOKCD from tb1 " +
            " left join [TESC].[dbo].[M0100] " +
            " on tb1.KABUH=[M0100].ZAICD " +
            " left join [TESC].[dbo].[D5000] " +
            " on  [D5000].JYUNO=tb1.JYUNO " +
            " left join [10.121.21.12].[TESCex].[dbo].[D1000exLog] as tb2 " +
            " on tb2.JYUNO=tb1.JYUNO " +
            " where  [M0100].ZAIKB='B' and [D5000].JSKBN='J' and tb2.JYUNO is null ";

        if (checkBothFactory == 1)
        {
            if (searchWhere == "1")
            {
                sqlquery = sqlselect + sqlwhereNOUKI + sqlselectApart + sqlorder;
            }
            else if (searchWhere == "2")
            {
                sqlquery = sqlselect + sqlwhereYDATE + sqlselectApart + sqlorder;
            }

        }
        else if (checkBothFactory == 2)
        {
            if (searchWhere == "1")
            {
                sqlquery = sqlselect + sqlwhereNOUKI + sqlselectApart + sqlorder;
            }
            else if (searchWhere == "2")
            {
                sqlquery = sqlselect + sqlwhereYDATE + sqlselectApart + sqlorder;
            }
        }

        return sqlquery;
    }

    //logic kiem tra xem trong kho co ton tai hang hay khong
    private static dataProviderF1 dataf1 = new dataProviderF1();//object ket noi va lay du lieu tu Data Source=10.121.21.11
    private static dataProviderF2 dataf2 = new dataProviderF2();//object ket noi va lay du lieu tu Data Source=10.121.21.12

    [WebMethod]
    public static DataTable CheckPO01(ArrayList s1, string id, ArrayList lstkz)
    {
        var tablecheck1 = new DataTable();
        tablecheck1.Columns.Add("NYUSU", typeof(string));
        tablecheck1.Columns.Add("SYUSU", typeof(string));

        for (int i = 0; i < s1.size(); i++)
        {
            string target = s1.get(i).ToString();
            string facid = lstkz.get(i).ToString();
            //check sql table1
            string sqlcheck1 = " WITH tb1 AS (SELECT [D5900].JYUNO,[D5900].ZAICD,[D5900].NSDAT " +
" FROM [TESC].[dbo].[D1000] inner join [TESC].[dbo].[D1010] " +
" on [D1010].JYUNO=[D1000].JYUNO  " +
" inner join [TESC].[dbo].[D5900] " +
" on [D1010].JYUNO=[D5900].JYUNO  " +
" where [D1000].CYUNO='" + target + "' ) " +
"  SELECT SUM(NYUSU) as NYUSU , SUM(SYUSU) as SYUSU FROM [TESC].[dbo].[D5900] " +
"  inner join tb1 on [D5900].ZAICD=tb1.ZAICD and [D5900].NSDAT<=tb1.NSDAT ";


            var result = new object[50][];
            if (facid == "1")
            {
                result = dataf1.get(sqlcheck1);
            }
            else
            {
                result = dataf2.get(sqlcheck1);
            }

            if (result.Length == 0)
            {
                tablecheck1.Rows.Add("");
            }
            else
            {
                for (int j = 0; j < result.Length; j++)
                {
                    var row = tablecheck1.NewRow();
                    for (int k = 0; k < result[j].Length; k++)
                    {
                        row[k] = result[j][k];
                    }
                    tablecheck1.Rows.Add(row);
                }
            }
        }

        return tablecheck1;
    }
    [WebMethod]
    public static DataTable CheckPO02(ArrayList s1, string id, ArrayList lstkz)
    {
        var tablecheck1 = new DataTable();
        tablecheck1.Columns.Add("ZAISU", typeof(string));
        tablecheck1.Columns.Add("ZUBAN", typeof(string));
        for (int i = 0; i < s1.size(); i++)
        {
            string target = s1.get(i).ToString();
            string facid = lstkz.get(i).ToString();
            string sqlcheck1 = "  WITH tb1 AS ( SELECT [D5900].JYUNO,[D5900].ZAICD,[D5900].NSDAT,[M0100].ZAISU,[M0120].KABUH " +
 "  FROM [TESC].[dbo].[D1000] " +
 "  inner join [TESC].[dbo].[D1010] " +
 "  on [D1010].JYUNO=[D1000].JYUNO  " +
 "  inner join [TESC].[dbo].[D5900] " +
 "  on [D1010].JYUNO=[D5900].JYUNO  " +
 "  inner join [TESC].[dbo].[M0100] " +
 "   on [M0100].ZAICD=[D5900].ZAICD  and  [M0100].ZAIKB='A' " +
 "   inner join [TESC].[dbo].[M0120] " +
 "  on [M0100].ZAICD=[M0120].ZAICD " +
 "  where [D1000].CYUNO='" + target + "' ) " +
 "  SELECT [M0100].ZAISU+tb1.ZAISU as SUMZAISU,[M0100].ZUBAN " +
 "  FROM [TESC].[dbo].[M0100] inner join tb1 " +
 "  on tb1.KABUH=[M0100].ZAICD where  [M0100].ZAIKB='B' ";
            var result = new object[50][];
            if (facid == "1")
            {
                result = dataf1.get(sqlcheck1);
            }
            else
            {
                result = dataf2.get(sqlcheck1);
            }

            if (result.Length == 0)
            {
                tablecheck1.Rows.Add("");
            }
            else
            {
                for (int j = 0; j < result.Length; j++)
                {
                    var row = tablecheck1.NewRow();
                    for (int k = 0; k < result[j].Length; k++)
                    {

                        row[k] = result[j][k];
                    }
                    tablecheck1.Rows.Add(row);
                }
            }

        }
        return tablecheck1;
    }
    //    try{
    //}catch(Exception ex){
    //    tb3.add("***");

    //}
    [WebMethod]
    public static ArrayList GetDataPO(ArrayList s1, string id, ArrayList lstkz)
    {
        DataTable tb1 = CheckPO01(s1, id, lstkz);
        DataTable tb2 = CheckPO02(s1, id, lstkz);
        ArrayList tb3 = new ArrayList();
        for (int i = 0; i < s1.size(); i++)
        {
            string syusutext = tb1.Rows[i][1].ToString();
            float syusu = -1;
            if (syusutext != "")
            {
                syusu = float.Parse(syusutext);
            }
            string nyusutext = tb1.Rows[i][0].ToString();
            float nyusu = -1;
            if (nyusutext != "")
            {
                nyusu = float.Parse(nyusutext);
            }

            string zaisutext = tb2.Rows[i][0].ToString();
            float zaisu = -1;
            if (zaisutext != "")
            {
                zaisu = float.Parse(zaisutext);
            }

            if (syusu == -1 || nyusu == -1 || zaisu == -1)
            {
                tb3.add("*在庫なし");
                continue;
            }
            else
            {
                if (zaisu >= syusu)
                {
                    tb3.add("在庫あり");
                }
                else
                {
                    string target = s1.get(i).ToString();
                    string facid = lstkz.get(i).ToString();

                    SqlConnection sqlconn = new SqlConnection();
                    if (facid == "1")
                    {
                        sqlconn = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
                    }
                    else if (facid == "2")
                    {
                        sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
                    }
                    try
                    {

                        // waitfor delay '00:00:02';
                        string sqlquery = " SELECT TOP(1) [D5000].KMNO " +
    "    FROM [TESC].[dbo].[D1000]  " +
    "    inner join  [TESC].[dbo].[M0120] " +
    "    on   [D1000].SEICD=[M0120].ZAICD " +
    "    inner join [TESC].[dbo].[D1010] " +
    "    on [D1000].JYUNO=[D1010].JYUNO " +
    "	inner join [TESC].[dbo].[D5000] " +
    "	on [D5000].JYUNO=[D1010].SIYNO and [D5000].BUHCD=[M0120].KABUH " +
    "   where [D1000].CYUNO='" + target + "' ";
                        sqlconn.Open();

                        SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);
                        sqlcomn.CommandTimeout = 3;
                        SqlDataReader sdr = sqlcomn.ExecuteReader();
                        int check_kmno = 0;
                        string mt = "";
                        while (sdr.Read())
                        {
                            check_kmno += 1;
                            mt = sdr["KMNO"].ToString();
                        }
                        sqlconn.Close();
                        if (check_kmno == 0)
                        {

                            var tablecheck2 = new DataTable();
                            tablecheck2.Columns.Add("KMNO", typeof(string));
                            string sqlcheck2 = "   WITH tb1 AS ( SELECT [M0120].KABUH,[D1010].SIYNO ,[D5900].NSDAT " +
    " FROM [TESC].[dbo].[D1000]  inner join  [TESC].[dbo].[M0120] on   [D1000].SEICD=[M0120].ZAICD " +
    "inner join [TESC].[dbo].[D1010]  on [D1000].JYUNO=[D1010].JYUNO  inner join [TESC].[dbo].[D5900] " +
    " on [D1010].SIYNO =[D5900].JYUNO  where [D1000].CYUNO='" + target + "'  ) " +
    " SELECT  TOP(1) [D5900].KMNO FROM [TESC].[dbo].[D5900] inner join tb1 on [D5900].ZAICD=tb1.KABUH and [D5900].NSDAT<=tb1.NSDAT " +
    " where    [D5900].KMNO !='' order by [D5900].NSDAT desc ";

                            var result2 = new object[50][];
                            if (facid == "1")
                            {
                                result2 = dataf1.get(sqlcheck2);
                            }
                            else
                            {
                                result2 = dataf2.get(sqlcheck2);
                            }

                            if (result2.Length == 0)
                            {
                                tb3.add("在庫なし");
                            }
                            else
                            {
                                for (int l = 0; l < result2.Length; l++)
                                {
                                    var row1 = tablecheck2.NewRow();
                                    for (int p = 0; p < result2[l].Length; p++)
                                    {
                                        row1[p] = result2[l][p];
                                    }
                                    tablecheck2.Rows.Add(row1);
                                }
                                tb3.add("*" + tablecheck2.Rows[0]["KMNO"].ToString() + "*");
                            }
                        }
                        else
                        {
                            tb3.add(mt);

                        }
                    }
                    catch (Exception ex)
                    {
                        tb3.add("***");
                        sqlconn.Close();
                    }

                }
            }

        }
        return tb3;
    }
    ////////// ////////////////////////////////////////////////////////////////////
    ///////////////////////////ket thuc phan tim kiem va hien thi danh sach ra bang
    ///////////////////////////////////////////////////////////////////////////////

    //click button de doc file va chia nho file ra 
    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            string valueSelect = "";
            string valueId = this.tb_Id.Text;
            if (valueId == "1")
            {
                valueSelect = this.tb_selectId1.Text.ToString().Trim();
                //ベックマン BC
                if (valueSelect == "0")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\BC_N\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startBC_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startBC_File\", "*.pdf"))
                    {
                        string destinationFileName = file;
                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_BC.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_bc(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
                if (valueSelect == "1")
                {
                    //キヤノン　マシナリー
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\CNNM_2\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCNNM_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCNNM_File\", "*.pdf"))
                    {
                        string destinationFileName = file;
                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Cnnm_2.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_canon(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
            }
            else
            {
                valueSelect = this.tb_selectId2.Text.ToString().Trim();
                //CMSC
                if (valueSelect == "0")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\CMSC_N\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCMSC_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCMSC_File\", "*.pdf"))
                    {
                        string destinationFileName = file;

                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Cmsc.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_cmsc(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
                //CANON
                if (valueSelect == "1")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\CANON_N\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCANON_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startCANON_File\", "*.pdf"))
                    {
                        string destinationFileName = file;

                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Canon.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_canon_2format(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
                //NIKON
                if (valueSelect == "2")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\NIKON_MIYAGI_2\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startNKO_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startNKO_File\", "*.pdf"))
                    {
                        string destinationFileName = file;

                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Nikon.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);

                            List<string> data_kazu = get_kazu_nikon(copyFileName);
                            string insert_kazu_file = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\nikon_temporary.pdf";
                            insert_kazu_nikon(copyFileName, insert_kazu_file, data_kazu);

                            split_pdf_nikon2(insert_kazu_file);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
                //DAINIKKOU
                if (valueSelect == "3")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\DAINIKKOU\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startDNK_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startDNK_File\", "*.pdf"))
                    {
                        string destinationFileName = file;

                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Dainikkou.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_dainikkou(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                   
                }
                //REON
                if (valueSelect == "4")
                {
                    HttpFileCollection uploadedFiles = Request.Files;
                    for (int i = 0; i < uploadedFiles.Count; i++)
                    {
                        HttpPostedFile userPostedFile = uploadedFiles[i];
                        if (userPostedFile.ContentLength > 0)
                        {
                            //save file with time process
                            string timeProcessed = DateTime.Now.ToString("yyyyMMdd-HHmmss").ToString().Trim();
                            string oldPdfPath = userPostedFile.FileName.Substring(0, userPostedFile.FileName.Length - 4);
                            userPostedFile.SaveAs(@"\\10.121.21.2\data\DeliveryNote\REON\SRC\" + timeProcessed + "-" + oldPdfPath + ".pdf");

                            userPostedFile.SaveAs(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startRE_File\" + userPostedFile.FileName);
                        }
                    }
                    foreach (string file in Directory.GetFiles(@"\\10.121.21.2\wwwroot\pdfedit\UploadPdfFile\startRE_File\", "*.pdf"))
                    {
                        string destinationFileName = file;
                        string copyFileName = @"\\10.121.21.2\data\DeliveryNote\SaveCopyPdfFile\copyPage_Reon.pdf";
                        PdfReader reader = new PdfReader(destinationFileName);
                        int numberOfPage = reader.NumberOfPages;
                        for (int i = 1; i <= numberOfPage; i++)
                        {
                            copyPerPage(destinationFileName, copyFileName, i);
                            split_pdf_reon(copyFileName);
                        }
                        reader.Close();
                        File.Delete(destinationFileName);
                    }
                }
            }


        }
        catch (Exception ex)
        {
            Response.Write("<script>alert('フォーマット の問題 !!!')</script>");
        }
    }

    //chia tu pdf ban dau thanh cac pdf nho voi ten la don dat hang cyuno

    //ham nay giup lay du lieu va chia tach thanh cac pdf nho voi ten la cyuno
    protected void extractPdf_withBC(string sourcePdfPath)
    {
        ArrayList arr = dataCyuno_withBC(sourcePdfPath);//get 注文NO from pdf 
        int countpage = 0;
        int checklastpage = 0;

        if (arr.size() % 2 == 0)
        {
            countpage = arr.size() / 2;
            checklastpage = 0;
        }
        else
        {
            countpage = (arr.size() + 1) / 2;
            checklastpage = 1;
        }

        for (int i = 1; i <= countpage; i++)
        {
            if (i == countpage)
            {
                if (checklastpage == 0)
                {
                    var namepdf = arr.get(i * 2 - 2).ToString();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + ".pdf";
                    if (File.Exists(outputPdfPath))
                    {
                        outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + "-1.pdf";
                    }
                    cutHalfpage_withBC(sourcePdfPath, outputPdfPath, i, 1);//split a page to small , with pdf name is cyuno 
                    var namepdf1 = arr.get(i * 2 - 1).ToString();
                    string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf1 + ".pdf";
                    if (File.Exists(outputPdfPath1))
                    {
                        outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf1 + "-1.pdf";
                    }
                    cutHalfpage_withBC(sourcePdfPath, outputPdfPath1, i, 2);
                }
                else
                {
                    var namepdf = arr.get(i * 2 - 2).ToString();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + ".pdf";
                    if (File.Exists(outputPdfPath))
                    {
                        outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + "-1.pdf";
                    }
                    cutHalfpage_withBC(sourcePdfPath, outputPdfPath, i, 1);
                }
            }
            else
            {
                var namepdf = arr.get(i * 2 - 2).ToString();
                string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + ".pdf";
                if (File.Exists(outputPdfPath))
                {
                    outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf + "-1.pdf";
                }
                cutHalfpage_withBC(sourcePdfPath, outputPdfPath, i, 1);
                var namepdf1 = arr.get(i * 2 - 1).ToString();
                string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf1 + ".pdf";
                if (File.Exists(outputPdfPath1))
                {
                    outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\BC\" + namepdf1 + "-1.pdf";
                }
                cutHalfpage_withBC(sourcePdfPath, outputPdfPath1, i, 2);
            }
        }
    }

    protected void extractPdf_withCMS(string sourcePdfPath)
    {
        ArrayList arr = dataCyuno_withCMS(sourcePdfPath);
        int countpage = 0;
        int checklastpage = 0;

        if (arr.size() % 2 == 0)
        {
            countpage = arr.size() / 2;
            checklastpage = 0;
        }
        else
        {
            countpage = (arr.size() + 1) / 2;
            checklastpage = 1;
        }

        for (int i = 1; i <= countpage; i++)
        {
            if (i == countpage)
            {
                if (checklastpage == 0)
                {
                    var namepdf = arr.get(i * 2 - 2).ToString();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\" + namepdf + ".pdf";
                    cutHalfpage_withCMS(sourcePdfPath, outputPdfPath, i, 1);
                    var namepdf1 = arr.get(i * 2 - 1).ToString();
                    string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\" + namepdf1 + ".pdf";
                    cutHalfpage_withCMS(sourcePdfPath, outputPdfPath1, i, 2);
                }
                else
                {
                    var namepdf = arr.get(i * 2 - 2).ToString();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\" + namepdf + ".pdf";
                    cutHalfpage_withCMS(sourcePdfPath, outputPdfPath, i, 1);
                }
            }
            else
            {
                var namepdf = arr.get(i * 2 - 2).ToString();
                string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\" + namepdf + ".pdf";
                cutHalfpage_withCMS(sourcePdfPath, outputPdfPath, i, 1);
                var namepdf1 = arr.get(i * 2 - 1).ToString();
                string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\" + namepdf1 + ".pdf";
                cutHalfpage_withCMS(sourcePdfPath, outputPdfPath1, i, 2);
            }
        }
    }

    protected void extractPdf_withNikon(string sourcePdfPath)
    {
        ArrayList arr = dataCyuno_withNKO(sourcePdfPath);
        for (int i = 1; i <= arr.size(); i++)
        {
            var namepdf = arr.get(i - 1).ToString().Trim();
            string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\NIKON\" + namepdf + ".pdf";
            cutPage_withNIKON(sourcePdfPath, outputPdfPath, i);
        }
    }

    //do Canon phai xu li dong thoi 2 loai format nen cach lay du lieu cung se khac so voi phia tren
    public void extractPdf_withCanon(string sourfile)
    {
        string outText = @"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\TEST\testcn.txt";
        string[] lines = TxtData(sourfile, outText);
        ArrayList s1 = new ArrayList();
        ArrayList s2 = new ArrayList();
        ArrayList s3 = new ArrayList();
        ArrayList s4 = new ArrayList();
        string line = string.Empty;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains("注文番号") == true)
            {
                int nx = i + 19;
                s2.add(i);
                if (lines[nx].Contains("注文番号") == true)
                {
                    s3.add("0");
                    int j = i;
                    j = i + 36;
                    line = lines[j];
                    string[] sp = new string[100];
                    string[] spliter = new string[] { " " };
                    sp = line.Split(spliter, 0);
                    s1.add(sp[0]);

                }
                else
                {
                    s3.add("1");
                    int j = i;
                    j = i + 17;
                    line = lines[j];
                    string[] sp = new string[100];
                    string[] spliter = new string[] { " " };
                    sp = line.Split(spliter, 0);
                    s1.add(sp[0]);
                }
            }
        }

        ArrayList sIndex = new ArrayList();
        for (int i = 1; i < s3.size(); i++)
        {
            if (s3.get(i).ToString().Contains(s3.get(i - 1).ToString()) == true)
            {
                sIndex.add(i);
            }
        }

        if (s3.get(0).ToString() == "1")
        {
            s3.set(0, "2");
        }

        for (int i = 0; i < sIndex.size(); i++)
        {
            s3.set((int)sIndex.get(i), "2");
        }


        s4 = dataCyuno_withCanon(s1, s2, s3, sourfile, outText);

        int leng1 = 0, leng2 = 0, leng3 = 0, leng4 = 0;

        ArrayList sList = new ArrayList();

        for (int i = 0; i < s1.size(); i++)
        {
            if (sList.contains(s1.get(i).ToString()) == false)
            {
                sList.add(s1.get(i));
            }
        }

        for (int i = 0; i < s1.size(); i++)
        {
            if (s1.get(i).ToString().Trim() == "キヤノン㈱阿見光機")
            {
                leng1 += 1;
            }
            else if (s1.get(i).ToString().Trim() == "キヤノンセミコンダクターエクィ")
            {
                leng2 += 1;
            }
            else if (s1.get(i).ToString().Trim() == "長浜キヤノン株式会社")
            {
                leng3 += 1;
            }
            else
            {
                leng4 += 1;
            }
        }

        ArrayList sPage = new ArrayList();
        for (int i = 0; i < sList.size(); i++)
        {
            if (sList.get(i).ToString().Trim() == "キヤノン㈱阿見光機")
            {
                sPage.add(leng1);
            }
            else if (sList.get(i).ToString().Trim() == "キヤノンセミコンダクターエクィ")
            {
                sPage.add(leng2);
            }
            else if (sList.get(i).ToString().Trim() == "長浜キヤノン株式会社")
            {
                sPage.add(leng3);
            }
            else
            {
                sPage.add(leng4);
            }
        }

        int numberpage = 0;
        for (int i = 0; i < sPage.size(); i++)
        {
            int countvalue = (int)sPage.get(i);
            int partpage = 0;
            if (countvalue % 2 == 0)
            {
                partpage = countvalue / 2;
            }
            else
            {
                partpage = (int)countvalue / 2 + 1;
            }
            numberpage += partpage;
        }

        string sourcePdfPath = sourfile;

        SplitNKS(sList, sPage, s4, sourcePdfPath);
        SplitGPH(sList, sPage, s4, sourcePdfPath, numberpage);
    }
    //phai chia 2 lan cat cung voi cac kieu cat cung khac nhau 2 phan va 6 phan
    public void SplitNKS(ArrayList sList, ArrayList sPage, ArrayList s4, string sourcePdfPath)
    {
        int numberpage = 0;
        int index = 0;
        for (int j = 0; j < sList.size(); j++)
        {
            int checklastpage = 0;
            int tpage = (int)sPage.get(j);
            int ppage = 0;
            if (tpage % 2 == 1)
            {
                ppage = (int)(tpage / 2) + 1;
                checklastpage = 1;
            }
            else
            {
                ppage = tpage / 2;
                checklastpage = 0;
            }
            numberpage += ppage;
            for (int i = numberpage - ppage + 1; i <= numberpage; i++)
            {
                if (i == numberpage)
                {
                    if (checklastpage == 0)
                    {
                        string namepdf = s4.get(index).ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\" + namepdf + ".pdf";
                        cutHalfpage_withCanon(sourcePdfPath, outputPdfPath, i, 1);
                        index += 1;
                        string namepdf1 = s4.get(index).ToString().Trim();
                        string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\" + namepdf1 + ".pdf";
                        cutHalfpage_withCanon(sourcePdfPath, outputPdfPath1, i, 2);
                        index += 1;
                    }
                    else
                    {
                        string namepdf = s4.get(index).ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\" + namepdf + ".pdf";
                        cutHalfpage_withCanon(sourcePdfPath, outputPdfPath, i, 1);
                        index += 1;
                    }
                }
                else
                {
                    string namepdf = s4.get(index).ToString().Trim();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\" + namepdf + ".pdf";
                    cutHalfpage_withCanon(sourcePdfPath, outputPdfPath, i, 1);
                    index = index + 1;
                    string namepdf1 = s4.get(index).ToString().Trim();
                    string outputPdfPath1 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\" + namepdf1 + ".pdf";
                    cutHalfpage_withCanon(sourcePdfPath, outputPdfPath1, i, 2);
                    index = index + 1;
                }

            }

        }
    }

    public void SplitGPH(ArrayList sList, ArrayList sPage, ArrayList s4, string sourcePdfPath, int numberpage)
    {
        int index = 0;
        for (int j = 0; j < sList.size(); j++)
        {
            int checklastpage = 0;
            int tpage = (int)sPage.get(j);
            int ppage = 0;
            if (tpage % 6 == 0)
            {
                ppage = tpage / 6;
                checklastpage = 0;
            }
            else if (tpage % 6 == 1)
            {
                ppage = (int)(tpage / 6) + 1;
                checklastpage = 1;
            }
            else if (tpage % 6 == 2)
            {
                ppage = (int)(tpage / 6) + 1;
                checklastpage = 2;
            }
            else if (tpage % 6 == 3)
            {
                ppage = (int)(tpage / 6) + 1;
                checklastpage = 3;
            }
            else if (tpage % 6 == 4)
            {
                ppage = (int)(tpage / 6) + 1;
                checklastpage = 4;
            }
            else if (tpage % 6 == 5)
            {
                ppage = (int)(tpage / 6) + 1;
                checklastpage = 5;
            }
            numberpage += ppage;
            for (int i = numberpage - ppage + 1; i <= numberpage; i++)
            {
                if (i == numberpage)
                {
                    if (checklastpage == 0)
                    {
                        checklastpage = 6;
                    }
                    for (int p = 1; p <= checklastpage; p++)
                    {
                        string namepdf = s4.get(index).ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\" + namepdf + ".pdf";
                        cutSixPartpage_withCanon(sourcePdfPath, outputPdfPath, i, p);
                        index = index + 1;
                    }
                }
                else
                {
                    for (int p = 1; p <= 6; p++)
                    {
                        string namepdf = s4.get(index).ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\" + namepdf + ".pdf";
                        cutSixPartpage_withCanon(sourcePdfPath, outputPdfPath, i, p);
                        index = index + 1;
                    }
                }

            }
        }
    }

    //su dung logic de lay cyono tu file txt
    protected ArrayList dataCyuno_withBC(string sourfile)
    {
        string outText = @"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\TEST\testpo.txt";
        string[] lines = TxtData(sourfile, outText);//get string
        ArrayList s1 = new ArrayList();
        string line = string.Empty;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains("$") == true && lines[i].Contains("*") == true)
            {
                int j = i + 1;
                line = lines[j];
                string[] sp = new string[100];
                string[] spliter = new string[] { " " };
                sp = line.Split(spliter, 0);
                s1.add(sp[0]);
            }
        }
        return s1;
    }

    protected ArrayList dataCyuno_withCMS(string sourfile)
    {
        string outText = @"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\TEST\testkn.txt";
        string[] lines = TxtData(sourfile, outText);
        ArrayList s1 = new ArrayList();
        string line = string.Empty;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains("（株）玉吉製作所") == true)
            {
                int j = i + 3;
                line = lines[j];
                s1.add(line);
            }
        }

        ArrayList s2 = new ArrayList();
        for (int i = 0; i < s1.size(); i++)
        {
            if (i % 2 == 0)
            {
                line = s1.get(i).ToString();
                string[] sp = new string[100];
                string[] spliter = new string[] { " " };
                sp = line.Split(spliter, 0);
                string name = sp[1].ToString().Trim();
                name = name.Substring(0, 8) + "-" + name.Substring(8, 2);
                s2.add(name);

            }
        }
        return s2;
    }

    protected ArrayList dataCyuno_withNKO(string sourfile)
    {
        string outText = @"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\TEST\testno.txt";
        string[] lines = TxtData(sourfile, outText);//get string
        ArrayList s1 = new ArrayList();
        string line = string.Empty;
        for (int i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains("棚番：") == true)
            {
                int j = i + 14;
                line = lines[j];
                string[] sp = new string[100];
                string[] spliter = new string[] { " " };
                sp = line.Split(spliter, 0);
                s1.add(sp[0]);
            }
        }
        return s1;
    }

    protected string Cyuno_withCanon(int j, string[] lines)
    {
        string line = string.Empty;
        line = lines[j];
        string[] sp = new string[100];
        string[] spliter = new string[] { " " };
        sp = line.Split(spliter, 0);
        string startCYUNO = sp[0].Substring(3, 14);
        string endCYUNO = startCYUNO.Substring(0, 11) + startCYUNO.Substring(13);
        return endCYUNO;
    }
    protected void arrCYUNO(int i, int position, string[] lines, ArrayList s2, ArrayList s4)
    {
        string line = string.Empty;
        for (int p = 0; p < 10; p++)
        {
            int j = (int)(s2.get(i)) + position + p;
            line = lines[j];
            string[] sp = new string[100];
            string[] spliter = new string[] { " " };
            sp = line.Split(spliter, 0);
            string checkString = sp[0].ToString();
            if (checkString.Length > 13)
            {
                string endCYUNO = Cyuno_withCanon(j, lines);
                s4.add(endCYUNO);
                break;
            }
        }
    }
    protected ArrayList dataCyuno_withCanon(ArrayList s1, ArrayList s2, ArrayList s3, string sourfile, string outText)
    {
        string[] lines = TxtData(sourfile, outText);
        ArrayList s4 = new ArrayList();
        string line = string.Empty;
        for (int i = 0; i < s1.size(); i++)
        {
            if (s1.get(i).ToString().Trim() == "キヤノンセミコンダクターエクィ")
            {
                if (s3.get(i).ToString() == "0")
                {
                    arrCYUNO(i, 46, lines, s2, s4);
                }
                else if (s3.get(i).ToString() == "2")
                {
                    arrCYUNO(i, 27, lines, s2, s4);
                }
                else
                {
                    arrCYUNO(i, 41, lines, s2, s4);
                }
            }
            else
            {
                if (s3.get(i).ToString() == "0")
                {
                    arrCYUNO(i, 45, lines, s2, s4);
                }
                else if (s3.get(i).ToString() == "2")
                {
                    arrCYUNO(i, 26, lines, s2, s4);
                }
                else
                {
                    arrCYUNO(i, 39, lines, s2, s4);
                }
            }
        }

        return s4;
    }
    //doc tu file txt de lay du lieu
    protected string[] TxtData(string sourfile, string outText)
    {
        string s = ExtractTextFromPdf(sourfile);
        StreamWriter gh = new StreamWriter(outText);
        gh.Flush();
        gh.WriteLine(s);
        gh.Close();
        string[] lines = File.ReadAllLines(outText);
        return lines;
    }

    //chuyen du lieu tu pdf sang txt de lay du lieu
    private static string ExtractTextFromPdf(string path)
    {
        PDDocument doc = null;
        try
        {
            doc = PDDocument.load(path);
            PDFTextStripper stripper = new PDFTextStripper();
            return stripper.getText(doc);
        }
        finally
        {
            if (doc != null)
            {
                doc.close();
            }
        }
    }
    //ham lay size cua pdf
    protected iTextSharp.text.Rectangle getHalfPageSize(iTextSharp.text.Rectangle pagesize)
    {
        float width = pagesize.Width;
        float height = pagesize.Height;
        return new iTextSharp.text.Rectangle(width, height);
    }

    //chia nho tu pdf to thanh pdf nho voi ten co san

    //day la cat nguyen trang khong chia
    protected void cutPage_withNIKON(string sourceFileName, string newFileName, int numberpage)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(pdfStream);
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            iTextSharp.text.Rectangle mediabox = new iTextSharp.text.Rectangle(getHalfPageSize(reader.GetPageSizeWithRotation(1)));
            PdfImportedPage page = writer.GetImportedPage(reader, numberpage);

            content.AddTemplate(page, 0, 0);
            document.SetPageSize(mediabox);
            document.NewPage();
            document.Close();
            reader.Close();
        }
    }

    //day la cach cat 2 phan
    protected void cutHalfpage_withBC(string sourceFileName, string newFileName, int numberpage, int updown)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(pdfStream);
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            iTextSharp.text.Rectangle mediabox = new iTextSharp.text.Rectangle(getHalfPageSize(reader.GetPageSizeWithRotation(1)));
            PdfImportedPage page = writer.GetImportedPage(reader, numberpage);

            if (updown == 1)
            {
                content.AddTemplate(page, 0, 0);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height / 2, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else
            {
                content.AddTemplate(page, 0, reader.GetPageSizeWithRotation(1).Height / 2 - 13);
                document.SetPageSize(mediabox);
                document.NewPage();
            }

            document.Close();
            reader.Close();
        }
    }

    protected void cutHalfpage_withCMS(string sourceFileName, string newFileName, int numberpage, int updown)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(pdfStream);
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            iTextSharp.text.Rectangle mediabox = new iTextSharp.text.Rectangle(getHalfPageSize(reader.GetPageSizeWithRotation(1)));
            PdfImportedPage page = writer.GetImportedPage(reader, numberpage);

            if (updown == 1)
            {
                content.AddTemplate(page, 0, 19);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 19);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 27, reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height / 2, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else
            {
                content.AddTemplate(page, 0, reader.GetPageSizeWithRotation(1).Height / 2 + 11);
                document.SetPageSize(mediabox);
                document.NewPage();
            }

            document.Close();
            reader.Close();
        }
    }

    protected void cutHalfpage_withCanon(string sourceFileName, string newFileName, int numberpage, int updown)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(pdfStream);
            Document document = new Document(reader.GetPageSizeWithRotation(1));
            PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            iTextSharp.text.Rectangle mediabox = new iTextSharp.text.Rectangle(getHalfPageSize(reader.GetPageSizeWithRotation(1)));
            PdfImportedPage page = writer.GetImportedPage(reader, numberpage);

            if (updown == 1)
            {
                content.AddTemplate(page, 0, 0);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height / 2, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else
            {
                content.AddTemplate(page, 0, reader.GetPageSizeWithRotation(1).Height / 2);
                document.SetPageSize(mediabox);
                document.NewPage();
            }

            document.Close();
            reader.Close();
        }
    }
    //day cach cat lam 6 phan
    protected void cutSixPartpage_withCanon(string sourceFileName, string newFileName, int numberpage, int position)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(pdfStream);
            Document document = new Document(reader.GetPageSizeWithRotation(numberpage));
            PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            iTextSharp.text.Rectangle mediabox = new iTextSharp.text.Rectangle(getHalfPageSize(reader.GetPageSizeWithRotation(numberpage)));
            PdfImportedPage page = writer.GetImportedPage(reader, numberpage);

            if (position == 1)
            {
                content.AddTemplate(page, -2, 0);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(-2, 0, reader.GetPageSizeWithRotation(numberpage).Width + 2, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 10, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else if (position == 2)
            {
                content.AddTemplate(page, reader.GetPageSizeWithRotation(numberpage).Width / 2 * (-1), 0);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(-2, 0, reader.GetPageSizeWithRotation(numberpage).Width + 2, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 10, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else if (position == 3)
            {
                content.AddTemplate(page, -2, reader.GetPageSizeWithRotation(numberpage).Height / 3 + 8);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(-2, 0, reader.GetPageSizeWithRotation(numberpage).Width + 2, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 10, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else if (position == 4)
            {
                content.AddTemplate(page, reader.GetPageSizeWithRotation(numberpage).Width / 2 * (-1), reader.GetPageSizeWithRotation(numberpage).Height / 3 + 8);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(-2, 0, reader.GetPageSizeWithRotation(numberpage).Width + 2, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 10, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else if (position == 5)
            {
                content.AddTemplate(page, -2, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 15);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSizeWithRotation(numberpage).Width, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            else if (position == 6)
            {
                content.AddTemplate(page, reader.GetPageSizeWithRotation(numberpage).Width / 2 * (-1), reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3 + 15);
                content.AddTemplate(content.CreateTemplate(100, 100), 0, 0);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSizeWithRotation(numberpage).Width, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                content.RoundRectangle(reader.GetPageSizeWithRotation(numberpage).Width / 2 - 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, reader.GetPageSizeWithRotation(numberpage).Width / 2 + 10, reader.GetPageSizeWithRotation(numberpage).Height * 2 / 3, 1);
                content.Fill();
                document.SetPageSize(mediabox);
                document.NewPage();
            }
            document.Close();
            reader.Close();
        }
    }
    ////////////////////////////////////////////////////////////////////////
    ////////////////////xong phan chia nho pdf thanh file nho de de quan li 
    ////////////////////////////////////////////////////////////////////////
    ///////////////////////////////////////////////////////////////////////

    //lay thong ti tu bang de them vao pdf
    [WebMethod]
    public static string GetPdf_withBC(string[][] empdetails)
    {
        // List<string> emp = new List<string>();
        string alert = "";
        ArrayList lstFiles = new ArrayList();
        ArrayList tana = new ArrayList();
        ArrayList kmno = new ArrayList();
        ArrayList jyuno = new ArrayList();
        ArrayList lstcyuno = new ArrayList();
        ArrayList lstkz = new ArrayList();
        ArrayList zaikou = new ArrayList();
        ArrayList recode = new ArrayList();
        string id = "";
        string repeat = "";
        for (int i = 0; i < empdetails.Length; i++)
        {
            string cyuno = empdetails[i][0].ToString().Trim();
            id = empdetails[i][4].ToString().Trim();
            repeat = empdetails[i][5].ToString().Trim();
            if (repeat == "1")
            {
                foreach (string file in Directory.GetFiles(@"\\10.121.21.2\data\DeliveryNote\BC\", "*.pdf"))
                {
                    string namefile = file.Substring(35, file.Length - 39);
                    if (namefile.Trim().ToString() == cyuno.Trim() || namefile.Trim().ToString() == cyuno.Trim() + "-1")
                    {
                        lstFiles.add(file);
                        lstcyuno.add(cyuno);
                        string tn = empdetails[i][1].ToString();
                        string cy = empdetails[i][2].ToString();
                        string jy = empdetails[i][3].ToString();
                        string zk = empdetails[i][10].ToString();
                        string rd = empdetails[i][11].ToString();
                        recode.add(rd);
                        zaikou.add(zk);
                        lstkz.add(id);
                        tana.add(tn);
                        kmno.add(cy);
                        jyuno.add(jy);
                    }

                }
            }
            else if (repeat == "2")
            {
                string file = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\" + cyuno + @".pdf";
                //emp.Add(file);
                lstFiles.add(file);
                lstcyuno.add(cyuno);
                string tn = empdetails[i][1].ToString();
                string cy = empdetails[i][2].ToString();
                string jy = empdetails[i][3].ToString();
                string zk = empdetails[i][10].ToString();
                string rd = empdetails[i][11].ToString();
                recode.add(rd);
                zaikou.add(zk);
                lstkz.add(id);
                tana.add(tn);
                kmno.add(cy);
                jyuno.add(jy);
                string file1 = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\" + cyuno + @"-1.pdf";
                if (File.Exists(file1))
                {
                    lstFiles.add(file1);
                    lstcyuno.add(cyuno);
                    lstkz.add(id);
                    tana.add(tn);
                    kmno.add(cy);
                    jyuno.add(jy);
                    zaikou.add(zk);
                    recode.add(rd);
                }
            }

        }
        if (lstFiles.size() > 0)
        {
            string newFileName = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\testpdf.pdf";
            mergerHalfpdf(newFileName, lstFiles, repeat, "11", 0, 5);
            string newFileName11 = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\endpdf.pdf";
            InsertTextToPdf_withBC(newFileName, newFileName11, tana, kmno, zaikou, repeat, recode);
            string idFactory = empdetails[0][4].ToString();
            updateDatabase(lstkz, jyuno);
            alert = "1";
        }
        else
        {
            alert = "0";
        }
        return alert;
    }

    [WebMethod]
    public static string GetPdf_withCMS(string[][] empdetails)
    {
        string alert = "";
        ArrayList lstFiles = new ArrayList();
        ArrayList tana = new ArrayList();
        ArrayList kmno = new ArrayList();
        ArrayList jyuno = new ArrayList();
        ArrayList lstcyuno = new ArrayList();
        ArrayList nouki = new ArrayList();
        ArrayList lstkz = new ArrayList();
        ArrayList zaikou = new ArrayList();
        ArrayList recode = new ArrayList();
        string id = "";
        string repeat = "";
        for (int i = 0; i < empdetails.Length; i++)
        {
            string cyuno = empdetails[i][0].ToString().Trim();
            id = empdetails[i][4].ToString().Trim();
            repeat = empdetails[i][5].ToString().Trim();
            if (repeat == "1")
            {
                foreach (string file in Directory.GetFiles(@"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\", "*.pdf"))
                {
                    string namefile = file.Substring(41, file.Length - 45);
                    if (namefile.Trim().ToString() == cyuno.Trim())
                    {
                        lstFiles.add(file);
                        lstcyuno.add(cyuno);
                        string tn = empdetails[i][1].ToString();
                        string cy = empdetails[i][2].ToString();
                        string jy = empdetails[i][3].ToString();
                        string nk = empdetails[i][6].ToString();
                        string rd = empdetails[i][11].ToString();
                        recode.add(rd);
                        nk = nk.Substring(2, 2) + "-" + nk.Substring(5, 2) + "-" + nk.Substring(8, 2);
                        string zk = empdetails[i][10].ToString();
                        zaikou.add(zk);
                        nouki.add(nk);
                        tana.add(tn);
                        kmno.add(cy);
                        jyuno.add(jy);
                        lstkz.add(empdetails[i][7].ToString());
                        break;
                    }

                }
            }
            else if (repeat == "2")
            {
                string file = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\" + cyuno + @".pdf";
                //emp.Add(file);
                lstFiles.add(file);
                lstcyuno.add(cyuno);
                string tn = empdetails[i][1].ToString();
                string cy = empdetails[i][2].ToString();
                string jy = empdetails[i][3].ToString();
                string nk = empdetails[i][6].ToString();
                string rd = empdetails[i][11].ToString();
                recode.add(rd);
                nk = nk.Substring(2, 2) + "-" + nk.Substring(5, 2) + "-" + nk.Substring(8, 2);
                string zk = empdetails[i][10].ToString();
                zaikou.add(zk);
                nouki.add(nk);
                tana.add(tn);
                kmno.add(cy);
                jyuno.add(jy);
                lstkz.add(empdetails[i][7].ToString());
            }

        }
        if (lstFiles.size() > 0)
        {
            string newFileName = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\testpdf.pdf";
            mergerHalfpdf(newFileName, lstFiles, repeat, "21", -18, -12);
            string newFileName11 = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\endpdf.pdf";
            InsertTextToPdf_withCMS(newFileName, newFileName11, tana, kmno, zaikou, nouki, lstkz, repeat, recode);
            string idFactory = empdetails[0][4].ToString();
            updateDatabase(lstkz, jyuno);

            alert = "1";
        }
        else
        {
            alert = "0";
        }
        return alert;
    }

    [WebMethod]
    public static string GetPdf_withCanon(string[][] empdetails)
    {
        string alert = "";
        ArrayList lstFiles = new ArrayList();
        ArrayList lstFiles1 = new ArrayList();
        ArrayList tana = new ArrayList();
        ArrayList kmno = new ArrayList();
        ArrayList jyuno = new ArrayList();
        ArrayList lstcyuno = new ArrayList();
        ArrayList nouki = new ArrayList();
        ArrayList lstkz = new ArrayList();
        ArrayList tokcd = new ArrayList();
        ArrayList recode = new ArrayList();
        string id = "";
        string repeat = "";
        for (int i = 0; i < empdetails.Length; i++)
        {
            string cyuno = empdetails[i][0].ToString().Trim();
            id = empdetails[i][4].ToString().Trim();
            repeat = empdetails[i][5].ToString().Trim();
            if (repeat == "1")
            {
                foreach (string file in Directory.GetFiles(@"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\", "*.pdf"))
                {
                    string namefile = file.Substring(47, file.Length - 51);
                    if (namefile.Trim().ToString() == cyuno.Trim())
                    {
                        lstFiles.add(file);
                        lstFiles1.add(@"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\" + namefile + ".pdf");
                        lstcyuno.add(cyuno);
                        string tn = empdetails[i][1].ToString();
                        string cy = empdetails[i][2].ToString();
                        string jy = empdetails[i][3].ToString();
                        string nk = empdetails[i][6].ToString();
                        string rd = empdetails[i][11].ToString();
                        recode.add(rd);
                        nk = nk.Substring(2, 2) + "-" + nk.Substring(5, 2) + "-" + nk.Substring(8, 2);
                        nouki.add(nk);
                        tana.add(tn);
                        kmno.add(cy);
                        jyuno.add(jy);
                        lstkz.add(empdetails[i][7].ToString());
                        tokcd.add(empdetails[i][9].ToString());
                        break;
                    }

                }
            }
            else if (repeat == "2")
            {
                string file = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\" + cyuno + @".pdf";
                string file1 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\Processed\" + cyuno + @".pdf";
                lstFiles.add(file);
                lstFiles1.add(file1);
                lstcyuno.add(cyuno);
                string tn = empdetails[i][1].ToString();
                string cy = empdetails[i][2].ToString();
                string jy = empdetails[i][3].ToString();
                string nk = empdetails[i][6].ToString();
                string rd = empdetails[i][11].ToString();
                recode.add(rd);
                nk = nk.Substring(2, 2) + "-" + nk.Substring(5, 2) + "-" + nk.Substring(8, 2);
                nouki.add(nk);
                tana.add(tn);
                kmno.add(cy);
                jyuno.add(jy);
                lstkz.add(empdetails[i][7].ToString());
                tokcd.add(empdetails[i][9].ToString());
            }
        }
        if (lstFiles.size() > 0)
        {
            // ArrayList zaikou = GetDataPO(lstcyuno, id, lstkz);//cai nay dung check kho co ton tai hay khong , de them thong tin o day thi khong can thiet
            int leng1 = 0; int leng2 = 0; int leng3 = 0; int leng4 = 0;
            for (int i = 0; i < tokcd.size(); i++)
            {
                if (tokcd.get(i).ToString().Trim() == "00012")
                {
                    leng1 += 1;
                }
                else if (tokcd.get(i).ToString().Trim() == "00095")
                {
                    leng2 += 1;
                }
                else if (tokcd.get(i).ToString().Trim() == "00099")
                {
                    leng3 += 1;
                }
                else if (tokcd.get(i).ToString().Trim() == "00207")
                {
                    leng4 += 1;
                }
            }

            ArrayList sPage = new ArrayList();
            sPage.add(leng1);
            sPage.add(leng2);
            sPage.add(leng3);
            sPage.add(leng4);

            string newFileName = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\testpdf.pdf";
            mergerHalfpdfCanon(newFileName, lstFiles, repeat, "22", 0, 0, sPage);

            string newFileName11 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\endpdf.pdf";
            InsertTextToPdf_withCanon(newFileName, newFileName11, tana, kmno, nouki, lstkz, sPage, repeat, recode);

            string newFileName1 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\Processed\testpdf.pdf";
            mergerSixApartpdf(newFileName1, lstFiles1, repeat, "22", sPage);
            // string newFileName22 = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\Processed\endpdf.pdf";//cai nay canon khong can vi khong them thong tin o phieu xuat
            // InsertTextToPdf_SixApart_withCanon(newFileName1, newFileName22, tana, kmno, zaikou, nouki, lstkz);
            string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\Canon\new.pdf";
            ArrayList lstMergerFile = new ArrayList();
            lstMergerFile.add(newFileName11);
            lstMergerFile.add(newFileName1);
            addMergerPdf(lstMergerFile, outputPdfPath);
            string idFactory = empdetails[0][4].ToString();
            updateDatabase(lstkz, jyuno);
            alert = "1";
        }
        else
        {
            alert = "0";
        }
        return alert;
    }

    [WebMethod]
    public static string GetPdf_withNIKON(string[][] empdetails)
    {
        string alert = "";
        ArrayList lstFiles = new ArrayList();
        ArrayList kazu = new ArrayList();
        ArrayList nouki = new ArrayList();
        ArrayList jyuno = new ArrayList();
        ArrayList lstcyuno = new ArrayList();
        ArrayList kmno = new ArrayList();
        ArrayList lstkz = new ArrayList();
        ArrayList zaikou = new ArrayList();
        ArrayList recode = new ArrayList();
        string id = "";
        string repeat = "";
        for (int i = 0; i < empdetails.Length; i++)
        {
            string cyuno = empdetails[i][0].ToString().Trim();
            id = empdetails[i][4].ToString().Trim();
            repeat = empdetails[i][5].ToString().Trim();
            if (repeat == "1")
            {
                foreach (string file in Directory.GetFiles(@"\\10.121.21.2\data\DeliveryNote\NIKON\", "*.pdf"))
                {
                    string namefile = file.Substring(38, file.Length - 42);
                    if (namefile.Trim().ToString() == cyuno.Trim())
                    {
                        lstFiles.add(file);
                        lstcyuno.add(cyuno);
                        string tn = empdetails[i][8].ToString();
                        string jy = empdetails[i][3].ToString();
                        string cy = empdetails[i][2].ToString();
                        string nk = empdetails[i][6].ToString();
                        string zk = empdetails[i][10].ToString();
                        string rd = empdetails[i][11].ToString();
                        recode.add(rd);
                        zaikou.add(zk);
                        nouki.add(nk);
                        kmno.add(cy);
                        kazu.add(tn);
                        jyuno.add(jy);
                        lstkz.add(empdetails[i][7].ToString());
                        break;
                    }

                }
            }
            else if (repeat == "2")
            {
                string file = @"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\" + cyuno + @".pdf";
                //emp.Add(file);
                lstFiles.add(file);
                lstcyuno.add(cyuno);
                string tn = empdetails[i][8].ToString();
                string jy = empdetails[i][3].ToString();
                string cy = empdetails[i][2].ToString();
                string nk = empdetails[i][6].ToString();
                string zk = empdetails[i][10].ToString();
                string rd = empdetails[i][11].ToString();
                recode.add(rd);
                zaikou.add(zk);
                nouki.add(nk);
                kmno.add(cy);
                kazu.add(tn);
                jyuno.add(jy);
                lstkz.add(empdetails[i][7].ToString());
            }

        }
        if (lstFiles.size() > 0)
        {

            try
            {
                string newFileName = @"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\testpdf.pdf";
                mergerPdf(newFileName, lstFiles, repeat, "23", 0, 0);
                string newFileName11 = @"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\endpdf.pdf";
                InsertTextToPdf_withNIKON(newFileName, newFileName11, kazu, lstkz, kmno, nouki, zaikou, repeat, recode);
                string idFactory = empdetails[0][4].ToString();
                updateDatabase(lstkz, jyuno);
                alert = "1";
            }
            catch (Exception ex)
            {
                alert = "0";
            }


        }
        else
        {
            alert = "0";
        }
        return alert;
    }

    //gop 2 hay nhieu pdf lai voi nha
    [WebMethod]
    public static void addMergerPdf(ArrayList lstFiles, string outputPdfPath)
    {
        //merger into one file
        PdfReader reader = null;
        Document sourceDocument = null;
        PdfCopy pdfCopyProvider = null;
        PdfImportedPage importedPage;
        DateTime now = DateTime.Now;
        string time = now.ToString("HHmmss");
        sourceDocument = new Document();
        pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
        //Open the output file
        sourceDocument.Open();
        //Loop through the files list
        for (int f = 0; f < lstFiles.size(); f++)
        {
            int pages = get_pageCcount(lstFiles.get(f).ToString());
            reader = new PdfReader(lstFiles.get(f).ToString());
            for (int i = 1; i <= pages; i++)
            {
                importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                pdfCopyProvider.AddPage(importedPage);
            }
            reader.Close();
        }
        sourceDocument.Close();
    }
    //dem so trang de gop pdf
    [WebMethod]
    public static int get_pageCcount(string file)
    {
        using (StreamReader sr = new StreamReader(File.OpenRead(file)))
        {
            Regex regex = new Regex(@"/Type\s*/Page[^s]");
            MatchCollection matches = regex.Matches(sr.ReadToEnd());
            return matches.Count;
        }
    }

    //them vao co so du lieu sau khi in
    [WebMethod]
    public static void updateDatabase(ArrayList lstkz, ArrayList jyuno)
    {
        SqlConnection sqlconn = new SqlConnection();

        for (int i = 0; i < jyuno.size(); i++)
        {
            string idFactory = lstkz.get(i).ToString();
            if (idFactory == "1")
            {
                sqlconn = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
            }
            else if (idFactory == "2")
            {
                sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
            }
            sqlconn.Open();
            string sqlquery =
   "  If EXISTS (SELECT 1  FROM [TESCex].[dbo].[D1000exLog] where [JYUNO]='" + jyuno.get(i).ToString() + "' )" +
"  BEGIN " +
" insert into [TESCex].[dbo].[D1000exLog]([JYUNO],[PrintStatus],[PrintGroup],[PrintRowOrder],[PrintDate]) " +
" SELECT  '" + jyuno.get(i).ToString() + "',( SELECT Count([JYUNO]) FROM [TESCex].[dbo].[D1000exLog] where [JYUNO]='" + jyuno.get(i).ToString() + "' ),'0', '0',GETDATE(); " +
" Update [TESCex].[dbo].[D1000ex] set [PrintDate]=GETDATE(), [PrintStatus]='1' where [JYUNO]='" + jyuno.get(i).ToString() + "'" +
" END " +
" ELSE " +
" BEGIN " +
" Insert into [TESCex].[dbo].[D1000exLog]([JYUNO],[PrintStatus],[PrintGroup],[PrintRowOrder],[PrintDate]) values('" + jyuno.get(i).ToString() + "','0','0','0',GETDATE()); " +
" Insert into [TESCex].[dbo].[D1000ex]([JYUNO],[PrintStatus],[PrintGroup],[PrintRowOrder],[PrintDate],[PrintDateFirst]) values('" + jyuno.get(i).ToString() + "','0','0','0',GETDATE(),GETDATE()) " +
" END";

            SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);

            sqlcomn.ExecuteNonQuery();
            sqlconn.Close();
        }
    }

    //tap hop cac file pdf nho thanh 1 file voi 1 phan duy nhat
    //tap hop cac file pdf nho thanh 1 file voi 2 phan trong 1 trang
    [WebMethod]
    public static void mergerHalfpdfCanon(string newFileName1, ArrayList lstFiles1, string repeat, string companyTokcd, float distanceFirst, float distanceLast, ArrayList sPage)
    {
        int k = 0;
        ArrayList connectFile = new ArrayList();
        for (int i = 0; i < sPage.size(); i++)
        {
            if ((int)sPage.get(i) == 0)
            {
                k += 1;
                continue;
            }
            else
            {
                string newFileName = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\testpdf" + k.ToString() + ".pdf";
                connectFile.add(newFileName);
                k += 1;
                Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite);
                PdfReader reader0 = new PdfReader(@"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\Test\tuananh.pdf");
                Document document = new Document(reader0.GetPageSizeWithRotation(1));
                PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
                document.Open();

                ArrayList lstFiles = new ArrayList();
                int start = 0;
                for (int q = 0; q < i; q++)
                {
                    start += (int)sPage.get(q);
                }
                for (int p = start; p < start + (int)sPage.get(i); p++)
                {
                    lstFiles.add(lstFiles1.get(p));
                }

                int filepage = 0;
                if (lstFiles.size() % 2 == 0)
                {
                    filepage = lstFiles.size() / 2;
                }
                else
                {
                    filepage = (lstFiles.size() + 1) / 2;
                }
                for (int f = 1; f <= filepage; f++)
                {
                    if (f == filepage)
                    {
                        string file = lstFiles.get(f * 2 - 2).ToString();
                        PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                        PdfImportedPage page = writer.GetImportedPage(reader, 1);
                        writer.DirectContentUnder.AddTemplate(page, 0, distanceFirst);
                        if (lstFiles.size() % 2 == 0)
                        {
                            string file1 = lstFiles.get(f * 2 - 1).ToString();
                            PdfReader reader1 = new PdfReader(System.IO.File.ReadAllBytes(file1));
                            PdfImportedPage page1 = writer.GetImportedPage(reader1, 1);
                            writer.DirectContentUnder.AddTemplate(page1, 0, reader1.GetPageSizeWithRotation(1).Height / 2 * (-1) + distanceLast);
                        }
                    }
                    else
                    {
                        string file = lstFiles.get(f * 2 - 2).ToString();
                        string file1 = lstFiles.get(f * 2 - 1).ToString();
                        PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                        PdfImportedPage page = writer.GetImportedPage(reader, 1);
                        writer.DirectContentUnder.AddTemplate(page, 0, distanceFirst);
                        PdfReader reader1 = new PdfReader(System.IO.File.ReadAllBytes(file1));
                        PdfImportedPage page1 = writer.GetImportedPage(reader1, 1);
                        writer.DirectContentUnder.AddTemplate(page1, 0, reader1.GetPageSizeWithRotation(1).Height / 2 * (-1) + distanceLast);
                        document.SetPageSize(reader0.GetPageSizeWithRotation(1));
                        document.NewPage();
                    }
                }
                document.Close();

            }
        }
        string output = newFileName1;
        addMergerPdf(connectFile, output);

        if (repeat == "1")
        {
            if (companyTokcd == "22")
            {
                for (int i = 0; i < lstFiles1.size(); i++)
                {
                    string deletenamefile = lstFiles1.get(i).ToString();
                    string namefile = deletenamefile.Substring(47);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\" + namefile;
                    File.Copy(deletenamefile, destine);
                    File.Delete(deletenamefile);
                }
            }
        }
    }


    [WebMethod]
    public static void mergerPdf(string newFileName, ArrayList lstFiles, string repeat, string companyTokcd, float distanceFirst, float distanceLast)
    {

        Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite);
        PdfReader reader0 = new PdfReader(@"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\Test\tuananh1.pdf");
        Document document = new Document(reader0.GetPageSizeWithRotation(1));
        PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
        document.Open();

        for (int f = 1; f <= lstFiles.size(); f++)
        {
            string file = lstFiles.get(f - 1).ToString();
            PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
            PdfImportedPage page = writer.GetImportedPage(reader, 1);
            writer.DirectContentUnder.AddTemplate(page, 0, distanceFirst);
            document.SetPageSize(reader0.GetPageSizeWithRotation(1));
            document.NewPage();

        }
        document.Close();

        if (repeat == "1")
        {
            if (companyTokcd == "23")
            {
                for (int i = 0; i < lstFiles.size(); i++)
                {
                    string deletenamefile = lstFiles.get(i).ToString();
                    string namefile = deletenamefile.Substring(38);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\" + namefile;
                    File.Copy(deletenamefile, destine);
                    File.Delete(deletenamefile);
                }
            }
        }
    }


    //tap hop cac file pdf nho thanh 1 file voi 2 phan trong 1 trang
    [WebMethod]
    public static void mergerHalfpdf(string newFileName, ArrayList lstFiles, string repeat, string companyTokcd, float distanceFirst, float distanceLast)
    {

        Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite);
        PdfReader reader0 = new PdfReader(@"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\Test\tuananh.pdf");
        Document document = new Document(reader0.GetPageSizeWithRotation(1));
        PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
        document.Open();
        int filepage = 0;
        if (lstFiles.size() % 2 == 0)
        {
            filepage = lstFiles.size() / 2;
        }
        else
        {
            filepage = (lstFiles.size() + 1) / 2;
        }
        for (int f = 1; f <= filepage; f++)
        {
            if (f == filepage)
            {
                string file = lstFiles.get(f * 2 - 2).ToString();
                PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                writer.DirectContentUnder.AddTemplate(page, 0, distanceFirst);
                if (lstFiles.size() % 2 == 0)
                {
                    string file1 = lstFiles.get(f * 2 - 1).ToString();
                    PdfReader reader1 = new PdfReader(System.IO.File.ReadAllBytes(file1));
                    PdfImportedPage page1 = writer.GetImportedPage(reader1, 1);
                    writer.DirectContentUnder.AddTemplate(page1, 0, reader1.GetPageSizeWithRotation(1).Height / 2 * (-1) + distanceLast);
                }
            }
            else
            {
                string file = lstFiles.get(f * 2 - 2).ToString();
                string file1 = lstFiles.get(f * 2 - 1).ToString();
                PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                writer.DirectContentUnder.AddTemplate(page, 0, distanceFirst);
                PdfReader reader1 = new PdfReader(System.IO.File.ReadAllBytes(file1));
                PdfImportedPage page1 = writer.GetImportedPage(reader1, 1);
                writer.DirectContentUnder.AddTemplate(page1, 0, reader1.GetPageSizeWithRotation(1).Height / 2 * (-1) + distanceLast);
                document.SetPageSize(reader0.GetPageSizeWithRotation(1));
                document.NewPage();
            }
        }
        document.Close();

        if (repeat == "1")
        {
            if (companyTokcd == "11")
            {
                for (int i = 0; i < lstFiles.size(); i++)
                {
                    string deletenamefile = lstFiles.get(i).ToString();
                    string namefile = deletenamefile.Substring(35);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\" + namefile;
                    //  string checkPath = @"\\10.121.21.2\data\DeliveryNote\BC1\FixFile1\" + namefile;
                    File.Copy(deletenamefile, destine);
                    //  File.Copy(deletenamefile, checkPath);
                    File.Delete(deletenamefile);
                }
            }
            else if (companyTokcd == "21")
            {
                for (int i = 0; i < lstFiles.size(); i++)
                {
                    string deletenamefile = lstFiles.get(i).ToString();
                    string namefile = deletenamefile.Substring(41);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\" + namefile;
                    File.Copy(deletenamefile, destine);
                    File.Delete(deletenamefile);
                }
            }
            else if (companyTokcd == "22")
            {
                for (int i = 0; i < lstFiles.size(); i++)
                {
                    string deletenamefile = lstFiles.get(i).ToString();
                    string namefile = deletenamefile.Substring(47);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\" + namefile;
                    File.Copy(deletenamefile, destine);
                    File.Delete(deletenamefile);
                }
            }
        }

    }

    //tap hop file voi 6 phan trong 1 trang
    [WebMethod]
    public static void mergerSixApartpdf(string newFileName11, ArrayList lstFiles1, string repeat, string companyTokcd, ArrayList sPage)
    {
        if (companyTokcd == "22")
        {
            ArrayList connectFile = new ArrayList();
            int k = 0;
            for (int i = 0; i < sPage.size(); i++)
            {
                if ((int)sPage.get(i) == 0)
                {
                    k += 1;
                    continue;
                }
                else
                {
                    string newFileName = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\Processed\testpdf" + k.ToString() + ".pdf";
                    connectFile.add(newFileName);
                    k += 1;
                    Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite);
                    PdfReader reader0 = new PdfReader(@"\\LD-FUJINOMIYA\DateCentre\SrvWork\PdfEdit\Test\tuananh1.pdf");
                    Document document = new Document(reader0.GetPageSizeWithRotation(1));
                    PdfWriter writer = PdfWriter.GetInstance(document, newpdfStream);
                    document.Open();
                    ArrayList lstFiles = new ArrayList();
                    int start = 0;
                    for (int q = 0; q < i; q++)
                    {
                        start += (int)sPage.get(q);
                    }
                    for (int p = start; p < start + (int)sPage.get(i); p++)
                    {
                        lstFiles.add(lstFiles1.get(p));
                    }

                    int filepage = 0;
                    if ((int)sPage.get(i) % 6 == 0)
                    {
                        filepage = (int)sPage.get(i) / 6;
                    }
                    else
                    {
                        filepage = (int)((int)sPage.get(i) / 6) + 1;
                    }
                    for (int f = 1; f <= filepage; f++)
                    {

                        if (f == filepage)
                        {
                            var p = lstFiles.size() % 6;
                            int max = (int)p;
                            if (max == 0)
                            {
                                max = 6;
                            }
                            for (int j = 6; j > 6 - max; j--)
                            {
                                string file = lstFiles.get(f * 6 - j).ToString();
                                PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                                float x0 = 0;
                                float y0 = 0;
                                if (j % 2 == 0)
                                {
                                    if (j == 6)
                                    {
                                        x0 = 0;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6 - 3;
                                    }
                                    else if (j == 4)
                                    {
                                        x0 = 0;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6 - 8;
                                    }
                                    else
                                    {
                                        x0 = 0;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6 - 13;
                                    }
                                }
                                else
                                {
                                    if (j == 5)
                                    {
                                        x0 = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6 - 3;
                                    }
                                    else if (j == 3)
                                    {
                                        x0 = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6 - 8;
                                    }
                                    else
                                    {
                                        x0 = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y0 = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6 - 13;
                                    }

                                    PdfContentByte cb = writer.DirectContent;
                                    cb.MoveTo(reader.GetPageSizeWithRotation(1).Width / 2, 0);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width / 2, reader.GetPageSizeWithRotation(1).Height);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();

                                    cb.MoveTo(0, reader.GetPageSizeWithRotation(1).Height / 3 - 4);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height / 3 - 4);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();

                                    cb.MoveTo(0, reader.GetPageSizeWithRotation(1).Height * 2 / 3 + 4);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height * 2 / 3 + 4);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();
                                }
                                writer.DirectContentUnder.AddTemplate(page, x0, y0);
                            }
                        }
                        else
                        {
                            for (int j = 6; j > 0; j--)
                            {
                                string file = lstFiles.get(f * 6 - j).ToString();
                                PdfReader reader = new PdfReader(System.IO.File.ReadAllBytes(file));
                                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                                float x = 0;
                                float y = 0;
                                if (j % 2 == 0)
                                {
                                    if (j == 6)
                                    {
                                        x = 0;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6;
                                    }
                                    else if (j == 4)
                                    {
                                        x = 0;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6 - 8;
                                    }
                                    else
                                    {
                                        x = 0;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j) / 6 - 13;
                                    }

                                }
                                else
                                {
                                    if (j == 5)
                                    {
                                        x = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6;
                                    }
                                    else if (j == 3)
                                    {
                                        x = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6 - 8;
                                    }
                                    else
                                    {
                                        x = reader.GetPageSizeWithRotation(1).Width / 2;
                                        y = reader.GetPageSizeWithRotation(1).Height * (-1) * (6 - j - 1) / 6 - 13;
                                    }


                                    PdfContentByte cb = writer.DirectContent;
                                    cb.MoveTo(reader.GetPageSizeWithRotation(1).Width / 2, 0);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width / 2, reader.GetPageSizeWithRotation(1).Height);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();
                                    cb.MoveTo(0, reader.GetPageSizeWithRotation(1).Height / 3 - 4);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height / 3 - 4);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();
                                    cb.MoveTo(0, reader.GetPageSizeWithRotation(1).Height * 2 / 3 + 4);
                                    cb.LineTo(reader.GetPageSizeWithRotation(1).Width, reader.GetPageSizeWithRotation(1).Height * 2 / 3 + 4);
                                    cb.SetLineDash(5f, 5f);
                                    cb.Stroke();
                                }
                                writer.DirectContentUnder.AddTemplate(page, x, y);
                            }

                            document.SetPageSize(reader0.GetPageSizeWithRotation(1));
                            document.NewPage();
                        }
                    }
                    document.Close();
                }
            }
            addMergerPdf(connectFile, newFileName11);
            if (repeat == "1")
            {
                for (int i = 0; i < lstFiles1.size(); i++)
                {
                    string deletenamefile = lstFiles1.get(i).ToString();
                    string namefile = deletenamefile.Substring(47);
                    string destine = @"\\10.121.21.2\data\DeliveryNote\Canon\CanonGPH\Processed\" + namefile;
                    File.Copy(deletenamefile, destine);
                    File.Delete(deletenamefile);
                }
            }
        }
    }
    //them cac thong tin can thiet vao file moi duoc tap hop lai
    [WebMethod]
    public static void InsertTextToPdf_withBC(string sourceFileName, string newFileName, ArrayList tana, ArrayList kmno, ArrayList zaikou, string repeat, ArrayList recode)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            if (tana.size() == 0)
            {

            }
            else
            {
                for (int l = 0; l < tana.size(); l++)
                {
                    if (tana.get(l).ToString().Trim() == "")
                    {
                        tana.set(l, "なし");
                    }
                }
                int page = 1;
                if (tana.size() % 2 == 0)
                {
                    page = tana.size() / 2;
                }
                else
                {
                    page = (tana.size() + 1) / 2;
                }
                PdfReader pdfReader = new PdfReader(pdfStream);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, newpdfStream);
                string time = DateTime.Now.ToString("MMdd");
                for (int i = 1; i < page; i++)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(i);
                    int j = i * 2;

                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();

                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(j - 2).ToString());
                    image1.SetAbsolutePosition(390, 511);
                    pdf.AddImage(image1);
                    iTextSharp.text.Image image2 = AddBarcode(pdf, kmno.get(j - 1).ToString());
                    image2.SetAbsolutePosition(390, 218);
                    pdf.AddImage(image2);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 2).ToString(), 500, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 1).ToString(), 500, 283, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(j - 2).ToString(), 400, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(j - 1).ToString(), 400, 283, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(j - 2).ToString(), 50, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(j - 1).ToString(), 50, 283, 0);

                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(j - 2).ToString(), 300, 575, 0);
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(j - 1).ToString(), 300, 283, 0);
                    }

                    pdf.EndText();

                }
                /// <summary>
                /// check the number of records on the last page
                /// </summary>
                /// <param name="sourceFileName"></param>
                /// <param name="newFileName"></param>

                if (tana.size() % 2 == 0)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();
                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(page * 2 - 2).ToString());
                    image1.SetAbsolutePosition(390, 511);
                    pdf.AddImage(image1);
                    iTextSharp.text.Image image2 = AddBarcode(pdf, kmno.get(page * 2 - 1).ToString());
                    image2.SetAbsolutePosition(390, 218);
                    pdf.AddImage(image2);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 2).ToString(), 500, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 1).ToString(), 500, 283, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 2).ToString(), 400, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 1).ToString(), 400, 283, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 2).ToString(), 50, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 1).ToString(), 50, 283, 0);

                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 2).ToString(), 300, 575, 0);
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 1).ToString(), 300, 283, 0);
                    }
                    pdf.EndText();
                }
                else
                {
                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();
                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(page * 2 - 2).ToString());
                    image1.SetAbsolutePosition(390, 511);
                    pdf.AddImage(image1);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 2).ToString(), 500, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 2).ToString(), 400, 575, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 2).ToString(), 50, 575, 0);

                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 2).ToString(), 300, 575, 0);
                    }
                    pdf.EndText();
                }
                pdfStamper.Close();
            }
        }
    }

    [WebMethod]
    public static void InsertTextToPdf_withCMS(string sourceFileName, string newFileName, ArrayList tana, ArrayList kmno, ArrayList zaikou, ArrayList nouki, ArrayList factory, string repeat, ArrayList recode)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            if (tana.size() == 0)
            {

            }
            else
            {
                for (int l = 0; l < tana.size(); l++)
                {
                    if (tana.get(l).ToString().Trim() == "")
                    {
                        tana.set(l, "なし");
                    }
                }
                int page = 1;
                if (tana.size() % 2 == 0)
                {
                    page = tana.size() / 2;
                }
                else
                {
                    page = (tana.size() + 1) / 2;
                }
                PdfReader pdfReader = new PdfReader(pdfStream);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, newpdfStream);
                string time = DateTime.Now.ToString("MMdd");
                for (int i = 1; i < page; i++)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(i);
                    int j = i * 2;

                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();

                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(j - 2).ToString());
                    image1.SetAbsolutePosition(720, 337);
                    pdf.AddImage(image1);
                    iTextSharp.text.Image image2 = AddBarcode(pdf, kmno.get(j - 1).ToString());
                    image2.SetAbsolutePosition(720, 45);
                    pdf.AddImage(image2);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 2).ToString(), 750, 325, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 1).ToString(), 750, 32, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(j - 2).ToString(), 774, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(j - 1).ToString(), 774, 81, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(j - 2).ToString(), 717, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(j - 1).ToString(), 717, 81, 0);
                    pdf.SetFontAndSize(baseFont, 10);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(j - 2).ToString(), 65, 367, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(j - 1).ToString(), 65, 76, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(j - 2).ToString(), 465, 367, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(j - 1).ToString(), 465, 76, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(j - 2).ToString() == "1" ? "F" : "O", 705, 321, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(j - 1).ToString() == "1" ? "F" : "O", 705, 30, 0);

                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(j - 2).ToString(), 650, 585, 0);
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(j - 1).ToString(), 650, 300, 0);
                    }
                    pdf.EndText();

                }
                /// <summary>
                /// check the number of records on the last page
                /// </summary>
                /// <param name="sourceFileName"></param>
                /// <param name="newFileName"></param>

                if (tana.size() % 2 == 0)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);


                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();

                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(page * 2 - 2).ToString());
                    image1.SetAbsolutePosition(720, 337);
                    pdf.AddImage(image1);
                    iTextSharp.text.Image image2 = AddBarcode(pdf, kmno.get(page * 2 - 1).ToString());
                    image2.SetAbsolutePosition(720, 45);
                    pdf.AddImage(image2);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 2).ToString(), 750, 325, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 1).ToString(), 750, 32, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 2).ToString(), 774, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 1).ToString(), 774, 81, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 2).ToString(), 717, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 1).ToString(), 717, 81, 0);
                    pdf.SetFontAndSize(baseFont, 10);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 2).ToString(), 65, 367, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 1).ToString(), 65, 76, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 2).ToString(), 465, 367, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 1).ToString(), 465, 76, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(page * 2 - 2).ToString() == "1" ? "F" : "O", 705, 321, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(page * 2 - 1).ToString() == "1" ? "F" : "O", 705, 30, 0);
                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 2).ToString(), 650, 585, 0);
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 1).ToString(), 650, 300, 0);
                    }
                    pdf.EndText();
                }
                else
                {
                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 12);
                    pdf.BeginText();

                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(page * 2 - 2).ToString());
                    image1.SetAbsolutePosition(720, 337);
                    pdf.AddImage(image1);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 2 - 2).ToString(), 750, 325, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(page * 2 - 2).ToString(), 774, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, tana.get(page * 2 - 2).ToString(), 717, 372, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(page * 2 - 2).ToString() == "1" ? "F" : "O", 705, 321, 0);
                    pdf.SetFontAndSize(baseFont, 10);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 2).ToString(), 65, 367, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nouki.get(page * 2 - 2).ToString(), 465, 367, 0);
                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, time + "-" + recode.get(page * 2 - 2).ToString(), 650, 585, 0);
                    }
                    pdf.EndText();
                }
                pdfStamper.Close();
            }

        }
    }

    [WebMethod]
    public static void InsertTextToPdf_withCanon(string sourceFileName, string newFileName, ArrayList tana, ArrayList kmno, ArrayList nouki, ArrayList factory, ArrayList sPage, string repeat, ArrayList recode)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            if (tana.size() == 0)
            {

            }
            else
            {
                PdfReader pdfReader = new PdfReader(pdfStream);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, newpdfStream);
                int checkIndex = 0;
                for (int p = 0; p < sPage.size(); p++)
                {
                    if ((int)sPage.get(p) == 0)
                    {
                        continue;
                    }
                    else
                    {
                        int countstart = 1;
                        int start = 1;
                        for (int q = 0; q < p; q++)
                        {
                            if ((int)sPage.get(q) % 2 == 0)
                            {
                                countstart = (int)sPage.get(q) / 2;
                            }
                            else
                            {
                                countstart = ((int)sPage.get(q) + 1) / 2;
                            }
                            start += countstart;
                        }

                        int rankpage = 0;
                        if ((int)sPage.get(p) % 2 == 0)
                        {
                            rankpage = (int)sPage.get(p) / 2;
                        }
                        else
                        {
                            rankpage = ((int)sPage.get(p) + 1) / 2;
                        }

                        int page = start + rankpage - 1;
                        for (int i = start; i <= page; i++)
                        {
                            if (i == page)
                            {
                                if ((int)sPage.get(p) % 2 == 1)
                                {
                                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                    pdf.SetColorFill(BaseColor.BLACK);
                                    pdf.SetFontAndSize(baseFont, 11);
                                    pdf.BeginText();
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(checkIndex).ToString(), 750, 300, 0);
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(checkIndex).ToString() == "1" ? "F" : "O", 810, 300, 0);
                                    if (repeat == "2")
                                    {
                                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(checkIndex).ToString(), 700, 300, 0);
                                    }

                                    checkIndex += 1;
                                    pdf.EndText();
                                }
                                else
                                {
                                    PdfContentByte pdf = pdfStamper.GetOverContent(i);
                                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                    pdf.SetColorFill(BaseColor.BLACK);
                                    pdf.SetFontAndSize(baseFont, 11);
                                    pdf.BeginText();
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(checkIndex).ToString(), 750, 300, 0);
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(checkIndex).ToString() == "1" ? "F" : "O", 810, 300, 0);
                                    if (repeat == "2")
                                    {
                                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(checkIndex).ToString(), 700, 300, 0);
                                    }

                                    checkIndex += 1;
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(checkIndex).ToString(), 750, 5, 0);
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(checkIndex).ToString() == "1" ? "F" : "O", 810, 5, 0);
                                    if (repeat == "2")
                                    {
                                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(checkIndex).ToString(), 700, 5, 0);
                                    }

                                    checkIndex += 1;
                                    pdf.EndText();
                                }
                            }
                            else
                            {
                                PdfContentByte pdf = pdfStamper.GetOverContent(i);

                                BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                pdf.SetColorFill(BaseColor.BLACK);
                                pdf.SetFontAndSize(baseFont, 11);
                                pdf.BeginText();
                                pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(checkIndex).ToString(), 750, 300, 0);
                                pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(checkIndex).ToString() == "1" ? "F" : "O", 810, 300, 0);
                                if (repeat == "2")
                                {
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(checkIndex).ToString(), 700, 300, 0);
                                }

                                checkIndex += 1;
                                pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(checkIndex).ToString(), 750, 5, 0);
                                pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(checkIndex).ToString() == "1" ? "F" : "O", 810, 5, 0);
                                if (repeat == "2")
                                {
                                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(checkIndex).ToString(), 700, 5, 0);
                                }

                                checkIndex += 1;
                                pdf.EndText();
                            }

                        }

                    }
                }
                pdfStamper.Close();
            }
        }
    }

    [WebMethod]
    public static void InsertTextToPdf_withNIKON(string sourceFileName, string newFileName, ArrayList kazu, ArrayList factory, ArrayList kmno, ArrayList nouki, ArrayList zaikou, string repeat, ArrayList recode)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            if (kazu.size() == 0)
            {

            }
            else
            {
                PdfReader pdfReader = new PdfReader(pdfStream);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, newpdfStream);
                for (int i = 1; i <= kazu.size(); i++)
                {
                    string nk = nouki.get(i - 1).ToString();
                    string nam = nk.Substring(0, 4).ToString();
                    string thang = nk.Substring(5, 2).ToString();
                    string ngay = nk.Substring(8, 2).ToString();
                    PdfContentByte pdf = pdfStamper.GetOverContent(i);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 10);
                    pdf.BeginText();

                    iTextSharp.text.Image image1 = AddBarcode(pdf, kmno.get(i - 1).ToString());
                    image1.SetAbsolutePosition(440, 646);
                    pdf.AddImage(image1);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(i - 1).ToString(), 465, 634, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, zaikou.get(i - 1).ToString(), 400, 634, 0);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kazu.get(i - 1).ToString(), 501, 399, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kazu.get(i - 1).ToString(), 255, 138, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kazu.get(i - 1).ToString(), 225, 138, 0);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "1", 255, 174, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "1", 225, 174, 0);


                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, thang, 218, 205, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nam, 180, 205, 0);

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, nam, 420, 497, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, thang, 453, 497, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, factory.get(i - 1).ToString() == "1" ? "F" : "O", 540, 40, 0);
                    if (repeat == "2")
                    {
                        pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, recode.get(i - 1).ToString(), 540, 80, 0);
                    }
                    pdf.EndText();
                }
                pdfStamper.Close();
            }
        }
    }

    [WebMethod]
    public static void InsertTextToPdf_SixApart_withCanon(string sourceFileName, string newFileName, ArrayList tana, ArrayList kmno, ArrayList zaikou, ArrayList nouki, ArrayList factory)
    {
        using (Stream pdfStream = new FileStream(sourceFileName, FileMode.Open))
        using (Stream newpdfStream = new FileStream(newFileName, FileMode.Create, FileAccess.ReadWrite))
        {
            if (tana.size() == 0)
            {

            }
            else
            {
                for (int l = 0; l < tana.size(); l++)
                {
                    if (tana.get(l).ToString().Trim() == "")
                    {
                        tana.set(l, "なし");
                    }
                }
                int page = 1;
                if (tana.size() % 6 == 0)
                {
                    page = tana.size() / 6;
                }
                else
                {
                    page = (int)(tana.size() / 6) + 1;
                }
                PdfReader pdfReader = new PdfReader(pdfStream);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, newpdfStream);

                for (int i = 1; i < page; i++)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(i);
                    int j = i * 6;

                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 5).ToString(), 560, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 4).ToString(), 260, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 3).ToString(), 560, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 2).ToString(), 260, 17, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(j - 1).ToString(), 560, 17, 0);

                    pdf.EndText();

                }
                if (tana.size() % 6 == 1)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.EndText();
                }
                else if (tana.size() % 6 == 2)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();

                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 5).ToString(), 560, 578, 0);
                    pdf.EndText();
                }
                else if (tana.size() % 6 == 3)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 5).ToString(), 560, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 4).ToString(), 260, 297, 0);
                    pdf.EndText();
                }
                else if (tana.size() % 6 == 4)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 5).ToString(), 560, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 4).ToString(), 260, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 3).ToString(), 560, 297, 0);
                    pdf.EndText();
                }
                else if (tana.size() % 6 == 5)
                {

                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 5).ToString(), 560, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 4).ToString(), 260, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 3).ToString(), 560, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 2).ToString(), 260, 17, 0);
                    pdf.EndText();
                }
                else if (tana.size() % 6 == 0)
                {
                    PdfContentByte pdf = pdfStamper.GetOverContent(page);
                    BaseFont baseFont = BaseFont.CreateFont("c:\\windows\\fonts\\msmincho.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    pdf.SetColorFill(BaseColor.BLACK);
                    pdf.SetFontAndSize(baseFont, 8);
                    pdf.BeginText();
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 6).ToString(), 260, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 5).ToString(), 560, 578, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 4).ToString(), 260, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 3).ToString(), 560, 297, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 2).ToString(), 260, 17, 0);
                    pdf.ShowTextAligned(PdfContentByte.ALIGN_CENTER, kmno.get(page * 6 - 1).ToString(), 560, 17, 0);
                    pdf.EndText();
                }
                pdfStamper.Close();
            }

        }
    }
    //chuong trinh rieng biet de them barcode cho file pdf
    [WebMethod]
    public static iTextSharp.text.Image AddBarcode(PdfContentByte pdf, String s)
    {
        Barcode128 c = new Barcode128();
        c.Font = null;
        c.Code = s;
        c.BarHeight = 15;
        iTextSharp.text.Image image = c.CreateImageWithBarcode(pdf, null, BaseColor.WHITE);
        return image;
    }

    /////////////////////////////////////////////////////////
    ///////////////////ket thuc phan ghep thanh file pdf , va them thong tin vao
    ///////////////////////////////////////////////////////////

    //gio la tong hop va in ra cac pdf da in truoc do
    [WebMethod]
    public static List<string> GetRecode_withBC(string empdetails)
    {
        List<string> emp = new List<string>();
        string cyuno = empdetails.ToString().Trim();
        if (File.Exists(@"\\10.121.21.2\data\DeliveryNote\BC\Processed\" + empdetails.ToString().Trim() + @".pdf"))
        {
            emp = getListRecodeApart(cyuno, emp, 1, "11");
        }
        return emp;
    }

    [WebMethod]
    public static List<string> GetRecode_withCMS(string empdetails)
    {
        List<string> emp = new List<string>();
        string cyuno = empdetails.ToString().Trim();
        if (File.Exists(@"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\" + empdetails.ToString().Trim() + @".pdf"))
        {
            emp = getListRecodeBoth(cyuno, emp, "21");
        }
        return emp;
    }

    [WebMethod]
    public static List<string> GetRecode_withCanon(string empdetails)
    {
        List<string> emp = new List<string>();
        string cyuno = empdetails.ToString().Trim();
        if (File.Exists(@"\\10.121.21.2\data\DeliveryNote\Canon\CanonNHS\Processed\" + empdetails.ToString().Trim() + @".pdf"))
        {
            emp = getListRecodeBoth(cyuno, emp, "22");
        }
        return emp;
    }

    [WebMethod]
    public static List<string> GetRecode_withNikon(string empdetails)
    {
        List<string> emp = new List<string>();
        string cyuno = empdetails.ToString().Trim();
        if (File.Exists(@"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\" + empdetails.ToString().Trim() + @".pdf"))
        {
            emp = getListRecodeBoth(cyuno, emp, "23");
        }
        return emp;
    }

    [WebMethod]
    public static List<string> getListRecodeApart(string cyuno, List<string> emp1, int idFactory, string companyTokcd)
    {
        ArrayList lstcyuno = new ArrayList();
        ArrayList lstkz = new ArrayList();
        List<string> emp = new List<string>();
        SqlConnection sqlconn = new SqlConnection();
        if (idFactory == 1)
        {
            sqlconn = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
        }
        else if (idFactory == 2)
        {
            sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
        }
        string sqlquery = "  WITH tb1 AS(SELECT [D1000].CYUNO,[M0120].KABUH,[D1000].JYUNO " +
" ,[D1000].JYUSU,[D1000].ZUBAN,[D1000].YDATE,[D1000].NOUKI,[D1000].TOKCD " +
" FROM [TESC].[dbo].[D1000] " +
" left join  [TESC].[dbo].[M0100] " +
" on  [D1000].SEICD=[M0100].ZAICD " +
"  left join [TESC].[dbo].[M0120] " +
" on  [M0100].ZAICD=[M0120].ZAICD " +
"  where [D1000].TOKCD='00164' and [D1000].[CYUNO] ='" + cyuno + "' and [M0120].JUNJ='001'  and [M0100].ZAIKB='A' ) " +
"  select tb1.CYUNO,[M0100].TANA,[D5000].KMNO " +
" ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD,Max([TESCex].[dbo].[D1000exLog].[PrintStatus]) as PrintStatus  from tb1 " +
" left join [TESC].[dbo].[M0100] " +
" on tb1.KABUH=[M0100].ZAICD " +
" left join [TESC].[dbo].[D5000] " +
" on  [D5000].JYUNO=tb1.JYUNO " +
" inner join [TESCex].[dbo].[D1000exLog] " +
" on [TESCex].[dbo].[D1000exLog].[JYUNO]=tb1.JYUNO " +
" where  [M0100].ZAIKB='B' and [D5000].JSKBN='J' " +
" group by  tb1.CYUNO,[M0100].TANA,[D5000].KMNO ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD " +
" order by tb1.NOUKI,[M0100].TANA ";
        sqlconn.Open();
        SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);
        SqlDataReader sdr = sqlcomn.ExecuteReader();
        while (sdr.Read())
        {
            lstcyuno.add(sdr["CYUNO"].ToString());
            lstkz.add(idFactory.ToString());
            string mt = sdr["CYUNO"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/" + sdr["KMNO"].ToString() + "/*/" + sdr["JYUNO"].ToString() + "/*/"
         + sdr["JYUSU"].ToString() + "/*/" + sdr["ZUBAN"].ToString() + "/*/" + sdr["YDATE"].ToString() + "/*/" + sdr["NOUKI"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/1/*/" + sdr["TOKCD"].ToString() + "/*/" + sdr["PrintStatus"].ToString();
            emp.Add(mt);
        }
        sqlconn.Close();
        ArrayList zaikou = GetDataPO(lstcyuno, idFactory.ToString(), lstkz);
        for (int i = 0; i < zaikou.size(); i++)
        {
            string text_zaikou = zaikou.get(i).ToString().Trim();
            emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);

        }
        return emp1;
    }

    [WebMethod]
    public static List<string> getListRecodeBoth(string cyuno, List<string> emp1, string companyTokcd)
    {
        ArrayList lstcyuno = new ArrayList();
        ArrayList lstkz = new ArrayList();
        List<string> emp = new List<string>();
        SqlConnection sqlconn = new SqlConnection(@"Data Source=10.121.21.12;Initial Catalog=TESC;User ID=tescwin; Password=''");
        string sqlwhere = "";
        if (companyTokcd == "21")
        {
            sqlwhere = "  where [D1000].TOKCD='00002' and [D1000].[CYUNO] ='" + cyuno + "' and [M0120].JUNJ='001'  and [M0100].ZAIKB='A' ) ";
        }
        else if (companyTokcd == "22")
        {
            sqlwhere = "  where ( [D1000].TOKCD='00012' or [D1000].TOKCD='00095' or [D1000].TOKCD='00099' or [D1000].TOKCD='00207' ) and [D1000].[CYUNO] ='" + cyuno + "' and [M0120].JUNJ='001'  and [M0100].ZAIKB='A' ) ";
        }
        else if (companyTokcd == "23")
        {
            sqlwhere = "  where ( [ [D1000].TOKCD='00164'  ) and [D1000].[CYUNO] ='" + cyuno + "' and [M0120].JUNJ='001'  and [M0100].ZAIKB='A' ) ";
        }
        string sqlquery = "  WITH tb1 AS(SELECT [D1000].CYUNO,[M0120].KABUH,[D1000].JYUNO " +
" ,[D1000].JYUSU,[D1000].ZUBAN,[D1000].YDATE,[D1000].NOUKI,[D1000].TOKCD " +
" FROM [TESC].[dbo].[D1000] " +
" left join  [TESC].[dbo].[M0100] " +
" on  [D1000].SEICD=[M0100].ZAICD " +
"  left join [TESC].[dbo].[M0120] " +
" on  [M0100].ZAICD=[M0120].ZAICD " +
sqlwhere +
"  select tb1.CYUNO,[M0100].TANA,[D5000].KMNO " +
" ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD,Max([TESCex].[dbo].[D1000exLog].[PrintStatus]) as PrintStatus  from tb1 " +
" left join [TESC].[dbo].[M0100] " +
" on tb1.KABUH=[M0100].ZAICD " +
" left join [TESC].[dbo].[D5000] " +
" on  [D5000].JYUNO=tb1.JYUNO " +
" inner join [TESCex].[dbo].[D1000exLog] " +
" on [TESCex].[dbo].[D1000exLog].[JYUNO]=tb1.JYUNO " +
" where  [M0100].ZAIKB='B' and [D5000].JSKBN='J' " +
" group by  tb1.CYUNO,[M0100].TANA,[D5000].KMNO ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD " +
" order by tb1.NOUKI,[M0100].TANA ";


        sqlconn.Open();
        SqlCommand sqlcomn = new SqlCommand(sqlquery, sqlconn);
        SqlDataReader sdr = sqlcomn.ExecuteReader();
        while (sdr.Read())
        {
            lstcyuno.add(sdr["CYUNO"].ToString());
            lstkz.add("2");
            string mt = sdr["CYUNO"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/" + sdr["KMNO"].ToString() + "/*/" + sdr["JYUNO"].ToString() + "/*/"
         + sdr["JYUSU"].ToString() + "/*/" + sdr["ZUBAN"].ToString() + "/*/" + sdr["YDATE"].ToString() + "/*/" + sdr["NOUKI"].ToString() + "/*/" + sdr["TANA"].ToString() + "/*/2/*/" + sdr["TOKCD"].ToString() + "/*/" + sdr["PrintStatus"].ToString();
            emp.Add(mt);
        }
        sqlconn.Close();
        ArrayList zaikou = GetDataPO(lstcyuno, "2", lstkz);
        for (int i = 0; i < zaikou.size(); i++)
        {
            string text_zaikou = zaikou.get(i).ToString().Trim();
            emp1.Add(emp[i].ToString() + "/*/" + text_zaikou);

        }

        if (emp.Count < 1)
        {
            ArrayList lstcyuno1 = new ArrayList();
            ArrayList lstkz1 = new ArrayList();
            SqlConnection sqlconn1 = new SqlConnection(@"Data Source=10.121.21.11;Initial Catalog=TESC;User ID=tescwin; Password=''");
            sqlconn1.Open();

            if (companyTokcd == "23")
            {
                sqlwhere = "  where ( [D1000].TOKCD='00203' ) and [D1000].[CYUNO] ='" + cyuno + "' and [M0120].JUNJ='001'  and [M0100].ZAIKB='A' ) ";
                sqlquery = "  WITH tb1 AS(SELECT [D1000].CYUNO,[M0120].KABUH,[D1000].JYUNO " +
" ,[D1000].JYUSU,[D1000].ZUBAN,[D1000].YDATE,[D1000].NOUKI,[D1000].TOKCD " +
" FROM [TESC].[dbo].[D1000] " +
" left join  [TESC].[dbo].[M0100] " +
" on  [D1000].SEICD=[M0100].ZAICD " +
"  left join [TESC].[dbo].[M0120] " +
" on  [M0100].ZAICD=[M0120].ZAICD " +
sqlwhere +
"  select tb1.CYUNO,[M0100].TANA,[D5000].KMNO " +
" ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD,Max([TESCex].[dbo].[D1000exLog].[PrintStatus]) as PrintStatus  from tb1 " +
" left join [TESC].[dbo].[M0100] " +
" on tb1.KABUH=[M0100].ZAICD " +
" left join [TESC].[dbo].[D5000] " +
" on  [D5000].JYUNO=tb1.JYUNO " +
" inner join [TESCex].[dbo].[D1000exLog] " +
" on [TESCex].[dbo].[D1000exLog].[JYUNO]=tb1.JYUNO " +
" where  [M0100].ZAIKB='B' and [D5000].JSKBN='J' " +
" group by  tb1.CYUNO,[M0100].TANA,[D5000].KMNO ,tb1.JYUNO,tb1.JYUSU,tb1.ZUBAN,tb1.YDATE,tb1.NOUKI,tb1.TOKCD " +
" order by tb1.NOUKI,[M0100].TANA ";
            }

            SqlCommand sqlcomn1 = new SqlCommand(sqlquery, sqlconn1);
            SqlDataReader sdr1 = sqlcomn1.ExecuteReader();
            while (sdr1.Read())
            {
                lstcyuno1.add(sdr1["CYUNO"].ToString());
                lstkz1.add("1");
                string mt1 = sdr1["CYUNO"].ToString() + "/*/" + sdr1["TANA"].ToString() + "/*/" + sdr1["KMNO"].ToString() + "/*/" + sdr1["JYUNO"].ToString() + "/*/"
              + sdr1["JYUSU"].ToString() + "/*/" + sdr1["ZUBAN"].ToString() + "/*/" + sdr1["YDATE"].ToString() + "/*/" + sdr1["NOUKI"].ToString() + "/*/" + sdr1["TANA"].ToString() + "/*/1/*/" + sdr1["TOKCD"].ToString() + "/*/" + sdr1["PrintStatus"].ToString();
                emp.Add(mt1);
            }
            sqlconn1.Close();
            ArrayList zaikou1 = GetDataPO(lstcyuno1, "1", lstkz1);
            for (int i = 0; i < zaikou1.size(); i++)
            {
                string text_zaikou1 = zaikou1.get(i).ToString().Trim();
                emp1.Add(emp[i].ToString() + "/*/" + text_zaikou1);
            }
        }
        return emp1;
    }

    //////////////////////////////////////////////////////
    /////////////ket thuc phan in lai pd/////////////////
    ////////////////////////////////////////////////////////

    //hien thi pdf sau khi tong hop
    protected void showBC(object sender, EventArgs e)
    {
        //backup file
        string DatePrint = DateTime.Now.ToString("yyyyMMdd");
        string HourPrint = DateTime.Now.ToString("HHmmss");
        string source = @"\\10.121.21.2\data\DeliveryNote\BC\Processed\endpdf.pdf";

        string destine = @"\\10.121.21.2\data\DeliveryNote\Backup\BC_" + DatePrint + "_" + HourPrint + ".pdf";
        File.Copy(source, destine);
        //hien thi ra ngoai man hinh
        Response.ContentType = "application/octet-stream";
        string st = DateTime.Now.ToString("yyyyMMdd");
        string st1 = "BC" + st + ".pdf";
        Response.AppendHeader("Content-Disposition", "attachment;filename=" + st1);
        Response.TransmitFile(@"\\10.121.21.2\data\DeliveryNote\BC\Processed\endpdf.pdf");
        Response.End();
    }

    protected void showCMS(object sender, EventArgs e)
    {
        //backup file
        string DatePrint = DateTime.Now.ToString("yyyyMMdd");
        string HourPrint = DateTime.Now.ToString("HHmmss");
        string source = @"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\endpdf.pdf";
        string destine = @"\\10.121.21.2\data\DeliveryNote\Backup\CMSC_" + DatePrint + "_" + HourPrint + ".pdf";
        File.Copy(source, destine);
        //hien thi ra ngoai man hinh
        Response.ContentType = "application/octet-stream";
        string st = DateTime.Now.ToString("yyyyMMdd");
        string st1 = "CMSC" + st + ".pdf";
        Response.AppendHeader("Content-Disposition", "attachment;filename=" + st1);
        Response.TransmitFile(@"\\10.121.21.2\data\DeliveryNote\CMSC\GPH\Processed\endpdf.pdf");
        Response.End();
    }

    protected void showCanon(object sender, EventArgs e)
    {
        //backup file
        string DatePrint = DateTime.Now.ToString("yyyyMMdd");
        string HourPrint = DateTime.Now.ToString("HHmmss");
        string source = @"\\10.121.21.2\data\DeliveryNote\Canon\new.pdf";
        string destine = @"\\10.121.21.2\data\DeliveryNote\Backup\Canon_" + DatePrint + "_" + HourPrint + ".pdf";
        File.Copy(source, destine);
        //hien thi ra ngoai man hinh
        Response.ContentType = "application/octet-stream";
        string st = DateTime.Now.ToString("yyyyMMdd");
        string st1 = "Canon" + st + ".pdf";
        Response.AppendHeader("Content-Disposition", "attachment;filename=" + st1);
        Response.TransmitFile(@"\\10.121.21.2\data\DeliveryNote\Canon\new.pdf");
        Response.End();
    }

    protected void showNikon(object sender, EventArgs e)
    {
        //backup file
        string DatePrint = DateTime.Now.ToString("yyyyMMdd");
        string HourPrint = DateTime.Now.ToString("HHmmss");
        string source = @"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\endpdf.pdf";
        string destine = @"\\10.121.21.2\data\DeliveryNote\Backup\Nikon_" + DatePrint + "_" + HourPrint + ".pdf";
        File.Copy(source, destine);
        //hien thi ra ngoai man hinh
        Response.ContentType = "application/octet-stream";
        string st = DateTime.Now.ToString("yyyyMMdd");
        string st1 = "Nikon" + st + ".pdf";
        Response.AppendHeader("Content-Disposition", "attachment;filename=" + st1);
        Response.TransmitFile(@"\\10.121.21.2\data\DeliveryNote\NIKON\Processed\endpdf.pdf");
        Response.End();
    }


    ///////////////////Binhさんのソースコード
    public void copyPerPage(string sourceFile, string destineFile, int numberPage)
    {
        using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
        using (PdfReader reader = new PdfReader(sourceFile))
        {
            iTextSharp.text.Rectangle pageSize = reader.GetPageSizeWithRotation(numberPage);
            Document document = new Document(pageSize);
            PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            PdfImportedPage page = writer.GetImportedPage(reader, numberPage);

            content.AddTemplate(page, 0, 0);
            content.Fill();

            document.SetPageSize(pageSize);
            document.NewPage();
            document.Close();
            reader.Close();
            newPdfStream.Close();
        }
    }
    //Cut Canon File
    //public void split_pdf_canon(string sourcePdfFile)
    //{
    //    List<string> list = getDataCyuno_Canon(sourcePdfFile);
    //    if (list.Count > 0)
    //    {
    //        int splitPageTo = 3;
    //        for (int i = 1; i <= splitPageTo; i++)
    //        {
    //            var namepdf = list[0].Trim().ToString();
    //            string outputPdfPath = "";
    //            if (i == 2)
    //            {
    //                outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CNNM\GPH\" + namepdf + ".pdf";
    //                Cut_3page_Canon(sourcePdfFile, outputPdfPath, 1, i, splitPageTo);
    //            }
    //            else if (i == 3)
    //            {
    //                outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CNNM\NHS\" + namepdf + ".pdf";
    //                Cut_3page_Canon(sourcePdfFile, outputPdfPath, 1, i, splitPageTo);
    //            }
    //        }
    //    }
    //    else
    //    {
    //        return;
    //    }
    //}
    public void split_pdf_canon(string sourcePdfFile)
    {
        List<string> list = getDataCyuno_Canon(sourcePdfFile);
        int split_to_page = 2;
        for (int i = 1; i <= split_to_page; i++)
        {
            string namePdf = list[0].ToString().Trim();
            if (i == 2)
            {
                string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CNNM_2\GPH\" + namePdf + ".pdf";
                Cut_2page_CNNM(sourcePdfFile, outputPdfPath, 1, i, split_to_page);
            }
        }
    }

    protected List<string> getDataCyuno_Canon(string sourcePdfFile)
    {
        PDDocument doc = PDDocument.load(sourcePdfFile);
        PDFTextStripperByArea stripper = new PDFTextStripperByArea();
        List<string> list = new List<string>();
        int x = 490;
        int y = 340;
        int w = 70;
        int h = 20;
        stripper.addRegion("testRegion", new java.awt.Rectangle(x, y, w, h));
        stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
        string cyuno = stripper.getTextForRegion("testRegion").Trim();
        if (cyuno.Trim() != "")
        {
            list.Add(cyuno.ToString());
        }
        doc.close();
        return list;
    }

    public void Cut_3page_Canon(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
    {
        using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
        using (PdfReader reader = new PdfReader(sourceFile))
        {
            iTextSharp.text.Rectangle pageSize = reader.GetPageSizeWithRotation(numberPage);
            Document document = new Document(pageSize);
            PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            PdfImportedPage page = writer.GetImportedPage(reader, numberPage);

            content.AddTemplate(page, 0, reader.GetPageSizeWithRotation(numberPage).Height * (position - 1) / cut_to);
            content.SetColorFill(BaseColor.WHITE);
            content.RoundRectangle(0, 0, reader.GetPageSizeWithRotation(numberPage).Width, reader.GetPageSizeWithRotation(numberPage).Height * (cut_to - 1) / cut_to, 0);
            content.Fill();

            document.SetPageSize(pageSize);
            document.NewPage();
            document.Close();
            reader.Close();
            newPdfStream.Close();
        }
    }

    public void Cut_2page_CNNM(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
    {
        using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
        {
            PdfReader reader = new PdfReader(sourceFile);
            iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
            Document document = new Document(pageSize);
            PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
            document.Open();
            PdfContentByte content = writer.DirectContent;
            PdfImportedPage page = writer.GetImportedPage(reader, 1);
            if (position == 2)
            {
                content.AddTemplate(page, 0, reader.GetPageSize(1).Height * 1 / 3);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 3, 0);
                content.Fill();
            }
            document.SetPageSize(pageSize);
            document.NewPage();
            document.Close();
            reader.Close();
            newPdfStream.Close();
        }
    }

    //Cut Dainikkou File
        public void split_pdf_dainikkou(string sourcePdfFile)
        {
            List<string> list = getDataCyuno_Dainikkou(sourcePdfFile);
            if (list.Count > 0)
            {
                int split_to_page = list.Count;
                for (int i = 1; i <= split_to_page; i++)
                {
                    string namePdf = list[i - 1].ToString().Trim();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\DAINIKKOU\NHS\" + namePdf + ".pdf";
                    Cut_2page_Dainikkou(sourcePdfFile, outputPdfPath, 1, i, 2);
                }
            }
            else
            {
                return;
            }
        }

        protected List<string> getDataCyuno_Dainikkou(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1
            int x1 = 300;
            int y1 = 60;
            int w1 = 100;
            int h1 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x1, y1, w1, h1));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text1 = stripper.getTextForRegion("testRegion");
            if (text1.Trim() != "")
            {
                list.Add(text1);
            }
            //position 2
            int x2 = 300;
            int y2 = 350;
            int w2 = 100;
            int h2 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x2, y2, w2, h2));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text2 = stripper.getTextForRegion("testRegion");
            if (text2.Trim() != "")
            {
                list.Add(text2);
            }
            doc.close();
            return list;
        }

        public void Cut_2page_Dainikkou(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                content.AddTemplate(page, 0, reader.GetPageSize(1).Height * (position - 1) / cut_to);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * (cut_to - 1) / cut_to, 0);
                content.Fill();

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        //Cut Reon File
        public void split_pdf_reon(string sourcePdfFile)
        {
            List<string> list = getDataCyuno_Reon(sourcePdfFile);
            if (list.Count > 0)
            {
                int split_to_page = list.Count * 2;
                for (int i = 1; i <= split_to_page; i++)
                {
                    if (i == 1)
                    {
                        string namePdf = list[0].ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\REON2\NHS\\" + namePdf + ".pdf";
                        Cut_4page_Reon(sourcePdfFile, outputPdfPath, 1, i, 2);
                    }
                    if (i == 2)
                    {
                        string namePdf = list[0].ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\REON2\GPH\\" + namePdf + ".pdf";
                        Cut_4page_Reon(sourcePdfFile, outputPdfPath, 1, i, 2);
                    }
                    if (i == 3)
                    {
                        string namePdf = list[1].ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\REON2\NHS\\" + namePdf + ".pdf";
                        Cut_4page_Reon(sourcePdfFile, outputPdfPath, 1, i, 2);
                    }
                    if (i == 4)
                    {
                        string namePdf = list[1].ToString().Trim();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\REON2\GPH\\" + namePdf + ".pdf";
                        Cut_4page_Reon(sourcePdfFile, outputPdfPath, 1, i, 2);
                    }
                }
            }
        }

        protected List<string> getDataCyuno_Reon(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1
            int x1 = 405;
            int y1 = 70;
            int w1 = 100;
            int h1 = 20;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x1, y1, w1, h1));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text1 = stripper.getTextForRegion("testRegion");
            if (text1.Trim() != "")
            {
                list.Add(text1);
            }
            //position 2
            int x2 = 405;
            int y2 = 370;
            int w2 = 100;
            int h2 = 20;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x2, y2, w2, h2));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text2 = stripper.getTextForRegion("testRegion");
            if (text2.Trim() != "")
            {
                list.Add(text2);
            }
            doc.close();
            return list;
        }

        public void Cut_2page_Reon(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                content.AddTemplate(page, 0, (reader.GetPageSize(1).Height * (position - 1) / cut_to) + 4);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, (reader.GetPageSize(1).Height * (cut_to - 1) / cut_to) + 4, 0);
                content.Fill();

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        public void Cut_4page_Reon(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                if (position == 1)
                {
                    content.AddTemplate(page, 0, 0);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 2, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 2 / 3, 0, reader.GetPageSize(1).Width * 1 / 3, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                if (position == 2)
                {
                    content.AddTemplate(page, -reader.GetPageSize(1).Width * 2 / 3, 0);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 2, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 3, 0, reader.GetPageSize(1).Width * 2 / 3, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                if (position == 3)
                {
                    content.AddTemplate(page, 0, reader.GetPageSize(1).Height * 1 / 2);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 2, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 2 / 3, 0, reader.GetPageSize(1).Width * 1 / 3, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                if (position == 4)
                {
                    content.AddTemplate(page, -reader.GetPageSize(1).Width * 2 / 3, reader.GetPageSize(1).Height * 1 / 2);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 2, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 3, 0, reader.GetPageSize(1).Width * 2 / 3, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        //Cut Nikon 2/3 File
        public void split_pdf_nikon2(string sourcePdfFile)
        {
            List<string> list = getDataCyuno_Nikon2(sourcePdfFile);
            if (list.Count > 0)
            {
                int split_to_page = 2;
                for (int i = 1; i <= split_to_page; i++)
                {
                    string namePdf = list[0].ToString().Trim();
                    if (i == 2)
                    {
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\NIKON_MIYAGI_2\GPH\" + namePdf + ".pdf";
                        Cut_2page_Nikon(sourcePdfFile, outputPdfPath, 1, i, split_to_page);
                    }
                }
            }
            else
            {
                return;
            }
        }

        protected List<string> getDataCyuno_Nikon2(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            int x = 110;
            int y = 90;
            int w = 100;
            int h = 20;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x, y, w, h));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text = stripper.getTextForRegion("testRegion");
            if (text.Trim() != "")
            {
                list.Add(text);
            }
            doc.close();
            return list;
        }

        public void Cut_2page_Nikon(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                if (position == 2)
                {
                    content.AddTemplate(page, 0, reader.GetPageSize(1).Height * 1 / 3);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 3, 0);
                    content.Fill();
                }
                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        public List<string> get_kazu_nikon(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //datetime
            int xdate_time = 350;
            int ydate_time = 430;
            int wdate_time = 70;
            int hdate_time = 20;
            stripper.addRegion("testRegion", new java.awt.Rectangle(xdate_time, ydate_time, wdate_time, hdate_time));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string dateTime = stripper.getTextForRegion("testRegion");
            if (dateTime.Trim() != "")
            {
                list.Add(dateTime);
            }
            //kazu
            int xkazu = 440;
            int ykazu = 430;
            int wkazu = 30;
            int hkazu = 20;
            stripper.addRegion("testRegion", new java.awt.Rectangle(xkazu, ykazu, wkazu, hkazu));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string kazu = stripper.getTextForRegion("testRegion");
            if (kazu.Trim() != "")
            {
                list.Add(kazu);
            }
            doc.close();
            return list;
        }

        public void insert_kazu_nikon(string sourceFile, string destineFile, List<string> data)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            using (PdfStamper stamper = new PdfStamper(reader, newPdfStream))
            {
                PdfContentByte content = stamper.GetOverContent(1);
                content.SetColorFill(BaseColor.BLACK);
                content.SetFontAndSize(BaseFont.CreateFont("c:\\windows\\fonts\\msgothic.ttc,0", BaseFont.IDENTITY_H, BaseFont.EMBEDDED), 9);

                content.BeginText();

                List<string> list = data;
                if (list.Count == 2)
                {
                    string dateTime = list[0];
                    string kazu = list[1];
                    string year = dateTime.Substring(0, 4);
                    string month = dateTime.Substring(5, 2);

                    content.ShowTextAligned(1, "1", 226, 175, 0);
                    content.ShowTextAligned(1, "1", 256, 175, 0);

                    content.ShowTextAligned(1, kazu, 226, 139, 0);
                    content.ShowTextAligned(1, kazu, 256, 139, 0);
                    content.ShowTextAligned(1, kazu, 480, 398, 0);

                    content.ShowTextAligned(1, year, 180, 206, 0);
                    content.ShowTextAligned(1, month, 213, 206, 0);

                    content.ShowTextAligned(1, year, 425, 498, 0);
                    content.ShowTextAligned(1, month, 453, 498, 0);
                }

                content.EndText();
                stamper.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        //Cut CMSC File
        public void split_pdf_cmsc(string sourcePdfFile)
        {
            List<string> list = getDataCyuno_CMSC(sourcePdfFile);
            if (list.Count > 0)
            {
                int splitPageTo = list.Count;
                for (int i = 1; i <= splitPageTo; i++)
                {
                    string namePdf = list[i - 1].ToString().Trim();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CMSC_N\GPH\" + namePdf + ".pdf";
                    Cut_2page_CMSC(sourcePdfFile, outputPdfPath, 1, i, 2);
                }
            }
            else
            {
                return;
            }
        }

        protected List<string> getDataCyuno_CMSC(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1
            int x1 = 310;
            int y1 = 110;
            int w1 = 100;
            int h1 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x1, y1, w1, h1));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text1 = stripper.getTextForRegion("testRegion");
            if (text1.Trim() != "")
            {
                text1 = text1.Substring(0, 8) + "-" + text1.Substring(8, 2);
                list.Add(text1);
            }
            //position 2
            int x2 = 310;
            int y2 = 390;
            int w2 = 100;
            int h2 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x2, y2, w2, h2));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text2 = stripper.getTextForRegion("testRegion");
            if (text2.Trim() != "")
            {
                text2 = text2.Substring(0, 8) + "-" + text2.Substring(8, 2);
                list.Add(text2);
            }
            doc.close();
            return list;
        }

        public void Cut_2page_CMSC(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, numberPage);

                content.AddTemplate(page, 0, reader.GetPageSizeWithRotation(numberPage).Height * (position - 1) / cut_to);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * (cut_to - 1) / cut_to, 0);
                content.Fill();

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        //Cut CANON from 2 format to 2 part (2 part and 6 part)
        public void split_pdf_canon_2format(string sourcePdfFile)
        {
            Boolean checkNHSCanon = checkNHS_Canon(sourcePdfFile);
            if (checkNHSCanon)
            {
                List<string> data_nhs = getDataCyuno_Canon_2part(sourcePdfFile);
                if (data_nhs.Count > 0)
                {
                    int split_to_page = data_nhs.Count;
                    for (int i = 1; i <= split_to_page; i++)
                    {
                        string namePdf = data_nhs[i - 1].Trim().ToString();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CANON_N\NHS\" + namePdf + ".pdf";
                        Cut_Canon_2Part(sourcePdfFile, outputPdfPath, 1, i, split_to_page);
                    }
                }
                else
                {
                    return;
                }
            }
            else
            {
                List<string> data_gph = getDataCyuno_Canon_6part(sourcePdfFile);
                if (data_gph.Count > 0)
                {
                    int split_to_page = data_gph.Count;
                    for (int i = 1; i <= split_to_page; i++)
                    {
                        string namePdf = data_gph[i - 1].Trim().ToString();
                        string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\CANON_N\GPH\" + namePdf + ".pdf";
                        Cut_Canon_6Part(sourcePdfFile, outputPdfPath, 1, i, split_to_page);
                    }
                }
                else
                {
                    return;
                }
            }
        }

        protected List<string> getDataCyuno_Canon_2part(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1
            int x1 = 630;
            int y1 = 50;
            int w1 = 140;
            int h1 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x1, y1, w1, h1));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text1 = stripper.getTextForRegion("testRegion");
            if (text1.Trim() != "")
            {
                text1 = text1.Substring(1, 11) + text1.Substring(14, 1);
                list.Add(text1);
            }
            //position 2
            int x2 = 630;
            int y2 = 350;
            int w2 = 140;
            int h2 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x2, y2, w2, h2));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text2 = stripper.getTextForRegion("testRegion");
            if (text2.Trim() != "")
            {
                text2 = text2.Substring(1, 11) + text2.Substring(14, 1);
                list.Add(text2);
            }
            doc.close();
            return list;
        }

        protected List<string> getDataCyuno_Canon_6part(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1
            int x1 = 40;
            int y1 = 140;
            int w1 = 120;
            int h1 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x1, y1, w1, h1));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text1 = stripper.getTextForRegion("testRegion");
            if (text1.Trim() != "")
            {
                text1 = text1.Substring(1, 11) + text1.Substring(14, 1);
                list.Add(text1);
            }
            //position 2
            int x2 = 340;
            int y2 = 140;
            int w2 = 120;
            int h2 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x2, y2, w2, h2));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text2 = stripper.getTextForRegion("testRegion");
            if (text2.Trim() != "")
            {
                text2 = text2.Substring(1, 11) + text2.Substring(14, 1);
                list.Add(text2);
            }
            //position 3
            int x3 = 40;
            int y3 = 430;
            int w3 = 120;
            int h3 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x3, y3, w3, h3));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text3 = stripper.getTextForRegion("testRegion");
            if (text3.Trim() != "")
            {
                text3 = text3.Substring(1, 11) + text3.Substring(14, 1);
                list.Add(text3);
            }
            //position 4
            int x4 = 340;
            int y4 = 430;
            int w4 = 120;
            int h4 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x4, y4, w4, h4));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text4 = stripper.getTextForRegion("testRegion");
            if (text4.Trim() != "")
            {
                text4 = text4.Substring(1, 11) + text4.Substring(14, 1);
                list.Add(text4);
            }
            //position 5
            int x5 = 40;
            int y5 = 720;
            int w5 = 120;
            int h5 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x5, y5, w5, h5));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text5 = stripper.getTextForRegion("testRegion");
            if (text5.Trim() != "")
            {
                text5 = text5.Substring(1, 11) + text5.Substring(14, 1);
                list.Add(text5);
            }
            //position 6
            int x6 = 340;
            int y6 = 720;
            int w6 = 120;
            int h6 = 15;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x6, y6, w6, h6));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text6 = stripper.getTextForRegion("testRegion");
            if (text6.Trim() != "")
            {
                text6 = text6.Substring(1, 11) + text6.Substring(14, 1);
                list.Add(text6);
            }
            doc.close();
            return list;
        }

        public Boolean checkNHS_Canon(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            int x = 270;
            int y = 40;
            int w = 50;
            int h = 10;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x, y, w, h));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text = stripper.getTextForRegion("testRegion");
            doc.close();
            if (text.Contains("注文番号") == true)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void Cut_Canon_2Part(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                content.AddTemplate(page, 0, reader.GetPageSize(1).Height * (position - 1) / cut_to);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * (cut_to - 1) / cut_to, 0);
                content.Fill();

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        public void Cut_Canon_6Part(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                if (position == 1)
                {
                    content.AddTemplate(page, 0, 0);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                else if (position == 2)
                {
                    content.AddTemplate(page, -reader.GetPageSize(1).Width * 1 / 2, 0);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                else if (position == 3)
                {
                    content.AddTemplate(page, 0, reader.GetPageSize(1).Height * 1 / 3);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                else if (position == 4)
                {
                    content.AddTemplate(page, -reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height * 1 / 3);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                else if (position == 5)
                {
                    content.AddTemplate(page, 0, reader.GetPageSize(1).Height * 2 / 3);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }
                else if (position == 6)
                {
                    content.AddTemplate(page, -reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height * 2 / 3);
                    content.SetColorFill(BaseColor.WHITE);
                    content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 2 / 3, 0);
                    content.RoundRectangle(reader.GetPageSize(1).Width * 1 / 2, 0, reader.GetPageSize(1).Width * 1 / 2, reader.GetPageSize(1).Height, 0);
                    content.Fill();
                }

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

        //Cut BC 2 part
        public void split_pdf_bc(string sourcePdfFile)
        {
            List<string> list = getDataCyuno_BC(sourcePdfFile);
            if (list.Count > 0)
            {
                int split_to_page = list.Count;
                for (int i = 1; i <= split_to_page; i++)
                {
                    string namePdf = list[i - 1].Trim().ToString();
                    string outputPdfPath = @"\\10.121.21.2\data\DeliveryNote\BC_N\GPH\" + namePdf + ".pdf";
                    Cut_2page_BC(sourcePdfFile, outputPdfPath, 1, i, split_to_page);
                }
            }
            else
            {
                return;
            }
        }

        protected List<string> getDataCyuno_BC(string sourcePdfFile)
        {
            PDDocument doc = PDDocument.load(sourcePdfFile);
            PDFTextStripperByArea stripper = new PDFTextStripperByArea();
            List<string> list = new List<string>();
            //position 1-1
            int x11 = 335;
            int y11 = 35;
            int w11 = 60;
            int h11 = 30;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x11, y11, w11, h11));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text11 = stripper.getTextForRegion("testRegion").Trim();
            //position 1-2
            int x12 = 440;
            int y12 = 35;
            int w12 = 30;
            int h12 = 30;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x12, y12, w12, h12));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text12 = stripper.getTextForRegion("testRegion").Trim();
            //merge text11 & text12
            if (text11.Trim() != "" && text12.Trim() != "")
            {
                string text1 = text11.Trim() + "$I" + text12.Trim();
                list.Add(text1.Trim());
            }
            //position 2-1
            int x21 = 335;
            int y21 = 325;
            int w21 = 60;
            int h21 = 30;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x21, y21, w21, h21));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text21 = stripper.getTextForRegion("testRegion").Trim();
            //position 2-2
            int x22 = 440;
            int y22 = 325;
            int w22 = 30;
            int h22 = 30;
            stripper.addRegion("testRegion", new java.awt.Rectangle(x22, y22, w22, h22));
            stripper.extractRegions((PDPage)doc.getDocumentCatalog().getAllPages().get(0));
            string text22 = stripper.getTextForRegion("testRegion").Trim();
            //merge text21 & text22
            if (text21.Trim() != "" && text22.Trim() != "")
            {
                string text2 = text21.Trim() + "$I" + text22.Trim();
                list.Add(text2.Trim());
            }
            doc.close();
            return list;
        }

        public void Cut_2page_BC(string sourceFile, string destineFile, int numberPage, int position, int cut_to)
        {
            using (FileStream newPdfStream = new FileStream(destineFile, FileMode.Create, FileAccess.ReadWrite))
            using (PdfReader reader = new PdfReader(sourceFile))
            {
                iTextSharp.text.Rectangle pageSize = reader.GetPageSize(1);
                Document document = new Document(pageSize);
                PdfWriter writer = PdfWriter.GetInstance(document, newPdfStream);
                document.Open();
                PdfContentByte content = writer.DirectContent;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                content.AddTemplate(page, 0, reader.GetPageSize(1).Height * (position - 1) / cut_to);
                content.SetColorFill(BaseColor.WHITE);
                content.RoundRectangle(0, 0, reader.GetPageSize(1).Width, reader.GetPageSize(1).Height * 1 / 2, 0);
                content.Fill();

                document.SetPageSize(pageSize);
                document.NewPage();
                document.Close();
                reader.Close();
                newPdfStream.Close();
            }
        }

    //function extract text from pdf to txt file and read txt file
    protected string[] getTextFromTxt(string pdfFile, string txtFile)
    {
        //read pdf file and get text
        PDDocument document = PDDocument.load(pdfFile);
        PDFTextStripper stripper = new PDFTextStripper();
        string txt_pdf = stripper.getText(document);
        //write data into txt file
        File.WriteAllText(txtFile, txt_pdf);
        //get text from txt file
        string[] lines = File.ReadAllLines(txtFile);
        document.close();
        return lines;
    }

    //function extract text from pdf
    private static string getTextFromPdf(string PdfFile)
    {
        PDDocument doc = null;
        try
        {
            doc = PDDocument.load(PdfFile);
            PDFTextStripper stripper = new PDFTextStripper();
            return stripper.getText(doc);
        }
        finally
        {
            if (doc != null)
            {
                doc.close();
            }
        }
    }

    /////Binhさんのエンド


}

