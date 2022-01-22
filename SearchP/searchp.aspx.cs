using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
namespace SearchP
{
    public partial class searchp : System.Web.UI.Page
    {
      
        protected void Page_Load(object sender, EventArgs e)
        {
         
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
  

            string mainconn = ConfigurationManager.ConnectionStrings["Myconnection"].ConnectionString;
            SqlConnection sqlconn = new SqlConnection(mainconn);
            SqlCommand sqlcomm = new SqlCommand("testsearchdate002", sqlconn);
            sqlcomm.CommandType = CommandType.StoredProcedure;
            sqlcomm.Parameters.AddWithValue("@YYMM", TextBox1.Text);
            sqlconn.Open();
            GridView1.DataSource = sqlcomm.ExecuteReader();
     
            GridView1.DataBind();
         

                sqlconn.Close();


        }

        //protected void Button2_Click(object sender, EventArgs e)
        //{
        //    Response.ClearContent();
        //    Response.AppendHeader("content-disposition", "attachment;filename=excel1.xls");
        //    Response.Charset = "";
        //    Response.ContentType = "application/excel";


        //    StringWriter stringWriter = new StringWriter();
        //    HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);
        //    GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

        //    // Set background color of each cell of GridView1 header row
        //    foreach (TableCell tableCell in GridView1.HeaderRow.Cells)
        //    {
        //        tableCell.Style["background-color"] = "#A55129";
        //    }

        //    // Set background color of each cell of each data row of GridView1
        //    foreach (GridViewRow gridViewRow in GridView1.Rows)
        //    {
        //        gridViewRow.BackColor = System.Drawing.Color.White;
        //        foreach (TableCell gridViewRowTableCell in gridViewRow.Cells)
        //        {
        //            gridViewRowTableCell.Style["background-color"] = "#FFF7E7";
        //        }
        //    }
        //    GridView1.RenderControl(htmlTextWriter);
        //    string style = @"<style> td { mso-number-format:\@;} </style>";
        //    Response.Write(style);
        //    Response.Write(stringWriter.ToString());
        //    Response.End();
        //}
        //public override void VerifyRenderingInServerForm(Control control)
        //{

           
        //}

  

        protected void Button3_Click1(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
           
            dt.Columns.Add("regno");
            dt.Columns.Add("COMPANY_CODE");
            dt.Columns.Add("mg_code");
            dt.Columns.Add("NATIONAL_NO");
            dt.Columns.Add("EDUCAT_LEVEL");
            dt.Columns.Add("EMPLOYM_TYPE");
            dt.Columns.Add("sabeghe");
            dt.Columns.Add("Date");
            dt.Columns.Add("SystemName");
            dt.Columns.Add("5185");
            dt.Columns.Add("5192");
            dt.Columns.Add("5207");
            dt.Columns.Add("6177");
            dt.Columns.Add("6178");
            dt.Columns.Add("6206");
            dt.Columns.Add("6207");
            dt.Columns.Add("6299");
            dt.Columns.Add("6356");
            //dt.Columns.Add("6358");
            //dt.Columns.Add("6362");
            //dt.Columns.Add("6363");
            //dt.Columns.Add("6364");
            //dt.Columns.Add("6365");
            //dt.Columns.Add("6367");
            //dt.Columns.Add("6368");
            //dt.Columns.Add("6369");
            //dt.Columns.Add("6372");
            //dt.Columns.Add("6626");
            foreach (GridViewRow row in GridView1.Rows)
            {
                int i = row.RowIndex;
                for (int t = 11; t <= 20; t++)
                {
                    if (GridView1.Rows[i].Cells[t].Text == "&nbsp;")
                        GridView1.Rows[i].Cells[t].Text = " ";

                }

                DataRow dr = dt.NewRow();
   
                dr["regno"] = GridView1.Rows[i].Cells[0].Text;
                dr["COMPANY_CODE"] = GridView1.Rows[i].Cells[1].Text;
                dr["mg_code"] = GridView1.Rows[i].Cells[2].Text;
                dr["NATIONAL_NO"] = GridView1.Rows[i].Cells[3].Text;
                dr["EDUCAT_LEVEL"] = GridView1.Rows[i].Cells[4].Text;
                dr["EMPLOYM_TYPE"] = GridView1.Rows[i].Cells[5].Text;
                dr["sabeghe"] = GridView1.Rows[i].Cells[6].Text;
                dr["Date"] = GridView1.Rows[i].Cells[7].Text;
                dr["SystemName"] = GridView1.Rows[i].Cells[8].Text;
                dr["5185"] = GridView1.Rows[i].Cells[9].Text;
                dr["5192"] = GridView1.Rows[i].Cells[10].Text;
                dr["5207"] = GridView1.Rows[i].Cells[11].Text;
                dr["6177"] = GridView1.Rows[i].Cells[12].Text;
                dr["6178"] = GridView1.Rows[i].Cells[13].Text;
                dr["6206"] = GridView1.Rows[i].Cells[14].Text;
                dr["6207"] = GridView1.Rows[i].Cells[15].Text;
                dr["6299"] = GridView1.Rows[i].Cells[16].Text;
                dr["6356"] = GridView1.Rows[i].Cells[17].Text;
                //dr["6358"] = GridView1.Rows[i].Cells[18].Text;
                //dr["6362"] = GridView1.Rows[i].Cells[19].Text;
                //dr["6363"] = GridView1.Rows[i].Cells[20].Text;
                //dr["6364"] = GridView1.Rows[i].Cells[21].Text;
                //dr["6365"] = GridView1.Rows[i].Cells[22].Text;
                //dr["6367"] = GridView1.Rows[i].Cells[23].Text;
                //dr["6368"] = GridView1.Rows[i].Cells[24].Text;
                //dr["6369"] = GridView1.Rows[i].Cells[25].Text;
                //dr["6372"] = GridView1.Rows[i].Cells[26].Text;
                //dr["6626"] = GridView1.Rows[i].Cells[27].Text;
                               dt.Rows.Add(dr);
            }
            GridView gv = new GridView();
            gv.DataSource = dt;
            gv.DataBind();
            Response.ClearContent();
            Response.AppendHeader("content-disposition", "attachment;filename=excel1.xls");
            Response.ContentType = "application/excel";

            StringWriter stringWriter = new StringWriter();
            HtmlTextWriter htmlTextWriter = new HtmlTextWriter(stringWriter);
            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

            // Set background color of each cell of GridView1 header row
            foreach (TableCell tableCell in GridView1.HeaderRow.Cells)
            {
                tableCell.Style["background-color"] = "#A55129";
                tableCell.Text.Replace("&nbsp;", " ");
               
            }

            // Set background color of each cell of each data row of GridView1
            foreach (GridViewRow gridViewRow in GridView1.Rows)
            {
                gridViewRow.BackColor = System.Drawing.Color.White;
                foreach (TableCell gridViewRowTableCell in gridViewRow.Cells)
                {
                    gridViewRowTableCell.Style["background-color"] = "#FFF7E7";
                    gridViewRowTableCell.Style.Add("border", "solid 1px #c1d8f1");
                    gridViewRowTableCell.Text.Replace("&nbsp;", " ");
                    gridViewRowTableCell.Attributes.Add("style", "mso-number-format:\\@");
                    
                }
            }
            gv.RenderControl(htmlTextWriter);
            string style = @"<style> td { mso-number-format:\@;} </style>";
            Response.Write(style);
            Response.Write(stringWriter.ToString());
            Response.End();
        }









        protected void Button4_Click(object sender, EventArgs e)
        {


            DataTable dt2 = new DataTable();
            dt2.Columns.Add("regno");
            dt2.Columns.Add("COMPANY_CODE");
            dt2.Columns.Add("mg_code");
            dt2.Columns.Add("NATIONAL_NO");
            dt2.Columns.Add("EDUCAT_LEVEL");
            dt2.Columns.Add("EMPLOYM_TYPE");
            dt2.Columns.Add("sabeghe");
            dt2.Columns.Add("Date");
            dt2.Columns.Add("SystemName");
            //dt.Columns.Add("5185");
            //dt.Columns.Add("5192");
            //dt.Columns.Add("5207");
            //dt.Columns.Add("6177");
            //dt.Columns.Add("6178");
            //dt.Columns.Add("6206");
            //dt.Columns.Add("6207");
            //dt.Columns.Add("6299");
            //dt.Columns.Add("6356");
            dt2.Columns.Add("6358");
            dt2.Columns.Add("6362");
            dt2.Columns.Add("6363");
            dt2.Columns.Add("6364");
            dt2.Columns.Add("6365");
            dt2.Columns.Add("6367");
            dt2.Columns.Add("6368");
            dt2.Columns.Add("6369");
            dt2.Columns.Add("6372");
            dt2.Columns.Add("6626");
            foreach (GridViewRow row2 in GridView1.Rows)
            {
                int j = row2.RowIndex;
                for (int k = 11; k <= 27; k++)
                {
                    if (GridView1.Rows[j].Cells[k].Text == "&nbsp;")
                        GridView1.Rows[j].Cells[k].Text = " ";
                }
                DataRow dr2 = dt2.NewRow();
                dr2["regno"] = GridView1.Rows[j].Cells[0].Text;
                dr2["COMPANY_CODE"] = GridView1.Rows[j].Cells[1].Text;
                dr2["mg_code"] = GridView1.Rows[j].Cells[2].Text;
                dr2["NATIONAL_NO"] = GridView1.Rows[j].Cells[3].Text;
                dr2["EDUCAT_LEVEL"] = GridView1.Rows[j].Cells[4].Text;
                dr2["EMPLOYM_TYPE"] = GridView1.Rows[j].Cells[5].Text;
                dr2["sabeghe"] = GridView1.Rows[j].Cells[6].Text;
                dr2["Date"] = GridView1.Rows[j].Cells[7].Text;
                dr2["SystemName"] = GridView1.Rows[j].Cells[8].Text;
                //dr["5185"] = GridView1.Rows[i].Cells[9].Text;
                //dr["5192"] = GridView1.Rows[i].Cells[10].Text;
                //dr["5207"] = GridView1.Rows[i].Cells[11].Text;
                //dr["6177"] = GridView1.Rows[i].Cells[12].Text;
               // dr["6178"] = GridView1.Rows[i].Cells[13].Text;
                //dr["6206"] = GridView1.Rows[i].Cells[14].Text;
                //dr["6207"] = GridView1.Rows[i].Cells[15].Text;
                //dr["6299"] = GridView1.Rows[i].Cells[16].Text;
                //dr["6356"] = GridView1.Rows[i].Cells[17].Text;
                dr2["6358"] = GridView1.Rows[j].Cells[18].Text;
                dr2["6362"] = GridView1.Rows[j].Cells[19].Text;
                dr2["6363"] = GridView1.Rows[j].Cells[20].Text;
                dr2["6364"] = GridView1.Rows[j].Cells[21].Text;
                dr2["6365"] = GridView1.Rows[j].Cells[22].Text;
                dr2["6367"] = GridView1.Rows[j].Cells[23].Text;
                dr2["6368"] = GridView1.Rows[j].Cells[24].Text;
                dr2["6369"] = GridView1.Rows[j].Cells[25].Text;
                dr2["6372"] = GridView1.Rows[j].Cells[26].Text;
                dr2["6626"] = GridView1.Rows[j].Cells[27].Text;
                dt2.Rows.Add(dr2);
            }
            GridView gv2 = new GridView();
            gv2.DataSource = dt2;
            gv2.DataBind();
            Response.ClearContent();
            Response.AppendHeader("content-disposition", "attachment;filename=excel1.xls");
            Response.ContentType = "application/excel";

            StringWriter stringWriter2 = new StringWriter();
            HtmlTextWriter htmlTextWriter2 = new HtmlTextWriter(stringWriter2);
            GridView1.HeaderRow.Style.Add("background-color", "#FFFFFF");

            // Set background color of each cell of GridView1 header row
            foreach (TableCell tableCell in GridView1.HeaderRow.Cells)
            {
                tableCell.Style["background-color"] = "#A55129";
            }

            // Set background color of each cell of each data row of GridView1
            foreach (GridViewRow gridViewRow in GridView1.Rows)
            {
                gridViewRow.BackColor = System.Drawing.Color.White;
                foreach (TableCell gridViewRowTableCell in gridViewRow.Cells)
                {
                    gridViewRowTableCell.Style["background-color"] = "#FFF7E7";
                }
            }
            gv2.RenderControl(htmlTextWriter2);
            string style = @"<style> td { mso-number-format:\@;} </style>";
            Response.Write(style);
            Response.Write(stringWriter2.ToString());
            Response.End();
        }
    }



}


