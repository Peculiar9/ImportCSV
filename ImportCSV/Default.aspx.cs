using ImportCSV;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ImportSample
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateData();
                lblMessage.Text = "Current Database Data!";
            }
        }

        private void PopulateData()
        {
            using (MuDatabaseEntities1 dc = new MuDatabaseEntities1())
            {
                gvData.DataSource = dc.EmployeeMasters.ToList();
                gvData.DataBind();
            }
        }



        protected void btnImport_Click(object sender, EventArgs e)
        {
            string fileName = Path.Combine(Server.MapPath("~/ImportDocument"), Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));

            FileUpload1.PostedFile.SaveAs(fileName);
            string conString = "";



            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" || FileUpload1.PostedFile.ContentType == "text/csv" ||
              FileUpload1.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                try
                {



                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\""; ;
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    }
                    else if (ext.Trim() == ".csv")
                    {
                        conString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source= " + fileName + ";Extended Properties = \"text;HDR=Yes;FMT=Delimited\"";
                    }


                    string query = "Select [Employee ID],[Company Name], [Contact Name],[Contact Title],[Employee Address],[Postal Code] from [Sheet1$]";
                    OleDbConnection con = new OleDbConnection(conString);  // you want to leave this ? Yes
                    if (con.State == System.Data.ConnectionState.Closed)
                    {
                        con.Open();
                    }
                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();

                    // Import to Database
                    using (MuDatabaseEntities1 dc = new MuDatabaseEntities1())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string empID = dr["Employee ID"].ToString();
                            var v = dc.EmployeeMasters.Where(a => a.EmployeeID.Equals(empID)).FirstOrDefault();
                            if (v != null)
                            {
                                // Update here
                                v.CompanyName = dr["Company Name"].ToString();
                                v.ContactName = dr["Contact Name"].ToString();
                                v.ContactTitle = dr["Contact Title"].ToString();
                                v.EmployeeAddress = dr["Employee Address"].ToString();
                                v.PostalCode = dr["Postal Code"].ToString();
                            }
                            else
                            {
                                // Insert
                                dc.EmployeeMasters.Add(new EmployeeMaster
                                {
                                    EmployeeID = dr["Employee ID"].ToString(),
                                    CompanyName = dr["Company Name"].ToString(),
                                    ContactName = dr["Contact Name"].ToString(),
                                    ContactTitle = dr["Contact Title"].ToString(),
                                    EmployeeAddress = dr["Employee Address"].ToString(),
                                    PostalCode = dr["Postal Code"].ToString()
                                });
                            }
                        }

                        dc.SaveChanges();
                    }

                    PopulateData();
                    lblMessage.Text = "Successfully data import done!";
                }
                catch (System.Data.Entity.Validation.DbEntityValidationException dbEx)
                {
                    Exception raise = dbEx;
                    foreach (var validationErrors in dbEx.EntityValidationErrors)
                    {
                        foreach (var validationError in validationErrors.ValidationErrors)
                        {
                            string message = string.Format("{0}:{1}",
                                validationErrors.Entry.Entity.ToString(),
                                validationError.ErrorMessage);
                            // raise a new exception nesting  
                            // the current instance as InnerException  
                            raise = new InvalidOperationException(message, raise);
                        }
                    }
                    throw raise;
                }
            }
        }

        protected void btnExport_Click(object sender, EventArgs e)
        {
            using (MuDatabaseEntities1 dc = new MuDatabaseEntities1())
            {
                List<EmployeeMaster> emList = dc.EmployeeMasters.ToList();
                StringBuilder sb = new StringBuilder();

                if (emList.Count > 0)
                {
                    string fileName = Path.Combine(Server.MapPath("~/ImportDocument"), DateTime.Now.ToString("ddMMyyyyhhmmss") + ".xlsx");
                    string conString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'";
                    using (OleDbConnection con = new OleDbConnection(conString))
                    {
                        string strCreateTab = "Create table EmployeeData (" +
                            " [Employee ID] varchar(50), " +
                            " [Company Name] varchar(200), " +
                            " [Contact Name] varchar(200), " +
                            " [Contact Title] varchar(200), " +
                            " [Employee Address] varchar(200), " +
                            " [Postal Code] varchar(50))";
                        if (con.State == ConnectionState.Closed)
                        {
                            con.Open();
                        }

                        OleDbCommand cmd = new OleDbCommand(strCreateTab, con);
                        cmd.ExecuteNonQuery();

                        string strInsert = "Insert into EmployeeData([Employee ID],[Company Name]," +
                            " [Contact Name], [Contact Title], [Employee Address], [Postal Code]" +
                            ") values(?,?,?,?,?,?)";
                        OleDbCommand cmdIns = new OleDbCommand(strInsert, con);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 50);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 200);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 200);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 200);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 200);
                        cmdIns.Parameters.Add("?", OleDbType.VarChar, 50);

                        foreach (var i in emList)
                        {
                            cmdIns.Parameters[0].Value = i.EmployeeID;
                            cmdIns.Parameters[1].Value = i.CompanyName;
                            cmdIns.Parameters[2].Value = i.ContactName;
                            cmdIns.Parameters[3].Value = i.ContactTitle;
                            cmdIns.Parameters[4].Value = i.EmployeeAddress;
                            cmdIns.Parameters[5].Value = i.PostalCode;

                            cmdIns.ExecuteNonQuery();
                        }
                    }

                    // Create Downloadable file
                    byte[] content = File.ReadAllBytes(fileName);
                    HttpContext context = HttpContext.Current;

                    context.Response.BinaryWrite(content);
                    context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    context.Response.AppendHeader("Content-Disposition", "attachment; filename=test.xlsx");
                    Context.Response.End();
                }
            }
        }

        protected void btnUpdate_Click(object sender, EventArgs e)
        {
            Response.Redirect("About.aspx");
        }


    }
}