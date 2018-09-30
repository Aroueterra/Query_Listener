using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.Data.OleDb;
using System.Diagnostics;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;
using SD = System.Data;
using GemBox.Spreadsheet;
using System.Xml;

namespace Query_Listener
{
    public partial class Dashboard : Form
    {
        /// <summary>
        /// Software created by August Bryan N. Florese  
        /// Contact: Aroueterra@gmail.com
        /// For: Tax team member Bryan Rucio of Convergys Finance department
        /// Author and developer of the code retains the program property rights
        /// </summary>
        /// <remarks> In great need of a refactor...</remarks>
        public Dashboard()
        {
            InitializeComponent();
        }
        public string Con = ConfigurationManager.ConnectionStrings["Con"].ConnectionString;
        Form Printer = new Printer();
        public static string Firstnamae = "";
        public static string Lastnamae = "";
        public static string Gross = "";
        public static string LessTNT = "";
        public static string TCI = "";
        public static string ADDTI = "";
        public static string GTI = "";
        public static string LessTE = "";
        public static string LessPPH = "";
        public static string NetTax = "";
        public static string TD = "";
        public static string HeldTaxDue = "";
        public static string HeldTaxPE = "";
        public static string HeldTaxCE = "";
        public static string TotalTax = "";
        public static string PersonID = "";
        public static string sheetName = "";
        public static string TIN_Printer = "";
        public static string From = "";
        public static string To = "";
        public static string CTC = "";
        public static string POI = "";
        public static string DOI = "";
        public static string AMT = "";
        public static string Middlenamae;
        //Printer Equity
        private Excel.Application Xls;
        private Excel.Workbooks WBs;
        private Excel.Workbook WB;
        private Excel.Worksheet WS;
        private Excel.Sheets SS;
        private string Persona;
        private string PersonaID;
        public  string Tin;
        private string lblPathings = "";
        object misValue = System.Reflection.Missing.Value;
        private string finalformat;
        SD.DataTable tbContainer = new SD.DataTable();
        OleDbConnection OleDbcon;
        int MaxSchema = 0;
        string OFDFile = "";
        int counter;
        string MatchString = "";
        Excel.Range cellsRange = null;
        long numberOfColumns = 0;
        long numberOfRows = 0;
        string DateOverride = "";
        string Foundfile = "";
        DataTable VirtualTable;
        int MaxItems;

        MicroWorker MW = new MicroWorker();

        private OleDbConnection Globalcon = new OleDbConnection();

        private async void Dashboard_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'aCDBDataSet1.ACTB' table. You can move, or remove it, as needed.
            this.aCTBTableAdapter1.Fill(this.aCDBDataSet1.ACTB);
            Initial_Refresh();
            await PutTaskDelay2();
            lblNotice.Text = "";
        }

        public void Initial_Refresh()
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    using (OleDbTransaction Scope = connection.BeginTransaction(SD.IsolationLevel.ReadCommitted))
                    {
                        try
                        {
                            string QueryEntry = "SELECT TOP 1000 * FROM ACTB";
                            OleDbDataAdapter oda = new OleDbDataAdapter(QueryEntry, connection);
                            SD.DataTable dt = new SD.DataTable();
                            oda.SelectCommand.Transaction = Scope;
                            oda.Fill(dt);
                            Scope.Commit();
                            DGVmain.DataSource = dt;
                        }
                        catch (OleDbException odx)
                        {
                            MessageBox.Show(odx.Message);
                            Scope.Rollback();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SQL error: " + ex);
            }
            finally
            {
                lblSuccess.Text = "Refreshed the view!";
            }
        }

        private void txtpcsd_TextChanged(object sender, EventArgs e)
        {
            if (cmbOmnisearch.Text == "ID")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where ID like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "FIRST_NAME")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where FIRST_NAME like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "Tax Due")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where TaxDue like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "GROSS_COMP_INCOME")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where GROSS_COMP_INCOME like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "LAST_NAME")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where LAST_NAME like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "STARTED")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where STARTED like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "TIN")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where TIN like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "ENDED")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where ENDED like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "Ended")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where Ended like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
        } //OMNI BOX

        #region Button Highlight

        private void btnDTAdd_Click(object sender, EventArgs e)
        {
            sheetName = txtSheetName.Text;
            if (string.IsNullOrEmpty(sheetName)) { sheetName = txtSheetName.Text; }
            if (btnDTAdd.FlatStyle == FlatStyle.Flat)
                btnDTAdd.FlatStyle = FlatStyle.Standard;
            else
            {
                btnDTAdd.FlatStyle = FlatStyle.Flat;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                try
                {
                    OleDbcon =
                        new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog.FileName +
                                            ";Extended Properties='Excel 12.0 Xml; HDR = Yes; IMEX = 1'");
                    OleDbcon.Open();
                    OleDbCommand cmd = new OleDbCommand();
                    OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), OleDbcon);
                    DataSet ds = new DataSet();
                    oda.Fill(ds);
                    tbContainer = ds.Tables[0];
                    DGVExcel.DataSource = tbContainer;
                    InsertRowsIntoTempTable(tbContainer);
                    OleDbcon.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Compile error: " + ex.Message);
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
        }

        #endregion // 6 Buttons //6 Buttons

        public void InsertRowIntoTempTable(string FirstName, string LastName, int PersonIDs)
        {
            using (OleDbConnection Globalcon = new OleDbConnection(Con))
            {
                var cmd = new OleDbCommand();
                cmd.Connection = Globalcon;
                cmd.CommandText = String.Format("INSERT INTO ACTB VALUES({0},{1},{2})", FirstName, LastName, PersonIDs);
                Globalcon.Open();
                cmd.ExecuteNonQuery();
                Globalcon.Close();
            }
        }

        public void InsertRowsIntoTempTable(SD.DataTable tbContainer)
        {
            foreach (DataRow row in tbContainer.Rows)
            {
                foreach (var item in row.ItemArray) // Loop over the items.
                {
                    Console.WriteLine(item); // Invokes ToString abstract method.
                }
            }
            //InsertRowIntoTempTable(row., row[1], row[1]);
        }

        public void CheckForNullsin(TextBox textbox)
        {
            
        }

        private void Insert_Single_Row_Click(object sender, EventArgs e)
        {
            Tin = txtTIN.Text;
            foreach (Control tb in panel15.Controls)
            {
                if (tb is TextBox)
                {
                    TextBox tbx = tb as TextBox;
                    if (String.IsNullOrWhiteSpace(tb.Text))
                    {
                        MessageBox.Show("Fill in the text boxes before proceeding!");
                        return;
                    }
                }
            }
            if (Tin.Length < 12)
            {

                MessageBox.Show("Taxpayer identification number must be at least 12 digits!");
                return;
            }            
            else
            {
                OleDbConnection Dbcon = new OleDbConnection(Con);
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText =
                "Insert INTO ACTB (ID, LAST_NAME, FIRST_NAME, MIDDLE_NAME, GROSS_COMP_INCOME, PRES_NONTAX_13TH_MONTH, PRES_NONTAX_DE_MINIMIS, PRES_NONTAX_SSS_ETS, PRES_NONTAX_SALARIES, TOTAL_NONTAX_COMP_INCOME, PRES_TAXABLE_BASIC_SALARY, PRES_TAXABLE_13TH_MONTH, PRES_TAXABLE_SALARIES, TOTAL_TAXABLE_COMP_INCOME, EXMPN_CODE, EXPMN_AMT, PREMIUM_PAID, NET_TAXABLE_COMP_INCOME, TAX_DUE, PRES_TAX_WTHLD, AMT_WTHLD_DEC, OVER_WTHLD, ACTUAL_AMT_WTHLD, SUBS_FILING, TIN, STARTED, ENDED, CTC, POI, DOI, AMT) " +
                          "VALUES(@ID, @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @GROSS_COMP_INCOME, @PRES_NONTAX_13TH_MONTH, @PRES_NONTAX_DE_MINIMIS, @PRES_NONTAX_SSS_ETS, @PRES_NONTAX_SALARIES, @TOTAL_NONTAX_COMP_INCOME, @PRES_TAXABLE_BASIC_SALARY, @PRES_TAXABLE_13TH_MONTH, @PRES_TAXABLE_SALARIES, @TOTAL_TAXABLE_COMP_INCOME, @EXMPN_CODE, @EXPMN_AMT, @PREMIUM_PAID, @NET_TAXABLE_COMP_INCOME, @TAX_DUE, @PRES_TAX_WTHLD, @AMT_WTHLD_DEC, @OVER_WTHLD, @ACTUAL_AMT_WTHLD, @SUBS_FILING, @TIN, @STARTED, @ENDED, @CTC, @POI, @DOI, @AMT)";
                cmd.Parameters.AddWithValue("@ID", txtID.Text);
                cmd.Parameters.AddWithValue("@LAST_NAME", txtLast.Text);
                cmd.Parameters.AddWithValue("@FIRST_NAME", txtFirst.Text);
                cmd.Parameters.AddWithValue("@MIDDLE_NAME", txtMiddle.Text);
                cmd.Parameters.AddWithValue("@GROSS_COMP_INCOME", String.IsNullOrWhiteSpace(txtGross.Text) ? 0 : Convert.ToDouble(txtGross.Text));
                cmd.Parameters.AddWithValue("@PRES_NONTAX_13TH_MONTH", String.IsNullOrWhiteSpace(txt13a.Text) ? 0 : Convert.ToDouble(txt13a.Text));
                cmd.Parameters.AddWithValue("@PRES_NONTAX_DE_MINIMIS", String.IsNullOrWhiteSpace(txtDemini.Text) ? 0 : Convert.ToDouble(txtDemini.Text));
                cmd.Parameters.AddWithValue("@PRES_NONTAX_SSS_ETS", String.IsNullOrWhiteSpace(txtsss.Text) ? 0 : Convert.ToDouble(txtsss.Text));
                cmd.Parameters.AddWithValue("@PRES_NONTAX_SALARIES", String.IsNullOrWhiteSpace(txtSala.Text) ? 0 : Convert.ToDouble(txtSala.Text));
                cmd.Parameters.AddWithValue("@TOTAL_NONTAX_COMP_INCOME", String.IsNullOrWhiteSpace(txtLessTNT.Text) ? 0 : Convert.ToDouble(txtLessTNT.Text));
                cmd.Parameters.AddWithValue("@PRES_TAXABLE_BASIC_SALARY", String.IsNullOrWhiteSpace(txtBasal.Text) ? 0 : Convert.ToDouble(txtBasal.Text));
                cmd.Parameters.AddWithValue("@PRES_TAXABLE_13TH_MONTH", String.IsNullOrWhiteSpace(txt13b.Text) ? 0 : Convert.ToDouble(txt13b.Text));
                cmd.Parameters.AddWithValue("@PRES_TAXABLE_SALARIES", String.IsNullOrWhiteSpace(txtSalb.Text) ? 0 : Convert.ToDouble(txtSalb.Text));
                cmd.Parameters.AddWithValue("@TOTAL_TAXABLE_COMP_INCOME", String.IsNullOrWhiteSpace(txtTCI.Text) ? 0 : Convert.ToDouble(txtTCI.Text));
                cmd.Parameters.AddWithValue("@EXMPN_CODE", txtExCode.Text);
                cmd.Parameters.AddWithValue("@EXPMN_AMT", (txtExAMT.Text));
                cmd.Parameters.AddWithValue("@PREMIUM_PAID", txtPPH.Text);
                cmd.Parameters.AddWithValue("@NET_TAXABLE_COMP_INCOME", txtNTI.Text);
                cmd.Parameters.AddWithValue("@TAX_DUE", txtTD.Text);
                cmd.Parameters.AddWithValue("@PRES_TAX_WTHLD", String.IsNullOrWhiteSpace(txtPresent.Text) ? 0 : Convert.ToDouble(txtPresent.Text));
                cmd.Parameters.AddWithValue("@AMT_WTHLD_DEC", String.IsNullOrWhiteSpace(txtDEC.Text) ? 0 : Convert.ToDouble(txtDEC.Text));
                cmd.Parameters.AddWithValue("@OVER_WTHLD", String.IsNullOrWhiteSpace(txtOver.Text) ? 0 : Convert.ToDouble(txtOver.Text));
                cmd.Parameters.AddWithValue("@ACTUAL_AMT_WTHLD", String.IsNullOrWhiteSpace(txtATW.Text) ? 0 : Convert.ToDouble(txtATW.Text));
                cmd.Parameters.AddWithValue("@SUBS_FILING", String.IsNullOrWhiteSpace(txtSubs.Text) ? 0 : Convert.ToDouble(txtSubs.Text));
                cmd.Parameters.AddWithValue("@TIN", txtTIN.Text);
                cmd.Parameters.AddWithValue("@STARTED", txtFrom.Text);
                cmd.Parameters.AddWithValue("@ENDED", txtTo.Text);
                cmd.Parameters.AddWithValue("@CTC", String.IsNullOrWhiteSpace(txtCTC.Text) ? 0 : Convert.ToDouble(txtCTC.Text));
                cmd.Parameters.AddWithValue("@POI", txtPOI.Text);
                cmd.Parameters.AddWithValue("@DOI", txtDOI.Text);
                cmd.Parameters.AddWithValue("@AMT", String.IsNullOrWhiteSpace(txtAMT.Text) ? 0 : Convert.ToDouble(txtAMT.Text));
                Dbcon.Open();
                cmd.Connection = Dbcon;
                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    lblSuccess.Text = "Successfully inserted the records!";
                    Dbcon.Close();
                }
            }
        }

        private void txtFirst_MouseClick(object sender, MouseEventArgs e)
        {
            SelectAll(txtFirst);
        }

        private void txtLast_MouseClick(object sender, MouseEventArgs e)
        {
            SelectAll(txtLast);
        }

        public void SelectAll(TextBox textbox)
        {
            textbox.SelectAll();
        }

        private void btnDeleterow_Click(object sender, EventArgs e)
        {
            if (txtID.Text != "" || txtID.Text != " ")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();

                        OleDbCommand cmd = new OleDbCommand("DELETE FROM ACTB WHERE ID = @ID", connection);
                        cmd.Parameters.AddWithValue("@ID", txtID.Text);
                        cmd.ExecuteNonQuery();

                    }
                    catch (OleDbException es)
                    {
                        MessageBox.Show("SQL error" + es);
                    }
                    finally
                    {
                        OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB", connection);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        sda.Fill(dt);
                        DGVmain.DataSource = dt;
                        connection.Close();
                    }
                }
            }
        }

        private void Refresh_Main()
        {
            string query = "SELECT * From ACTB";
            using (OleDbConnection conn = new OleDbConnection(Con))
            {
                conn.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    DGVmain.DataSource = ds.Tables[0];
                }
            }
            lblSuccess.Text = "Refreshed the view";
        }

        private void Refresh_Excel()
        {
            string query = "SELECT * From ACXL";
            using (OleDbConnection conn = new OleDbConnection(Con))
            {
                conn.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    DGVExcel.DataSource = ds.Tables[0];
                }
            }
        }

        private void txtDrag_Click(object sender, EventArgs e)
        {
            if (DGVmain.SelectedCells.Count > 0)
            {
                int selectedrowindex = DGVmain.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = DGVmain.Rows[selectedrowindex];
                string drags = Convert.ToString(selectedRow.Cells[0].Value);
                string dragsy = Convert.ToString(selectedRow.Cells[1].Value);
                string drag = Convert.ToString(selectedRow.Cells[2].Value);
                string drag2 = Convert.ToString(selectedRow.Cells[3].Value);
                string drag3 = Convert.ToString(selectedRow.Cells[4].Value);
                string drag4 = Convert.ToString(selectedRow.Cells[5].Value);
                string drag5 = Convert.ToString(selectedRow.Cells[6].Value);
                string drag6 = Convert.ToString(selectedRow.Cells[7].Value);
                string drag7 = Convert.ToString(selectedRow.Cells[8].Value);
                string drag8 = Convert.ToString(selectedRow.Cells[9].Value);
                string drag9 = Convert.ToString(selectedRow.Cells[10].Value);
                string drag10 = Convert.ToString(selectedRow.Cells[11].Value);
                string drag11 = Convert.ToString(selectedRow.Cells[12].Value);
                string drag12 = Convert.ToString(selectedRow.Cells[13].Value);
                string drag13 = Convert.ToString(selectedRow.Cells[14].Value);
                string drag14 = Convert.ToString(selectedRow.Cells[15].Value);
                string drag15 = Convert.ToString(selectedRow.Cells[16].Value);
                string drag16 = Convert.ToString(selectedRow.Cells[17].Value);
                string drag17 = Convert.ToString(selectedRow.Cells[18].Value);
                string drag18 = Convert.ToString(selectedRow.Cells[19].Value);
                string drag19 = Convert.ToString(selectedRow.Cells[20].Value);
                string drag20 = Convert.ToString(selectedRow.Cells[21].Value);
                string drag21 = Convert.ToString(selectedRow.Cells[22].Value);
                string drag22 = Convert.ToString(selectedRow.Cells[23].Value);
                string drag23 = Convert.ToString(selectedRow.Cells[24].Value);
                string drag24 = Convert.ToString(selectedRow.Cells[25].Value);
                string drag25 = Convert.ToString(selectedRow.Cells[26].Value);
                string drag26 = Convert.ToString(selectedRow.Cells[27].Value);
                string drag27 = Convert.ToString(selectedRow.Cells[28].Value);
                string drag28 = Convert.ToString(selectedRow.Cells[29].Value);
                string drag29 = Convert.ToString(selectedRow.Cells[30].Value);
                txtID.Text = drags;
                txtLast.Text = dragsy;
                txtFirst.Text = drag;
                txtMiddle.Text = drag2;
                txtGross.Text = drag3;
                txt13a.Text = drag4;
                txtDemini.Text = drag5;
                txtsss.Text = drag6;
                txtSala.Text = drag7;
                txtLessTNT.Text = drag8;
                txtBasal.Text = drag9;
                txt13b.Text = drag10;
                txtSalb.Text = drag11;
                txtTCI.Text = drag12;
                txtExCode.Text = drag13;
                txtExAMT.Text = drag14;
                txtPPH.Text = drag15;
                txtNTI.Text = drag16;
                txtTD.Text = drag17;
                txtPresent.Text = drag18;
                txtDEC.Text = drag19;
                txtOver.Text = drag20;
                txtATW.Text = drag21;
                txtSubs.Text = drag22;
                txtTIN.Text = drag23;
                txtFrom.Text = drag24;
                txtTo.Text = drag25;
                txtCTC.Text = drag26;
                txtPOI.Text = drag27;
                txtDOI.Text = drag28;
                txtAMT.Text = drag29;

                lblSuccess.Text = "Successfully dragged the records!";
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    OleDbCommand cmd = new OleDbCommand("UPDATE ACTB SET LAST_NAME=@LAST_NAME, FIRST_NAME=@FIRST_NAME, MIDDLE_NAME=@MIDDLE_NAME, GROSS_COMP_INCOME=@GROSS_COMP_INCOME, PRES_NONTAX_13TH_MONTH=@PRES_NONTAX_13TH_MONTH, PRES_NONTAX_DE_MINIMIS=@PRES_NONTAX_DE_MINIMIS, PRES_NONTAX_SSS_ETS=@PRES_NONTAX_SSS_ETS, PRES_NONTAX_SALARIES=@PRES_NONTAX_SALARIES, TOTAL_NONTAX_COMP_INCOME=@TOTAL_NONTAX_COMP_INCOME, PRES_TAXABLE_BASIC_SALARY=@PRES_TAXABLE_BASIC_SALARY, PRES_TAXABLE_13TH_MONTH=@PRES_TAXABLE_13TH_MONTH, PRES_TAXABLE_SALARIES=@PRES_TAXABLE_SALARIES, TOTAL_TAXABLE_COMP_INCOME= @TOTAL_TAXABLE_COMP_INCOME, EXMPN_CODE=@EXMPN_CODE, EXPMN_AMT=@EXPMN_AMT, PREMIUM_PAID=@PREMIUM_PAID, NET_TAXABLE_COMP_INCOME=@NET_TAXABLE_COMP_INCOME, TAX_DUE=@TAX_DUE, PRES_TAX_WTHLD=@PRES_TAX_WTHLD, AMT_WTHLD_DEC=@AMT_WTHLD_DEC, OVER_WTHLD=@OVER_WTHLD, ACTUAL_AMT_WTHLD=@ACTUAL_AMT_WTHLD, SUBS_FILING=@SUBS_FILING, TIN=@TIN, STARTED=@STARTED, ENDED= @ENDED, CTC=@CTC, POI=@POI, DOI=@DOI, AMT=@AMT WHERE ID = @ID", connection);
                    cmd.Connection = connection;
                    connection.Open();
                    cmd.Parameters.AddWithValue("@LAST_NAME", txtLast.Text);
                    cmd.Parameters.AddWithValue("@FIRST_NAME", txtFirst.Text);
                    cmd.Parameters.AddWithValue("@MIDDLE_NAME", txtMiddle.Text);
                    cmd.Parameters.AddWithValue("@GROSS_COMP_INCOME", Convert.ToDouble(txtGross.Text));
                    cmd.Parameters.AddWithValue("@PRES_NONTAX_13TH_MONTH", Convert.ToDouble(txt13a.Text));
                    cmd.Parameters.AddWithValue("@PRES_NONTAX_DE_MINIMIS", Convert.ToDouble(txtDemini.Text));
                    cmd.Parameters.AddWithValue("@PRES_NONTAX_SSS_ETS", Convert.ToDouble(txtsss.Text));
                    cmd.Parameters.AddWithValue("@PRES_NONTAX_SALARIES", Convert.ToDouble(txtSala.Text));
                    cmd.Parameters.AddWithValue("@TOTAL_NONTAX_COMP_INCOME", Convert.ToDouble(txtNTI.Text));
                    cmd.Parameters.AddWithValue("@PRES_TAXABLE_BASIC_SALARY", Convert.ToDouble(txtBasal.Text));
                    cmd.Parameters.AddWithValue("@PRES_TAXABLE_13TH_MONTH", Convert.ToDouble(txt13b.Text));
                    cmd.Parameters.AddWithValue("@PRES_TAXABLE_SALARIES", Convert.ToDouble(txtSalb.Text));
                    cmd.Parameters.AddWithValue("@TOTAL_TAXABLE_COMP_INCOME", Convert.ToDouble(txtTCI.Text));
                    cmd.Parameters.AddWithValue("@EXMPN_CODE", txtExCode.Text);
                    cmd.Parameters.AddWithValue("@EXPMN_AMT", txtExAMT.Text);
                    cmd.Parameters.AddWithValue("@PREMIUM_PAID", txtPPH.Text);
                    cmd.Parameters.AddWithValue("@NET_TAXABLE_COMP_INCOME", txtTCI.Text);
                    cmd.Parameters.AddWithValue("@TAX_DUE", txtTD.Text);
                    cmd.Parameters.AddWithValue("@PRES_TAX_WTHLD", Convert.ToDouble(txtPresent.Text));
                    cmd.Parameters.AddWithValue("@AMT_WTHLD_DEC", Convert.ToDouble(txtDEC.Text));
                    cmd.Parameters.AddWithValue("@OVER_WTHLD", Convert.ToDouble(txtOver.Text));
                    cmd.Parameters.AddWithValue("@ACTUAL_AMT_WTHLD", Convert.ToDouble(txtATW.Text));
                    cmd.Parameters.AddWithValue("@SUBS_FILING", Convert.ToDouble(txtSubs.Text));
                    cmd.Parameters.AddWithValue("@TIN", txtTIN.Text);
                    cmd.Parameters.AddWithValue("@STARTED", txtFrom.Text);
                    cmd.Parameters.AddWithValue("@ENDED", txtTo.Text);
                    cmd.Parameters.AddWithValue("@CTC", Convert.ToDouble(txtCTC.Text));
                    cmd.Parameters.AddWithValue("@POI", txtPOI.Text);
                    cmd.Parameters.AddWithValue("@DOI", txtDOI.Text);
                    cmd.Parameters.AddWithValue("@AMT", Convert.ToDouble(txtAMT.Text));
                    cmd.Parameters.AddWithValue("@ID", txtID.Text);
                    cmd.ExecuteNonQuery();
                    connection.Close();
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                lblSuccess.Text = "Successfully updated the records!";
                Refresh_Main();
            }
        }
        
        private void CMBsheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from [" + CMBsheets.Text + "$]", OleDbcon);
            System.Data.DataTable dt = new System.Data.DataTable();
            oledbDa.Fill(dt);
            if (dt == null)
            {
                MessageBox.Show("Invalid Selection! Try renaming the sheet.");
                return;
            }
            DGVExcel.DataSource = dt;
        }

        #region Digit Only 
        private void txtGross_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtLessTNT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtTCI_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtADDTI_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtGTI_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtLessTE_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtLessPPH_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtLessNTI_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtTD_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtTWCE_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtTWPE_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void txtTATW_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
(e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as System.Windows.Forms.TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }
        #endregion  //  // Keypress Event

        private void btnReturn_Click(object sender, EventArgs e)
        {
            Form Help = new Help();
            Help.ShowDialog();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            Lastnamae = txtLast.Text;
            Firstnamae = txtFirst.Text;
            Middlenamae = txtMiddle.Text;
            Gross = txtGross.Text;
            LessTNT = txtLessTNT.Text;
            TCI = txtTCI.Text;
            ADDTI = "0";
            GTI = txtTCI.Text;
            LessTE = txtExAMT.Text;
            LessPPH = txtPPH.Text;
            NetTax = txtNTI.Text;
            TD = txtTD.Text;
            HeldTaxDue = txtPresent.Text;
            HeldTaxPE = txtExCode.Text;
            HeldTaxCE = "";
            TotalTax = txtATW.Text;
            PersonID = txtID.Text;
            sheetName = txtSheetName.Text;
            TIN_Printer = txtTIN.Text;
            From = txtFrom.Text;
            To = txtTo.Text;
            CTC = txtCTC.Text;
            POI = txtPOI.Text;
            DOI = txtDOI.Text;
            AMT = txtAMT.Text;
            if (Tin.Length < 12)
            {
                MessageBox.Show("Taxpayer identification number must be at least 12 digits long!");
                return;
            }
            Printer = new Printer();
            Printer.ShowDialog();
        }

        public void Validator()
        {
            List<string> ExcludedColumnsList = new List<string> { "ID", "LAST_NAME", "FIRST_NAME", "MIDDLE_NAME", "EXMPN_CODE", "SUBS_FILING", "TIN", "STARTED", "ENDED", "POI", "DOI" };

            foreach (DataGridViewRow row in DGVExcel.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    if (!ExcludedColumnsList.Contains(DGVExcel.Columns[i].Name))
                    {
                        if (row.Cells[i].Value == null || row.Cells[i].Value == DBNull.Value ||
                             String.IsNullOrWhiteSpace(row.Cells[i].Value.ToString()))
                        {
                            row.Cells[i].Value = 0;
                            //DGVExcel.RefreshEdit();
                        }
                    }
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            if (DGVExcel.Rows.Count == 0)
            {
                return;
            }
            DialogResult dr = MessageBox.Show(
                "Warning: when importing data into the Access database, ensure that the field columns match the Access fields (TIN must be the last field). Do you still wish to proceed?", "Import caution",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                lblSuccess.Text = "Loading, please wait!";
                this.Enabled = false;
                if (VirtualTable != null)
                {
                    VirtualTable.Clear(); 
                }
                GetDataTableFromDGV(DGVExcel);
                Transfer_Worker.RunWorkerAsync();
            }
        }

        private void btnColumn_Click(object sender, EventArgs e)
        {
            if (btnColumn.FlatStyle == FlatStyle.Flat)
                btnColumn.FlatStyle = FlatStyle.Standard;
            else
            {
                btnColumn.FlatStyle = FlatStyle.Flat;
            }

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            string pather = openFileDialog.FileName;

            if (pather == "" || pather == " " || pather == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                try
                {
                    Xls = new Excel.Application();
                    WBs = Xls.Workbooks;
                    WB = WBs.Open(pather, 0, false, 5, "", "", true,
                        Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    SS = WB.Worksheets;
                    WS = SS.get_Item(1);
                    cellsRange = WS.UsedRange;
                    numberOfColumns = cellsRange.Columns.CountLarge;
                    numberOfRows = cellsRange.Rows.CountLarge;
                    //long lastCell = numberOfColumns+1;
                    //WS.Cells[1, lastCell] = "Tax Formula";
                    for (long i = 2; i <= numberOfRows; i++)
                    {
                        string quarry =
                            "=IF(R" + i + "<10001,R" + i + "*0.05,IF(R" + i + "<30001,(R" + i + "-10000)*0.1+500,IF(R" + i + "<70001,(R" + i + "-30000)*0.15+2500,IF(R" + i + "<140001,(R" + i + "-70000)*0.2+8500,IF(R" + i + "<250001,(R" + i + "-140000)*0.25+22500,IF(R" + i + "<500001,(R" + i + "-250000)*0.3+50000,IF(R" + i + ">500000,(R" + i + "-500000)*0.32+125000)))))))";
                        WS.Cells[i, 12].Formula = quarry;
                        // =IF(B4<400,B4*7%,IF(B4<750,B4*10%,IF(B4<1000,B4*12.5%,B4*16%)))
                        //WS.Cells[i, LastCell].Value = (WS.Cells[i, 12] + WS.Cells[i, 13]) * WS.Cells[i, 14];
                    }
                    //==========================
                    WB.Save();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Write Excel: " + ex.Message);
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    try
                    {
                        WB.Close();
                        Xls.Quit();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Closing Excel: " + ex.Message);
                    }

                    releaseObject(cellsRange);
                    releaseObject(SS);
                    releaseObject(WS);
                    releaseObject(WBs);
                    releaseObject(WB);
                    releaseObject(Xls);
                }
                MessageBox.Show("Finished Updating File", "Task complete");
                lblSuccess.Text = "Successfully updated the records!";
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }

        }

        private void txtSheetName_MouseClick(object sender, MouseEventArgs e)
        {
            txtSheetName.SelectAll();
        }

        private void btnConcatenate_Click(object sender, EventArgs e)
        {
            sheetName = txtSheetName.Text;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            if ( txtSheetName.Text == "Enter a sheet name"|| (string.IsNullOrWhiteSpace(txtSheetName.Text))){
                MessageBox.Show("Invalid sheet name!");
                return;
            }
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                MatchString = openFileDialog.FileName;
                lblSuccess.Text = "Processing match... please wait.";
                this.Enabled = false;
                MatchWorker.RunWorkerAsync();
            }
        }

        private void btnDeleteAll_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(
                "Warning: All data in the table view and Access file will be permanently deleted. The data will no longer be recoverable, do you wish to proceed?", "Disposing records",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    try
                    {
                        connection.Open();
                        OleDbCommand cmd = new OleDbCommand("DELETE FROM ACTB", connection);
                        cmd.ExecuteNonQuery();

                    }
                    catch (OleDbException es)
                    {
                        MessageBox.Show("SQL error: " + es);
                    }
                    finally
                    {
                        OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB", connection);
                        System.Data.DataTable dt = new System.Data.DataTable();
                        sda.Fill(dt);
                        DGVmain.DataSource = dt;
                        connection.Close();
                        Refresh_Main();
                        MessageBox.Show("All the records in the excel file have been disposed of!");
                    }
                }
            }
        }

        private void txtTIN_TextChanged(object sender, EventArgs e)
        {
            pbChecks.Visible = false;
            if (txtTIN.TextLength >= 12)
            {
                pbChecks.Visible = true;
            }
        }

        private void DGVmain_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnPeriod_Click(object sender, EventArgs e)
        {
            sheetName = txtSheetName.Text;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            if (txtSheetName.Text == "Enter a sheet name" || (string.IsNullOrWhiteSpace(txtSheetName.Text)))
            {
                MessageBox.Show("Invalid sheet name!");
                return;
            }
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    try
                    {
                        conn.Open();
                        string sqls = @"INSERT INTO ABCD SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text + "$];";
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = sqls;
                        cmd.ExecuteNonQuery();
                        string updater =
                            @"UPDATE ACTB " + @"INNER JOIN ABCD on ACTB.ID = ABCD.ID " +
                            @"SET ACTB.Started = ABCD.Started, " +
                            @"ACTB.Ended = ABCD.Ended ";
                        cmd.CommandText = updater;
                        cmd.ExecuteNonQuery();
                        string deleter = @"DELETE from ABCD";
                        cmd.CommandText = deleter;
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Combine error: " + ex);
                    }
                    finally
                    {
                        Refresh_Main();
                        Refresh_Excel();
                        lblSuccess.Text = "Successfully imported the records!";
                        MessageBox.Show("Action Complete");
                    }
                }
            }
        }

        private void btnIssue_Click(object sender, EventArgs e)
        {
            sheetName = txtSheetName.Text;
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            if (txtSheetName.Text == "Enter a sheet name" || (string.IsNullOrWhiteSpace(txtSheetName.Text)))
            {
                MessageBox.Show("Invalid sheet name!");
                return;
            }
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    try
                    {
                        conn.Open();
                        string sqls = @"INSERT INTO ACTC SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text + "$];";
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = sqls;
                        cmd.ExecuteNonQuery();

                        string updater =
                            @" UPDATE ACTB INNER JOIN ACTC on ACTB.[ID] = ACTC.[ID]
                               SET ACTB.[CTC] = ACTC.[CTC], 
                                   ACTB.[POI] = ACTC.[POI], 
                                   ACTB.[DOI] = ACTC.[DOI]
                                      ";
                        cmd.CommandText = updater;
                        cmd.ExecuteNonQuery();
                        string deleter = @"DELETE from ACTC";
                        cmd.CommandText = deleter;
                        cmd.ExecuteNonQuery();
                        conn.Close();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Combine error: " + ex);
                    }
                    finally
                    {
                        Refresh_Main();
                        Refresh_Excel();
                        lblSuccess.Text = "Successfully imported the records!";
                        MessageBox.Show("Action Complete");
                    }
                }
            }
        }

        private void btnMacro_Click(object sender, EventArgs e)
        {
            if (lblIncrement.Text == "x1k")
            {
                if (counter == 0)
                {
                    DGVmain.CurrentCell = DGVmain.Rows[1].Cells[0];
                    int CurrentRowIndex = DGVmain.CurrentCell.RowIndex;
                    if ((CurrentRowIndex + 1000) > DGVmain.RowCount)
                    {
                        return;
                    }
                    for (int i = CurrentRowIndex; i < CurrentRowIndex + 1000; i++)
                    {
                        DGVmain.Rows[i].Selected = true;
                    }
                }
            }
            if (lblIncrement.Text == "x10k")
            {
                if (counter == 0)
                {
                    DGVmain.CurrentCell = DGVmain.Rows[1].Cells[0];
                    for (int i = 0; i < 10000; i++)
                    {
                        DGVmain.Rows[i].Selected = true;
                    }
                    counter++;
                }
                else if (counter == 1)
                {
                    DGVmain.CurrentCell = DGVmain.Rows[10001].Cells[0];
                    for (int i = 10000; i < 20000; i++)
                    {
                        DGVmain.Rows[i].Selected = true;
                    }
                    counter++;
                }
                else if (counter == 2)
                {
                    DGVmain.CurrentCell = DGVmain.Rows[20000].Cells[0];
                    for (int i = 20000; i < DGVmain.RowCount; i++)
                    {
                        DGVmain.Rows[i].Selected = true;
                    }
                    counter = 0;
                }
            }
        }

        public void generateID()
        {
            string append = Path.GetDirectoryName(lblSuccess.Text) + Path.DirectorySeparatorChar;
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var random = new Random();
            var result = new string(
                Enumerable.Repeat(chars, 3)
                          .Select(s => s[random.Next(s.Length)])
                          .ToArray());
            finalformat = append + PersonaID + "-" + Persona + "-" + Tin + "-" + DateOverride + ".pdf";
        }

        private void releaseObjects(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
        }

        public void Blanker(OfficeOpenXml.ExcelWorksheet WS)
        {

            WS.Cells[11, 9].Value = "";
            WS.Cells[11, 12].Value = "";
            WS.Cells[11, 15].Value = "";
            WS.Cells[11, 18].Value = "";
            WS.Cells[45, 9].Value = "";
            WS.Cells[45, 12].Value = "";
            WS.Cells[45, 15].Value = "";
            WS.Cells[45, 18].Value = "";
            WS.Cells[64, 12].Value = "";
            WS.Cells[66, 12].Value = "";
            WS.Cells[68, 12].Value = "";
            WS.Cells[70, 12].Value = "";
            WS.Cells[72, 12].Value = "";
            WS.Cells[74, 12].Value = "";
            WS.Cells[76, 12].Value = "";
            WS.Cells[78, 12].Value = "";
            WS.Cells[80, 12].Value = "";
            WS.Cells[82, 12].Value = "";
            WS.Cells[84, 12].Value = "";
            WS.Cells[86, 12].Value = "";
            WS.Cells[33, 31].Value = "";
            WS.Cells[35, 31].Value = "";
            WS.Cells[39, 31].Value = "";
            WS.Cells[42, 31].Value = "";
            WS.Cells[47, 31].Value = "";
            WS.Cells[86, 31].Value = "";
            WS.Cells[29, 5].Value = "";
            WS.Cells[29, 11].Value = "";
            WS.Cells[97, 5].Value = "";
            WS.Cells[97, 15].Value = "";
            WS.Cells[97, 24].Value = "";
        }

        public void Tanker(GemBox.Spreadsheet.ExcelWorksheet WS)
        {
            WS.Cells[11, 9].Style.NumberFormat = "#.##0,00";
            WS.Cells[11, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[11, 15].Style.NumberFormat = "#.##0,00";
            WS.Cells[11, 18].Style.NumberFormat = "#.##0,00";
            WS.Cells[45, 9].Style.NumberFormat = "#.##0,00";
            WS.Cells[45, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[45, 15].Style.NumberFormat = "#.##0,00";
            WS.Cells[45, 18].Style.NumberFormat = "#.##0,00";
            WS.Cells[64, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[66, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[68, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[70, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[72, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[74, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[76, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[78, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[80, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[82, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[84, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[86, 12].Style.NumberFormat = "#.##0,00";
            WS.Cells[33, 31].Style.NumberFormat = "#.##0,00";
            WS.Cells[35, 31].Style.NumberFormat = "#.##0,00";
            WS.Cells[39, 31].Style.NumberFormat = "#.##0,00";
            WS.Cells[42, 31].Style.NumberFormat = "#.##0,00";
            WS.Cells[47, 31].Style.NumberFormat = "#.##0,00";
            WS.Cells[86, 31].Style.NumberFormat = "#.##0,00";
        }

        private void btnMicro_Click(object sender, EventArgs e)
        {

        }

        private void btnSelecta_Click(object sender, EventArgs e)
        {
            DGVmain.CurrentCell = DGVmain.Rows[1].Cells[0];
            for (int i = 0; i < DGVmain.RowCount / 2; i++)
            {
                DGVmain.Rows[i].Selected = true;
            }

        }

        private void btnSelecto_Click(object sender, EventArgs e)
        {
            for (int i = DGVmain.RowCount / 2; i < DGVmain.RowCount; i++)
            {
                int roko = DGVmain.RowCount / 2;
                DGVmain.CurrentCell = DGVmain.Rows[roko].Cells[0];
                //DGVmain.CurrentRow.Selected = false;
                DGVmain.Rows[i].Selected = true;
            }
        }

        private void btnEmptyRows_Click(object sender, EventArgs e)
        {
            for (int i = DGVExcel.Rows.Count - 1; i > -1; i--)
            {
                DataGridViewRow row = DGVExcel.Rows[i];
                if (!row.IsNewRow && row.Cells[0].Value == null)
                {
                    DGVExcel.Rows.RemoveAt(i);
                }
            }
        }

        private void DGVmain_SelectionChanged(object sender, EventArgs e)
        {
            if (DGVmain.RowCount > 0 && DGVmain.SelectedRows.Count > 0)
            {
                lblItem.Text = DGVmain.SelectedRows.Count + " : " + DGVmain.RowCount.ToString();
            }
        }

        private void ChkBxAll_CheckedChanged(object sender, EventArgs e)
        {
            if (ChkBxAll.Checked)
            {
                DGVmain.SelectAll();
            }
            else
            {
                ChkBxAll.Checked = false;
                DGVmain.ClearSelection();
            }
        }

        private void lblIncrement_DoubleClick(object sender, EventArgs e)
        {
            if (!(lblIncrement.Text == "x10k"))
            {
                lblIncrement.Text = "x10k";
            }
            else { lblIncrement.Text = "x1k"; }
        }

        

        private async void btnSum_Click(object sender, EventArgs e)
        {
            if (DGVmain.RowCount > 0)
            {
                OpenFileDialog sdf = new OpenFileDialog();
                sdf.Filter = "Excel Files|*.xls;*.xlsx";
                DialogResult dr = sdf.ShowDialog();
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                if (dr != DialogResult.OK)
                {
                    return;
                }
                Foundfile = sdf.FileName;
                lblSuccess.Text = Foundfile;
                if (txtFrom.Text == "" || txtTo.Text == "" || txtDOI.Text == "")
                {
                    txtFrom.BackColor = Color.Goldenrod; txtTo.BackColor = Color.Goldenrod; txtDOI.BackColor = Color.Goldenrod;
                    MessageBox.Show("Please fill in the missing form data on the left pane!");
                    await PutTaskDelay();
                    txtFrom.BackColor = Color.White; txtTo.BackColor = Color.White; txtDOI.BackColor = Color.White;
                    return;
                }
                else
                {
                    MessageBox.Show("Macro initiated, this may take several minutes...");
                    MW.MaxRows = DGVmain.SelectedRows.Count;
                    lblItem.Text = "Processing: " + MW.MaxRows;
                    this.Enabled = false;
                    Micro_Worker.RunWorkerAsync();
                }
            }
        }

        public class MicroWorker
        {
            public int MaxRows { get; set; }
        }



        public void Micro_Void(FileInfo file, int current)
        {

            for(int rowindex = 0; rowindex < VirtualTable.Rows.Count; rowindex++)
            {
                Stopwatch ws = new Stopwatch();
                ws.Start();
                current++;
                string CTC = "";
                string POI = "";
                string DOI = "";
                string SUBS = "";
                string Started = "";
                string Ended = "";
                string AMT = "";
                List<string> StringRows = new List<string>();
                List<string> CTCRows = new List<string>();
                List<string> DateRows = new List<string>();
                List<Double> DoubleRows = new List<Double>();
                List<string> ID = new List<string>();
                lblPathings = Path.ChangeExtension(Foundfile, null);

                ID.Add(Convert.ToString(VirtualTable.Rows[rowindex][0]));
                for (int i = 1; i < 4; i++)
                {
                    String StringValue = VirtualTable.Rows[rowindex][i]== DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][i]);
                    if (StringValue == null || StringValue == "") StringValue = " ";
                    StringRows.Add(StringValue); //0-2
                }
                String StringValues = "";
                StringRows.Add(StringValues = VirtualTable.Rows[rowindex][14]== DBNull.Value ? " " : Convert.ToString(VirtualTable.Rows[rowindex][14])); //3 Code
                for (int i = 23; i < 25; i++)
                {
                    String StringValue = VirtualTable.Rows[rowindex][i]== DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][i]);
                    StringRows.Add(StringValue);
                }
                for (int i = 4; i < 14; i++)
                {
                    Double DoubleValue = VirtualTable.Rows[rowindex][i]== DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][i]);
                    DoubleRows.Add(DoubleValue);
                }
                for (int i = 15; i < 23; i++)
                {
                    Double DoubleValue = VirtualTable.Rows[rowindex][i]== DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[rowindex][i]);
                    DoubleRows.Add(DoubleValue);
                }
                if (chkbxCTC.Checked == true)
                {
                    for (int i = 27; i < 30; i++)
                    {
                        String StringValue = "";
                        CTCRows.Add(StringValue = VirtualTable.Rows[rowindex][i] == DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][i]));
                    }
                }
                if (chkbxDate.Checked == true)
                {
                    String DateValue = "";
                    DateRows.Add(DateValue = VirtualTable.Rows[rowindex][23] == DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][23]));
                    DateRows.Add(DateValue = VirtualTable.Rows[rowindex][25] == DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][25]));
                    DateRows.Add(DateValue = VirtualTable.Rows[rowindex][26] == DBNull.Value ? "" : Convert.ToString(VirtualTable.Rows[rowindex][26]));
                }
                string PIN = StringRows.ElementAt(5);
                string LastName = StringRows.ElementAt(0);
                string FirstName = StringRows.ElementAt(1);
                string MiddleName = StringRows.ElementAt(2);
                string Expn = StringRows.ElementAt(3);
                if (chkbxCTC.Checked)
                {
                    CTC = CTCRows.ElementAt(0);
                    POI = CTCRows.ElementAt(1);
                    DOI = CTCRows.ElementAt(2);
                    AMT = CTCRows.ElementAt(3);
                }
                if (chkbxDate.Checked)
                {
                    SUBS = DateRows.ElementAt(0);
                    Started = DateRows.ElementAt(1);
                    Ended = DateRows.ElementAt(2);
                }
                if (PIN.Length == 8) PIN = PIN + "0";
                if (PIN.Length == 7) PIN = PIN + "00";
                if (PIN.Length == 6) PIN = PIN + "000";
                if (PIN.Length == 5) PIN = PIN + "0000";
                if (PIN.Length == 4) PIN = PIN + "00000";
                if (PIN.Length == 3) PIN = PIN + "000000";
                if (PIN.Length == 2) PIN = PIN + "0000000";
                if (PIN.Length == 1) PIN = PIN + "00000000";
                Tin = PIN;
                Console.WriteLine(Tin);
                Persona = Convert.ToString(VirtualTable.Rows[rowindex][1]);
                for (int i = 0; i < Persona.Length; i++)
                {
                    Persona = Persona.Split(' ')[0];
                }
                DateOverride = txtDOI.Text;
                DateOverride = DateOverride.Replace('/', '.');
                DateOverride = DateOverride.Replace('-', '.');
                DateOverride = DateOverride.Replace(' ', '.');
                PersonaID = Convert.ToString(VirtualTable.Rows[rowindex][0]);
                generateID();
                using (ExcelPackage p = new ExcelPackage(file))
                {
                    ExcelWorkbook WBook = p.Workbook;
                    if (WBook != null)
                    {
                        if (WBook.Worksheets.Count > 0)
                        {
                            OfficeOpenXml.ExcelWorksheet WS = WBook.Worksheets.First();
                            //
                            string input = Tin;
                            string digit1 = "0";
                            string digit2 = "0";
                            string digit3 = "0";
                            string digit4 = "0";
                            if ((input.Length < 3 || input == "0" || input == "" || input == " "))
                            {
                                WS.Cells[11, 9].Value = 000;
                                WS.Cells[11, 12].Value = 000;
                                WS.Cells[11, 15].Value = 000;
                                WS.Cells[11, 18].Value = 000;
                            }
                            else
                            {
                                StringBuilder sb = new StringBuilder();
                                StringBuilder partBuilder = new StringBuilder();
                                int partsSplitted = 0;
                                for (int i = 1; i <= input.Length; i++)
                                {
                                    partBuilder.Append(input[i - 1]);
                                    if (i % 3 == 0 && partsSplitted <= 3)
                                    {
                                        sb.Append(' ');
                                        sb.Append(partBuilder.ToString());
                                        partBuilder = new StringBuilder();
                                        partsSplitted++;
                                    }
                                }
                                if (input.Length < 10)
                                {
                                    partBuilder.Append(" 000");
                                }
                                sb.Append(partBuilder.ToString());
                                string formatted = sb.ToString().TrimStart();
                                string[] formatCollection = formatted.Split(' ');
                                digit1 = formatCollection[0];
                                digit2 = formatCollection[1];
                                digit3 = formatCollection[2];
                                digit4 = formatCollection[3];
                            }
                            //Names
                            if (FirstName == null || String.IsNullOrWhiteSpace(FirstName.ToString()))
                            {
                                WS.Cells[14, 2].Value = LastName;
                            }
                            else if (chkBxMiddle.Checked)
                            {
                                WS.Cells[14, 2].Value = LastName + ",  " + FirstName + " " + MiddleName;
                            }
                            else
                            {
                                WS.Cells[14, 2].Value = LastName + ",  " + FirstName + " ";
                            }
                            //Blanker(WS.Name);
                            //============================== Year & Period
                            if (chkbxDate.Checked == true)
                            {
                                WS.Cells[8, 8].Value = txtDOI.Text;
                                WS.Cells[8, 29].Value = Started;
                                WS.Cells[8, 34].Value = Ended;
                            }
                            else
                            {
                                string strong = txtDOI.Text;
                                string[] col;
                                strong = strong.Replace('/', '.');
                                strong = strong.Replace('-', '.');
                                strong = strong.Replace(' ', '.');
                                col = strong.Split('.');
                                int Indexes = col.Count();
                                if (Indexes == 0 || Indexes > 3)
                                {
                                    MessageBox.Show("Invalid date format! Correct format: 7.7.2077");
                                    return;
                                }
                                WS.Cells[8, 8].Value = col.Last();
                                WS.Cells[8, 29].Value = txtFrom.Text;
                                WS.Cells[8, 34].Value = txtTo.Text;
                            }
                            WS.Cells[11, 9].Value = digit1;
                            WS.Cells[11, 12].Value = digit2;
                            WS.Cells[11, 15].Value = digit3;
                            WS.Cells[11, 18].Value = digit4;
                            //WS.Cells[45, 9].Value = digit1;
                            //WS.Cells[45, 12].Value = digit2;
                            //WS.Cells[45, 15].Value = digit3;
                            //WS.Cells[45, 18].Value = digit4;
                            char InitChar = StringRows.ElementAt(3)[0];
                            WS.Cells[29, 5].Value = "";
                            WS.Cells[29, 11].Value = "";
                            if (Expn.StartsWith("S") || Expn.StartsWith("s"))
                                WS.Cells[29, 5].Value = "X";
                            if (Expn.StartsWith("M") || Expn.StartsWith("m"))
                                WS.Cells[29, 11].Value = "X";
                            WS.Cells[64, 12].Value = DoubleRows.ElementAt(0);
                            WS.Cells[64, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[66, 12].Value = DoubleRows.ElementAt(5);
                            WS.Cells[68, 12].Value = DoubleRows.ElementAt(9);
                            WS.Cells[66, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[68, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[70, 12].Value = "";
                            WS.Cells[72, 12].Value = DoubleRows.ElementAt(9);
                            WS.Cells[74, 12].Value = DoubleRows.ElementAt(10);
                            WS.Cells[72, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[74, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[76, 12].Value = "";
                            WS.Cells[78, 12].Value = DoubleRows.ElementAt(12);
                            WS.Cells[80, 12].Value = DoubleRows.ElementAt(13);
                            WS.Cells[78, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[80, 12].Style.Numberformat.Format = "#,##0.00";
                            //========================== Last 3
                            WS.Cells[82, 12].Value = DoubleRows.ElementAt(14);
                            WS.Cells[82, 12].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[84, 12].Value = "";
                            WS.Cells[86, 12].Value = DoubleRows.ElementAt(17);
                            WS.Cells[86, 12].Style.Numberformat.Format = "#,##0.00";
                            //========================== Form
                            if (chkbxCTC.Checked)
                            {
                                WS.Cells[97, 5].Value = CTC;
                                WS.Cells[97, 15].Value = POI;
                                WS.Cells[97, 24].Value = POI;
                                WS.Cells[97, 33].Value = AMT;
                            }
                            //========================== Right-Seid
                            WS.Cells[32, 31].Value = DoubleRows.ElementAt(2);
                            WS.Cells[35, 31].Value = DoubleRows.ElementAt(3);
                            WS.Cells[38, 31].Value = DoubleRows.ElementAt(5);
                            WS.Cells[41, 31].Value = DoubleRows.ElementAt(5);
                            WS.Cells[46, 31].Value = DoubleRows.ElementAt(9);
                            WS.Cells[86, 31].Value = DoubleRows.ElementAt(9);
                            WS.Cells[32, 31].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[35, 31].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[38, 31].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[41, 31].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[46, 31].Style.Numberformat.Format = "#,##0.00";
                            WS.Cells[86, 31].Style.Numberformat.Format = "#,##0.00";
                        }
                    }
                    if (ChkbxGen.CheckState == CheckState.Checked)
                    {
                        string recarm = Path.ChangeExtension(finalformat, null) + ".xlsx";
                        FileInfo formatFinal = new FileInfo(recarm);
                        p.SaveAs(formatFinal);
                    }
                    else { p.SaveAs(file); }
                    ws.Stop();
                    Console.WriteLine("XLSX took: " + ws.Elapsed.Seconds.ToString());
                }
                try
                {
                    Stopwatch sw = new Stopwatch();
                    sw.Start();
                    var workbook = ExcelFile.Load(Foundfile);
                    foreach (var worksheet in workbook.Worksheets)
                    {
                        var printOptions = worksheet.PrintOptions;
                        printOptions.LeftMargin = .4;
                        printOptions.RightMargin = 0;
                        printOptions.TopMargin = 0;
                        printOptions.BottomMargin = 0;
                        //printOptions.AutomaticPageBreakScalingFactor = 120;
                        printOptions.FitWorksheetWidthToPages = 1;
                    }
                    var saveOptions = new PdfSaveOptions();
                    saveOptions.SelectionType = SelectionType.EntireFile;
                    workbook.Save(finalformat, saveOptions);
                    sw.Stop();
                    Console.WriteLine("PDF took: "+sw.Elapsed.Seconds.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("PDF Conversion failed: " + ex.Message);
                } //PDF
                Micro_Worker.ReportProgress(current);
            }

        }

        private DataTable GetDataTableFromDGV(DataGridView dgv)
        {
            VirtualTable = new DataTable();
            foreach (DataGridViewColumn column in dgv.Columns)
            {
                if (column.Visible)
                {
                    VirtualTable.Columns.Add();
                }
            }

            object[] cellValues = new object[dgv.Columns.Count];
            foreach (DataGridViewRow row in dgv.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                VirtualTable.Rows.Add(cellValues);
            }
            return VirtualTable;
        }

        private void btnNameAppend_Click(object sender, EventArgs e)
        {


        }

        

        async Task PutTaskDelay()
        {
            await Task.Delay(5000);
        }

        async Task PutTaskDelay2()
        {
            await Task.Delay(10000);
        }

        private void DGVExcel_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void chkbxDate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxDate.Checked == true)
            {
                chkbxCTC.Checked = true;
            }
        }

        private void chkbxCTC_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxDate.Checked == true)
            {
                chkbxCTC.Checked = true;
            }
        }

        private void ReplaceDB_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Here");
            OpenFileDialog OFD = new OpenFileDialog();
            OFD.Filter = "Access Database|*.accdb|All files(*.*)|*.*";
            OFD.ShowDialog();
            string FoundFile = OFD.FileName;
            if (!string.IsNullOrEmpty(FoundFile))
            {
                try
                {
                    string Src = FoundFile;
                    string Init = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + Src;
                    ChangeConnectionString("Con", Init, "System.Data.OleDb", "Query Listener");
                    lblSuccess.Text = "Successfully replaced the data source!";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error reading file: " + ex.Message);
                    return;
                }
            }
        }

        public static bool ChangeConnectionString(string Name, string value, string providerName, string AppName)
        {
            bool retVal = false;
            try
            {
                string FILE_NAME = string.Concat(System.Windows.Forms.Application.StartupPath, "\\", AppName.Trim(), ".exe.Config"); //the application configuration file name
                XmlTextReader reader = new XmlTextReader(FILE_NAME);
                XmlDocument doc = new XmlDocument();
                doc.Load(reader);
                reader.Close();
                string nodeRoute = string.Concat("connectionStrings/add");

                XmlNode cnnStr = null;
                XmlElement root = doc.DocumentElement;
                XmlNodeList Settings = root.SelectNodes(nodeRoute);

                for (int i = 0; i < Settings.Count; i++)
                {
                    cnnStr = Settings[i];
                    if (cnnStr.Attributes["name"].Value.Equals(Name))
                        break;
                    cnnStr = null;
                }

                cnnStr.Attributes["connectionString"].Value = value;
                cnnStr.Attributes["providerName"].Value = providerName;
                doc.Save(FILE_NAME);
                retVal = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                retVal = false;
            }
            finally
            {
                MessageBox.Show(
                "You must restart the program in order for the changes to take effect!", "Data source was changed",
                MessageBoxButtons.OK, MessageBoxIcon.Stop);
                System.Windows.Forms.Application.Exit();
            }
            return retVal;
        }

        private void btnRefresher_Click(object sender, EventArgs e)
        {
            lblSuccess.Text = "Loading...";
            Refresh_Main();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            DialogResult DR = openFileDialog.ShowDialog();
            if (DR == DialogResult.OK)
            {
                lblSuccess.Text = "Loading the file into the program, please wait...";
                OFDFile = openFileDialog.FileName;
                CMBsheets.Items.Clear();
                this.Enabled = false;
                Import_Worker.RunWorkerAsync();
            }
        }

#region Background Workers
        private void Import_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                SD.DataTable dts;
                OleDbcon = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + OFDFile + ";Extended Properties=  'Excel 12.0 Xml;HDR=Yes; IMEX = 1;TypeGuessRows=0;ImportMixedTypes=Text';");
                OleDbcon.Open();
                dts = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                MaxItems = dts.Rows.Count;
                OleDbcon.Close();
                for (int i = 0; i < dts.Rows.Count; i++)
                {
                    String sheetName = dts.Rows[i]["TABLE_NAME"].ToString();
                    sheetName = sheetName.Substring(0, sheetName.Length - 1);
                    CMBsheets.Invoke((MethodInvoker)delegate
                    {
                        CMBsheets.Items.Add(sheetName);
                    });
                    Import_Worker.ReportProgress(i);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Write Excel: " + ex.Message);
            }
            finally
            {
                Import_Worker.ReportProgress(MaxItems);
            }
        }
        private async void Import_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Red;
                lblSuccess.Text = "Error: The thread aborted";
            }
            else
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Lime;
                lblItem.Text = "";
            }
            lblSuccess.Text = "Successfully opened the file!";
            MessageBox.Show("Browse the file by selecting a sheet from the combo box above.");
            this.Enabled = true;
            await PutTaskDelay();
            pbar.Value = 0;
        }
        private void Import_Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbar.Value = (e.ProgressPercentage * 100) / MaxItems;
        }

        private void Transfer_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {

                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    using (OleDbCommand cmd = new OleDbCommand())
                    {
                        cmd.Connection = conn;
                        conn.Open();
                        if (!chkbxCTC.Checked == true)
                        {
                            cmd.CommandText =
                                "Insert INTO ACTB (ID, LAST_NAME, FIRST_NAME, MIDDLE_NAME, GROSS_COMP_INCOME, PRES_NONTAX_13TH_MONTH, PRES_NONTAX_DE_MINIMIS, PRES_NONTAX_SSS_ETS, PRES_NONTAX_SALARIES, TOTAL_NONTAX_COMP_INCOME, PRES_TAXABLE_BASIC_SALARY, PRES_TAXABLE_13TH_MONTH, PRES_TAXABLE_SALARIES, TOTAL_TAXABLE_COMP_INCOME, EXMPN_CODE, EXPMN_AMT, PREMIUM_PAID, NET_TAXABLE_COMP_INCOME, TAX_DUE, PRES_TAX_WTHLD, AMT_WTHLD_DEC, OVER_WTHLD, ACTUAL_AMT_WTHLD, SUBS_FILING, TIN) " +
                                "VALUES(@ID, @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @GROSS_COMP_INCOME, @PRES_NONTAX_13TH_MONTH, @PRES_NONTAX_DE_MINIMIS, @PRES_NONTAX_SSS_ETS, @PRES_NONTAX_SALARIES, @TOTAL_NONTAX_COMP_INCOME, @PRES_TAXABLE_BASIC_SALARY, @PRES_TAXABLE_13TH_MONTH, @PRES_TAXABLE_SALARIES, @TOTAL_TAXABLE_COMP_INCOME, @EXMPN_CODE, @EXPMN_AMT, @PREMIUM_PAID, @NET_TAXABLE_COMP_INCOME, @TAX_DUE, @PRES_TAX_WTHLD, @AMT_WTHLD_DEC, @OVER_WTHLD, @ACTUAL_AMT_WTHLD, @SUBS_FILING,@TIN)";
                        }
                        else if ((chkbxCTC.Checked == true) && (!chkbxDate.Checked == true))
                        {
                            cmd.CommandText =
                                "Insert INTO ACTB (ID, LAST_NAME, FIRST_NAME, MIDDLE_NAME, GROSS_COMP_INCOME, PRES_NONTAX_13TH_MONTH, PRES_NONTAX_DE_MINIMIS, PRES_NONTAX_SSS_ETS, PRES_NONTAX_SALARIES, TOTAL_NONTAX_COMP_INCOME, PRES_TAXABLE_BASIC_SALARY, PRES_TAXABLE_13TH_MONTH, PRES_TAXABLE_SALARIES, TOTAL_TAXABLE_COMP_INCOME, EXMPN_CODE, EXPMN_AMT, PREMIUM_PAID, NET_TAXABLE_COMP_INCOME, TAX_DUE, PRES_TAX_WTHLD, AMT_WTHLD_DEC, OVER_WTHLD, ACTUAL_AMT_WTHLD, SUBS_FILING, TIN, CTC, POI, DOI) " +
                                "VALUES(@ID, @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @GROSS_COMP_INCOME, @PRES_NONTAX_13TH_MONTH, @PRES_NONTAX_DE_MINIMIS, @PRES_NONTAX_SSS_ETS, @PRES_NONTAX_SALARIES, @TOTAL_NONTAX_COMP_INCOME, @PRES_TAXABLE_BASIC_SALARY, @PRES_TAXABLE_13TH_MONTH, @PRES_TAXABLE_SALARIES, @TOTAL_TAXABLE_COMP_INCOME, @EXMPN_CODE, @EXPMN_AMT, @PREMIUM_PAID, @NET_TAXABLE_COMP_INCOME, @TAX_DUE, @PRES_TAX_WTHLD, @AMT_WTHLD_DEC, @OVER_WTHLD, @ACTUAL_AMT_WTHLD, @SUBS_FILING, @TIN, @CTC, @POI, @DOI)";
                        }
                        else if ((chkbxCTC.Checked == true) && (chkbxDate.Checked == true))
                        {
                            cmd.CommandText =
                                "Insert INTO ACTB (ID, LAST_NAME, FIRST_NAME, MIDDLE_NAME, GROSS_COMP_INCOME, PRES_NONTAX_13TH_MONTH, PRES_NONTAX_DE_MINIMIS, PRES_NONTAX_SSS_ETS, PRES_NONTAX_SALARIES, TOTAL_NONTAX_COMP_INCOME, PRES_TAXABLE_BASIC_SALARY, PRES_TAXABLE_13TH_MONTH, PRES_TAXABLE_SALARIES, TOTAL_TAXABLE_COMP_INCOME, EXMPN_CODE, EXPMN_AMT, PREMIUM_PAID, NET_TAXABLE_COMP_INCOME, TAX_DUE, PRES_TAX_WTHLD, AMT_WTHLD_DEC, OVER_WTHLD, ACTUAL_AMT_WTHLD, SUBS_FILING, TIN, STARTED, ENDED, CTC, POI, DOI, AMT) " +
                                "VALUES(@ID, @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @GROSS_COMP_INCOME, @PRES_NONTAX_13TH_MONTH, @PRES_NONTAX_DE_MINIMIS, @PRES_NONTAX_SSS_ETS, @PRES_NONTAX_SALARIES, @TOTAL_NONTAX_COMP_INCOME, @PRES_TAXABLE_BASIC_SALARY, @PRES_TAXABLE_13TH_MONTH, @PRES_TAXABLE_SALARIES, @TOTAL_TAXABLE_COMP_INCOME, @EXMPN_CODE, @EXPMN_AMT, @PREMIUM_PAID, @NET_TAXABLE_COMP_INCOME, @TAX_DUE, @PRES_TAX_WTHLD, @AMT_WTHLD_DEC, @OVER_WTHLD, @ACTUAL_AMT_WTHLD, @SUBS_FILING, @TIN, @STARTED, @ENDED, @CTC, @POI, @DOI, @AMT)";
                        }

                        MaxSchema = VirtualTable.Rows.Count;
                        Console.WriteLine(MaxSchema);
                        for (int s = 0; s < VirtualTable.Rows.Count; s++)
                        {
                            cmd.Parameters.Clear();
                            cmd.Parameters.AddWithValue("@ID", VirtualTable.Rows[s][0]);
                            cmd.Parameters.AddWithValue("@LAST_NAME", VirtualTable.Rows[s][1]);
                            cmd.Parameters.AddWithValue("@FIRST_NAME", VirtualTable.Rows[s][2]);
                            cmd.Parameters.AddWithValue("@MIDDLE_NAME", VirtualTable.Rows[s][3]);
                            cmd.Parameters.AddWithValue("@GROSS_COMP_INCOME", (VirtualTable.Rows[s][4] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][4])));
                            cmd.Parameters.AddWithValue("@PRES_NONTAX_13TH_MONTH", (VirtualTable.Rows[s][5] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][5])));
                            cmd.Parameters.AddWithValue("@PRES_NONTAX_DE_MINIMIS", (VirtualTable.Rows[s][6] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][6])));
                            cmd.Parameters.AddWithValue("@PRES_NONTAX_SSS_ETS", (VirtualTable.Rows[s][7] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][7])));
                            cmd.Parameters.AddWithValue("@PRES_NONTAX_SALARIES", (VirtualTable.Rows[s][8] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][8])));
                            cmd.Parameters.AddWithValue("@TOTAL_NONTAX_COMP_INCOME", (VirtualTable.Rows[s][9] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][9])));
                            cmd.Parameters.AddWithValue("@PRES_TAXABLE_BASIC_SALARY", (VirtualTable.Rows[s][10] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][10])));
                            cmd.Parameters.AddWithValue("@PRES_TAXABLE_13TH_MONTH", (VirtualTable.Rows[s][11] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][11])));
                            cmd.Parameters.AddWithValue("@PRES_TAXABLE_SALARIES", (VirtualTable.Rows[s][12] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][12])));
                            cmd.Parameters.AddWithValue("@TOTAL_TAXABLE_COMP_INCOME", (VirtualTable.Rows[s][13] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][13])));
                            cmd.Parameters.AddWithValue("@EXMPN_CODE", (VirtualTable.Rows[s][14] == DBNull.Value ? " " : Convert.ToString(VirtualTable.Rows[s][14])));
                            cmd.Parameters.AddWithValue("@EXPMN_AMT", (VirtualTable.Rows[s][15] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][15])));
                            cmd.Parameters.AddWithValue("@PREMIUM_PAID", (VirtualTable.Rows[s][16] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][16])));
                            cmd.Parameters.AddWithValue("@NET_TAXABLE_COMP_INCOME", (VirtualTable.Rows[s][17] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][17])));
                            cmd.Parameters.AddWithValue("@TAX_DUE", (VirtualTable.Rows[s][18] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][18])));
                            cmd.Parameters.AddWithValue("@PRES_TAX_WTHLD", (VirtualTable.Rows[s][19] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][19])));
                            cmd.Parameters.AddWithValue("@AMT_WTHLD_DEC", (VirtualTable.Rows[s][20] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][20])));
                            cmd.Parameters.AddWithValue("@OVER_WTHLD", (VirtualTable.Rows[s][21] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][21])));
                            cmd.Parameters.AddWithValue("@ACTUAL_AMT_WTHLD", (VirtualTable.Rows[s][22] == DBNull.Value ? 0 : Convert.ToDouble(VirtualTable.Rows[s][22])));
                            cmd.Parameters.AddWithValue("@Subs_Filing", (VirtualTable.Rows[s][23] == DBNull.Value ? 0 : Convert.ToInt32(VirtualTable.Rows[s][23])));
                            cmd.Parameters.AddWithValue("@TIN", VirtualTable.Rows[s][24]);
                            if ((chkbxDate.Checked == true) && (!chkbxCTC.Checked == true))
                            {
                                cmd.Parameters.AddWithValue("@CTC", VirtualTable.Rows[s][27]);
                                cmd.Parameters.AddWithValue("@POI", VirtualTable.Rows[s][28]);
                                cmd.Parameters.AddWithValue("@DOI", VirtualTable.Rows[s][29]);
                            }
                            if ((chkbxDate.Checked == true) && (chkbxCTC.Checked == true))
                            {
                                cmd.Parameters.AddWithValue("@STARTED", VirtualTable.Rows[s][25]);
                                cmd.Parameters.AddWithValue("@ENDED", VirtualTable.Rows[s][26]);
                                cmd.Parameters.AddWithValue("@AMT", VirtualTable.Rows[s][30]);
                            }
                            cmd.ExecuteNonQuery();
                            Transfer_Worker.ReportProgress(s);
                        }
                    }
                }
                Transfer_Worker.ReportProgress(MaxSchema);
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Import error: " + ex);
            }
        }
        private async void Transfer_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Red;
                lblSuccess.Text = "Error: The thread aborted";
            }
            else
            {
                //pbar.Value = 100;
                pbar.BackColor = Color.Lime;
                lblItem.Text = "";
                MessageBox.Show("Import completed without incident!");
            }
            Refresh_Main();
            lblSuccess.Text = "Successfully imported the records!";
            this.Enabled = true;
            await PutTaskDelay();
            pbar.Value = 0;
            DGVExcel.DataSource = null;
        }
        private void Transfer_Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbar.Value = (e.ProgressPercentage * 100) / MaxSchema;
        }

        private void MatchWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (sheetName == null || sheetName == "Enter a sheet name" || MatchString == null)
            {
                MessageBox.Show("Either you failed to write the sheet name in the box above or your file selection was invalid!");
                return;
            }
            using (OleDbConnection conn = new OleDbConnection(Con))
            {
                try
                {
                    conn.Open();
                    string sql = @"insert * into ACTB from [Excel 12.0;HDR=YES;DATABASE=" + MatchString + "].[" + sheetName +
                                 "$]s;";
                    string sqls = @"INSERT INTO ACXL SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" + MatchString + "].[" + sheetName + "$];";
                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = sqls;
                    cmd.ExecuteNonQuery();
                    string updater =
                        @"UPDATE ACTB " + @"INNER JOIN ACXL on ACTB.ID = ACXL.ID " +
                        @"SET ACTB.GROSS_COMP_INCOME = ACTB.GROSS_COMP_INCOME + ACXL.GROSS_COMP_INCOME, " +
                        @"ACTB.PRES_NONTAX_13TH_MONTH = ACTB.PRES_NONTAX_13TH_MONTH + ACXL.PRES_NONTAX_13TH_MONTH, " +
                        @"ACTB.PRES_NONTAX_DE_MINIMIS = ACTB.PRES_NONTAX_DE_MINIMIS + ACXL.PRES_NONTAX_DE_MINIMIS, " +
                        @"ACTB.PRES_NONTAX_SSS_ETS = ACTB.PRES_NONTAX_SSS_ETS + ACXL.PRES_NONTAX_SSS_ETS, " +
                        @"ACTB.PRES_NONTAX_SALARIES = ACTB.PRES_NONTAX_SALARIES + ACXL.PRES_NONTAX_SALARIES, " +
                        @"ACTB.TOTAL_NONTAX_COMP_INCOME = ACTB.TOTAL_NONTAX_COMP_INCOME + ACXL.TOTAL_NONTAX_COMP_INCOME, " +
                        @"ACTB.PRES_TAXABLE_BASIC_SALARY = ACTB.PRES_TAXABLE_BASIC_SALARY + ACXL.PRES_TAXABLE_BASIC_SALARY, " +
                        @"ACTB.PRES_TAXABLE_13TH_MONTH = ACTB.PRES_TAXABLE_13TH_MONTH + ACXL.PRES_TAXABLE_13TH_MONTH, " +
                        @"ACTB.PRES_TAXABLE_SALARIES = ACTB.PRES_TAXABLE_SALARIES + ACXL.PRES_TAXABLE_SALARIES, " +
                        @"ACTB.TOTAL_TAXABLE_COMP_INCOME = ACTB.TOTAL_TAXABLE_COMP_INCOME + ACXL.TOTAL_TAXABLE_COMP_INCOME, " +
                        @"ACTB.TAX_DUE = ACTB.TAX_DUE + ACXL.TAX_DUE ";
                    cmd.CommandText = updater;
                    cmd.ExecuteNonQuery();
                    string deleter = @"DELETE from ACXL";
                    cmd.CommandText = deleter;
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Match error: " + ex);
                }
            }
        }
        private async void MatchWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Red;
                lblSuccess.Text = "Error: The thread aborted";
            }
            else
            {
                //pbar.Value = 100;
                pbar.BackColor = Color.Lime;
                lblItem.Text = "";
                MessageBox.Show("Match completed without incident!");
            }
            Refresh_Main();
            lblSuccess.Text = "Successfully imported the records!";
            this.Enabled = true;
            await PutTaskDelay();
            pbar.Value = 0;
        }

        private void Micro_Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            FileInfo file = new FileInfo(Foundfile);
            int current = new int();
            current = 0;
            if (chkbxFull.Checked == true)
            {
                GetDataTableFromDGV(DGVmain);
                Console.WriteLine("." + VirtualTable.Rows.Count);
                Micro_Void(file, current);
                Micro_Worker.ReportProgress(VirtualTable.Rows.Count);
            }
            else
            {
                Console.WriteLine("." + MW.MaxRows);
                foreach (DataGridViewRow selectedRow in DGVmain.SelectedRows)
                {
                    current++;
                    string CTC = "";
                    string POI = "";
                    string DOI = "";
                    string SUBS = "";
                    string Started = "";
                    string Ended = "";
                    string AMT = "";
                    List<string> StringRows = new List<string>();
                    List<string> CTCRows = new List<string>();
                    List<string> DateRows = new List<string>();
                    List<Double> DoubleRows = new List<Double>();
                    List<string> ID = new List<string>();
                    lblPathings = Path.ChangeExtension(Foundfile, null);

                    ID.Add(Convert.ToString(selectedRow.Cells[0].Value));
                    for (int i = 1; i < 4; i++)
                    {
                        String StringValue = selectedRow.Cells[i].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[i].Value);
                        if (StringValue == null || StringValue == "") StringValue = " ";
                        StringRows.Add(StringValue); //0-2
                    }
                    String StringValues = "";
                    StringRows.Add(StringValues = selectedRow.Cells[14].Value == DBNull.Value ? " " : Convert.ToString(selectedRow.Cells[14].Value)); //3 Code
                    for (int i = 23; i < 25; i++)
                    {
                        String StringValue = selectedRow.Cells[i].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[i].Value);
                        StringRows.Add(StringValue);
                    }
                    for (int i = 4; i < 14; i++)
                    {
                        Double DoubleValue = selectedRow.Cells[i].Value == DBNull.Value ? 0 : Convert.ToDouble(selectedRow.Cells[i].Value);
                        DoubleRows.Add(DoubleValue);
                    }
                    for (int i = 15; i < 23; i++)
                    {
                        Double DoubleValue = selectedRow.Cells[i].Value == DBNull.Value ? 0 : Convert.ToDouble(selectedRow.Cells[i].Value);
                        DoubleRows.Add(DoubleValue);
                    }
                    if (chkbxCTC.Checked == true)
                    {
                        for (int i = 27; i < 30; i++)
                        {
                            String StringValue = "";
                            CTCRows.Add(StringValue = selectedRow.Cells[i].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[i].Value));
                        }
                    }
                    if (chkbxDate.Checked == true)
                    {
                        String DateValue = "";
                        DateRows.Add(DateValue = selectedRow.Cells[23].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[23].Value));
                        DateRows.Add(DateValue = selectedRow.Cells[25].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[25].Value));
                        DateRows.Add(DateValue = selectedRow.Cells[26].Value == DBNull.Value ? "" : Convert.ToString(selectedRow.Cells[26].Value));
                    }
                    string PIN = StringRows.ElementAt(5);
                    string LastName = StringRows.ElementAt(0);
                    string FirstName = StringRows.ElementAt(1);
                    string MiddleName = StringRows.ElementAt(2);
                    string Expn = StringRows.ElementAt(3);
                    if (chkbxCTC.Checked)
                    {
                        CTC = CTCRows.ElementAt(0);
                        POI = CTCRows.ElementAt(1);
                        DOI = CTCRows.ElementAt(2);
                        AMT = CTCRows.ElementAt(3);
                    }
                    if (chkbxDate.Checked)
                    {
                        SUBS = DateRows.ElementAt(0);
                        Started = DateRows.ElementAt(1);
                        Ended = DateRows.ElementAt(2);
                    }
                    if (PIN.Length == 8) PIN = PIN + "0";
                    if (PIN.Length == 7) PIN = PIN + "00";
                    if (PIN.Length == 6) PIN = PIN + "000";
                    if (PIN.Length == 5) PIN = PIN + "0000";
                    if (PIN.Length == 4) PIN = PIN + "00000";
                    if (PIN.Length == 3) PIN = PIN + "000000";
                    if (PIN.Length == 2) PIN = PIN + "0000000";
                    if (PIN.Length == 1) PIN = PIN + "00000000";
                    Tin = PIN;
                    Console.WriteLine(Tin);
                    Persona = Convert.ToString(selectedRow.Cells[1].Value);
                    for (int i = 0; i < Persona.Length; i++)
                    {
                        Persona = Persona.Split(' ')[0];
                    }
                    DateOverride = txtDOI.Text;
                    DateOverride = DateOverride.Replace('/', '.');
                    DateOverride = DateOverride.Replace('-', '.');
                    DateOverride = DateOverride.Replace(' ', '.');
                    PersonaID = Convert.ToString(selectedRow.Cells[0].Value);
                    generateID();
                    using (ExcelPackage p = new ExcelPackage(file))
                    {
                        ExcelWorkbook WBook = p.Workbook;
                        if (WBook != null)
                        {
                            if (WBook.Worksheets.Count > 0)
                            {
                                OfficeOpenXml.ExcelWorksheet WS = WBook.Worksheets.First();
                                //
                                string input = Tin;
                                string digit1 = "0";
                                string digit2 = "0";
                                string digit3 = "0";
                                string digit4 = "0";
                                if ((input.Length < 3 || input == "0" || input == "" || input == " "))
                                {
                                    WS.Cells[11, 9].Value = 000;
                                    WS.Cells[11, 12].Value = 000;
                                    WS.Cells[11, 15].Value = 000;
                                    WS.Cells[11, 18].Value = 000;
                                }
                                else
                                {
                                    StringBuilder sb = new StringBuilder();
                                    StringBuilder partBuilder = new StringBuilder();
                                    int partsSplitted = 0;
                                    for (int i = 1; i <= input.Length; i++)
                                    {
                                        partBuilder.Append(input[i - 1]);
                                        if (i % 3 == 0 && partsSplitted <= 3)
                                        {
                                            sb.Append(' ');
                                            sb.Append(partBuilder.ToString());
                                            partBuilder = new StringBuilder();
                                            partsSplitted++;
                                        }
                                    }
                                    if (input.Length < 10)
                                    {
                                        partBuilder.Append(" 000");
                                    }
                                    sb.Append(partBuilder.ToString());
                                    string formatted = sb.ToString().TrimStart();
                                    string[] formatCollection = formatted.Split(' ');
                                    digit1 = formatCollection[0];
                                    digit2 = formatCollection[1];
                                    digit3 = formatCollection[2];
                                    digit4 = formatCollection[3];
                                }
                                //Names
                                if (FirstName == null || String.IsNullOrWhiteSpace(FirstName.ToString()))
                                {
                                    WS.Cells[14, 2].Value = LastName;
                                }
                                else if (chkBxMiddle.Checked)
                                {
                                    WS.Cells[14, 2].Value = LastName + ",  " + FirstName + " " + MiddleName;
                                }
                                else
                                {
                                    WS.Cells[14, 2].Value = LastName + ",  " + FirstName + " ";
                                }
                                //Blanker(WS.Name);
                                //============================== Year & Period
                                if (chkbxDate.Checked == true)
                                {
                                    WS.Cells[8, 8].Value = txtDOI.Text;
                                    WS.Cells[8, 29].Value = Started;
                                    WS.Cells[8, 34].Value = Ended;
                                }
                                else
                                {
                                    string strong = txtDOI.Text;
                                    string[] col;
                                    strong = strong.Replace('/', '.');
                                    strong = strong.Replace('-', '.');
                                    strong = strong.Replace(' ', '.');
                                    col = strong.Split('.');
                                    int Indexes = col.Count();
                                    if (Indexes == 0 || Indexes > 3)
                                    {
                                        MessageBox.Show("Invalid date format! Correct format: 7.7.2077");
                                        return;
                                    }
                                    WS.Cells[8, 8].Value = col.Last();
                                    WS.Cells[8, 29].Value = txtFrom.Text;
                                    WS.Cells[8, 34].Value = txtTo.Text;
                                }
                                WS.Cells[11, 9].Value = digit1;
                                WS.Cells[11, 12].Value = digit2;
                                WS.Cells[11, 15].Value = digit3;
                                WS.Cells[11, 18].Value = digit4;
                                //WS.Cells[45, 9].Value = digit1;
                                //WS.Cells[45, 12].Value = digit2;
                                //WS.Cells[45, 15].Value = digit3;
                                //WS.Cells[45, 18].Value = digit4;
                                char InitChar = StringRows.ElementAt(3)[0];
                                WS.Cells[29, 5].Value = "";
                                WS.Cells[29, 11].Value = "";
                                if (Expn.StartsWith("S") || Expn.StartsWith("s"))
                                    WS.Cells[29, 5].Value = "X";
                                if (Expn.StartsWith("M") || Expn.StartsWith("m"))
                                    WS.Cells[29, 11].Value = "X";
                                WS.Cells[64, 12].Value = DoubleRows.ElementAt(0);
                                WS.Cells[64, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[66, 12].Value = DoubleRows.ElementAt(5);
                                WS.Cells[68, 12].Value = DoubleRows.ElementAt(9);
                                WS.Cells[66, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[68, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[70, 12].Value = "";
                                WS.Cells[72, 12].Value = DoubleRows.ElementAt(9);
                                WS.Cells[74, 12].Value = DoubleRows.ElementAt(10);
                                WS.Cells[72, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[74, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[76, 12].Value = "";
                                WS.Cells[78, 12].Value = DoubleRows.ElementAt(12);
                                WS.Cells[80, 12].Value = DoubleRows.ElementAt(13);
                                WS.Cells[78, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[80, 12].Style.Numberformat.Format = "#,##0.00";
                                //========================== Last 3
                                WS.Cells[82, 12].Value = DoubleRows.ElementAt(14);
                                WS.Cells[82, 12].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[84, 12].Value = "";
                                WS.Cells[86, 12].Value = DoubleRows.ElementAt(17);
                                WS.Cells[86, 12].Style.Numberformat.Format = "#,##0.00";
                                //========================== Form
                                if (chkbxCTC.Checked)
                                {
                                    WS.Cells[97, 5].Value = CTC;
                                    WS.Cells[97, 15].Value = POI;
                                    WS.Cells[97, 24].Value = POI;
                                    WS.Cells[97, 33].Value = AMT;
                                }
                                //========================== Right-Seid
                                WS.Cells[32, 31].Value = DoubleRows.ElementAt(2);
                                WS.Cells[35, 31].Value = DoubleRows.ElementAt(3);
                                WS.Cells[38, 31].Value = DoubleRows.ElementAt(5);
                                WS.Cells[41, 31].Value = DoubleRows.ElementAt(5);
                                WS.Cells[46, 31].Value = DoubleRows.ElementAt(9);
                                WS.Cells[86, 31].Value = DoubleRows.ElementAt(9);
                                WS.Cells[32, 31].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[35, 31].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[38, 31].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[41, 31].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[46, 31].Style.Numberformat.Format = "#,##0.00";
                                WS.Cells[86, 31].Style.Numberformat.Format = "#,##0.00";
                            }
                        }
                        if (ChkbxGen.CheckState == CheckState.Checked)
                        {
                            string recarm = Path.ChangeExtension(finalformat, null) + ".xlsx";
                            FileInfo formatFinal = new FileInfo(recarm);
                            p.SaveAs(formatFinal);
                        }
                        else { p.SaveAs(file); }
                    }
                    try
                    {
                        var workbook = ExcelFile.Load(Foundfile);
                        foreach (var worksheet in workbook.Worksheets)
                        {
                            var printOptions = worksheet.PrintOptions;
                            printOptions.LeftMargin = .4;
                            printOptions.RightMargin = 0;
                            printOptions.TopMargin = 0;
                            printOptions.BottomMargin = 0;
                            //printOptions.AutomaticPageBreakScalingFactor = 120;
                            printOptions.FitWorksheetWidthToPages = 1;
                        }
                        var saveOptions = new PdfSaveOptions();
                        saveOptions.SelectionType = SelectionType.EntireFile;
                        workbook.Save(finalformat, saveOptions);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("PDF Conversion failed: " + ex.Message);
                    } //PDF
                    Micro_Worker.ReportProgress(current);
                }
                Micro_Worker.ReportProgress(MW.MaxRows);
            }

        }
        private void Micro_Worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            pbar.Value = (e.ProgressPercentage * 100) / MW.MaxRows;
        }
        private async void Micro_Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Red;
                lblSuccess.Text = "Error: The thread aborted";
            }
            else
            {
                pbar.Value = 100;
                pbar.BackColor = Color.Lime;
                lblSuccess.Text = "Process completed";
                lblItem.Text = "";
                MessageBox.Show("Macro completed");
            }
            panel1.Enabled = true;
            panel2.Enabled = true;
            panel6.Enabled = true;
            this.Enabled = true;
            await PutTaskDelay();
            pbar.Value = 0;
        }
        #endregion

        private void FindFolderPath(object sender, EventArgs e)
        {
            string FileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            try
            {
                Process.Start(FileName);
            }
            catch (Win32Exception win32Exception)
            {
                //The system cannot find the file specified...
                MessageBox.Show(win32Exception.Message);
            }
        }
        
    }
}
