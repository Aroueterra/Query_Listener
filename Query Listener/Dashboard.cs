using System;
using System.CodeDom;
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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using SD = System.Data;
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
        public static string ExcelOFD = "";
        public static string PersonID = "";
        public static string sheetName = "";
        public static string TIN = "";
        SD.DataTable tbContainer = new SD.DataTable();
        OleDbConnection OleDbcon;
        private OleDbConnection Globalcon = new OleDbConnection();
        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'aCDBDataSet1.ACTB' table. You can move, or remove it, as needed.
            this.aCTBTableAdapter1.Fill(this.aCDBDataSet1.ACTB);
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
            else if (cmbOmnisearch.Text == "First Name")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where FirstName like '" + txtSearchID.Text + "%'", connection);
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
            else if (cmbOmnisearch.Text == "Tax as Adjusted")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where TotalTax like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "Last Name")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where LastName like '" + txtSearchID.Text + "%'", connection);
                    System.Data.DataTable dt = new System.Data.DataTable();
                    sda.Fill(dt);
                    DGVmain.DataSource = dt;
                    connection.Close();
                }
            }
            else if (cmbOmnisearch.Text == "Gross Income")
            {
                using (OleDbConnection connection = new OleDbConnection(Con))
                {
                    connection.Open();
                    OleDbDataAdapter sda = new OleDbDataAdapter("SELECT * FROM ACTB where GrossIncome like '" + txtSearchID.Text + "%'", connection);
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
                    //===========================
                    /*
                    foreach (SD.DataTable table in ds.Tables)
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            foreach (var item in row.ItemArray) // Loop over the items.
                            {
                                Console.WriteLine(item); // Invokes ToString abstract method.
                            }
                        }
                        Console.WriteLine("=======");
                        foreach (DataColumn column in table.Columns)
                        {                           
                            var  collu = column.ToString();
                            Console.WriteLine(collu);
                            //Console.WriteLine(item);                        
                            // read column and item
                        }
                    }*/
                    //==========================
                    //OleDbCommand command = new OleDbCommand();
                    //command.Connection = OleDbcon;
                    //command.CommandType = CommandType.Text;
                    InsertRowsIntoTempTable(tbContainer);
                    /*
                    foreach (SD.DataTable table in ds.Tables)
                    {
                        Console.WriteLine(table.TableName);
                        String columnsCommandText = "(";
                        foreach (DataColumn column in table.Columns)
                        {
                            String columnName = column.ColumnName;
                            String dataTypeName = column.DataType.Name;
                            columnsCommandText += "[" + columnName + "] ,";
                        }
                        columnsCommandText = columnsCommandText.Remove(columnsCommandText.Length - 1);
                        columnsCommandText += ")";

                        command.CommandText = "Insert into ACXL " + table.TableName + columnsCommandText;

                        command.ExecuteNonQuery();
                    }
                    //This loop fills the database with all information
                    foreach (SD.DataTable table in ds.Tables)
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            String commandText = "INSERT INTO "+table+" VALUES (";
                            foreach (var item in row.ItemArray)
                            {
                                commandText += "'" + item.ToString() + "',";
                            }
                            commandText = commandText.Remove(commandText.Length - 1);
                            commandText += ")";
                            command.CommandText = commandText;
                            command.ExecuteNonQuery();
                            
                        }
                    }*/
                    //Restore();
                    //Restormy();
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

        void InsertRowIntoTempTable(string FirstName, string LastName, int PersonIDs)
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

        void InsertRowsIntoTempTable(SD.DataTable tbContainer)
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

        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            TIN = txtTIN.Text;
            //This could definitely do with a list collection revamp in the future
            if (txtFirst.Text == "" || txtLast.Text == "" || txtGross.Text == "" ||
                txtLessTNT.Text == "" || txtTCI.Text == "" || txtADDTI.Text == "" || txtGTI.Text == "" ||
                txtLessTE.Text == "" || txtLessPPH.Text == "" || txtLessNTI.Text == "" || txtTD.Text == "" ||
                txtTWCE.Text == "" || txtTWPE.Text == "" || txtTATW.Text == "" ||
                txtFirst.Text == " " || txtLast.Text == " " || txtGross.Text == " " ||
                txtLessTNT.Text == " " || txtTCI.Text == " " || txtADDTI.Text == " " || txtGTI.Text == " " ||
                txtLessTE.Text == " " || txtLessPPH.Text == " " || txtLessNTI.Text == " " || txtTD.Text == " " ||
                txtTWCE.Text == " " || txtTWPE.Text == " " || txtTATW.Text == " " || txtID.Text == "" ||
                txtID.Text == " " || txtTIN.Text == "" || txtTIN.Text == " "
            )


            { 
                MessageBox.Show("Cannot enter null values!");
                return;
            }
            else if (TIN.Length < 12)
            {

                MessageBox.Show("Taxpayer identification number must be at least 12 digits!");
                return;
            }
            else { 
            OleDbConnection Dbcon = new OleDbConnection(Con);
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText =
                    "Insert INTO ACTB (ID, FirstName, LastName, GrossIncome, LessTNT, TaxableIncomeCE, TaxableIncomePE, GrossTaxableIncome, LessTE, LessPPH, NetTax, TaxDue, HeldTaxCE, HeldTaxPE, TotalTax, TIN) " +
                    "VALUES(@ID, @First, @Last, @Gross, @LessTNT, @TCI, @ADDTI, @GTI, @LessTE, @LessPPH, @LessNTI, @TD, @TWCE, @TWPE, @TATW, @TIN)";
                cmd.Parameters.AddWithValue("@ID", txtID.Text);
                cmd.Parameters.AddWithValue("@First", txtFirst.Text);
                cmd.Parameters.AddWithValue("@Last", txtLast.Text);
                cmd.Parameters.AddWithValue("@Gross", Convert.ToDouble(txtGross.Text));
                cmd.Parameters.AddWithValue("@LessTNT", Convert.ToDouble(txtLessTNT.Text));
                cmd.Parameters.AddWithValue("@TCI", Convert.ToDouble(txtTCI.Text));
                cmd.Parameters.AddWithValue("@ADDTI", Convert.ToDouble(txtADDTI.Text));
                cmd.Parameters.AddWithValue("@GTI", Convert.ToDouble(txtGTI.Text));
                cmd.Parameters.AddWithValue("@LessTE", Convert.ToDouble(txtLessTE.Text));
                cmd.Parameters.AddWithValue("@LessPPH", Convert.ToDouble(txtLessPPH.Text));
                cmd.Parameters.AddWithValue("@LessNTI", Convert.ToDouble(txtLessNTI.Text));
                cmd.Parameters.AddWithValue("@TD", Convert.ToDouble(txtTD.Text));
                cmd.Parameters.AddWithValue("@TWCE", Convert.ToDouble(txtTWCE.Text));
                cmd.Parameters.AddWithValue("@TWPE", Convert.ToDouble(txtTWPE.Text));
                cmd.Parameters.AddWithValue("@TATW", Convert.ToDouble(txtTATW.Text));
                cmd.Parameters.AddWithValue("@TIN", txtTIN.Text);
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
                    string query = "SELECT * From ACTB";
                    try
                    {
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
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
            }
            }
        }

        private void txtFirst_MouseClick(object sender, MouseEventArgs e)
        {
            txtFirst.SelectAll();
        }

        private void txtLast_MouseClick(object sender, MouseEventArgs e)
        {
            txtLast.SelectAll();
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

        private void Restore()
        {
            string query = "SELECT * From ACTB";
            using (OleDbConnection conn = new OleDbConnection(Con))
            {
                conn.Open();
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    //DGVmain.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
                    // set it to false if not needed
                    //DGVmain.RowHeadersVisible = false;
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    DGVmain.DataSource = ds.Tables[0];
                }
            }
        }
        private void Restormy()
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
                string drag = Convert.ToString(selectedRow.Cells[0].Value);
                txtID.Text = drag;
            }
            if (DGVmain.SelectedCells.Count > 0)
            {
                int selectedrowindex = DGVmain.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = DGVmain.Rows[selectedrowindex];
                string drag = Convert.ToString(selectedRow.Cells[1].Value);
                txtFirst.Text = drag;
            }
            if (DGVmain.SelectedCells.Count > 0)
            {
                int selectedrowindex = DGVmain.SelectedCells[0].RowIndex;
                DataGridViewRow selectedRow = DGVmain.Rows[selectedrowindex];
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
                txtLast.Text = drag;
                txtGross.Text = drag2;
                txtLessTNT.Text = drag3;
                txtTCI.Text = drag4;
                txtADDTI.Text = drag5;
                txtGTI.Text = drag6;
                txtLessTE.Text = drag7;
                txtLessPPH.Text = drag8;
                txtLessNTI.Text = drag9;
                txtTD.Text = drag10;
                txtTWCE.Text = drag11;
                txtTWPE.Text = drag12;
                txtTATW.Text = drag13;
                txtTIN.Text = drag14;
                lblSuccess.Text = "Successfully dragged the records!";
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    
                    OleDbCommand cmd = new OleDbCommand("UPDATE ACTB SET FirstName = @FirstName, LastName = @LastName, GrossIncome = @GrossIncome, LessTNT = @LessTNT,  TaxableIncomeCE = @TCI, " +
                                                                        "TaxableIncomePE = @ADDTI, GrossTaxableIncome = @GTI, LessTE = @LessTE, LessPPH = @LessPPH, NetTax = @LessNTI," +
                                                                        "TaxDue = @TD, HeldTaxCE = @TWCE, HeldTaxPE = @TWPE, TotalTax = @TATW, TIN = @TIN WHERE ID = @ID", conn);
                    cmd.Connection = conn;
                    conn.Open();
                    cmd.Parameters.AddWithValue("@FirstName", txtFirst.Text);
                    cmd.Parameters.AddWithValue("@LastName", txtLast.Text);
                    cmd.Parameters.AddWithValue("@GrossIncome", Convert.ToDouble(txtGross.Text));
                    cmd.Parameters.AddWithValue("@LessTNT", Convert.ToDouble(txtLessTNT.Text));
                    cmd.Parameters.AddWithValue("@TCI", Convert.ToDouble(txtTCI.Text));
                    cmd.Parameters.AddWithValue("@ADDTI", Convert.ToDouble(txtADDTI.Text));
                    cmd.Parameters.AddWithValue("@GTI", Convert.ToDouble(txtGTI.Text));
                    cmd.Parameters.AddWithValue("@LessTE", Convert.ToDouble(txtLessTE.Text));
                    cmd.Parameters.AddWithValue("@LessPPH", Convert.ToDouble(txtLessPPH.Text));
                    cmd.Parameters.AddWithValue("@LessNTI", Convert.ToDouble(txtLessNTI.Text));
                    cmd.Parameters.AddWithValue("@TD", Convert.ToDouble(txtTD.Text));
                    cmd.Parameters.AddWithValue("@TWCE", Convert.ToDouble(txtTWCE.Text));
                    cmd.Parameters.AddWithValue("@TWPE", Convert.ToDouble(txtTWPE.Text));
                    cmd.Parameters.AddWithValue("@TATW", Convert.ToDouble(txtTATW.Text));
                    cmd.Parameters.AddWithValue("@TIN", txtTIN.Text);
                    cmd.Parameters.AddWithValue("@ID", txtID.Text);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                lblSuccess.Text = "Successfully updated the records!";
                Restore();
            }
        }
        
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            ExcelOFD = openFileDialog.FileName;
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                try
                {
                    OleDbcon =
                        new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + openFileDialog.FileName +
                                            ";Extended Properties='Excel 12.0 Xml; HDR = Yes; IMEX = 1'");
                    OleDbcon.Open();
                    System.Data.DataTable dt = OleDbcon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    OleDbcon.Close();
                    CMBsheets.Items.Clear();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        String sheetName = dt.Rows[i]["TABLE_NAME"].ToString();
                        sheetName = sheetName.Substring(0, sheetName.Length - 1);
                        CMBsheets.Items.Add(sheetName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Write Excel: " + ex.Message);
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    lblSuccess.Text = "Successfully opened the file!";
                }
                MessageBox.Show("File accessed. Browse the file in table view by selecting a sheet from the combo box below.");
            }
        }

        private void CMBsheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            OleDbDataAdapter oledbDa = new OleDbDataAdapter("Select * from [" + CMBsheets.Text + "$]", OleDbcon);
            System.Data.DataTable dt = new System.Data.DataTable();
            oledbDa.Fill(dt);
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

        private void btnreturn_Click(object sender, EventArgs e)
        {
            Form Help = new Help();
            Help.ShowDialog();
        }
        
        private void btnPrint_Click(object sender, EventArgs e)
        {

            Firstnamae = txtFirst.Text;
            Lastnamae = txtLast.Text;
            Gross = txtGross.Text;
            LessTNT = txtLessTNT.Text;
            TCI = txtTCI.Text;
            ADDTI = txtADDTI.Text;
            GTI = txtGTI.Text;
            LessTE = txtLessTE.Text;
            LessPPH = txtLessPPH.Text;
            NetTax = txtLessNTI.Text;
            TD = txtTD.Text;
            HeldTaxCE = txtTWCE.Text;
            HeldTaxPE = txtTWPE.Text;
            TotalTax = txtTATW.Text;
            TIN = txtTIN.Text;
            PersonID = txtID.Text;
            if (TIN.Length < 12)
            {
                MessageBox.Show("Taxpayer identification number must be at least 12 digits long!");
                return;
            }
            Printer = new Printer();
            Printer.ShowDialog();   
    }

        private void btnImport_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(
                "Warning: when importing data into the Access database, ensure that the field columns match Access's fields or the file may become corrupt. Do you still wish to proceed?","Import caution",
                MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (dr == DialogResult.OK)
            {
                try
                {

                    using (OleDbConnection conn = new OleDbConnection(Con))
                    {
                        using (OleDbCommand cmd = new OleDbCommand())
                        {
                            cmd.Connection = conn;
                            conn.Open();
                            cmd.CommandText =
                                    "Insert INTO ACTB (ID, FirstName, LastName, GrossIncome, LessTNT, TaxableIncomeCE, TaxableIncomePE, GrossTaxableIncome, LessTE, LessPPH, NetTax, TaxDue, HeldTaxCE, HeldTaxPE, TotalTax, TIN) " +
                                    "VALUES(@ID, @First, @Last, @Gross, @LessTNT, @TCI, @ADDTI, @GTI, @LessTE, @LessPPH, @LessNTI, @TD, @TWCE, @TWPE, @TATW, @TIN)";
                            for (int s = 0; s < DGVExcel.Rows.Count - 1; s++)
                            {
                                cmd.Parameters.Clear();
                                cmd.Parameters.AddWithValue("@ID", DGVExcel.Rows[s].Cells[0].Value);
                                cmd.Parameters.AddWithValue("@First", DGVExcel.Rows[s].Cells[1].Value);
                                cmd.Parameters.AddWithValue("@Last", DGVExcel.Rows[s].Cells[2].Value);
                                cmd.Parameters.AddWithValue("@Gross", Convert.ToDouble(DGVExcel.Rows[s].Cells[3].Value));
                                cmd.Parameters.AddWithValue("@LessTNT",
                                    Convert.ToDouble(DGVExcel.Rows[s].Cells[4].Value));
                                cmd.Parameters.AddWithValue("@TCI", Convert.ToDouble(DGVExcel.Rows[s].Cells[5].Value));
                                cmd.Parameters.AddWithValue("@ADDTI", Convert.ToDouble(DGVExcel.Rows[s].Cells[6].Value));
                                cmd.Parameters.AddWithValue("@GTI", Convert.ToDouble(DGVExcel.Rows[s].Cells[7].Value));
                                cmd.Parameters.AddWithValue("@LessTE", Convert.ToDouble(DGVExcel.Rows[s].Cells[8].Value));
                                cmd.Parameters.AddWithValue("@LessPPH",
                                    Convert.ToDouble(DGVExcel.Rows[s].Cells[9].Value));
                                cmd.Parameters.AddWithValue("@LessNTI",
                                    Convert.ToDouble(DGVExcel.Rows[s].Cells[10].Value));
                                cmd.Parameters.AddWithValue("@TD", Convert.ToDouble(DGVExcel.Rows[s].Cells[11].Value));
                                cmd.Parameters.AddWithValue("@TWCE", Convert.ToDouble(DGVExcel.Rows[s].Cells[12].Value));
                                cmd.Parameters.AddWithValue("@TWPE", Convert.ToDouble(DGVExcel.Rows[s].Cells[13].Value));
                                cmd.Parameters.AddWithValue("@TATW", Convert.ToDouble(DGVExcel.Rows[s].Cells[14].Value));
                                cmd.Parameters.AddWithValue("@TIN", Convert.ToDouble(DGVExcel.Rows[s].Cells[15].Value));
                                cmd.ExecuteNonQuery();
                            }
                        }

                    }
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("Import error: " + ex);
                }
                finally
                {
                    Restore();
                    lblSuccess.Text = "Successfully imported the records!";
                }

            }
        }

        private Excel.Application Xls;
        private Excel.Workbooks WBs;
        private Excel.Workbook WB;
        private Excel.Worksheet WS;
        private Excel.Sheets SS;
        Excel.Range cellsRange = null;
        long numberOfColumns = 0;
        long numberOfRows = 0;
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
                        XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
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
                            "=IF(K"+i+"<10001,K"+i+"*0.05,IF(K"+i+ "<30001,(K" + i + "-10000)*0.1+500,IF(K" + i + "<70001,(K" + i + "-30000)*0.15+2500,IF(K" + i + "<140001,(K" + i + "-70000)*0.2+8500,IF(K" + i + "<250001,(K" + i + "-140000)*0.25+22500,IF(K" + i + "<500001,(K" + i + "-250000)*0.3+50000,IF(K" + i + ">500000,(K" + i + "-500000)*0.32+125000)))))))";
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
                        MessageBox.Show("Closing Excel: " + ex.Message); }
                    
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
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    try
                    {
                        conn.Open();
                        string sql = @"insert * into ACTB from [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text +
                                     "$]s;";
                        string sqls = @"INSERT INTO ACXL SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text + "$];";
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = sqls;
                        cmd.ExecuteNonQuery();
                        string updater =
                            @"UPDATE ACTB " + @"INNER JOIN ACXL on ACTB.ID = ACXL.ID " +
                            @"SET ACTB.GrossIncome = ACTB.GrossIncome + ACXL.GrossIncome, " +
                            @"ACTB.LessTNT = ACTB.LessTNT + ACXL.LessTNT, " +
                            @"ACTB.TaxableIncomeCE = ACTB.TaxableIncomeCE + ACXL.TaxableIncomeCE, " +
                            @"ACTB.TaxableIncomePE = ACTB.TaxableIncomePE + ACXL.TaxableIncomePE, " +
                            @"ACTB.GrossTaxableIncome = ACTB.GrossTaxableIncome + ACXL.GrossTaxableIncome, " +
                            @"ACTB.LessTE = ACTB.LessTE + ACXL.LessTE, " +
                            @"ACTB.LessPPH = ACTB.LessPPH + ACXL.LessPPH, " +
                            @"ACTB.NetTax = ACTB.NetTax + ACXL.NetTax, " +
                            @"ACTB.TaxDue = ACTB.TaxDue + ACXL.TaxDue, " +
                            @"ACTB.HeldTaxCE = ACTB.HeldTaxCE + ACXL.HeldTaxCE, " +
                            @"ACTB.HeldTaxPE = ACTB.HeldTaxPE + ACXL.HeldTaxPE, " +
                            @"ACTB.TotalTax = ACTB.TotalTax + ACXL.TotalTax ";

                        cmd.CommandText = updater;
                        cmd.ExecuteNonQuery();
                        string deleter = @"DELETE from ACXL";
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
                        Restore();
                        Restormy();
                        lblSuccess.Text = "Successfully imported the records!";
                        MessageBox.Show("Succesfully imported the records");
                    }
                }
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
                    Restore();
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
            if (txtSheetName.Text == null)
            {
                return;
            }
            if (!string.IsNullOrEmpty(openFileDialog.FileName))
            {
                using (OleDbConnection conn = new OleDbConnection(Con))
                {
                    try
                    {
                        conn.Open();
                        string sql = @"insert * into ACTB from [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text +
                                     "$]s;";
                        string sqls = @"INSERT INTO ACXL SELECT * FROM [Excel 12.0;HDR=YES;DATABASE=" + openFileDialog.FileName + "].[" + txtSheetName.Text + "$];";
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        cmd.CommandText = sqls;
                        cmd.ExecuteNonQuery();
                        string updater =
                            @"UPDATE ACTB " + @"INNER JOIN ACXL on ACTB.ID = ACXL.ID " +
                            @"SET ACTB.GrossIncome = ACTB.GrossIncome + ACXL.GrossIncome, " +
                            @"ACTB.LessTNT = ACTB.LessTNT + ACXL.LessTNT, " +
                            @"ACTB.TaxableIncomeCE = ACTB.TaxableIncomeCE + ACXL.TaxableIncomeCE, " +
                            @"ACTB.TotalTax = ACTB.TotalTax + ACXL.TotalTax ";

                        cmd.CommandText = updater;
                        cmd.ExecuteNonQuery();
                        string deleter = @"DELETE from ACXL";
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
                        Restore();
                        Restormy();
                        lblSuccess.Text = "Successfully imported the records!";
                        MessageBox.Show("Succesfully imported the records");
                    }
                }
            }
        }
    }
}
