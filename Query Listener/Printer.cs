using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
namespace Query_Listener
{
    /// <summary>
    /// Software created by August Bryan N. Florese  
    /// Contact: Aroueterra@gmail.com
    /// For: Tax team member Bryan Rucio of Convergys Finance department
    /// Author and developer of the code retains the program property rights
    /// </summary>
    public partial class Printer : Form
    {
        private Excel.Application Xls;
        private Excel.Workbooks WBs;
        private Excel.Workbook WB;
        private Excel.Worksheet WS;
        private Excel.Sheets SS;
        private string Persona;
        private string lblPathing = "";
        object misValue = System.Reflection.Missing.Value;
        private string finalformat;
        private string Middlenamae;
        public Printer()
        {
            InitializeComponent();

        }

        private void Printer_Load(object sender, EventArgs e)
        {
            txtFirst.Text = Dashboard.Firstnamae;
            txtLast.Text = Dashboard.Lastnamae;
            txtGross.Text = Dashboard.Gross;
            txtLessTNT.Text = Dashboard.LessTNT;
            txtTCI.Text = Dashboard.TCI;
            txtADDTI.Text = Dashboard.ADDTI;
            txtGTI.Text = Dashboard.GTI;
            txtLessTE.Text = Dashboard.LessTE;
            txtLessPPH.Text = Dashboard.LessPPH;
            txtLessNTI.Text = Dashboard.NetTax;
            txtTD.Text = Dashboard.TD;
            txtTWCE.Text = Dashboard.HeldTaxCE;
            txtTWPE.Text = Dashboard.HeldTaxPE;
            txtTATW.Text = Dashboard.TotalTax;
            txtPersonnelID.Text = Dashboard.PersonID;
            txtTIN.Text = Dashboard.TIN_Printer;
            txtPeriod1.Text = Dashboard.From;
            txtPeriod2.Text = Dashboard.To;
            Persona = txtFirst.Text;
            txtCTC.Text = Dashboard.CTC;
            txtPOI.Text = Dashboard.POI;
            txtDOI.Text = Dashboard.DOI;
            txtAMT.Text = Dashboard.AMT;
            Middlenamae = Dashboard.Middlenamae;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (Controls.OfType<System.Windows.Forms.TextBox>().Any(t => t.Text == ""))
            {
                MessageBox.Show("One of the text boxes is empty! Make sure they have a value before proceeding.", "Null value detected", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (lblPath.Text == "" || lblPath.Text == " " || lblPath.Text == null)
            {
                return;
            }
            try
            {
                //Initialize the Excel File
                Xls = new Excel.Application();
                WBs = Xls.Workbooks;
                WB = WBs.Open(lblPath.Text, 0, false, 5, "", "", true,
                    XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                if (WB == null)
                { 
                    Xls.Quit();
                    Xls = null;
                    WB = null;
                    return;
                }
                SS = WB.Worksheets;
                WS = SS.get_Item(1);
                //Tin Algorithm
                string input = txtTIN.Text;
                string digit1 = "0";
                string digit2 = "0";
                string digit3 = "0";
                string digit4 = "0";
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
                sb.Append(partBuilder.ToString());
                string formatted = sb.ToString().TrimStart();
                string[] formatCollection = formatted.Split(' ');
                digit1 = formatCollection[0];
                digit2 = formatCollection[1];
                digit3 = formatCollection[2];
                digit4 = formatCollection[3];

                Console.WriteLine(formatted);
                //Names
                WS.Cells[14, 2] = txtLast.Text + ",  " + txtFirst.Text + " " + Middlenamae;
                WS.Cells[48, 2] = txtEmployer.Text;
                //Year & Period
                WS.Cells[8, 8] = txtYear.Text;
                WS.Cells[8, 29] = txtPeriod1.Text;
                WS.Cells[8, 34] = txtPeriod2.Text;
                //Tin
                WS.Cells[11, 9] = digit1;
                WS.Cells[11, 12] = digit2;
                WS.Cells[11, 15] = digit3;
                WS.Cells[11, 18] = digit4;
                WS.Cells[45, 9] = digit1;
                WS.Cells[45, 12] = digit2;
                WS.Cells[45, 15] = digit3;
                WS.Cells[45, 18] = digit4;
                //========================== Summary
                WS.Cells[64, 12] = txtGross.Text;
                WS.Cells[66, 12] = txtLessTNT.Text;
                WS.Cells[68, 12] = txtTCI.Text;
                WS.Cells[70, 12] = txtADDTI.Text;
                WS.Cells[72, 12] = txtGTI.Text;
                WS.Cells[74, 12] = txtLessTE.Text;
                WS.Cells[76, 12] = txtLessPPH.Text;
                WS.Cells[78, 12] = txtLessNTI.Text;
                WS.Cells[80, 12] = txtTD.Text;
                //==========================
                WS.Cells[82, 12] = txtTWCE.Text;
                WS.Cells[84, 12] = txtTWPE.Text;
                WS.Cells[86, 12] = txtTATW.Text;
                //==========================
                WS.Cells[95, 5] = txtCTC.Text;
                WS.Cells[95, 15] = txtPOI.Text;
                WS.Cells[95, 24] = txtDOI.Text;
                WS.Cells[95, 33] = txtAMT.Text;
                WB.Save();
                if (chkbxPDF.Checked == false)
                {
                    DialogResult dr = MessageBox.Show(
                        "Create PDF?", "Creates a PDF in the source file's directory",
                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.OK)
                    {
                        try
                        {
                            //var Unique = string.Format(@"{0}.pdf", Guid.NewGuid());            
                            WB.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, finalformat);

                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show("Error occurred: " + ex, "General error exception");
                        }
                    }
                }
                else
                {
                    try
                    {           
                        WB.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, finalformat);

                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Error occurred: " + ex, "General error exception");
                    }
                }
                MessageBox.Show("Finished Updating File", "Task complete");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Write Excel: " + ex.Message);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                WB.Close();
                Xls.Quit();
                releaseObject(SS);
                releaseObject(WS);
                releaseObject(WBs);
                releaseObject(WB);
                releaseObject(Xls);
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

 
        public void generateID()
        {
            string append = lblPathing.ToString();
            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            var random = new Random();
            var result = new string(
                Enumerable.Repeat(chars, 3)
                          .Select(s => s[random.Next(s.Length)])
                          .ToArray());
            finalformat = append + "-" + Persona + "-" + result + ".pdf";
        }
        private void btnSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
            openFileDialog.ShowDialog();
            lblPath.Text = openFileDialog.FileName;
            lblPathing = Path.ChangeExtension(openFileDialog.FileName, null);
            Console.WriteLine(lblPath);
            Console.WriteLine(lblPathing);
            generateID();
            Console.WriteLine(finalformat);
        }

        void EPPLUS()
        {
            // Taking existing file: 'Sample1.xlsx'. Here 'Sample1.xlsx' is treated as template file
            FileInfo templateFile = new FileInfo(@"Sample1.xlsx");
            // Making a new file 'Sample2.xlsx'
            FileInfo newFile = new FileInfo(@"Sample2.xlsx");

            // If there is any file having same name as 'Sample2.xlsx', then delete it first
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(@"Sample2.xlsx");
            }

            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                // Openning first Worksheet of the template file i.e. 'Sample1.xlsx'
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                // I'm adding 5th & 6th rows as 1st to 4th rows are already filled up with values in 'Sample1.xlsx'
                worksheet.InsertRow(5, 2);

                // Inserting values in the 5th row
                worksheet.Cells["A5"].Value = "12010";
                worksheet.Cells["B5"].Value = "Drill";
                worksheet.Cells["C5"].Value = 20;
                worksheet.Cells["D5"].Value = 8;

                // Inserting values in the 6th row
                worksheet.Cells["A6"].Value = "12011";
                worksheet.Cells["B6"].Value = "Crowbar";
                worksheet.Cells["C6"].Value = 7;
                worksheet.Cells["D6"].Value = 23.48;
            }
        }

        //void OPENXML()
        //{
        //    if (DGVmain.RowCount > 0)
        //    {
        //    string strfilepath = "C:\\Users\\m\\Desktop\\Employeedata.xlsx";
        //    using (ExcelPackage p = new ExcelPackage())
        //    {
        //        using (FileStream stream = new FileStream(strfilepath, FileMode.Open))
        //        {
        //            p.Load(stream);
        //            //deleting worksheet if already present in excel file
        //            var wk = p.Workbook.Worksheets.SingleOrDefault(x => x.Name == "Hola");
        //            if (wk != null) { p.Workbook.Worksheets.Delete(wk); }

        //            p.Workbook.Worksheets.Add("Hola");
        //            p.Workbook.Worksheets.MoveToEnd("Hola");
        //            ExcelWorksheet worksheet = p.Workbook.Worksheets[p.Workbook.Worksheets.Count];

        //            worksheet.InsertRow(5, 2);
        //            worksheet.Cells["A9"].LoadFromDataTable(dt1, true);
        //            // Inserting values in the 5th row
        //            worksheet.Cells["A5"].Value = "12010";
        //            worksheet.Cells["B5"].Value = "Drill";
        //            worksheet.Cells["C5"].Value = 20;
        //            worksheet.Cells["D5"].Value = 8;

        //            // Inserting values in the 6th row
        //            worksheet.Cells["A6"].Value = "12011";
        //            worksheet.Cells["B6"].Value = "Crowbar";
        //            worksheet.Cells["C6"].Value = 7;
        //            worksheet.Cells["D6"].Value = 23.48;
        //        }
        //        //p.Save() ;
        //        Byte[] bin = p.GetAsByteArray();
        //        File.WriteAllBytes(@"C:\Users\m\Desktop\Employeedata.xlsx", bin);
        //    }
        //}
        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
