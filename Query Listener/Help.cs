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

namespace Query_Listener
{

    public partial class Help : Form
    {
        /// <summary>
        /// Software created by August Bryan N. Florese  
        /// Contact: Aroueterra@gmail.com
        /// For: Tax team member Bryan Rucio of Convergys Finance department
        /// Author and developer of the code retains the program property rights
        /// </summary>

        public Help()
        {
            InitializeComponent();
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case 0x84:
                    base.WndProc(ref m);
                    if ((int) m.Result == 0x1)
                        m.Result = (IntPtr) 0x2;
                    return;
            }
            base.WndProc(ref m);
        }

        private void btnCrud_Click(object sender, EventArgs e)
        {
            RTB.Text = "To delete records, first select the [Drag] hand icon to select an ID to delete, then select the [DELETE] button to delete the item permanently." + Environment.NewLine +
           "To update a record, first select the [Drag] hand icon to select an ID to delete, then select the [UPDATE] button to update the records based on the active textboxes on the left pane. " + Environment.NewLine +
           "The left pane textboxes must be filled before you can insert the data into the table. All fields must be filled before you can proceed. Currency fields cannot accept string/text data types." + Environment.NewLine +
           "Finally, the [EXCEL SHEET] button functions to add an entire column to an existing excel sheet on COLUMN L, because of this, your Net Tax Field HAS to be on column K. Make sure there is no excel file open at the time, to make sure of this, open task manager and end the excel task. The feature works by performing a formula on columns 12, 13 and 14 as stated in the formula. ";

        }

        private void btnreturn_Click(object sender, EventArgs e)
        {
            this.Hide();
            btnBackup.Visible = false;
            btnRestore.Visible = false;
        }


        private void btnBackup_Click(object sender, EventArgs e)
        {
            string CurrentDatabasePath = Environment.CurrentDirectory + @"\ACDB.accdb";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string PathtobackUp = fbd.SelectedPath.ToString();
                File.Copy(CurrentDatabasePath, PathtobackUp + @"\BackUp.accdb", true);
                MessageBox.Show("Back Up Successfull! ");
            }
        }

        private void btnRestore_Click(object sender, EventArgs e)
        {
            string PathToRestoreDB = Environment.CurrentDirectory + @"\ACDB.accdb";
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                string Filetorestore = ofd.FileName;

                // Rename Current Database to .Bak

                //File.Move(PathToRestoreDB, PathToRestoreDB + ".bak");

                //Restore the Databse From Backup Folder
                File.Copy(Filetorestore, PathToRestoreDB, true);

            }
        }

        private void btnDb_Click(object sender, EventArgs e)
        {
            btnBackup.Visible = true;
            btnRestore.Visible = true;
            RTB.Text = "Select the [BACKUP] button to create a backup of your ACCDB file." + Environment.NewLine +
                       "Select the [RESTORE] button to transfer ACCDB file to your program folder. Be aware, this is an experimental feature and may have unexpected consequenses";
        }

        private void btnDrag_Click(object sender, EventArgs e)
        {
            RTB.Text = "Clicking the upper hand icon [DRAG]s the data of the selected table row into the active boxes for editing or printing.";
        }

        private void btnSheet_Click(object sender, EventArgs e)
        {
            RTB.Text = "Select the [IMPORT] button to transfer the entire contents of the visible Excel table into the Access table." + Environment.NewLine +
           "Select the [APPEND] button to update the value of the selected row with the value from another Excel file. The ID is the unique field of reference. This feature is built with identical records in mind. Make sure that the structure, not just the data, matches the table format.";
        }

        private void btnCrit_Click(object sender, EventArgs e)
        {
            RTB.Text = "Select a criteria from the [Drop-down menu] to search the active table for a specific value." + Environment.NewLine +
           "The Access and Excel tables have separate menu buttons";
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            RTB.Text = "The [Print] button grabs all of the data from the active textboxes on the left pane. (The Summary)" + Environment.NewLine +
           "After gathering the data, select the formatted blank Excel spreadsheet with the [SOURCE] button which you will use to print over a printed 2316 form. " + Environment.NewLine +
           "Now select the print button to transfer the data to Excel.";
        }

        private void btnMail_Click(object sender, EventArgs e)
        {
            RTB.Text = "For feedback and suggestions, contact the author of the software system: August Bryan N. Florese" + Environment.NewLine +
"Aroueterra@Gmail.com"
+ Environment.NewLine + Environment.NewLine + Environment.NewLine +
"If you wish to retrieve the installer, blank form or the driver installers, simply download them from the following Gdrive link: " + Environment.NewLine + " https://drive.google.com/open?id=0BxCGSzA16oq9OWZ5bmpya3M2am8 ";
            RTB.Enabled = true;
        }
    }
}
