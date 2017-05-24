using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;

namespace MyWindowsFormApplication
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void generateButton_Click(object sender, System.EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

                //Manipulate a variable number of columns for Quarterly Sales Data.
               // DisplayQuarterlySales(oSheet);

                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Select a Excel File";

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void openFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void pullButton_Click(object sender, EventArgs e)
        {

        }

        private void filePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void sendButton_Click(object sender, EventArgs e)
        {
            Outlook.Application outlookApplication = null;
            Outlook._NameSpace outlookNameSpace = null;
            Outlook.MailItem item = null;
            Outlook.MAPIFolder inbox = null;
            Outlook.MAPIFolder subfolder = null;

            try
            {
                outlookApplication = new Outlook.Application();
                outlookNameSpace = outlookApplication.GetNamespace("MAPI");
                outlookNameSpace.Logon(null, null, false, false);

                inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                subfolder = inbox.Folders["TestFolder"];
                Console.WriteLine("Folder Name: {0}, Entry ID: {1}", subfolder.Name, subfolder.EntryID);
                Console.WriteLine("Num Items: {0}", subfolder.Items.Count.ToString());

                foreach(Outlook.MailItem subfolderItem in subfolder.Items)
                {
                    item = subfolderItem;
                    MessageBox.Show($"Item: {item.ToString()}\nSubject: {item.Subject}\nSent: {item.SentOn.ToLongDateString()}\nCategories: {item.Categories}\nBody: {item.Body}\nHTMLBody: {item.HTMLBody}");
                    foreach(Outlook.Attachment att in item.Attachments)
                    {
                        att.SaveAsFile(@"C:\TestFolder\" + att.FileName);
                    }
                    
                }
            } catch (System.Runtime.InteropServices.COMException exception)
            {
                Console.WriteLine(exception.ToString());
            }
            outlookNameSpace = null;
            outlookApplication = null;
            inbox = null;
        }
    }
}
