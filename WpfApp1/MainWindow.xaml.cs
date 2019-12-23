using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;
using System.Diagnostics;

using Outlook = Microsoft.Office.Interop.Outlook;

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security;

using CredentialManagement;

namespace WpfApp1
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filename;
        public MainWindow()
        {
            
            InitializeComponent();
        }

            private void excel_button(object sender, RoutedEventArgs e)
        {
          
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filename);          
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[Int32.Parse(sheet_num.Text)];

            try
            {
                xlApp.Visible = true;
             
            }
            catch (Exception theException) {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, "Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);
                MessageBox.Show(errorMessage, "Error");
            }

            //copy range from excel to clipboard
            Excel.Range range1 = xlWorksheet.Range[start_rowInput.Text, rowInput.Text];
            range1.Copy();
            Clipboard.ContainsImage();
            

            //sending email
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com"); // connect with gmail server

                mail.From = new MailAddress(input_from_email.Text); // from email address
                mail.To.Add(input_to_email.Text); // to email address
                mail.Subject = "Testing Excel to Email";

                AlternateView alternate = AlternateView.CreateAlternateViewFromString(Clipboard.GetText());


                string strHtmlBody =

                //set body contents from clipboard in html format
                mail.Body = Clipboard.GetText(TextDataFormat.Html);
                //yes it is html
                mail.IsBodyHtml = true;
                
                //eamil server set up 
                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential(input_from_email.Text, input_from_email_password.Password);
                SmtpServer.EnableSsl = true;

                //send email
                SmtpServer.Send(mail);

                //Checking messeage
                MessageBox.Show("Email Sent");
                mail.Dispose();//clean up
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            //KillSpecificExcelFileProcess(filename);
        }

  
        public static int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
        private void KillSpecificExcelFileProcess(string excelFileName)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;

            foreach (var process in processes)
            {
                if (process.MainWindowTitle == "Microsoft Excel - " + excelFileName)
                    process.Kill();
            }
        }

       private void button_file_finder(object sender, EventArgs e)
        {
            int size = -1;
            Microsoft.Win32.OpenFileDialog openFileDialog1 = new Microsoft.Win32.OpenFileDialog();
            openFileDialog1.ShowDialog();
        
            Console.WriteLine(size); // <-- Shows file size in debugging mode.

            filename = openFileDialog1.FileName;
        }
           
          
  
    }

}
