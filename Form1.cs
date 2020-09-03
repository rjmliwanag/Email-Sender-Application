using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace EmailApps
{
    public partial class Form1 : Form
    {
        private XSSFWorkbook workBook = null;
        private ISheet sheet = null;
        private int rowIndex = 0;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!File.Exists("emailapp.xlsx"))
            {
                workBook = new XSSFWorkbook();
                sheet = (XSSFSheet)workBook.CreateSheet("emailsender");

                var row = sheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("EmailAdd:");
                row.CreateCell(1).SetCellValue("Subject:");
                row.CreateCell(2).SetCellValue("Message:");
                row.CreateCell(3).SetCellValue("Date and Time Sent:");

                using (var filedata = new FileStream("emailapp.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workBook.Write(filedata);
                }
            }

            rowIndex = 0;
            var emailAddresses = new List<string>();
            using (FileStream filedata = new FileStream("emailapp.xlsx", FileMode.Open, FileAccess.Read))
            {
                workBook = new XSSFWorkbook(filedata);
                sheet = workBook.GetSheet("emailsender");
                rowIndex = sheet.LastRowNum + 1;

                for (int i = 1; i < rowIndex; i++)
                {
                    emailAddresses.Add(Convert.ToString(sheet.GetRow(i).Cells[0]));
                }
            }

            foreach (var recipient in textBox1.Text.Split(','))
            {
                string pattern = "^([0-9a-zA-Z]([-\\.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";
                if (!Regex.IsMatch(recipient, pattern))
                {
                    button1.Enabled = false;
                    if (!string.IsNullOrEmpty(recipient)) MessageBox.Show($"{recipient} is an invalid email address.");
                    return;
                }

                string box_msg = $"{recipient} already received email.{Environment.NewLine}{Environment.NewLine}Remove Entry?";

                if (emailAddresses.Any(m => m.ToLower() == recipient.ToLower()))
                {
                    DialogResult = MessageBox.Show(box_msg, "Conditional", MessageBoxButtons.YesNo);
                    if (DialogResult == DialogResult.Yes)
                    {
                        var newRecepients = string.Join(",", textBox1.Text.Split(',').Where(m => m != recipient));
                        textBox1.Text = newRecepients;
                    }            
                    return;
                }
            }

            textBox2.Text = string.Empty;
            progressBar1.Value = 0;
            progressBar1.Maximum = textBox1.Text.Split(',').Length;
            backgroundWorker1.RunWorkerAsync();
            button1.Text = "Processing";
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                button1.Enabled = true;
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var result = string.Empty;

            foreach (var recipient in textBox1.Text.Split(','))
            {
                try
                {
                    MailMessage mm = new MailMessage(ConfigurationManager.AppSettings["Email"], recipient);
                    mm.Subject = ConfigurationManager.AppSettings["Subject"];
                    mm.Body = ConfigurationManager.AppSettings["EmailMessage"];
                    mm.Attachments.Add(new Attachment(ConfigurationManager.AppSettings["Attachment"]));
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = "smtp.gmail.com";
                    smtp.Port = 587;
                    smtp.UseDefaultCredentials = false;
                    smtp.EnableSsl = true;
                    NetworkCredential nc = new NetworkCredential(ConfigurationManager.AppSettings["Email"], ConfigurationManager.AppSettings["Password"]);
                    smtp.Credentials = nc;
                    smtp.Send(mm);
                    result = $"Mail has been sent succesfully to {recipient}{Environment.NewLine}";

                    var row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(recipient);
                    row.CreateCell(1).SetCellValue(ConfigurationManager.AppSettings["Subject"]);
                    row.CreateCell(2).SetCellValue(ConfigurationManager.AppSettings["EmailMessage"]);
                    row.CreateCell(3).SetCellValue(DateTime.Now.ToString("MM-dd-yyyy hh:ss"));
                    rowIndex++;

                }
                catch (TimeoutException ex)
                {
                    result = $"Timeout occurred sending to {recipient}, {ex.Message}{Environment.NewLine}";
                }
                catch (Exception ex)
                {
                    result = $"Error occurred sending to {recipient}, {ex.Message}{Environment.NewLine}";
                }
                          
                backgroundWorker1.ReportProgress(0, new WorkerStatus {Message = result});
            }

            using (FileStream filedata = new FileStream("emailapp.xlsx", FileMode.Create, FileAccess.Write))
            {
                workBook.Write(filedata);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value += 1;
            textBox2.Text += $"{((WorkerStatus)e.UserState).Message}{Environment.NewLine}";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Email Send Succesfully");
            button1.Text = "Send";
            progressBar1.Value = 0;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
        }
    }

    public class WorkerStatus
    {
        public string Message { get; set; }
    }
}
