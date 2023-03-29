using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net;
using System.Net.Mail;
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
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace DesktopMailingLisit
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? EmailFrom { get; set; }
        private string? PasswordFrom { get; set; }
        private string? Subject { get; set; }
        private string? EmailName { get; set; }
        private bool IsHtml { get; set; }
        private List<Email> Emails { get; set; } = new List<Email>();

        public MainWindow()
        {
            InitializeComponent();
            InitializeGrid();
            InitializeFields();

            Email.ItemsSource = Emails;
            SendButton.Click += SendEmails;
            AddEmailButton.Click += AddEmail;
            DeleteEmailButton.Click += DeleteEmail;
            SendTestButton.Click += SendTestEmail;
            IncludeAll.Click += IncludeAllEmails;
        }

        private void IncludeAllEmails(object sender, RoutedEventArgs e)
        {
            foreach (var email in Emails)
            {
                email.Include = true;
            }
            Email.Items.Refresh();
        }

        private void InitializeFields()
        {
            using (StreamReader r = new StreamReader("files/Settings.txt"))
            {
                var line = r.ReadLine();
                EmailTextBox.Text = line?.Substring(line.IndexOf(':') + 2);

                line = r.ReadLine();
                UTF8Encoding encoder = new UTF8Encoding();
                Decoder utf8Decode = encoder.GetDecoder();
                byte[] todecode_byte = Convert.FromBase64String(line?.Substring(line.IndexOf(':') + 2));
                int charCount = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length);
                char[] decoded_char = new char[charCount];
                utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0);
                PasswordTextBox.Password = new string(decoded_char);

                line = r.ReadLine();
                SubjectTextBox.Text = line?.Substring(line.IndexOf(':') + 2);

                line = r.ReadLine();
                NameTextBox.Text = line?.Substring(line.IndexOf(':') + 2);

                line = r.ReadLine();
                IsHtmlCheckBox.IsChecked = Convert.ToBoolean(line?.Substring(line.IndexOf(':') + 2));

            }
        }

        private void InitializeGrid()
        {
            FileInfo newFile = new FileInfo("files/Emails.xlsx");
            ExcelPackage pck = new ExcelPackage(newFile);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var sheet = pck.Workbook.Worksheets[0];
            var emails = new List<Email>();

            bool isEnd = false;
            int i = 1;
            while (!isEnd)
            {
                try
                {
                    if (sheet.Cells[i, 1].Value.ToString() == "END")
                    {
                        isEnd = true;
                        continue;
                    }
                    emails.Add(new Email()
                    {
                        Id = Convert.ToInt32(sheet.Cells[i, 1].Value.ToString()),
                        EmailString = sheet.Cells[i, 2].Value.ToString()?.Trim(),
                        Include = Convert.ToBoolean(sheet.Cells[i, 3].Value.ToString()),
                    });
                    i++;
                }
                catch
                {
                    isEnd = true;
                }
            }
            Emails = emails;
            pck.Dispose();
            return;
        }

        private void SendTestEmail(object sender, RoutedEventArgs e)
        {
            try
            {
                string body = File.ReadAllText("files/Message.txt").Replace("DATE", Month(DateTime.Now.Month));
                EmailFrom = EmailTextBox.Text;
                PasswordFrom = PasswordTextBox.Password;
                Subject = SubjectTextBox.Text;
                EmailName = NameTextBox.Text;

                MailAddress from = new MailAddress(EmailFrom, EmailName);
                MailAddress to = new MailAddress("trinity221078@gmail.com");
                MailMessage message = new MailMessage(from, to);
                message.Subject = Subject;
                message.BodyEncoding = Encoding.UTF8;
                if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                {
                    message.Body += "<html>";
                    message.Body += "<head>";
                    message.Body += "<meta charset=\"utf-8\">";
                    message.Body += "</head>";
                    message.Body += "<body>";
                }
                message.Body = body;
                if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                {
                    message.Body += "</body>";
                    message.Body += "</html>";
                }
                if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                    message.IsBodyHtml = true;
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.Credentials = new NetworkCredential(EmailFrom, PasswordFrom);
                smtp.EnableSsl = true;
                smtp.Send(message);
                MessageBox.Show("Тестове повідомлення успішно відправлене! Перевірте Вашу скриньку.", "Успіх!");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message, "Помилка!");
            }
        }

        private void DeleteEmail(object sender, RoutedEventArgs e)
        {
            try
            {
                var index = Convert.ToInt32(DeleteEmailTextBox.Text) - 1;
                Emails.RemoveAt(index);

                for (; index < Emails.Count; index++)
                {
                    Emails[index].Id--;
                }
                Email.Items.Refresh();
            }
            catch
            {
                return;
            }

        }

        private void AddEmail(object sender, RoutedEventArgs e)
        {
            Emails.Add(new Email { Id = Emails.Count + 1, EmailString = AddEmailTextBox.Text, Include = true });
            Email.Items.Refresh();
        }

        private void SendEmails(object sender, RoutedEventArgs e)
        {
            var i = 1;
            try
            {
                string body = File.ReadAllText("files/Message.txt").Replace("DATE", Month(DateTime.Now.Month));
                string messageBody = "";

                if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                {
                    messageBody = "<html>";
                    messageBody += "<html>";
                    messageBody += "<head>";
                    messageBody += "<meta charset=\"utf-8\">";
                    messageBody += "</head>";
                    messageBody += "<body>";
                }
                messageBody = body;
                if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                {
                    messageBody += "</body>";
                    messageBody += "</html>";
                }

                EmailFrom = EmailTextBox.Text;
                PasswordFrom = PasswordTextBox.Password;
                Subject = SubjectTextBox.Text;
                EmailName = NameTextBox.Text;

                MailAddress from = new MailAddress(EmailFrom, EmailName);
                SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                smtp.Credentials = new NetworkCredential(EmailFrom, PasswordFrom);
                smtp.EnableSsl = true;
                smtp.DeliveryFormat = SmtpDeliveryFormat.International;

                foreach (var email in Emails.Where(e => e.Include))
                {
                    MailAddress to = new MailAddress(email.EmailString);
                    MailMessage message = new MailMessage(from, to);

                    message.Subject = Subject;
                    message.BodyEncoding = Encoding.UTF8;
                    message.Body = messageBody;
                    if (IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value)
                        message.IsBodyHtml = true;
                    smtp.Send(message);
                    i++;
                    email.Include = false;
                }

                MessageBox.Show("Усі повідомлення відправлені! Перевірте Вашу скриньку.", "Успіх!");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message + "\nМожливо перевищено ліміт повідомлень. Останнє повідомлення під номером " + i.ToString() + ".\nПросто активуйте розсилку ще раз!", "Помилка!");
            }
        }

        private string Month(int number)
        {
            var name = "";
            switch (number)
            {
                case 1:
                    name = "січня";
                    break;
                case 2:
                    name = "лютого";
                    break;
                case 3:
                    name = "березня";
                    break;
                case 4:
                    name = "квітня";
                    break;
                case 5:
                    name = "травня";
                    break;
                case 6:
                    name = "червня";
                    break;
                case 7:
                    name = "липня";
                    break;
                case 8:
                    name = "серпня";
                    break;
                case 9:
                    name = "вересня";
                    break;
                case 10:
                    name = "жовтня";
                    break;
                case 11:
                    name = "листопада";
                    break;
                case 12:
                    name = "грудня";
                    break;
            }
            return name;
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (!(SaveDataCheckBox.IsChecked.HasValue && SaveDataCheckBox.IsChecked.Value))
            {
                return;
            }
            EmailFrom = EmailTextBox.Text;
            PasswordFrom = PasswordTextBox.Password;
            Subject = SubjectTextBox.Text;
            EmailName = NameTextBox.Text;
            IsHtml = IsHtmlCheckBox.IsChecked.HasValue && IsHtmlCheckBox.IsChecked.Value;

            byte[] encData_byte = new byte[PasswordFrom.Length];
            encData_byte = Encoding.UTF8.GetBytes(PasswordFrom);
            string encodedData = Convert.ToBase64String(encData_byte);

            string toWrite = "Email: " + EmailFrom + "\nPassword: " + encodedData + "\nSubject: " + Subject + "\nName: " + EmailName + "\nIsHtml: " + IsHtml.ToString();
            using (StreamWriter w = new StreamWriter("files/Settings.txt"))
            {
                w.Write(toWrite);
            }

            Emails = (List<Email>)Email.ItemsSource;
            FileInfo newFile = new FileInfo("files/Emails.xlsx");
            ExcelPackage pck = new ExcelPackage(newFile);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var sheet = pck.Workbook.Worksheets[0];

            for (int i = 1; i <= Emails.Count; i++)
            {
                sheet.Cells[i, 1].Value = Emails[i - 1].Id.ToString();
                sheet.Cells[i, 2].Value = Emails[i - 1].EmailString;
                sheet.Cells[i, 3].Value = Emails[i - 1].Include.ToString();
            }
            sheet.Cells[Emails.Count + 1, 1].Value = "END";
            pck.SaveAs(newFile);
            pck.Dispose();
        }
    }

    
}
