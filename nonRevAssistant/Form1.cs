using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using zoyobar.shared.panzer.web.ib;
using System.Web;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace nonRevAssistant
{
    public partial class Form1 : Form
    {
        private const int INTERNET_OPTION_END_BROWSER_SESSION = 42;
        [DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int lpdwBufferLength);
        [DllImport("wininet.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern bool InternetSetCookie(string lpszUrl, string lpszCookieName, string lpszCookieData);

        public string url = "";
        public string badge = "289527";
        public string confirmationNumber = "";
        public string departureCode = "";
        public IEBrowser ie = null;
        int windowtimes = 0;

        //args[0] is "US" or "AA", args[1] is 6-digit confirmation #, 
        //if "US", args[2] is departure airport
        //if "AA",
        public Form1(string cNum, string dAirport)
        {
            this.confirmationNumber = cNum;
            this.departureCode = dAirport;
            InitializeComponent();
            openUS();
        }

        private void openUS()
        {
            url = "https://travel.usairways.com/us/checkin/checkin.jsf";
            InternetSetOption(IntPtr.Zero, INTERNET_OPTION_END_BROWSER_SESSION, IntPtr.Zero, 0);
            webBrowser1.Navigate(url);
        }

        private void webBrowser_Complete(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            if (windowtimes == 0)
            {
                ie = new IEBrowser(this.webBrowser1);
                //string doc = this.webBrowser1.DocumentText;
                if ((e.Url != webBrowser1.Url) || (webBrowser1.ReadyState != WebBrowserReadyState.Complete))
                {
                    ((WebBrowser)sender).Document.Window.Error += new HtmlElementErrorEventHandler(Window_Error);
                    return;
                }
                string fillScript = "setTimeout(function(){document.getElementById('checkinForm:badgeNum').value = '" + badge;
                fillScript += "';document.getElementById('checkinForm:confirmCd').value='" + confirmationNumber;
                fillScript += "';document.getElementById('checkinForm:board').value='" + departureCode;
                fillScript += "';document.getElementById('checkinForm:btnLookup').click();},1300);";
                ie.ExecuteScript(fillScript);              
            }
            else if (windowtimes == 3)
            {
                //ie = new IEBrowser(this.webBrowser1);
                //string doc = this.webBrowser1.DocumentText;

                HtmlElementCollection checkbox = webBrowser1.Document.GetElementsByTagName("input");
                int indexOfCheckBox = 0;
                string fillScript = ""; 
                foreach (HtmlElement cb in checkbox)
                {
                    if (cb.GetAttribute("type").Equals("checkbox"))
                    {
                        cb.SetAttribute("id", "shoppersignin" + indexOfCheckBox);
                        fillScript += "document.getElementById('shoppersignin" + indexOfCheckBox++ + "').checked = true;";
                    }
                }
                ie.ExecuteScript("setTimeout(function(){" + fillScript + "document.getElementById('checkinForm:btnConfirm').click();},1000);");
                
            }
            else if (windowtimes == 6)
            {
                //send successful email then
                sendMail();
                Application.Exit();
            }
            windowtimes++;
        }

        private void Window_Error(object sender, HtmlElementErrorEventArgs e)
        {
            // Ignore the error and suppress the error dialog box.   
            e.Handled = false;
        }

        private void sendMail()
        {
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress("mmnnmit@gmail.com");
            msg.To.Add(new MailAddress("alexjhang@gmail.com"));

            msg.Subject = "US Non-rev checked in " + confirmationNumber;
            msg.Body = " ";
            SmtpClient smtp = new SmtpClient("smtp.gmail.com");
            smtp.UseDefaultCredentials = false;
            smtp.EnableSsl = true;
            smtp.Port = 587;
            smtp.Credentials = new NetworkCredential("mmnnmit@gmail.com", "laohutu123");
            smtp.Send(msg);
            
        }
    }
}
