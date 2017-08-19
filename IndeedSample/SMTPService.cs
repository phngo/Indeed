using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using static IndeedSample.Program;

namespace IndeedSample
{
    public partial class SMTPService : ServiceBase
    {
        private System.Diagnostics.EventLog eventLog;
        private int eventId;

        private SmtpClient mailer;

        private string SMTPserver;
        private int SMTPport;
        private string SMTPLogin;
        private string SMTPpassword;
        private string EmailAddress;

        private bool hasError;

        public SMTPService()
        {
            hasError = false;
            InitializeComponent();
            InitilizeConfigValue();

            //Initilize mailer
            mailer = new SmtpClient();
            mailer.Host = SMTPserver;
            mailer.Port = SMTPport;
            mailer.Credentials = new System.Net.NetworkCredential(SMTPLogin, SMTPpassword);

            //Initilize Log
            eventLog = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("MySource"))
            {
                System.Diagnostics.EventLog.CreateEventSource(
                    "MySource", "MyNewLog");
            }
            eventLog.Source = "MySource";
            eventLog.Log = "MyNewLog";
        }
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(IntPtr handle, ref ServiceStatus serviceStatus);
        
        //Initialize Parameter
        private void InitilizeConfigValue()
        {
            try
            {
                SMTPserver = ConfigurationManager.AppSettings["SMTPserver"];
                SMTPport = Int32.Parse(ConfigurationManager.AppSettings["SMTPport"]);
                SMTPLogin = ConfigurationManager.AppSettings["SMTPLogin"];
                SMTPpassword = ConfigurationManager.AppSettings["SMTPpassword"];
                EmailAddress = ConfigurationManager.AppSettings["EmailAddress"];
            }
            catch(Exception ex)
            {
                eventLog.WriteEntry("Error in InitilizeConfigValue : " + ex.ToString());
                hasError = true;
            }
        }
        protected override void OnStart(string[] args)
        {
            eventLog.WriteEntry("In OnStart");

            //Timer
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 300000; // 300 seconds  = 5 mins
            timer.Elapsed += new System.Timers.ElapsedEventHandler(this.OnTimer);
            timer.Start();

            // Update the service state to Start Pending.  
            ServiceStatus serviceStatus = new ServiceStatus();
            serviceStatus.dwCurrentState = ServiceState.SERVICE_START_PENDING;
            serviceStatus.dwWaitHint = 100000;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);

            // Update the service state to Running.  
            serviceStatus.dwCurrentState = ServiceState.SERVICE_RUNNING;
            SetServiceStatus(this.ServiceHandle, ref serviceStatus);
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            if (!hasError)
            {
                eventLog.WriteEntry("Monitoring the System", EventLogEntryType.Information, eventId++);

                //Preparing message
                MailMessage message = new MailMessage();
                DateTime dt = DateTime.Now;
                message.From = new System.Net.Mail.MailAddress("indeedsample@indeedsample.com", "Indeed Sample");
                message.To.Add(EmailAddress);
                message.Subject = "Indeed Sample [" + dt.ToString() + "]";
                message.Body = "This is a sample";
                try
                {
                    //First attempt
                    mailer.Send(message);
                }
                catch(Exception ex)
                {
                    try
                    {
                        //Re-try
                        mailer.Send(message);
                    }
                    catch(Exception ex1)
                    {
                        eventLog.WriteEntry("Error in Sending mail : " + ex1.ToString());
                    }
                }
            }
        }

        protected override void OnContinue()
        {
            eventLog.WriteEntry("In OnContinue.");
        }
        protected override void OnStop()
        {
            eventLog.WriteEntry("In OnStop.");
        }
    }
}
