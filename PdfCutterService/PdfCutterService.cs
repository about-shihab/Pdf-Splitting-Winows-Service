using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PdfCutterService
{
    public partial class PdfCutterService : ServiceBase
    {
        PdfManager pdfManager = new PdfManager();
        public PdfCutterService()
        {
            InitializeComponent();

        }
        public void OnDebug()
        {
            OnStart(null);
        }
        protected override void OnStart(string[] args)
        {
            pdfManager.SendMail("MT700 Service has been Stopped", "MT700 Service Alert");

            pdfManager.WriteToFile("\n ---------------------------------------------------------------------------------\n");
            pdfManager.WriteToFile("Pdf Split Service started {0}");

            
            this.ScheduleService();
            //Console.ReadLine();
        }

        protected override void OnStop()
        {
            pdfManager.SendMail("MT700 Service has been Stopped", "MT700 Service Alert");
            pdfManager.WriteToFile("Pdf Split Service stopped {0}\n");
            this.Schedular.Dispose();
        }


        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
               // pdfManager.WriteToFile("Service Mode: " + mode + " {0}");



                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);
                    pdfManager.WriteToFile("Service Interval time " + intervalMinutes + " minutes");

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                pdfManager.WriteToFile("Scheduled Service wil run at " + DateTime.Now.AddMinutes(Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"])));

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                pdfManager.WriteToFile("Pdf split Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {

            try
            {
                string type = "*.pdf";
                pdfManager.ExtractPdf(type);

                this.ScheduleService();
            }
            catch (Exception ex)
            {
                pdfManager.WriteToFile("Pdf Split Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("Pdf Split Log Service"))
                {
                    serviceController.Stop();
                }
            }
        }

        
    }
}
