using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

using System.Timers;
using GetCoreTempInfoNET;
using Microsoft.Management.Infrastructure;

namespace ClevoFanControl
{
    public partial class ClevoFanControlService : ServiceBase
    {
        private CoreTempInfo coreTempInfo;
        CimInstance searchInstance = null;
        private static readonly string clevoAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)+"/ClevoFanControl";
        private static readonly string clevoLog = clevoAppData + "/ClevoFanControl.log";
        Timer eventTimer;

        private StreamWriter logWriter;
        public ClevoFanControlService()
        {
            InitializeComponent();

            //Initiate CoreTempInfo class.
            coreTempInfo = new CoreTempInfo();
            //Sign up for an event reporting errors
            coreTempInfo.ReportError += new ErrorOccured(CTInfo_ReportError);


            string Namespace = @"root\WMI";
            string className = "CLEVO_GET";
            CimInstance clevo = new CimInstance(className, Namespace);

            CimSession mySession = CimSession.Create("localhost");
            CimInstance searchInstance = mySession.GetInstance(Namespace, clevo);

            // Set up a timer to trigger every minute.  
            eventTimer = new Timer
            {
                Interval = 1000 // 60 seconds  
            };
            eventTimer.Elapsed += new ElapsedEventHandler(this.OnTimer);
            eventTimer.Start();

            if (!Directory.Exists(clevoAppData))
            {
                Directory.CreateDirectory(clevoAppData);
            }

        }

        void CTInfo_ReportError(ErrorCodes ErrCode, string ErrMsg)
        {
            eventLog.WriteEntry(ErrMsg);
        }

        private void OnTimer(object sender, ElapsedEventArgs e)
        {

            //Attempt to read shared memory.
            bool readSuccess = coreTempInfo.GetData();

            //If read was successful the post the new info on the console.
            if (readSuccess)
            {
                var temps = coreTempInfo.GetTemp;
                var maxTemp = temps.Max();
                logWriter.WriteLine("Tmax {0}", maxTemp);
            }
            else
            {
               eventLog.WriteEntry("Internal error name: " + coreTempInfo.GetLastError);
               eventLog.WriteEntry("Internal error message: " + coreTempInfo.GetErrorMessage(coreTempInfo.GetLastError));
            }
        }

        protected override void OnStart(string[] args)
        {
            eventLog.WriteEntry("Starting clevo fan control.");
            logWriter = new StreamWriter(clevoLog);
            eventTimer.Start(); 
        }

        protected override void OnStop()
        {
            eventTimer.Stop();
            logWriter.Close();
        }
    }
}
