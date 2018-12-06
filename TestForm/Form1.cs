using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Inflectra.SpiraTest.PlugIns.FogBugzDataSync;

namespace TestForm
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            //Instantiate a new FogBugz class
            FogBugz fogBugz = new FogBugz("https://evildauphin.fogbugz.com");
            fogBugz.VerifyCertificate = false;

            //Verify the API
            if (!fogBugz.VerifyApi())
            {
                MessageBox.Show("FogBugz API not compatible with this plug-in");
            }
           
            //Now log-in
            fogBugz.Logon("evil_dauphin@yahoo.com", "zigzag");

            //Now create a new case
            FogBugzCase newCase = new FogBugzCase();
            newCase.Title = "Test Form Case";
            newCase.Project = 1;
            newCase.Priority = 3;
            newCase.Area = 4;
            newCase.FixFor = 2;
            newCase.Category = 1;
            newCase.Computer = "Windows 2003";
            newCase.Version = "Version 1";
            newCase.HrsCurrEst = 1;
            newCase.PersonAssignedTo = 2;
            newCase.Description = "Long Description";
            newCase.Due = DateTime.Now.AddDays(2);
            newCase = fogBugz.Add(newCase);

            //Display the id of the newly inserted case
            MessageBox.Show("Inserted new case " + newCase.Id.ToString() + ".");

            //Finally log-off
            fogBugz.Logoff();
        }

        private void btnRetrieve_Click(object sender, EventArgs e)
        {
            //Instantiate a new FogBugz class
            FogBugz fogBugz = new FogBugz("https://evildauphin.fogbugz.com");
            fogBugz.VerifyCertificate = false;

            //Verify the API
            if (!fogBugz.VerifyApi())
            {
                MessageBox.Show("FogBugz API not compatible with this plug-in");
            }

            //Now log-in
            fogBugz.Logon("evil_dauphin@yahoo.com", "zigzag");

            //Now retrieve a specific case
            FogBugzCase fogBugzCase = fogBugz.GetCase(Int32.Parse(this.txtCaseId.Text));

            //Display the id of the retrieved case
            MessageBox.Show("Retrieved case " + fogBugzCase.Id.ToString() + ".");

            //Finally log-off
            fogBugz.Logoff();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Instantiate a new FogBugz class
            FogBugz fogBugz = new FogBugz("https://evildolphin.fogbugz.com");
            fogBugz.VerifyCertificate = false;

            //Verify the API
            if (!fogBugz.VerifyApi())
            {
                MessageBox.Show("FogBugz API not compatible with this plug-in");
            }

            //Now log-in
            fogBugz.Logon("pack478sandman@verizon.net", "SemperF1");

            //Now retrieve a specific case
            FogBugzFixFor fogBugzFixFor = fogBugz.GetFixFor(Int32.Parse(this.txtCaseId.Text));

            //Display the id of the retrieved case
            MessageBox.Show("Retrieved fixfor " + fogBugzFixFor.Name.ToString() + ".");

            //Finally log-off
            fogBugz.Logoff();
        }
    }
}
