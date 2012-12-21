using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace CSTMgrProject
{
    public partial class ProgressWindow : Form
    {
        int iTotal;

        /// <summary>
        /// Constructor for the class
        /// </summary>
        /// <param name="sTitle"></param>
        /// <param name="pTotal"></param>
        public ProgressWindow(string sTitle, int pTotal)
        {
            InitializeComponent();
            this.Text = sTitle;
            iTotal = pTotal;
        }

        /// <summary>
        /// Overloaded Constructor, creates a marquee progress bar for inidicating work is being done
        /// </summary>
        /// <param name="sTitle"></param>
        public ProgressWindow(string sTitle)
        {
            InitializeComponent();
            this.Text = sTitle;
            pgbProgress.Style = ProgressBarStyle.Marquee;
            pgbProgress.MarqueeAnimationSpeed = 100;
            lblPercent.Visible = false;
        }

        /// <summary>
        /// Sets the progress bar to the percentage done based on the value passed in
        /// </summary>
        /// <param name="value">Percentage the task is complete</param>
        public void SetPercentDone(int doneSoFar)
        {
            int percentDone = 100*(doneSoFar+1)/iTotal;

            lblPercent.Text = percentDone.ToString() + "%";

            pgbProgress.Value = percentDone;
            //actually updates the progress bar on the screen
            pgbProgress.Update();
            this.Refresh();
        }

        /// <summary>
        /// When the marquee bar is used this method is used to update it
        /// </summary>
        public void Draw()
        {
            while (true)
            {
                this.Refresh();
                Thread.Sleep(100);
            }
        }
    }
}
