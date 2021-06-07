using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class ProgressBarForm : Form
    {
        private Thread thread;
        private Action initializer;

        public ProgressBarForm()
        {
            InitializeComponent();
            this.thread = new Thread(() => this.ShowDialog());
        }

        public void InitializeAndShow(int maximum)
        {
            this.thread.Start();
            initializer = () =>
            {
                this.progressBar.Minimum = 0;
                this.progressBar.Maximum = maximum;
                this.progressBar.Step = 1;
                this.progressBar.Value = 0;
                this.progressBar.Visible = true;
            };
            if (this.IsHandleCreated)
            {
                BeginInvoke(initializer);
            } else
            {
                HandleCreated += (s, e) => BeginInvoke(initializer);
            }
        }


        public void Step()
        {
            Action stepAction = () =>
            {
                this.progressBar.PerformStep();
                this.progressLabel.Text = ((this.progressBar.Value * 100) / this.progressBar.Maximum) + "% completed";
            };
            if (this.IsHandleCreated)
            {
                BeginInvoke(stepAction);
            } else
            {
                HandleCreated += (s, e) => BeginInvoke(stepAction);
            }
        }

        public void Done()
        {
            Action doneAction = () =>
            {
                this.Close();
                this.Dispose();
            };
            if (this.IsHandleCreated)
            {
                BeginInvoke(doneAction);
            } else
            {
                HandleCreated += (s, e) => BeginInvoke(doneAction);
            }

        }
    }
}
