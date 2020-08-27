using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class TimerBaseForm : Form
    {
        private TimerControl[] timers;
        private Worksheet sheet;

        public TimerBaseForm()
        {
            InitializeComponent();
        }

        private void TimerBase_Load(object sender, EventArgs e)
        {

        }

        public void PopulateTimers(Timers.TimerType[] timerTypes, Worksheet sheet)
        {
            this.timersPanel.RowCount = 1;
            this.timersPanel.ColumnCount = timerTypes.Length;
            this.timersPanel.ColumnStyles.Clear();
            this.sheet = sheet;
            timers = new TimerControl[timerTypes.Length];
            for (int i = 0; i < timerTypes.Length; i++)
            {
                int columnCount = Timers.GetColumnCountForTimerType(timerTypes[i]);
                SheetWriter writer = new SheetWriter(sheet,  0, i * columnCount);
                TimerControl control = Timers.GetTimerControl(timerTypes[i], writer, i + 1);
                this.timers[i] = control;
                this.timersPanel.Controls.Add(control, i, 0);
                this.timersPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                this.commentTimerSelectCombo.Items.Add(control.GetHeadingText());
            }
            this.storeCommentButton.Enabled = false;
        }

        private void StoreCommentButton_Click(object sender, EventArgs e)
        {
            TimerControl timer = this.timers[this.commentTimerSelectCombo.SelectedIndex];
            timer.AddComment(this.commentTextBox.Text);
            this.commentTextBox.Clear();
        }

        private void CommentTimerSelectCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.storeCommentButton.Enabled = this.commentTimerSelectCombo.SelectedIndex >= 0;
        }

        private void SaveWorksheetButton_Click(object sender, EventArgs e)
        {
            ThisAddIn.app.ActiveWorkbook.Save();
        }
    }
}
