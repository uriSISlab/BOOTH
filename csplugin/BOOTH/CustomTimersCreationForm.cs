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
using System.Windows.Forms.VisualStyles;

namespace BOOTH
{
    public partial class CustomTimersCreationForm : Form
    {
        private readonly TimerControl.TimerType[] allTypes;
        private readonly string[] allNiceNames;
        private Worksheet worksheet;
        private bool created;

        public CustomTimersCreationForm(Worksheet sheet)
        {
            InitializeComponent();
            this.allTypes = TimerControl.GetMainPanelTimerTypes();
            this.allNiceNames = new string[this.allTypes.Length];
            this.timerSelectComboBox.Items.Clear();
            for (int i = 0; i < this.allTypes.Length; i++)
            {
                this.allNiceNames[i] = TimerControl.GetNiceNameForTimerType(this.allTypes[i]);
                this.timerSelectComboBox.Items.Add(this.allNiceNames[i]);
            }
            this.timerAddButton.Enabled = false;
            this.upButton.Enabled = false;
            this.downButton.Enabled = false;
            this.deleteButton.Enabled = false;
            this.worksheet = sheet;
            this.created = false;
        }

        private void TimerAddButton_Click(object sender, EventArgs e)
        {
            TimerControl.TimerType timerType = this.allTypes[this.timerSelectComboBox.SelectedIndex];
            this.timersListBox.Items.Add(TimerControl.GetNiceNameForTimerType(timerType));
        }

        private void TimerSelectComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.timerAddButton.Enabled = this.timerSelectComboBox.SelectedIndex >= 0;
        }

        private void UpButton_Click(object sender, EventArgs e)
        {
            int itemIdx = this.timersListBox.SelectedIndex;
            string item = this.timersListBox.Items[itemIdx].ToString();
            if (itemIdx > 0)
            {
                this.timersListBox.Items.RemoveAt(itemIdx);
                this.timersListBox.Items.Insert(itemIdx - 1, item);
                this.timersListBox.SetSelected(itemIdx - 1, true);
            }
        }

        private void DownButton_Click(object sender, EventArgs e)
        {
            int itemIdx = this.timersListBox.SelectedIndex;
            string item = this.timersListBox.Items[itemIdx].ToString();
            if (itemIdx < this.timersListBox.Items.Count - 1)
            {
                this.timersListBox.Items.RemoveAt(itemIdx);
                this.timersListBox.Items.Insert(itemIdx + 1, item);
                this.timersListBox.SetSelected(itemIdx + 1, true);
            }
        }

        private void TimersListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool enable = this.timersListBox.SelectedIndex >= 0;
            this.upButton.Enabled = enable;
            this.downButton.Enabled = enable;
            this.deleteButton.Enabled = enable;
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            int itemIdx = this.timersListBox.SelectedIndex;
            this.timersListBox.Items.RemoveAt(itemIdx);
            if (itemIdx != 0)
            {
                this.timersListBox.SetSelected(itemIdx - 1, true);
            } else if (this.timersListBox.Items.Count > 0)
            {
                this.timersListBox.SetSelected(0, true);
            }
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            int count = this.timersListBox.Items.Count;
            if (count == 0 && !this.arrivalTimerCheckbox.Checked)
            {
                Util.MessageBox("Please select at least one timer.");
                return;
            }
            string[] niceNames = new string[count];
            this.timersListBox.Items.CopyTo(niceNames, 0);
            TimerControl.TimerType[] toCreateTypes =  niceNames.Select(
                name => this.allTypes[Array.IndexOf(this.allNiceNames, name)]).ToArray();
            TimerBaseForm.CreateWithTimerTypes(toCreateTypes, this.arrivalTimerCheckbox.Checked, this.worksheet,
                Properties.Resources.customTimersTitle).Show();
            this.created = true;
            this.Dispose();
        }

        private void CustomTimersCreationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!this.created)
            {
                this.worksheet.Delete();
            }
        }
    }
}
