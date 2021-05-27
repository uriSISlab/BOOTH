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
    public partial class HelpForm : Form
    {
        public HelpForm()
        {
            InitializeComponent();
        }

        public void AddHelpRow(string item, string description)
        {
            RowStyle rowStyle = new RowStyle(SizeType.AutoSize);
            this.helpTablePanel.RowStyles.Add(rowStyle);
            Label itemLabel = new Label { Text = item };
            Label descriptionLabel = new Label { Text = description };
            itemLabel.TextAlign = ContentAlignment.MiddleRight;
            descriptionLabel.TextAlign = ContentAlignment.MiddleLeft;
            this.helpTablePanel.Controls.Add(itemLabel, 0, this.helpTablePanel.RowCount - 1);
            this.helpTablePanel.Controls.Add(descriptionLabel, 1, this.helpTablePanel.RowCount - 1);
            this.helpTablePanel.RowCount++;
        }
    }
}
