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
        static private readonly Padding ROW_PADDING = new Padding(5, 5, 10, 10);

        public HelpForm()
        {
            InitializeComponent();
            this.helpTablePanel.RowCount = 0;
            this.helpTablePanel.ColumnCount = 2;
            this.helpTablePanel.AutoSize = true;
            this.helpTablePanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.helpTablePanel.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            this.AutoSize = true;
            this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Label itemHeader = new Label {
                Text = Properties.Resources.helpItemHeader,
                Font = new Font(Label.DefaultFont, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleRight,
            };
            Label descriptionHeader = new Label {
                Text = Properties.Resources.helpDescriptionHeader,
                Font = new Font(Label.DefaultFont, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
            };
            this.AddHelpRowFromLabels(itemHeader, descriptionHeader);
        }

        public void AddHelpRow(string item, string description)
        {
            Label itemLabel = new Label { Text = item };
            Label descriptionLabel = new Label { Text = description };
            itemLabel.TextAlign = ContentAlignment.MiddleRight;
            descriptionLabel.TextAlign = ContentAlignment.MiddleLeft;
            AddHelpRowFromLabels(itemLabel, descriptionLabel);
        }

        public void AddHelpRowFromLabels(Label left, Label right)
        {
            int rc = this.helpTablePanel.RowCount++;
            left.Dock = DockStyle.Fill;
            left.Margin = ROW_PADDING;
            right.Dock = DockStyle.Fill;
            right.Margin = ROW_PADDING;
            this.helpTablePanel.Controls.Add(left, 0, rc);
            this.helpTablePanel.Controls.Add(right, 1, rc);

            // Auto-size the rows
            this.helpTablePanel.RowStyles.Clear();
            for (int i = 0; i < this.helpTablePanel.RowCount; i++) {
                this.helpTablePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            // Auto-size the columns
            this.helpTablePanel.ColumnStyles.Clear();
            this.helpTablePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));
            this.helpTablePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70f));
        }
    }
}
