namespace BOOTH
{
    partial class TimerBaseForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.commentTextBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.storeCommentButton = new System.Windows.Forms.Button();
            this.commentTimerSelectCombo = new System.Windows.Forms.ComboBox();
            this.timersPanel = new System.Windows.Forms.TableLayoutPanel();
            this.saveWorksheetButton = new System.Windows.Forms.Button();
            this.leftPanel = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // commentTextBox
            // 
            this.commentTextBox.Location = new System.Drawing.Point(365, 378);
            this.commentTextBox.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.commentTextBox.Multiline = true;
            this.commentTextBox.Name = "commentTextBox";
            this.commentTextBox.Size = new System.Drawing.Size(419, 56);
            this.commentTextBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(363, 360);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Comment";
            // 
            // storeCommentButton
            // 
            this.storeCommentButton.Location = new System.Drawing.Point(687, 357);
            this.storeCommentButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.storeCommentButton.Name = "storeCommentButton";
            this.storeCommentButton.Size = new System.Drawing.Size(95, 20);
            this.storeCommentButton.TabIndex = 2;
            this.storeCommentButton.Text = "Store Comment";
            this.storeCommentButton.UseVisualStyleBackColor = true;
            this.storeCommentButton.Click += new System.EventHandler(this.StoreCommentButton_Click);
            // 
            // commentTimerSelectCombo
            // 
            this.commentTimerSelectCombo.FormattingEnabled = true;
            this.commentTimerSelectCombo.Location = new System.Drawing.Point(419, 359);
            this.commentTimerSelectCombo.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.commentTimerSelectCombo.Name = "commentTimerSelectCombo";
            this.commentTimerSelectCombo.Size = new System.Drawing.Size(82, 21);
            this.commentTimerSelectCombo.TabIndex = 4;
            this.commentTimerSelectCombo.SelectedIndexChanged += new System.EventHandler(this.CommentTimerSelectCombo_SelectedIndexChanged);
            // 
            // timersPanel
            // 
            this.timersPanel.AutoSize = true;
            this.timersPanel.ColumnCount = 1;
            this.timersPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.timersPanel.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.AddColumns;
            this.timersPanel.Location = new System.Drawing.Point(8, 8);
            this.timersPanel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.timersPanel.Name = "timersPanel";
            this.timersPanel.RowCount = 1;
            this.timersPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 49F));
            this.timersPanel.Size = new System.Drawing.Size(1151, 342);
            this.timersPanel.TabIndex = 5;
            // 
            // saveWorksheetButton
            // 
            this.saveWorksheetButton.Location = new System.Drawing.Point(975, 390);
            this.saveWorksheetButton.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.saveWorksheetButton.Name = "saveWorksheetButton";
            this.saveWorksheetButton.Size = new System.Drawing.Size(97, 31);
            this.saveWorksheetButton.TabIndex = 6;
            this.saveWorksheetButton.Text = "Save Worksheet";
            this.saveWorksheetButton.UseVisualStyleBackColor = true;
            this.saveWorksheetButton.Click += new System.EventHandler(this.SaveWorksheetButton_Click);
            // 
            // leftPanel
            // 
            this.leftPanel.Location = new System.Drawing.Point(8, 360);
            this.leftPanel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.leftPanel.Name = "leftPanel";
            this.leftPanel.Size = new System.Drawing.Size(351, 142);
            this.leftPanel.TabIndex = 7;
            // 
            // TimerBaseForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(1167, 510);
            this.Controls.Add(this.leftPanel);
            this.Controls.Add(this.saveWorksheetButton);
            this.Controls.Add(this.timersPanel);
            this.Controls.Add(this.commentTimerSelectCombo);
            this.Controls.Add(this.storeCommentButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.commentTextBox);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "TimerBaseForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox commentTextBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button storeCommentButton;
        private System.Windows.Forms.ComboBox commentTimerSelectCombo;
        private System.Windows.Forms.TableLayoutPanel timersPanel;
        private System.Windows.Forms.Button saveWorksheetButton;
        private System.Windows.Forms.Panel leftPanel;
    }
}