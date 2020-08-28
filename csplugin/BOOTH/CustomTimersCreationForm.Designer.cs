namespace BOOTH
{
    partial class CustomTimersCreationForm
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
            this.arrivalTimerCheckbox = new System.Windows.Forms.CheckBox();
            this.timerSelectComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.timerAddButton = new System.Windows.Forms.Button();
            this.timersListBox = new System.Windows.Forms.ListBox();
            this.upButton = new System.Windows.Forms.Button();
            this.downButton = new System.Windows.Forms.Button();
            this.deleteButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.createButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // arrivalTimerCheckbox
            // 
            this.arrivalTimerCheckbox.AutoSize = true;
            this.arrivalTimerCheckbox.Location = new System.Drawing.Point(35, 353);
            this.arrivalTimerCheckbox.Name = "arrivalTimerCheckbox";
            this.arrivalTimerCheckbox.Size = new System.Drawing.Size(177, 24);
            this.arrivalTimerCheckbox.TabIndex = 0;
            this.arrivalTimerCheckbox.Text = "Include Arrival Timer";
            this.arrivalTimerCheckbox.UseVisualStyleBackColor = true;
            // 
            // timerSelectComboBox
            // 
            this.timerSelectComboBox.FormattingEnabled = true;
            this.timerSelectComboBox.Items.AddRange(new object[] {
            "Check In",
            "BMD",
            "Voting Booth",
            "Ballot Scanning",
            "Throughput"});
            this.timerSelectComboBox.Location = new System.Drawing.Point(35, 49);
            this.timerSelectComboBox.Name = "timerSelectComboBox";
            this.timerSelectComboBox.Size = new System.Drawing.Size(195, 28);
            this.timerSelectComboBox.TabIndex = 1;
            this.timerSelectComboBox.SelectedIndexChanged += new System.EventHandler(this.TimerSelectComboBox_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(31, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 20);
            this.label1.TabIndex = 2;
            this.label1.Text = "Choose timer to add";
            // 
            // timerAddButton
            // 
            this.timerAddButton.Location = new System.Drawing.Point(236, 49);
            this.timerAddButton.Name = "timerAddButton";
            this.timerAddButton.Size = new System.Drawing.Size(75, 28);
            this.timerAddButton.TabIndex = 3;
            this.timerAddButton.Text = "Add";
            this.timerAddButton.UseVisualStyleBackColor = true;
            this.timerAddButton.Click += new System.EventHandler(this.TimerAddButton_Click);
            // 
            // timersListBox
            // 
            this.timersListBox.FormattingEnabled = true;
            this.timersListBox.ItemHeight = 20;
            this.timersListBox.Location = new System.Drawing.Point(35, 129);
            this.timersListBox.Name = "timersListBox";
            this.timersListBox.Size = new System.Drawing.Size(230, 204);
            this.timersListBox.TabIndex = 4;
            this.timersListBox.SelectedIndexChanged += new System.EventHandler(this.TimersListBox_SelectedIndexChanged);
            // 
            // upButton
            // 
            this.upButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.upButton.ForeColor = System.Drawing.Color.Blue;
            this.upButton.Location = new System.Drawing.Point(271, 129);
            this.upButton.Name = "upButton";
            this.upButton.Size = new System.Drawing.Size(40, 40);
            this.upButton.TabIndex = 5;
            this.upButton.Text = "⬆";
            this.upButton.UseVisualStyleBackColor = true;
            this.upButton.Click += new System.EventHandler(this.UpButton_Click);
            // 
            // downButton
            // 
            this.downButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.downButton.ForeColor = System.Drawing.Color.Blue;
            this.downButton.Location = new System.Drawing.Point(271, 175);
            this.downButton.Name = "downButton";
            this.downButton.Size = new System.Drawing.Size(40, 40);
            this.downButton.TabIndex = 6;
            this.downButton.Text = "⬇";
            this.downButton.UseVisualStyleBackColor = true;
            this.downButton.Click += new System.EventHandler(this.DownButton_Click);
            // 
            // deleteButton
            // 
            this.deleteButton.ForeColor = System.Drawing.Color.Red;
            this.deleteButton.Location = new System.Drawing.Point(271, 293);
            this.deleteButton.Name = "deleteButton";
            this.deleteButton.Size = new System.Drawing.Size(40, 40);
            this.deleteButton.TabIndex = 7;
            this.deleteButton.Text = "❌";
            this.deleteButton.UseVisualStyleBackColor = true;
            this.deleteButton.Click += new System.EventHandler(this.DeleteButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(31, 103);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(123, 20);
            this.label2.TabIndex = 8;
            this.label2.Text = "Selected timers:";
            // 
            // createButton
            // 
            this.createButton.Location = new System.Drawing.Point(236, 401);
            this.createButton.Name = "createButton";
            this.createButton.Size = new System.Drawing.Size(75, 30);
            this.createButton.TabIndex = 9;
            this.createButton.Text = "Create";
            this.createButton.UseVisualStyleBackColor = true;
            this.createButton.Click += new System.EventHandler(this.CreateButton_Click);
            // 
            // CustomTimersCreationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(353, 454);
            this.Controls.Add(this.createButton);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.deleteButton);
            this.Controls.Add(this.downButton);
            this.Controls.Add(this.upButton);
            this.Controls.Add(this.timersListBox);
            this.Controls.Add(this.timerAddButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.timerSelectComboBox);
            this.Controls.Add(this.arrivalTimerCheckbox);
            this.Name = "CustomTimersCreationForm";
            this.Text = "Create Custom Timer Form";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.CustomTimersCreationForm_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox arrivalTimerCheckbox;
        private System.Windows.Forms.ComboBox timerSelectComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button timerAddButton;
        private System.Windows.Forms.ListBox timersListBox;
        private System.Windows.Forms.Button upButton;
        private System.Windows.Forms.Button downButton;
        private System.Windows.Forms.Button deleteButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button createButton;
    }
}