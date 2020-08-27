namespace BOOTH
{
    partial class VotingBoothTimerControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.headingLabel = new System.Windows.Forms.Label();
            this.textbox = new System.Windows.Forms.TextBox();
            this.clearButton = new System.Windows.Forms.Button();
            this.startButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.undoLastButton = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // headingLabel
            // 
            this.headingLabel.AutoSize = true;
            this.headingLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headingLabel.Location = new System.Drawing.Point(12, 0);
            this.headingLabel.Name = "headingLabel";
            this.headingLabel.Size = new System.Drawing.Size(258, 32);
            this.headingLabel.TabIndex = 0;
            this.headingLabel.Text = "Voting Booth Timer";
            // 
            // textbox
            // 
            this.textbox.Location = new System.Drawing.Point(14, 291);
            this.textbox.Name = "textbox";
            this.textbox.Size = new System.Drawing.Size(175, 26);
            this.textbox.TabIndex = 2;
            // 
            // clearButton
            // 
            this.clearButton.Location = new System.Drawing.Point(195, 291);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(75, 26);
            this.clearButton.TabIndex = 3;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(14, 327);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(119, 41);
            this.startButton.TabIndex = 4;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            // 
            // stopButton
            // 
            this.stopButton.Location = new System.Drawing.Point(139, 327);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(131, 41);
            this.stopButton.TabIndex = 5;
            this.stopButton.Text = "Stop";
            this.stopButton.UseVisualStyleBackColor = true;
            // 
            // undoLastButton
            // 
            this.undoLastButton.Location = new System.Drawing.Point(54, 382);
            this.undoLastButton.Name = "undoLastButton";
            this.undoLastButton.Size = new System.Drawing.Size(160, 35);
            this.undoLastButton.TabIndex = 6;
            this.undoLastButton.Text = "Undo Last";
            this.undoLastButton.UseVisualStyleBackColor = true;
            this.undoLastButton.Click += new System.EventHandler(this.UndoLastButton_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::BOOTH.Properties.Resources.VotingBooth_resized;
            this.pictureBox1.Location = new System.Drawing.Point(39, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(201, 250);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // VotingBoothTimerControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.undoLastButton);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.textbox);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.headingLabel);
            this.Name = "VotingBoothTimerControl";
            this.Size = new System.Drawing.Size(282, 450);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label headingLabel;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textbox;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.Button undoLastButton;
    }
}
