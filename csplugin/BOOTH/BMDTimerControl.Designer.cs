namespace BOOTH
{
    partial class BMDTimerControl
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
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.clearButton = new System.Windows.Forms.Button();
            this.textbox = new System.Windows.Forms.TextBox();
            this.startButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.helpButton = new System.Windows.Forms.Button();
            this.undoLastButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // headingLabel
            // 
            this.headingLabel.AutoSize = true;
            this.headingLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.headingLabel.Location = new System.Drawing.Point(60, 0);
            this.headingLabel.Name = "headingLabel";
            this.headingLabel.Size = new System.Drawing.Size(156, 32);
            this.headingLabel.TabIndex = 0;
            this.headingLabel.Text = "BMD Timer";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::BOOTH.Properties.Resources.BMDLA_scaled;
            this.pictureBox1.Location = new System.Drawing.Point(30, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(214, 252);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // clearButton
            // 
            this.clearButton.Location = new System.Drawing.Point(180, 293);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(64, 26);
            this.clearButton.TabIndex = 2;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            // 
            // textbox
            // 
            this.textbox.Location = new System.Drawing.Point(30, 293);
            this.textbox.Name = "textbox";
            this.textbox.Size = new System.Drawing.Size(144, 26);
            this.textbox.TabIndex = 3;
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(16, 336);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(119, 41);
            this.startButton.TabIndex = 4;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // stopButton
            // 
            this.stopButton.Location = new System.Drawing.Point(141, 336);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(119, 41);
            this.stopButton.TabIndex = 5;
            this.stopButton.Text = "Stop";
            this.stopButton.UseVisualStyleBackColor = true;
            this.stopButton.Click += new System.EventHandler(this.StopButton_Click);
            // 
            // helpButton
            // 
            this.helpButton.Location = new System.Drawing.Point(16, 392);
            this.helpButton.Name = "helpButton";
            this.helpButton.Size = new System.Drawing.Size(86, 39);
            this.helpButton.TabIndex = 6;
            this.helpButton.Text = "Help";
            this.helpButton.UseVisualStyleBackColor = true;
            this.helpButton.Click += new System.EventHandler(this.HelpButton_Click);
            // 
            // undoLastButton
            // 
            this.undoLastButton.Location = new System.Drawing.Point(141, 392);
            this.undoLastButton.Name = "undoLastButton";
            this.undoLastButton.Size = new System.Drawing.Size(119, 39);
            this.undoLastButton.TabIndex = 7;
            this.undoLastButton.Text = "Undo Last";
            this.undoLastButton.UseVisualStyleBackColor = true;
            this.undoLastButton.Click += new System.EventHandler(this.UndoLastButton_Click);
            // 
            // BMDTimerControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.undoLastButton);
            this.Controls.Add(this.helpButton);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.textbox);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.headingLabel);
            this.Name = "BMDTimerControl";
            this.Size = new System.Drawing.Size(282, 450);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label headingLabel;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.TextBox textbox;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.Button helpButton;
        private System.Windows.Forms.Button undoLastButton;
    }
}
