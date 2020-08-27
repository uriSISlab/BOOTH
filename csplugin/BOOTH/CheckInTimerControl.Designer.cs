namespace BOOTH
{
    partial class CheckInTimerControl
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
            this.heading = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.textbox = new System.Windows.Forms.TextBox();
            this.clearButton = new System.Windows.Forms.Button();
            this.startButton = new System.Windows.Forms.Button();
            this.stopButton = new System.Windows.Forms.Button();
            this.vbmButton = new System.Windows.Forms.Button();
            this.startProvButton = new System.Windows.Forms.Button();
            this.endProvButton = new System.Windows.Forms.Button();
            this.undoLastButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // heading
            // 
            this.heading.AutoSize = true;
            this.heading.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.heading.Location = new System.Drawing.Point(38, 0);
            this.heading.Name = "heading";
            this.heading.Size = new System.Drawing.Size(204, 32);
            this.heading.TabIndex = 0;
            this.heading.Text = "Check In Timer";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::BOOTH.Properties.Resources.PollPad;
            this.pictureBox1.InitialImage = global::BOOTH.Properties.Resources.PollPad;
            this.pictureBox1.Location = new System.Drawing.Point(12, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(256, 196);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // textbox
            // 
            this.textbox.Location = new System.Drawing.Point(12, 237);
            this.textbox.Name = "textbox";
            this.textbox.Size = new System.Drawing.Size(175, 26);
            this.textbox.TabIndex = 2;
            // 
            // clearButton
            // 
            this.clearButton.Location = new System.Drawing.Point(193, 237);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(75, 26);
            this.clearButton.TabIndex = 3;
            this.clearButton.Text = "Clear";
            this.clearButton.UseVisualStyleBackColor = true;
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(12, 282);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(119, 41);
            this.startButton.TabIndex = 4;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            // 
            // stopButton
            // 
            this.stopButton.Location = new System.Drawing.Point(137, 282);
            this.stopButton.Name = "stopButton";
            this.stopButton.Size = new System.Drawing.Size(131, 41);
            this.stopButton.TabIndex = 5;
            this.stopButton.Text = "Stop";
            this.stopButton.UseVisualStyleBackColor = true;
            // 
            // vbmButton
            // 
            this.vbmButton.Location = new System.Drawing.Point(12, 329);
            this.vbmButton.Name = "vbmButton";
            this.vbmButton.Size = new System.Drawing.Size(61, 36);
            this.vbmButton.TabIndex = 6;
            this.vbmButton.Text = "VBM";
            this.vbmButton.UseVisualStyleBackColor = true;
            this.vbmButton.Click += new System.EventHandler(this.VbmButton_Click);
            // 
            // startProvButton
            // 
            this.startProvButton.Location = new System.Drawing.Point(79, 329);
            this.startProvButton.Name = "startProvButton";
            this.startProvButton.Size = new System.Drawing.Size(94, 36);
            this.startProvButton.TabIndex = 7;
            this.startProvButton.Text = "Start Prov";
            this.startProvButton.UseVisualStyleBackColor = true;
            this.startProvButton.Click += new System.EventHandler(this.StartProvButton_Click);
            // 
            // endProvButton
            // 
            this.endProvButton.Location = new System.Drawing.Point(179, 329);
            this.endProvButton.Name = "endProvButton";
            this.endProvButton.Size = new System.Drawing.Size(89, 36);
            this.endProvButton.TabIndex = 8;
            this.endProvButton.Text = "End Prov";
            this.endProvButton.UseVisualStyleBackColor = true;
            this.endProvButton.Click += new System.EventHandler(this.EndProvButton_Click);
            // 
            // undoLastButton
            // 
            this.undoLastButton.Location = new System.Drawing.Point(54, 382);
            this.undoLastButton.Name = "undoLastButton";
            this.undoLastButton.Size = new System.Drawing.Size(160, 35);
            this.undoLastButton.TabIndex = 9;
            this.undoLastButton.Text = "Undo Last";
            this.undoLastButton.UseVisualStyleBackColor = true;
            this.undoLastButton.Click += new System.EventHandler(this.UndoLastButton_Click);
            // 
            // CheckInTimerControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.undoLastButton);
            this.Controls.Add(this.endProvButton);
            this.Controls.Add(this.startProvButton);
            this.Controls.Add(this.vbmButton);
            this.Controls.Add(this.stopButton);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.textbox);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.heading);
            this.Name = "CheckInTimerControl";
            this.Size = new System.Drawing.Size(282, 450);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label heading;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox textbox;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.Button stopButton;
        private System.Windows.Forms.Button vbmButton;
        private System.Windows.Forms.Button startProvButton;
        private System.Windows.Forms.Button endProvButton;
        private System.Windows.Forms.Button undoLastButton;
    }
}
