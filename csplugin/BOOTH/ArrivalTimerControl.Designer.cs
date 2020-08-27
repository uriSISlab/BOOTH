namespace BOOTH
{
    partial class ArrivalTimerControl
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
            this.arrivalButton = new System.Windows.Forms.Button();
            this.vbmArrivalButton = new System.Windows.Forms.Button();
            this.undoLastButton = new System.Windows.Forms.Button();
            this.arrivalCountLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // arrivalButton
            // 
            this.arrivalButton.Location = new System.Drawing.Point(19, 19);
            this.arrivalButton.Name = "arrivalButton";
            this.arrivalButton.Size = new System.Drawing.Size(107, 45);
            this.arrivalButton.TabIndex = 0;
            this.arrivalButton.Text = "Arrival";
            this.arrivalButton.UseVisualStyleBackColor = true;
            this.arrivalButton.Click += new System.EventHandler(this.ArrivalButton_Click);
            // 
            // vbmArrivalButton
            // 
            this.vbmArrivalButton.Location = new System.Drawing.Point(19, 70);
            this.vbmArrivalButton.Name = "vbmArrivalButton";
            this.vbmArrivalButton.Size = new System.Drawing.Size(107, 46);
            this.vbmArrivalButton.TabIndex = 1;
            this.vbmArrivalButton.Text = "VBM Arrival";
            this.vbmArrivalButton.UseVisualStyleBackColor = true;
            this.vbmArrivalButton.Click += new System.EventHandler(this.VbmArrivalButton_Click);
            // 
            // undoLastButton
            // 
            this.undoLastButton.Location = new System.Drawing.Point(19, 144);
            this.undoLastButton.Name = "undoLastButton";
            this.undoLastButton.Size = new System.Drawing.Size(107, 31);
            this.undoLastButton.TabIndex = 2;
            this.undoLastButton.Text = "Undo Last";
            this.undoLastButton.UseVisualStyleBackColor = true;
            this.undoLastButton.Click += new System.EventHandler(this.UndoLastButton_Click);
            // 
            // arrivalCountLabel
            // 
            this.arrivalCountLabel.AutoSize = true;
            this.arrivalCountLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.arrivalCountLabel.Location = new System.Drawing.Point(154, 44);
            this.arrivalCountLabel.Name = "arrivalCountLabel";
            this.arrivalCountLabel.Size = new System.Drawing.Size(35, 37);
            this.arrivalCountLabel.TabIndex = 3;
            this.arrivalCountLabel.Text = "0";
            // 
            // ArrivalTimerControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.arrivalCountLabel);
            this.Controls.Add(this.undoLastButton);
            this.Controls.Add(this.vbmArrivalButton);
            this.Controls.Add(this.arrivalButton);
            this.Name = "ArrivalTimerControl";
            this.Size = new System.Drawing.Size(225, 186);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button arrivalButton;
        private System.Windows.Forms.Button vbmArrivalButton;
        private System.Windows.Forms.Button undoLastButton;
        private System.Windows.Forms.Label arrivalCountLabel;
    }
}
