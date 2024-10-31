namespace tekyaz7
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnSelectFile = new Button();
            btnConvert = new Button();
            txtFilePath = new TextBox();
            progressBar = new ProgressBar();
            rtbLogs = new RichTextBox();

            SuspendLayout();
            // 
            // btnSelectFile
            // 
            btnSelectFile.Location = new Point(12, 12);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(120, 30);
            btnSelectFile.TabIndex = 0;
            btnSelectFile.Text = "Dosya Seç";
            btnSelectFile.UseVisualStyleBackColor = true;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // btnConvert
            // 
            btnConvert.Location = new Point(180, 54);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new Size(120, 30);
            btnConvert.TabIndex = 1;
            btnConvert.Text = "Dönüştür";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += btnConvert_Click;
            // 
            // txtFilePath
            // 
            txtFilePath.Location = new Point(138, 16);
            txtFilePath.Name = "txtFilePath";
            txtFilePath.Size = new Size(330, 23);
            txtFilePath.TabIndex = 2;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 96);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(456, 23);
            progressBar.TabIndex = 3;
            // 
            // rtbLogs
            // 
            rtbLogs.Location = new Point(12, 125);
            rtbLogs.Name = "rtbLogs";
            rtbLogs.Size = new Size(456, 200);
            rtbLogs.TabIndex = 4;
            rtbLogs.Text = "";
            rtbLogs.ReadOnly = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(480, 340);
            Controls.Add(rtbLogs);
            Controls.Add(progressBar);
            Controls.Add(txtFilePath);
            Controls.Add(btnConvert);
            Controls.Add(btnSelectFile);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "SolidWorks DWG Dönüştürücü";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnSelectFile;
        private Button btnConvert;
        private TextBox txtFilePath;
        private ProgressBar progressBar;
        private RichTextBox rtbLogs;
    }
}