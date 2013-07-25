namespace TranscriptFromGrades
{
    partial class MainForm
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
            this.openFileButton = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.excelFileTextBox = new System.Windows.Forms.TextBox();
            this.generateTranscriptButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.rtfTemplateFileTextBox = new System.Windows.Forms.TextBox();
            this.templateFileButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileButton
            // 
            this.openFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.openFileButton.Location = new System.Drawing.Point(396, 60);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(30, 23);
            this.openFileButton.TabIndex = 0;
            this.openFileButton.Text = "...";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.openFileButton_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            // 
            // excelFileTextBox
            // 
            this.excelFileTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.excelFileTextBox.Location = new System.Drawing.Point(13, 60);
            this.excelFileTextBox.Name = "excelFileTextBox";
            this.excelFileTextBox.Size = new System.Drawing.Size(377, 20);
            this.excelFileTextBox.TabIndex = 1;
            // 
            // generateTranscriptButton
            // 
            this.generateTranscriptButton.Location = new System.Drawing.Point(13, 149);
            this.generateTranscriptButton.Name = "generateTranscriptButton";
            this.generateTranscriptButton.Size = new System.Drawing.Size(75, 23);
            this.generateTranscriptButton.TabIndex = 2;
            this.generateTranscriptButton.Text = "Generate Transcript";
            this.generateTranscriptButton.UseVisualStyleBackColor = true;
            this.generateTranscriptButton.Click += new System.EventHandler(this.generateTranscriptButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Grade File:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 98);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Transcript Template:";
            // 
            // templateFileTextBox
            // 
            this.rtfTemplateFileTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtfTemplateFileTextBox.Location = new System.Drawing.Point(19, 115);
            this.rtfTemplateFileTextBox.Name = "templateFileTextBox";
            this.rtfTemplateFileTextBox.Size = new System.Drawing.Size(371, 20);
            this.rtfTemplateFileTextBox.TabIndex = 5;
            // 
            // templateFileButton
            // 
            this.templateFileButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.templateFileButton.Location = new System.Drawing.Point(397, 115);
            this.templateFileButton.Name = "templateFileButton";
            this.templateFileButton.Size = new System.Drawing.Size(29, 23);
            this.templateFileButton.TabIndex = 6;
            this.templateFileButton.Text = "...";
            this.templateFileButton.UseVisualStyleBackColor = true;
            this.templateFileButton.Click += new System.EventHandler(this.templateFileButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(438, 198);
            this.Controls.Add(this.templateFileButton);
            this.Controls.Add(this.rtfTemplateFileTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.generateTranscriptButton);
            this.Controls.Add(this.excelFileTextBox);
            this.Controls.Add(this.openFileButton);
            this.Name = "MainForm";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TextBox excelFileTextBox;
        private System.Windows.Forms.Button generateTranscriptButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox rtfTemplateFileTextBox;
        private System.Windows.Forms.Button templateFileButton;
    }
}

