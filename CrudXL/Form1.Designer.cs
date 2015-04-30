namespace CrudXL
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.changeRequest = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataCentre = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.redirect301 = new System.Windows.Forms.RadioButton();
            this.redirect302 = new System.Windows.Forms.RadioButton();
            this.statusOk = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.button1.Location = new System.Drawing.Point(16, 209);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Create";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.ForeColor = System.Drawing.SystemColors.Desktop;
            this.button2.Location = new System.Drawing.Point(114, 209);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "Read";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.DarkRed;
            this.button3.Location = new System.Drawing.Point(305, 209);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(75, 23);
            this.button3.TabIndex = 2;
            this.button3.Text = "Exit";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // changeRequest
            // 
            this.changeRequest.AcceptsTab = true;
            this.changeRequest.AccessibleDescription = "CR Number Input Field";
            this.changeRequest.AccessibleName = "CRInput";
            this.changeRequest.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.changeRequest.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.changeRequest.Location = new System.Drawing.Point(16, 25);
            this.changeRequest.Name = "changeRequest";
            this.changeRequest.Size = new System.Drawing.Size(228, 20);
            this.changeRequest.TabIndex = 3;
            this.changeRequest.UseWaitCursor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 48);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(346, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Please enter the CR/Ticket # above, this will be prefixed to the filename";
            // 
            // dataCentre
            // 
            this.dataCentre.AcceptsTab = true;
            this.dataCentre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.dataCentre.Location = new System.Drawing.Point(16, 79);
            this.dataCentre.Name = "dataCentre";
            this.dataCentre.Size = new System.Drawing.Size(100, 20);
            this.dataCentre.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 106);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(243, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Please enter the 3 character datacentre reference";
            // 
            // redirect301
            // 
            this.redirect301.AutoSize = true;
            this.redirect301.Checked = true;
            this.redirect301.Location = new System.Drawing.Point(19, 137);
            this.redirect301.Name = "redirect301";
            this.redirect301.Size = new System.Drawing.Size(85, 17);
            this.redirect301.TabIndex = 7;
            this.redirect301.TabStop = true;
            this.redirect301.Text = "301 [Moved]";
            this.redirect301.UseVisualStyleBackColor = true;
            this.redirect301.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // redirect302
            // 
            this.redirect302.AutoSize = true;
            this.redirect302.Location = new System.Drawing.Point(19, 160);
            this.redirect302.Name = "redirect302";
            this.redirect302.Size = new System.Drawing.Size(92, 17);
            this.redirect302.TabIndex = 8;
            this.redirect302.TabStop = true;
            this.redirect302.Text = "302 [Redirect]";
            this.redirect302.UseVisualStyleBackColor = true;
            // 
            // statusOk
            // 
            this.statusOk.AutoSize = true;
            this.statusOk.Location = new System.Drawing.Point(159, 137);
            this.statusOk.Name = "statusOk";
            this.statusOk.Size = new System.Drawing.Size(66, 17);
            this.statusOk.TabIndex = 9;
            this.statusOk.TabStop = true;
            this.statusOk.Text = "200 [Ok]";
            this.statusOk.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AcceptButton = this.button1;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.button2;
            this.ClientSize = new System.Drawing.Size(446, 246);
            this.Controls.Add(this.statusOk);
            this.Controls.Add(this.redirect302);
            this.Controls.Add(this.redirect301);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataCentre);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.changeRequest);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.RightToLeftLayout = true;
            this.Text = "Find the source excel data file";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox changeRequest;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox dataCentre;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton redirect301;
        private System.Windows.Forms.RadioButton redirect302;
        private System.Windows.Forms.RadioButton statusOk;
    }
}

