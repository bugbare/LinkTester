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
            this.createButton = new System.Windows.Forms.Button();
            this.readButton = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.changeRequest = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataCentre = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.status301 = new System.Windows.Forms.RadioButton();
            this.status302 = new System.Windows.Forms.RadioButton();
            this.statusOk = new System.Windows.Forms.RadioButton();
            this.inputPageUrl = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // createButton
            // 
            this.createButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.createButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.createButton.Location = new System.Drawing.Point(15, 79);
            this.createButton.Name = "createButton";
            this.createButton.Size = new System.Drawing.Size(75, 23);
            this.createButton.TabIndex = 0;
            this.createButton.Text = "Create";
            this.createButton.UseVisualStyleBackColor = true;
            this.createButton.Click += new System.EventHandler(this.createButton_Click);
            // 
            // readButton
            // 
            this.readButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.readButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.readButton.ForeColor = System.Drawing.SystemColors.Desktop;
            this.readButton.Location = new System.Drawing.Point(12, 303);
            this.readButton.Name = "readButton";
            this.readButton.Size = new System.Drawing.Size(75, 23);
            this.readButton.TabIndex = 1;
            this.readButton.Text = "Read";
            this.readButton.UseVisualStyleBackColor = true;
            this.readButton.Click += new System.EventHandler(this.readButton_Click);
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.ForeColor = System.Drawing.Color.DarkRed;
            this.button3.Location = new System.Drawing.Point(349, 303);
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
            this.changeRequest.Location = new System.Drawing.Point(12, 134);
            this.changeRequest.Name = "changeRequest";
            this.changeRequest.Size = new System.Drawing.Size(228, 20);
            this.changeRequest.TabIndex = 3;
            this.changeRequest.UseWaitCursor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 157);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(346, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Please enter the CR/Ticket # above, this will be prefixed to the filename";
            // 
            // dataCentre
            // 
            this.dataCentre.AcceptsTab = true;
            this.dataCentre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.dataCentre.Location = new System.Drawing.Point(12, 188);
            this.dataCentre.Name = "dataCentre";
            this.dataCentre.Size = new System.Drawing.Size(100, 20);
            this.dataCentre.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 215);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(243, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Please enter the 3 character datacentre reference";
            // 
            // status301
            // 
            this.status301.AutoSize = true;
            this.status301.Checked = true;
            this.status301.Location = new System.Drawing.Point(15, 246);
            this.status301.Name = "status301";
            this.status301.Size = new System.Drawing.Size(85, 17);
            this.status301.TabIndex = 7;
            this.status301.TabStop = true;
            this.status301.Text = "301 [Moved]";
            this.status301.UseVisualStyleBackColor = true;
            this.status301.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // status302
            // 
            this.status302.AutoSize = true;
            this.status302.Location = new System.Drawing.Point(15, 269);
            this.status302.Name = "status302";
            this.status302.Size = new System.Drawing.Size(92, 17);
            this.status302.TabIndex = 8;
            this.status302.TabStop = true;
            this.status302.Text = "302 [Redirect]";
            this.status302.UseVisualStyleBackColor = true;
            // 
            // statusOk
            // 
            this.statusOk.AutoSize = true;
            this.statusOk.Location = new System.Drawing.Point(155, 246);
            this.statusOk.Name = "statusOk";
            this.statusOk.Size = new System.Drawing.Size(66, 17);
            this.statusOk.TabIndex = 9;
            this.statusOk.TabStop = true;
            this.statusOk.Text = "200 [Ok]";
            this.statusOk.UseVisualStyleBackColor = true;
            // 
            // inputPageUrl
            // 
            this.inputPageUrl.Location = new System.Drawing.Point(15, 25);
            this.inputPageUrl.Name = "inputPageUrl";
            this.inputPageUrl.Size = new System.Drawing.Size(374, 20);
            this.inputPageUrl.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 52);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(412, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "To create a spreadsheet that contains all links on any given page, enter the URL " +
    "here";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // Form1
            // 
            this.AcceptButton = this.createButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.readButton;
            this.ClientSize = new System.Drawing.Size(446, 353);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.inputPageUrl);
            this.Controls.Add(this.statusOk);
            this.Controls.Add(this.status302);
            this.Controls.Add(this.status301);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataCentre);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.changeRequest);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.readButton);
            this.Controls.Add(this.createButton);
            this.Name = "Form1";
            this.RightToLeftLayout = true;
            this.Text = "Find the source excel data file";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button createButton;
        private System.Windows.Forms.Button readButton;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox changeRequest;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox dataCentre;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton status301;
        private System.Windows.Forms.RadioButton status302;
        private System.Windows.Forms.RadioButton statusOk;
        private System.Windows.Forms.TextBox inputPageUrl;
        private System.Windows.Forms.Label label3;
    }
}

