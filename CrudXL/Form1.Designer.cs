namespace CrudXL
{
    partial class BugBareSmoke
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
            this.CreateButton = new System.Windows.Forms.Button();
            this.ReadButton = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            this.readFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.ChangeRequest = new System.Windows.Forms.TextBox();
            this.TixRefLabel = new System.Windows.Forms.Label();
            this.DataCentre = new System.Windows.Forms.TextBox();
            this.DCRefLabel = new System.Windows.Forms.Label();
            this.Status301 = new System.Windows.Forms.RadioButton();
            this.Status302 = new System.Windows.Forms.RadioButton();
            this.StatusOk = new System.Windows.Forms.RadioButton();
            this.InputPageUrl = new System.Windows.Forms.TextBox();
            this.TestGenHelpTextLabel = new System.Windows.Forms.Label();
            this.TestButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // createButton
            // 
            this.CreateButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CreateButton.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.CreateButton.Location = new System.Drawing.Point(15, 79);
            this.CreateButton.Name = "createButton";
            this.CreateButton.Size = new System.Drawing.Size(75, 23);
            this.CreateButton.TabIndex = 0;
            this.CreateButton.Text = "Create";
            this.CreateButton.UseVisualStyleBackColor = true;
            this.CreateButton.Click += new System.EventHandler(this.CreateButton_Click);
            // 
            // Read Button
            // 
            this.ReadButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.ReadButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ReadButton.ForeColor = System.Drawing.SystemColors.Desktop;
            this.ReadButton.Location = new System.Drawing.Point(12, 303);
            this.ReadButton.Name = "readButton";
            this.ReadButton.Size = new System.Drawing.Size(75, 23);
            this.ReadButton.TabIndex = 1;
            this.ReadButton.Text = "Read";
            this.ReadButton.UseVisualStyleBackColor = true;
            this.ReadButton.Click += new System.EventHandler(this.ReadButton_Click);
            // 
            // Exit Button
            // 
            this.Exit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Exit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Exit.ForeColor = System.Drawing.Color.DarkRed;
            this.Exit.Location = new System.Drawing.Point(314, 303);
            this.Exit.Name = "exit";
            this.Exit.Size = new System.Drawing.Size(75, 23);
            this.Exit.TabIndex = 2;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_BugBareSmoke);
            // 
            // CR Input Text Field
            // 
            this.ChangeRequest.AcceptsTab = true;
            this.ChangeRequest.AccessibleDescription = "CR Number Input Field";
            this.ChangeRequest.AccessibleName = "CRInput";
            this.ChangeRequest.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ChangeRequest.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.ChangeRequest.Location = new System.Drawing.Point(12, 134);
            this.ChangeRequest.Name = "changeRequest";
            this.ChangeRequest.Size = new System.Drawing.Size(228, 20);
            this.ChangeRequest.TabIndex = 3;
            this.ChangeRequest.UseWaitCursor = true;
            // 
            // Ticket Reference Label
            // 
            this.TixRefLabel.AutoSize = true;
            this.TixRefLabel.Location = new System.Drawing.Point(9, 157);
            this.TixRefLabel.Name = "label1";
            this.TixRefLabel.Size = new System.Drawing.Size(346, 13);
            this.TixRefLabel.TabIndex = 4;
            this.TixRefLabel.Text = "Please enter the CR/Ticket # above, this will be prefixed to the filename";
            // 
            // Data Centre Input Text Field
            // 
            this.DataCentre.AcceptsTab = true;
            this.DataCentre.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.DataCentre.Location = new System.Drawing.Point(12, 188);
            this.DataCentre.Name = "dataCentre";
            this.DataCentre.Size = new System.Drawing.Size(100, 20);
            this.DataCentre.TabIndex = 5;
            // 
            // DataCentre Reference Label
            // 
            this.DCRefLabel.AutoSize = true;
            this.DCRefLabel.Location = new System.Drawing.Point(12, 215);
            this.DCRefLabel.Name = "label2";
            this.DCRefLabel.Size = new System.Drawing.Size(243, 13);
            this.DCRefLabel.TabIndex = 6;
            this.DCRefLabel.Text = "Please enter the 3 character datacentre reference";
            // 
            // status301
            // 
            this.Status301.AutoSize = true;
            this.Status301.Checked = true;
            this.Status301.Location = new System.Drawing.Point(15, 246);
            this.Status301.Name = "status301";
            this.Status301.Size = new System.Drawing.Size(85, 17);
            this.Status301.TabIndex = 7;
            this.Status301.TabStop = true;
            this.Status301.Text = "301 [Moved]";
            this.Status301.UseVisualStyleBackColor = true;
            this.Status301.CheckedChanged += new System.EventHandler(this.RedirectMode_CheckBox_Click);
            // 
            // status302
            // 
            this.Status302.AutoSize = true;
            this.Status302.Location = new System.Drawing.Point(15, 269);
            this.Status302.Name = "status302";
            this.Status302.Size = new System.Drawing.Size(92, 17);
            this.Status302.TabIndex = 8;
            this.Status302.TabStop = true;
            this.Status302.Text = "302 [Redirect]";
            this.Status302.UseVisualStyleBackColor = true;
            // 
            // statusOk
            // 
            this.StatusOk.AutoSize = true;
            this.StatusOk.Location = new System.Drawing.Point(155, 246);
            this.StatusOk.Name = "statusOk";
            this.StatusOk.Size = new System.Drawing.Size(66, 17);
            this.StatusOk.TabIndex = 9;
            this.StatusOk.TabStop = true;
            this.StatusOk.Text = "200 [Ok]";
            this.StatusOk.UseVisualStyleBackColor = true;
            // 
            // inputPageUrl
            // 
            this.InputPageUrl.Location = new System.Drawing.Point(15, 25);
            this.InputPageUrl.Name = "inputPageUrl";
            this.InputPageUrl.Size = new System.Drawing.Size(374, 20);
            this.InputPageUrl.TabIndex = 10;
            // 
            // Test Generation Help Text
            // 
            this.TestGenHelpTextLabel.AutoSize = true;
            this.TestGenHelpTextLabel.Location = new System.Drawing.Point(12, 52);
            this.TestGenHelpTextLabel.Name = "label3";
            this.TestGenHelpTextLabel.Size = new System.Drawing.Size(412, 13);
            this.TestGenHelpTextLabel.TabIndex = 11;
            this.TestGenHelpTextLabel.Text = "To create a spreadsheet that contains all links on any given page, enter the URL " +
    "here";
            this.TestGenHelpTextLabel.Click += new System.EventHandler(this.TestBuilder_Description_Click);
            // 
            // TestButton Designer References
            // 
            this.TestButton.BackColor = System.Drawing.SystemColors.ScrollBar;
            this.TestButton.FlatAppearance.BorderSize = 2;
            this.TestButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.TestButton.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.TestButton.Location = new System.Drawing.Point(314, 79);
            this.TestButton.Name = "testButton";
            this.TestButton.Size = new System.Drawing.Size(75, 23);
            this.TestButton.TabIndex = 12;
            this.TestButton.Text = "Test";
            this.TestButton.UseVisualStyleBackColor = false;
            this.TestButton.Click += new System.EventHandler(this.TestButton_Click);
            // 
            // Form1
            // 
            this.AcceptButton = this.CreateButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Exit;
            this.ClientSize = new System.Drawing.Size(429, 353);
            this.Controls.Add(this.TestButton);
            this.Controls.Add(this.TestGenHelpTextLabel);
            this.Controls.Add(this.InputPageUrl);
            this.Controls.Add(this.StatusOk);
            this.Controls.Add(this.Status302);
            this.Controls.Add(this.Status301);
            this.Controls.Add(this.DCRefLabel);
            this.Controls.Add(this.DataCentre);
            this.Controls.Add(this.TixRefLabel);
            this.Controls.Add(this.ChangeRequest);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.ReadButton);
            this.Controls.Add(this.CreateButton);
            this.Name = "Form1";
            this.RightToLeftLayout = true;
            this.Text = "Find the source excel data file";
            this.Load += new System.EventHandler(this.BugBareSmoke_LoadFile);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CreateButton;
        private System.Windows.Forms.Button ReadButton;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.OpenFileDialog readFileDialog;
        private System.Windows.Forms.TextBox ChangeRequest;
        private System.Windows.Forms.Label TixRefLabel;
        private System.Windows.Forms.TextBox DataCentre;
        private System.Windows.Forms.Label DCRefLabel;
        private System.Windows.Forms.RadioButton Status301;
        private System.Windows.Forms.RadioButton Status302;
        private System.Windows.Forms.RadioButton StatusOk;
        private System.Windows.Forms.TextBox InputPageUrl;
        private System.Windows.Forms.Label TestGenHelpTextLabel;
        private System.Windows.Forms.Button TestButton;
    }
}

