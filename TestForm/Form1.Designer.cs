namespace TestForm
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
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnRetrieve = new System.Windows.Forms.Button();
            this.txtCaseId = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCreate
            // 
            this.btnCreate.Location = new System.Drawing.Point(220, 67);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(150, 35);
            this.btnCreate.TabIndex = 1;
            this.btnCreate.Text = "Create Case";
            this.btnCreate.UseVisualStyleBackColor = true;
            this.btnCreate.Click += new System.EventHandler(this.btnCreate_Click);
            // 
            // btnRetrieve
            // 
            this.btnRetrieve.Location = new System.Drawing.Point(220, 28);
            this.btnRetrieve.Name = "btnRetrieve";
            this.btnRetrieve.Size = new System.Drawing.Size(148, 33);
            this.btnRetrieve.TabIndex = 2;
            this.btnRetrieve.Text = "Retrieve Case";
            this.btnRetrieve.UseVisualStyleBackColor = true;
            this.btnRetrieve.Click += new System.EventHandler(this.btnRetrieve_Click);
            // 
            // txtCaseId
            // 
            this.txtCaseId.Location = new System.Drawing.Point(66, 28);
            this.txtCaseId.Name = "txtCaseId";
            this.txtCaseId.Size = new System.Drawing.Size(128, 20);
            this.txtCaseId.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(218, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(150, 35);
            this.button1.TabIndex = 4;
            this.button1.Text = "Retrieve FixFor";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(523, 180);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtCaseId);
            this.Controls.Add(this.btnRetrieve);
            this.Controls.Add(this.btnCreate);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreate;
        private System.Windows.Forms.Button btnRetrieve;
        private System.Windows.Forms.TextBox txtCaseId;
        private System.Windows.Forms.Button button1;
    }
}

