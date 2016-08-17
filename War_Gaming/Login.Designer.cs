namespace WindowsFormsApplication2
{
    partial class Login
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Login));
            this.panel1 = new System.Windows.Forms.Panel();
            this.pictureLogin = new System.Windows.Forms.PictureBox();
            this.pictureBox10 = new System.Windows.Forms.PictureBox();
            this.textPassword = new System.Windows.Forms.TextBox();
            this.textUser = new System.Windows.Forms.TextBox();
            this.label37 = new System.Windows.Forms.Label();
            this.label38 = new System.Windows.Forms.Label();
            this.label39 = new System.Windows.Forms.Label();
            this.butLogin = new System.Windows.Forms.Button();
            this.butExit = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogin)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox10)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightSkyBlue;
            this.panel1.Controls.Add(this.pictureLogin);
            this.panel1.Controls.Add(this.pictureBox10);
            this.panel1.Location = new System.Drawing.Point(12, 39);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(241, 251);
            this.panel1.TabIndex = 44;
            // 
            // pictureLogin
            // 
            this.pictureLogin.Image = ((System.Drawing.Image)(resources.GetObject("pictureLogin.Image")));
            this.pictureLogin.InitialImage = ((System.Drawing.Image)(resources.GetObject("pictureLogin.InitialImage")));
            this.pictureLogin.Location = new System.Drawing.Point(69, 67);
            this.pictureLogin.Name = "pictureLogin";
            this.pictureLogin.Size = new System.Drawing.Size(101, 104);
            this.pictureLogin.TabIndex = 45;
            this.pictureLogin.TabStop = false;
            // 
            // pictureBox10
            // 
            this.pictureBox10.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox10.Image")));
            this.pictureBox10.Location = new System.Drawing.Point(3, 208);
            this.pictureBox10.Name = "pictureBox10";
            this.pictureBox10.Size = new System.Drawing.Size(232, 37);
            this.pictureBox10.TabIndex = 44;
            this.pictureBox10.TabStop = false;
            // 
            // textPassword
            // 
            this.textPassword.Location = new System.Drawing.Point(397, 158);
            this.textPassword.Name = "textPassword";
            this.textPassword.Size = new System.Drawing.Size(170, 20);
            this.textPassword.TabIndex = 49;
            this.textPassword.UseSystemPasswordChar = true;
            // 
            // textUser
            // 
            this.textUser.Location = new System.Drawing.Point(397, 101);
            this.textUser.Name = "textUser";
            this.textUser.Size = new System.Drawing.Size(170, 20);
            this.textUser.TabIndex = 48;
            // 
            // label37
            // 
            this.label37.AutoSize = true;
            this.label37.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label37.Location = new System.Drawing.Point(285, 158);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(91, 20);
            this.label37.TabIndex = 47;
            this.label37.Text = "Password:";
            // 
            // label38
            // 
            this.label38.AutoSize = true;
            this.label38.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label38.Location = new System.Drawing.Point(277, 101);
            this.label38.Name = "label38";
            this.label38.Size = new System.Drawing.Size(103, 20);
            this.label38.TabIndex = 46;
            this.label38.Text = "User Name:";
            // 
            // label39
            // 
            this.label39.AutoSize = true;
            this.label39.Font = new System.Drawing.Font("Microsoft Sans Serif", 20F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label39.Location = new System.Drawing.Point(317, 39);
            this.label39.Name = "label39";
            this.label39.Size = new System.Drawing.Size(261, 31);
            this.label39.TabIndex = 45;
            this.label39.Text = "War Gaming Login ";
            // 
            // butLogin
            // 
            this.butLogin.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butLogin.Location = new System.Drawing.Point(485, 255);
            this.butLogin.Name = "butLogin";
            this.butLogin.Size = new System.Drawing.Size(82, 35);
            this.butLogin.TabIndex = 51;
            this.butLogin.Text = "Login";
            this.butLogin.UseVisualStyleBackColor = true;
            this.butLogin.Click += new System.EventHandler(this.butLogin_Click_1);
            // 
            // butExit
            // 
            this.butExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.butExit.Location = new System.Drawing.Point(347, 253);
            this.butExit.Name = "butExit";
            this.butExit.Size = new System.Drawing.Size(82, 35);
            this.butExit.TabIndex = 50;
            this.butExit.Text = "Exit";
            this.butExit.UseVisualStyleBackColor = true;
            this.butExit.Click += new System.EventHandler(this.butExit_Click_1);
            // 
            // Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.InactiveCaption;
            this.ClientSize = new System.Drawing.Size(583, 301);
            this.Controls.Add(this.butLogin);
            this.Controls.Add(this.butExit);
            this.Controls.Add(this.textPassword);
            this.Controls.Add(this.textUser);
            this.Controls.Add(this.label37);
            this.Controls.Add(this.label38);
            this.Controls.Add(this.label39);
            this.Controls.Add(this.panel1);
            this.Name = "Login";
            this.Text = "Login";
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureLogin)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox10)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureLogin;
        private System.Windows.Forms.PictureBox pictureBox10;
        private System.Windows.Forms.TextBox textPassword;
        private System.Windows.Forms.TextBox textUser;
        private System.Windows.Forms.Label label37;
        private System.Windows.Forms.Label label38;
        private System.Windows.Forms.Label label39;
        private System.Windows.Forms.Button butLogin;
        private System.Windows.Forms.Button butExit;
    }
}