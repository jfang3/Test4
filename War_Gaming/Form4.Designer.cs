namespace WindowsFormsApplication2
{
    partial class Form4
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
            this.labelReceiver = new System.Windows.Forms.Label();
            this.labelAttach = new System.Windows.Forms.Label();
            this.labelSubject = new System.Windows.Forms.Label();
            this.labelBody = new System.Windows.Forms.Label();
            this.textReciever = new System.Windows.Forms.TextBox();
            this.textAttachment = new System.Windows.Forms.TextBox();
            this.textSubject = new System.Windows.Forms.TextBox();
            this.richBody = new System.Windows.Forms.RichTextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(743, 52);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(80, 28);
            this.button1.TabIndex = 0;
            this.button1.Text = "Browser";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // labelReceiver
            // 
            this.labelReceiver.AutoSize = true;
            this.labelReceiver.Location = new System.Drawing.Point(6, 15);
            this.labelReceiver.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelReceiver.Name = "labelReceiver";
            this.labelReceiver.Size = new System.Drawing.Size(133, 17);
            this.labelReceiver.TabIndex = 2;
            this.labelReceiver.Text = "Receiver\'s Email:";
            // 
            // labelAttach
            // 
            this.labelAttach.AutoSize = true;
            this.labelAttach.Location = new System.Drawing.Point(12, 60);
            this.labelAttach.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelAttach.Name = "labelAttach";
            this.labelAttach.Size = new System.Drawing.Size(127, 17);
            this.labelAttach.TabIndex = 3;
            this.labelAttach.Text = "Add Attachment:";
            // 
            // labelSubject
            // 
            this.labelSubject.AutoSize = true;
            this.labelSubject.Location = new System.Drawing.Point(72, 93);
            this.labelSubject.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelSubject.Name = "labelSubject";
            this.labelSubject.Size = new System.Drawing.Size(67, 17);
            this.labelSubject.TabIndex = 4;
            this.labelSubject.Text = "Subject:";
            // 
            // labelBody
            // 
            this.labelBody.AutoSize = true;
            this.labelBody.Location = new System.Drawing.Point(90, 140);
            this.labelBody.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelBody.Name = "labelBody";
            this.labelBody.Size = new System.Drawing.Size(49, 17);
            this.labelBody.TabIndex = 5;
            this.labelBody.Text = "Body:";
            // 
            // textReciever
            // 
            this.textReciever.Location = new System.Drawing.Point(154, 12);
            this.textReciever.Name = "textReciever";
            this.textReciever.Size = new System.Drawing.Size(129, 23);
            this.textReciever.TabIndex = 8;
            this.textReciever.Text = "jfang3@gmail.com";
            // 
            // textAttachment
            // 
            this.textAttachment.Location = new System.Drawing.Point(154, 57);
            this.textAttachment.Name = "textAttachment";
            this.textAttachment.Size = new System.Drawing.Size(582, 23);
            this.textAttachment.TabIndex = 9;
            // 
            // textSubject
            // 
            this.textSubject.Location = new System.Drawing.Point(154, 90);
            this.textSubject.Name = "textSubject";
            this.textSubject.Size = new System.Drawing.Size(582, 23);
            this.textSubject.TabIndex = 10;
            // 
            // richBody
            // 
            this.richBody.Location = new System.Drawing.Point(154, 153);
            this.richBody.Name = "richBody";
            this.richBody.Size = new System.Drawing.Size(638, 155);
            this.richBody.TabIndex = 11;
            this.richBody.Text = "";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(680, 338);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 28);
            this.button2.TabIndex = 12;
            this.button2.Text = "Send";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(837, 379);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.richBody);
            this.Controls.Add(this.textSubject);
            this.Controls.Add(this.textAttachment);
            this.Controls.Add(this.textReciever);
            this.Controls.Add(this.labelBody);
            this.Controls.Add(this.labelSubject);
            this.Controls.Add(this.labelAttach);
            this.Controls.Add(this.labelReceiver);
            this.Controls.Add(this.button1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form4";
            this.Text = "Email Sender";
            this.Load += new System.EventHandler(this.Form4_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label labelReceiver;
        private System.Windows.Forms.Label labelAttach;
        private System.Windows.Forms.Label labelSubject;
        private System.Windows.Forms.Label labelBody;
        private System.Windows.Forms.TextBox textReciever;
        private System.Windows.Forms.TextBox textAttachment;
        private System.Windows.Forms.TextBox textSubject;
        private System.Windows.Forms.RichTextBox richBody;
        private System.Windows.Forms.Button button2;
    }
}