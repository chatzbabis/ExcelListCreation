
namespace ExcelListCreation
{
    partial class UsersListCreation
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
            this.components = new System.ComponentModel.Container();
            this.SelectInitialList = new System.Windows.Forms.Button();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.CreateTheList = new System.Windows.Forms.Button();
            this.SelectPath = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.Help = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SelectInitialList
            // 
            this.SelectInitialList.Location = new System.Drawing.Point(18, 46);
            this.SelectInitialList.Name = "SelectInitialList";
            this.SelectInitialList.Size = new System.Drawing.Size(122, 20);
            this.SelectInitialList.TabIndex = 0;
            this.SelectInitialList.Text = "Select initial list";
            this.SelectInitialList.UseVisualStyleBackColor = true;
            this.SelectInitialList.Click += new System.EventHandler(this.Button1_Click_1);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(169, 47);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(144, 20);
            this.textBox1.TabIndex = 1;
            // 
            // CreateTheList
            // 
            this.CreateTheList.Location = new System.Drawing.Point(83, 176);
            this.CreateTheList.Name = "CreateTheList";
            this.CreateTheList.Size = new System.Drawing.Size(159, 22);
            this.CreateTheList.TabIndex = 2;
            this.CreateTheList.Text = "Create the list";
            this.CreateTheList.UseVisualStyleBackColor = true;
            this.CreateTheList.Click += new System.EventHandler(this.Button2_Click);
            // 
            // SelectPath
            // 
            this.SelectPath.Location = new System.Drawing.Point(18, 106);
            this.SelectPath.Name = "SelectPath";
            this.SelectPath.Size = new System.Drawing.Size(126, 20);
            this.SelectPath.TabIndex = 3;
            this.SelectPath.Text = "Select path";
            this.SelectPath.UseVisualStyleBackColor = true;
            this.SelectPath.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(169, 106);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(144, 20);
            this.textBox2.TabIndex = 4;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(18, 250);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(295, 10);
            this.progressBar1.TabIndex = 5;
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Help
            // 
            this.Help.Location = new System.Drawing.Point(278, 12);
            this.Help.Name = "Help";
            this.Help.Size = new System.Drawing.Size(46, 22);
            this.Help.TabIndex = 6;
            this.Help.Text = "Help";
            this.Help.UseVisualStyleBackColor = true;
            this.Help.Click += new System.EventHandler(this.Help_Click);
            // 
            // UsersListCreation
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(336, 262);
            this.Controls.Add(this.Help);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.SelectPath);
            this.Controls.Add(this.CreateTheList);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.SelectInitialList);
            this.Name = "UsersListCreation";
            this.Text = "Users List Creation";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button SelectInitialList;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button CreateTheList;
        private System.Windows.Forms.Button SelectPath;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Button Help;
    }
}

