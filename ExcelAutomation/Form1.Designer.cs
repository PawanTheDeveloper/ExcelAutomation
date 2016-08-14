namespace ExcelAutomation
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
            this.textBox_firstfilename = new System.Windows.Forms.TextBox();
            this.button_firstfile = new System.Windows.Forms.Button();
            this.listBox_populatecolumn = new System.Windows.Forms.ListBox();
            this.button_add = new System.Windows.Forms.Button();
            this.button_remove = new System.Windows.Forms.Button();
            this.listBox_selectedcolumn = new System.Windows.Forms.ListBox();
            this.button_generatefile = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // textBox_firstfilename
            // 
            this.textBox_firstfilename.Location = new System.Drawing.Point(24, 36);
            this.textBox_firstfilename.Name = "textBox_firstfilename";
            this.textBox_firstfilename.Size = new System.Drawing.Size(437, 20);
            this.textBox_firstfilename.TabIndex = 0;
            // 
            // button_firstfile
            // 
            this.button_firstfile.Location = new System.Drawing.Point(467, 33);
            this.button_firstfile.Name = "button_firstfile";
            this.button_firstfile.Size = new System.Drawing.Size(75, 23);
            this.button_firstfile.TabIndex = 2;
            this.button_firstfile.Text = "Browse....";
            this.button_firstfile.UseVisualStyleBackColor = true;
            this.button_firstfile.Click += new System.EventHandler(this.button1_Click);
            // 
            // listBox_populatecolumn
            // 
            this.listBox_populatecolumn.FormattingEnabled = true;
            this.listBox_populatecolumn.Location = new System.Drawing.Point(24, 68);
            this.listBox_populatecolumn.Name = "listBox_populatecolumn";
            this.listBox_populatecolumn.Size = new System.Drawing.Size(189, 420);
            this.listBox_populatecolumn.TabIndex = 7;
            // 
            // button_add
            // 
            this.button_add.Location = new System.Drawing.Point(243, 199);
            this.button_add.Name = "button_add";
            this.button_add.Size = new System.Drawing.Size(75, 23);
            this.button_add.TabIndex = 8;
            this.button_add.Text = "Add";
            this.button_add.UseVisualStyleBackColor = true;
            this.button_add.Click += new System.EventHandler(this.button_add_Click);
            // 
            // button_remove
            // 
            this.button_remove.Location = new System.Drawing.Point(243, 246);
            this.button_remove.Name = "button_remove";
            this.button_remove.Size = new System.Drawing.Size(75, 23);
            this.button_remove.TabIndex = 9;
            this.button_remove.Text = "Remove";
            this.button_remove.UseVisualStyleBackColor = true;
            this.button_remove.Click += new System.EventHandler(this.button_remove_Click);
            // 
            // listBox_selectedcolumn
            // 
            this.listBox_selectedcolumn.FormattingEnabled = true;
            this.listBox_selectedcolumn.Location = new System.Drawing.Point(365, 175);
            this.listBox_selectedcolumn.Name = "listBox_selectedcolumn";
            this.listBox_selectedcolumn.Size = new System.Drawing.Size(177, 160);
            this.listBox_selectedcolumn.TabIndex = 10;
            // 
            // button_generatefile
            // 
            this.button_generatefile.Location = new System.Drawing.Point(243, 378);
            this.button_generatefile.Name = "button_generatefile";
            this.button_generatefile.Size = new System.Drawing.Size(324, 55);
            this.button_generatefile.TabIndex = 12;
            this.button_generatefile.Text = "Generate File";
            this.button_generatefile.UseVisualStyleBackColor = true;
            this.button_generatefile.Click += new System.EventHandler(this.button_generatefile_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(243, 453);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(320, 23);
            this.progressBar1.TabIndex = 13;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(575, 493);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button_generatefile);
            this.Controls.Add(this.listBox_selectedcolumn);
            this.Controls.Add(this.button_remove);
            this.Controls.Add(this.button_add);
            this.Controls.Add(this.listBox_populatecolumn);
            this.Controls.Add(this.button_firstfile);
            this.Controls.Add(this.textBox_firstfilename);
            this.Name = "Form1";
            this.Text = "Excel Addin";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox_firstfilename;
        private System.Windows.Forms.Button button_firstfile;
        private System.Windows.Forms.ListBox listBox_populatecolumn;
        private System.Windows.Forms.Button button_add;
        private System.Windows.Forms.Button button_remove;
        private System.Windows.Forms.ListBox listBox_selectedcolumn;
        private System.Windows.Forms.Button button_generatefile;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}

