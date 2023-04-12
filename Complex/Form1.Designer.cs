namespace Complex
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
            button1 = new Button();
            button2 = new Button();
            textBox1 = new TextBox();
            dataGridView1 = new DataGridView();
            WhatEdit = new DataGridViewTextBoxColumn();
            EditTo = new DataGridViewTextBoxColumn();
            FolderPath = new Label();
            button3 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(25, 86);
            button1.Name = "button1";
            button1.Size = new Size(114, 23);
            button1.TabIndex = 0;
            button1.Text = "Select files";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(25, 129);
            button2.Name = "button2";
            button2.Size = new Size(681, 23);
            button2.TabIndex = 1;
            button2.Text = "Bulk change rows in Word and Excel";
            button2.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            textBox1.Location = new Point(25, 47);
            textBox1.Name = "textBox1";
            textBox1.ReadOnly = true;
            textBox1.Size = new Size(671, 23);
            textBox1.TabIndex = 2;
            // 
            // dataGridView1
            // 
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { WhatEdit, EditTo });
            dataGridView1.Location = new Point(25, 158);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.RowHeadersWidth = 100;
            dataGridView1.RowTemplate.Height = 25;
            dataGridView1.Size = new Size(681, 123);
            dataGridView1.TabIndex = 3;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            // 
            // WhatEdit
            // 
            WhatEdit.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            WhatEdit.HeaderText = "What edit";
            WhatEdit.Name = "WhatEdit";
            // 
            // EditTo
            // 
            EditTo.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            EditTo.HeaderText = "Edit to";
            EditTo.Name = "EditTo";
            // 
            // FolderPath
            // 
            FolderPath.AutoSize = true;
            FolderPath.Location = new Point(25, 23);
            FolderPath.Name = "FolderPath";
            FolderPath.Size = new Size(67, 15);
            FolderPath.TabIndex = 4;
            FolderPath.Text = "Folder path";
            // 
            // button3
            // 
            button3.Location = new Point(417, 86);
            button3.Name = "button3";
            button3.Size = new Size(279, 23);
            button3.TabIndex = 5;
            button3.Text = "Bulk convert from Word to PDF";
            button3.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(button3);
            Controls.Add(FolderPath);
            Controls.Add(dataGridView1);
            Controls.Add(textBox1);
            Controls.Add(button2);
            Controls.Add(button1);
            Name = "Form1";
            Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private Button button2;
        private TextBox textBox1;
        private DataGridView dataGridView1;
        private DataGridViewTextBoxColumn WhatEdit;
        private DataGridViewTextBoxColumn EditTo;
        private Label FolderPath;
        private Button button3;
    }
}