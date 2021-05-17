
namespace WinFormsApp1
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
            this.ans = new System.Windows.Forms.DataGridView();
            this.fileselect = new System.Windows.Forms.Button();
            this.filetext = new System.Windows.Forms.TextBox();
            this.draw = new System.Windows.Forms.DataGridView();
            this.export = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ans)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.draw)).BeginInit();
            this.SuspendLayout();
            // 
            // ans
            // 
            this.ans.AllowDrop = true;
            this.ans.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ans.Location = new System.Drawing.Point(81, 100);
            this.ans.Name = "ans";
            this.ans.RowTemplate.Height = 25;
            this.ans.Size = new System.Drawing.Size(403, 519);
            this.ans.TabIndex = 0;
            this.ans.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // fileselect
            // 
            this.fileselect.Location = new System.Drawing.Point(560, 26);
            this.fileselect.Name = "fileselect";
            this.fileselect.Size = new System.Drawing.Size(75, 23);
            this.fileselect.TabIndex = 1;
            this.fileselect.Text = "選擇檔案";
            this.fileselect.UseVisualStyleBackColor = true;
            this.fileselect.Click += new System.EventHandler(this.button1_Click);
            // 
            // filetext
            // 
            this.filetext.Location = new System.Drawing.Point(81, 26);
            this.filetext.Name = "filetext";
            this.filetext.Size = new System.Drawing.Size(460, 23);
            this.filetext.TabIndex = 2;
            // 
            // draw
            // 
            this.draw.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.draw.Location = new System.Drawing.Point(529, 100);
            this.draw.Name = "draw";
            this.draw.RowTemplate.Height = 25;
            this.draw.Size = new System.Drawing.Size(742, 519);
            this.draw.TabIndex = 3;
            // 
            // export
            // 
            this.export.Location = new System.Drawing.Point(671, 25);
            this.export.Name = "export";
            this.export.Size = new System.Drawing.Size(75, 23);
            this.export.TabIndex = 4;
            this.export.Text = "匯出excel";
            this.export.UseVisualStyleBackColor = true;
            this.export.Click += new System.EventHandler(this.export_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1566, 714);
            this.Controls.Add(this.export);
            this.Controls.Add(this.draw);
            this.Controls.Add(this.filetext);
            this.Controls.Add(this.fileselect);
            this.Controls.Add(this.ans);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.ans)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.draw)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView ans;
        private System.Windows.Forms.Button fileselect;
        private System.Windows.Forms.TextBox filetext;
        private System.Windows.Forms.DataGridView draw;
        private System.Windows.Forms.Button export;
    }
}

