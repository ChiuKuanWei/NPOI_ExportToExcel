namespace EXCEL_Export
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.lblBefore = new System.Windows.Forms.Label();
            this.lblAfter = new System.Windows.Forms.Label();
            this.dtpBefore = new System.Windows.Forms.DateTimePicker();
            this.dtpAfter = new System.Windows.Forms.DateTimePicker();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnMerge = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtPath2 = new System.Windows.Forms.TextBox();
            this.txtPath1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnSelect2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSelect1 = new System.Windows.Forms.Button();
            this.dlg1 = new System.Windows.Forms.OpenFileDialog();
            this.dlg2 = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblBefore
            // 
            this.lblBefore.AutoSize = true;
            this.lblBefore.Location = new System.Drawing.Point(21, 73);
            this.lblBefore.Name = "lblBefore";
            this.lblBefore.Size = new System.Drawing.Size(17, 12);
            this.lblBefore.TabIndex = 20;
            this.lblBefore.Text = "到";
            // 
            // lblAfter
            // 
            this.lblAfter.AutoSize = true;
            this.lblAfter.Location = new System.Drawing.Point(21, 38);
            this.lblAfter.Name = "lblAfter";
            this.lblAfter.Size = new System.Drawing.Size(17, 12);
            this.lblAfter.TabIndex = 19;
            this.lblAfter.Text = "從";
            // 
            // dtpBefore
            // 
            this.dtpBefore.Location = new System.Drawing.Point(56, 68);
            this.dtpBefore.Name = "dtpBefore";
            this.dtpBefore.Size = new System.Drawing.Size(200, 22);
            this.dtpBefore.TabIndex = 18;
            // 
            // dtpAfter
            // 
            this.dtpAfter.Location = new System.Drawing.Point(56, 33);
            this.dtpAfter.Name = "dtpAfter";
            this.dtpAfter.Size = new System.Drawing.Size(200, 22);
            this.dtpAfter.TabIndex = 17;
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(277, 38);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(96, 47);
            this.btnExport.TabIndex = 16;
            this.btnExport.Text = "匯出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnMerge
            // 
            this.btnMerge.Location = new System.Drawing.Point(148, 154);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new System.Drawing.Size(96, 29);
            this.btnMerge.TabIndex = 21;
            this.btnMerge.Text = "合併";
            this.btnMerge.UseVisualStyleBackColor = true;
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnExport);
            this.groupBox1.Controls.Add(this.dtpAfter);
            this.groupBox1.Controls.Add(this.lblBefore);
            this.groupBox1.Controls.Add(this.dtpBefore);
            this.groupBox1.Controls.Add(this.lblAfter);
            this.groupBox1.Location = new System.Drawing.Point(3, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(399, 100);
            this.groupBox1.TabIndex = 22;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "匯出功能";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtPath2);
            this.groupBox2.Controls.Add(this.txtPath1);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.btnSelect2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.btnSelect1);
            this.groupBox2.Controls.Add(this.btnMerge);
            this.groupBox2.Location = new System.Drawing.Point(3, 108);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(399, 198);
            this.groupBox2.TabIndex = 23;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "合併檔案";
            // 
            // txtPath2
            // 
            this.txtPath2.Location = new System.Drawing.Point(33, 115);
            this.txtPath2.Name = "txtPath2";
            this.txtPath2.ReadOnly = true;
            this.txtPath2.Size = new System.Drawing.Size(327, 22);
            this.txtPath2.TabIndex = 27;
            // 
            // txtPath1
            // 
            this.txtPath1.Location = new System.Drawing.Point(33, 51);
            this.txtPath1.Name = "txtPath1";
            this.txtPath1.ReadOnly = true;
            this.txtPath1.Size = new System.Drawing.Size(327, 22);
            this.txtPath1.TabIndex = 26;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(125, 91);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 12);
            this.label2.TabIndex = 25;
            this.label2.Text = "合併檔案2:";
            // 
            // btnSelect2
            // 
            this.btnSelect2.Location = new System.Drawing.Point(193, 85);
            this.btnSelect2.Name = "btnSelect2";
            this.btnSelect2.Size = new System.Drawing.Size(73, 24);
            this.btnSelect2.TabIndex = 24;
            this.btnSelect2.Text = "請選擇檔案";
            this.btnSelect2.UseVisualStyleBackColor = true;
            this.btnSelect2.Click += new System.EventHandler(this.btnSelect2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(125, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(62, 12);
            this.label1.TabIndex = 23;
            this.label1.Text = "合併檔案1:";
            // 
            // btnSelect1
            // 
            this.btnSelect1.Location = new System.Drawing.Point(193, 21);
            this.btnSelect1.Name = "btnSelect1";
            this.btnSelect1.Size = new System.Drawing.Size(73, 24);
            this.btnSelect1.TabIndex = 22;
            this.btnSelect1.Text = "請選擇檔案";
            this.btnSelect1.UseVisualStyleBackColor = true;
            this.btnSelect1.Click += new System.EventHandler(this.btnSelect1_Click);
            // 
            // dlg1
            // 
            this.dlg1.FileName = "openFileDialog1";
            // 
            // dlg2
            // 
            this.dlg2.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(407, 318);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "NPOI_EXCEL匯出Demo";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lblBefore;
        private System.Windows.Forms.Label lblAfter;
        private System.Windows.Forms.DateTimePicker dtpBefore;
        private System.Windows.Forms.DateTimePicker dtpAfter;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnSelect2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSelect1;
        private System.Windows.Forms.OpenFileDialog dlg1;
        private System.Windows.Forms.OpenFileDialog dlg2;
        private System.Windows.Forms.TextBox txtPath2;
        private System.Windows.Forms.TextBox txtPath1;
    }
}

