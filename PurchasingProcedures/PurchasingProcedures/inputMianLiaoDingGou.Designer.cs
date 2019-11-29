namespace PurchasingProcedures
{
    partial class inputMianLiaoDingGou
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(inputMianLiaoDingGou));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.cb_jgc = new System.Windows.Forms.ComboBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.txt_ml = new System.Windows.Forms.ComboBox();
            this.txt_ks = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cb_cd = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(14, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "面料号";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 81);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 0;
            this.label2.Text = "款式号";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 109);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(41, 12);
            this.label3.TabIndex = 0;
            this.label3.Text = "加工厂";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(134, 146);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "确定";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // cb_jgc
            // 
            this.cb_jgc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_jgc.FormattingEnabled = true;
            this.cb_jgc.Location = new System.Drawing.Point(70, 109);
            this.cb_jgc.Name = "cb_jgc";
            this.cb_jgc.Size = new System.Drawing.Size(139, 20);
            this.cb_jgc.TabIndex = 3;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // txt_ml
            // 
            this.txt_ml.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.txt_ml.FormattingEnabled = true;
            this.txt_ml.Location = new System.Drawing.Point(70, 49);
            this.txt_ml.Name = "txt_ml";
            this.txt_ml.Size = new System.Drawing.Size(139, 20);
            this.txt_ml.TabIndex = 4;
            // 
            // txt_ks
            // 
            this.txt_ks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.txt_ks.FormattingEnabled = true;
            this.txt_ks.Location = new System.Drawing.Point(70, 78);
            this.txt_ks.Name = "txt_ks";
            this.txt_ks.Size = new System.Drawing.Size(139, 20);
            this.txt_ks.TabIndex = 4;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(14, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "裁单号";
            // 
            // cb_cd
            // 
            this.cb_cd.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_cd.FormattingEnabled = true;
            this.cb_cd.Location = new System.Drawing.Point(70, 23);
            this.cb_cd.Name = "cb_cd";
            this.cb_cd.Size = new System.Drawing.Size(139, 20);
            this.cb_cd.TabIndex = 4;
            // 
            // inputMianLiaoDingGou
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(230, 181);
            this.Controls.Add(this.txt_ks);
            this.Controls.Add(this.cb_cd);
            this.Controls.Add(this.txt_ml);
            this.Controls.Add(this.cb_jgc);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "inputMianLiaoDingGou";
            this.Text = "面辅料订购";
            this.Load += new System.EventHandler(this.inputMianLiaoDingGou_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cb_jgc;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ComboBox txt_ml;
        private System.Windows.Forms.ComboBox txt_ks;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cb_cd;
    }
}