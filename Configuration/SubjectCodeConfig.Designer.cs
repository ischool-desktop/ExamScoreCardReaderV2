namespace ExamScoreCardReaderV2
{
    partial class SubjectCodeConfig
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgv = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chSubject = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(786, 511);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(867, 511);
            // 
            // picWaiting
            // 
            this.picWaiting.Location = new System.Drawing.Point(453, 239);
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(6, 511);
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(62, 511);
            // 
            // dgv
            // 
            this.dgv.AllowUserToResizeRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.BackgroundColor = System.Drawing.Color.White;
            this.dgv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chSubject,
            this.chCode});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgv.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgv.HighlightSelectedColumnHeaders = false;
            this.dgv.Location = new System.Drawing.Point(6, 5);
            this.dgv.Name = "dgv";
            this.dgv.RowHeadersWidth = 25;
            this.dgv.Size = new System.Drawing.Size(926, 500);
            this.dgv.TabIndex = 4;
            this.dgv.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_CellEndEdit);
            // 
            // chSubject
            // 
            this.chSubject.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chSubject.HeaderText = "科目";
            this.chSubject.Name = "chSubject";
            // 
            // chCode
            // 
            this.chCode.HeaderText = "代碼";
            this.chCode.Name = "chCode";
            // 
            // labelX3
            // 
            this.labelX3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(118, 513);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(382, 20);
            this.labelX3.TabIndex = 13;
            this.labelX3.Text = "同一代碼匯入多個科目成績時，可用\"<font color=\"#000000\">,</font><font color=\"#ED1C24\"></font>\"分隔，E" +
    "x：英文,英文選修。";
            // 
            // SubjectCodeConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(938, 539);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.dgv);
            this.MaximizeBox = true;
            this.Name = "SubjectCodeConfig";
            this.Text = "科目代碼設定";
            this.Controls.SetChildIndex(this.btnImport, 0);
            this.Controls.SetChildIndex(this.btnExport, 0);
            this.Controls.SetChildIndex(this.dgv, 0);
            this.Controls.SetChildIndex(this.btnSave, 0);
            this.Controls.SetChildIndex(this.btnClose, 0);
            this.Controls.SetChildIndex(this.picWaiting, 0);
            this.Controls.SetChildIndex(this.labelX3, 0);
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgv;
        private System.Windows.Forms.DataGridViewTextBoxColumn chSubject;
        private System.Windows.Forms.DataGridViewTextBoxColumn chCode;
        private DevComponents.DotNetBar.LabelX labelX3;
    }
}