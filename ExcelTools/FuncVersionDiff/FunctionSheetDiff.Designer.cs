namespace ExcelTools
{
    partial class FunctionSheetDiff
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose( bool disposing )
        {
            if ( disposing && ( components != null ) )
            {
                components.Dispose();
            }
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblOldExcelPath = new System.Windows.Forms.Label();
            this.lblNewExcelPath = new System.Windows.Forms.Label();
            this.btnPushNew = new System.Windows.Forms.Button();
            this.btnPushOld = new System.Windows.Forms.Button();
            this.btnSelectOld = new System.Windows.Forms.Button();
            this.btnSelectNew = new System.Windows.Forms.Button();
            this.lvwOldExcelDiff = new ExcelTools.SyncListView();
            this.lvwNewExcelDiff = new ExcelTools.SyncListView();
            this.SuspendLayout();
            // 
            // lblOldExcelPath
            // 
            this.lblOldExcelPath.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lblOldExcelPath.AutoSize = true;
            this.lblOldExcelPath.Location = new System.Drawing.Point(10, 9);
            this.lblOldExcelPath.Name = "lblOldExcelPath";
            this.lblOldExcelPath.Size = new System.Drawing.Size(41, 12);
            this.lblOldExcelPath.TabIndex = 3;
            this.lblOldExcelPath.Text = "label1";
            // 
            // lblNewExcelPath
            // 
            this.lblNewExcelPath.Anchor = System.Windows.Forms.AnchorStyles.Top;
            this.lblNewExcelPath.AutoSize = true;
            this.lblNewExcelPath.Location = new System.Drawing.Point(960, 9);
            this.lblNewExcelPath.Name = "lblNewExcelPath";
            this.lblNewExcelPath.Size = new System.Drawing.Size(41, 12);
            this.lblNewExcelPath.TabIndex = 3;
            this.lblNewExcelPath.Text = "label1";
            // 
            // btnPushNew
            // 
            this.btnPushNew.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnPushNew.Location = new System.Drawing.Point(870, 247);
            this.btnPushNew.Name = "btnPushNew";
            this.btnPushNew.Size = new System.Drawing.Size(75, 23);
            this.btnPushNew.TabIndex = 4;
            this.btnPushNew.Text = "==>";
            this.btnPushNew.UseVisualStyleBackColor = true;
            this.btnPushNew.Click += new System.EventHandler(this.btnPushNew_Click);
            // 
            // btnPushOld
            // 
            this.btnPushOld.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.btnPushOld.Location = new System.Drawing.Point(870, 383);
            this.btnPushOld.Name = "btnPushOld";
            this.btnPushOld.Size = new System.Drawing.Size(75, 23);
            this.btnPushOld.TabIndex = 4;
            this.btnPushOld.Text = "<==";
            this.btnPushOld.UseVisualStyleBackColor = true;
            this.btnPushOld.Click += new System.EventHandler(this.btnPushOld_Click);
            // 
            // btnSelectOld
            // 
            this.btnSelectOld.Location = new System.Drawing.Point(763, 4);
            this.btnSelectOld.Name = "btnSelectOld";
            this.btnSelectOld.Size = new System.Drawing.Size(100, 23);
            this.btnSelectOld.TabIndex = 6;
            this.btnSelectOld.Text = "新增行全选";
            this.btnSelectOld.UseVisualStyleBackColor = true;
            this.btnSelectOld.Click += new System.EventHandler(this.btnSelectOld_Click);
            // 
            // btnSelectNew
            // 
            this.btnSelectNew.Location = new System.Drawing.Point(1713, 4);
            this.btnSelectNew.Name = "btnSelectNew";
            this.btnSelectNew.Size = new System.Drawing.Size(100, 23);
            this.btnSelectNew.TabIndex = 7;
            this.btnSelectNew.Text = "新增行全选";
            this.btnSelectNew.UseVisualStyleBackColor = true;
            this.btnSelectNew.Click += new System.EventHandler(this.btnSelectNew_Click);
            // 
            // lvwOldExcelDiff
            // 
            this.lvwOldExcelDiff.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.lvwOldExcelDiff.FullRowSelect = true;
            this.lvwOldExcelDiff.HideSelection = false;
            this.lvwOldExcelDiff.Location = new System.Drawing.Point(1, 29);
            this.lvwOldExcelDiff.Name = "lvwOldExcelDiff";
            this.lvwOldExcelDiff.Size = new System.Drawing.Size(862, 751);
            this.lvwOldExcelDiff.TabIndex = 2;
            this.lvwOldExcelDiff.UseCompatibleStateImageBehavior = false;
            this.lvwOldExcelDiff.View = System.Windows.Forms.View.Details;
            this.lvwOldExcelDiff.VirtualMode = true;
            this.lvwOldExcelDiff.TopItemChanged += new System.EventHandler(this.lvwOldExcelDiff_TopItemChanged);
            this.lvwOldExcelDiff.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvwOldExcelDiff_ColumnClick);
            this.lvwOldExcelDiff.ColumnWidthChanging += new System.Windows.Forms.ColumnWidthChangingEventHandler(this.lvwOldExcelDiff_ColumnWidthChanging);
            this.lvwOldExcelDiff.RetrieveVirtualItem += new System.Windows.Forms.RetrieveVirtualItemEventHandler(this.lvwOldExcelDiff_RetrieveVirtualItem);
            this.lvwOldExcelDiff.SelectedIndexChanged += new System.EventHandler(this.lvwOldExcelDiff_SelectedIndexChanged);
            this.lvwOldExcelDiff.VirtualItemsSelectionRangeChanged += new System.Windows.Forms.ListViewVirtualItemsSelectionRangeChangedEventHandler(this.lvwOldExcelDiff_VirtualItemsSelectionRangeChanged);
            this.lvwOldExcelDiff.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lvwOldExcelDiff_MouseDown);
            // 
            // lvwNewExcelDiff
            // 
            this.lvwNewExcelDiff.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.lvwNewExcelDiff.FullRowSelect = true;
            this.lvwNewExcelDiff.HideSelection = false;
            this.lvwNewExcelDiff.Location = new System.Drawing.Point(951, 29);
            this.lvwNewExcelDiff.Name = "lvwNewExcelDiff";
            this.lvwNewExcelDiff.Size = new System.Drawing.Size(862, 751);
            this.lvwNewExcelDiff.TabIndex = 2;
            this.lvwNewExcelDiff.UseCompatibleStateImageBehavior = false;
            this.lvwNewExcelDiff.View = System.Windows.Forms.View.Details;
            this.lvwNewExcelDiff.TopItemChanged += new System.EventHandler(this.lvwNewExcelDiff_TopItemChanged);
            this.lvwNewExcelDiff.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvwNewExcelDiff_ColumnClick);
            this.lvwNewExcelDiff.ColumnWidthChanging += new System.Windows.Forms.ColumnWidthChangingEventHandler(this.lvwNewExcelDiff_ColumnWidthChanging);
            this.lvwNewExcelDiff.RetrieveVirtualItem += new System.Windows.Forms.RetrieveVirtualItemEventHandler(this.lvwNewExcelDiff_RetrieveVirtualItem);
            this.lvwNewExcelDiff.SelectedIndexChanged += new System.EventHandler(this.lvwNewExcelDiff_SelectedIndexChanged);
            this.lvwNewExcelDiff.VirtualItemsSelectionRangeChanged += new System.Windows.Forms.ListViewVirtualItemsSelectionRangeChangedEventHandler(this.lvwNewExcelDiff_VirtualItemsSelectionRangeChanged);
            this.lvwNewExcelDiff.MouseDown += new System.Windows.Forms.MouseEventHandler(this.lvwNewExcelDiff_MouseDown);
            // 
            // FunctionSheetDiff
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(1814, 782);
            this.Controls.Add(this.btnSelectNew);
            this.Controls.Add(this.btnSelectOld);
            this.Controls.Add(this.btnPushOld);
            this.Controls.Add(this.btnPushNew);
            this.Controls.Add(this.lblNewExcelPath);
            this.Controls.Add(this.lblOldExcelPath);
            this.Controls.Add(this.lvwNewExcelDiff);
            this.Controls.Add(this.lvwOldExcelDiff);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "FunctionSheetDiff";
            this.Text = "FunctionSheetDiff";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FunctionSheetDiff_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private SyncListView lvwOldExcelDiff;
        private System.Windows.Forms.Label lblOldExcelPath;
        private SyncListView lvwNewExcelDiff;
        private System.Windows.Forms.Label lblNewExcelPath;
        private System.Windows.Forms.Button btnPushNew;
        private System.Windows.Forms.Button btnPushOld;
        private System.Windows.Forms.Button btnSelectOld;
        private System.Windows.Forms.Button btnSelectNew;

    }
}