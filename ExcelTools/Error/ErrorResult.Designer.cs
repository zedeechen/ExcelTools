namespace ExcelTools
{
    partial class ErrorResult
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
            this.lvwErrorResult = new System.Windows.Forms.ListView();
            this.lvwErrorHeaderIndex = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lvwErrorHeaderReason = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // lvwErrorResult
            // 
            this.lvwErrorResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lvwErrorResult.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lvwErrorHeaderIndex,
            this.lvwErrorHeaderReason});
            this.lvwErrorResult.FullRowSelect = true;
            this.lvwErrorResult.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lvwErrorResult.Location = new System.Drawing.Point(13, 13);
            this.lvwErrorResult.MultiSelect = false;
            this.lvwErrorResult.Name = "lvwErrorResult";
            this.lvwErrorResult.Size = new System.Drawing.Size(500, 484);
            this.lvwErrorResult.TabIndex = 0;
            this.lvwErrorResult.UseCompatibleStateImageBehavior = false;
            this.lvwErrorResult.View = System.Windows.Forms.View.Details;
            // 
            // lvwErrorHeaderIndex
            // 
            this.lvwErrorHeaderIndex.Text = "序号";
            this.lvwErrorHeaderIndex.Width = 50;
            // 
            // lvwErrorHeaderReason
            // 
            this.lvwErrorHeaderReason.Text = "出错原因";
            this.lvwErrorHeaderReason.Width = 440;
            // 
            // ErrorResult
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(525, 509);
            this.Controls.Add(this.lvwErrorResult);
            this.Name = "ErrorResult";
            this.Text = "错误列表";
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ListView lvwErrorResult;
        private System.Windows.Forms.ColumnHeader lvwErrorHeaderIndex;
        private System.Windows.Forms.ColumnHeader lvwErrorHeaderReason;


    }
}