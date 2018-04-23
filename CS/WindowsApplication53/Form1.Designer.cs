namespace WindowsApplication53
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
            this.pivotGridControl1 = new DevExpress.XtraPivotGrid.PivotGridControl();
            this.dataSet1 = new System.Data.DataSet();
            this.dataTable1 = new System.Data.DataTable();
            this.dataColumn1 = new System.Data.DataColumn();
            this.dataColumn2 = new System.Data.DataColumn();
            this.dataColumn3 = new System.Data.DataColumn();
            this.fiedName = new DevExpress.XtraPivotGrid.PivotGridField();
            this.fieldDate = new DevExpress.XtraPivotGrid.PivotGridField();
            this.fieldValue = new DevExpress.XtraPivotGrid.PivotGridField();
            this.fieldValue1 = new DevExpress.XtraPivotGrid.PivotGridField();
            this.fieldValue2 = new DevExpress.XtraPivotGrid.PivotGridField();
            this.fieldDate1 = new DevExpress.XtraPivotGrid.PivotGridField();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.pivotGridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).BeginInit();
            this.SuspendLayout();
            // 
            // pivotGridControl1
            // 
            this.pivotGridControl1.Cursor = System.Windows.Forms.Cursors.Default;
            this.pivotGridControl1.DataMember = "Data";
            this.pivotGridControl1.DataSource = this.dataSet1;
            this.pivotGridControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pivotGridControl1.Fields.AddRange(new DevExpress.XtraPivotGrid.PivotGridField[] {
            this.fiedName,
            this.fieldDate,
            this.fieldValue,
            this.fieldValue1,
            this.fieldValue2,
            this.fieldDate1});
            this.pivotGridControl1.Location = new System.Drawing.Point(0, 0);
            this.pivotGridControl1.Name = "pivotGridControl1";
            this.pivotGridControl1.OptionsPrint.PrintColumnHeaders = DevExpress.Utils.DefaultBoolean.False;
            this.pivotGridControl1.OptionsPrint.PrintDataHeaders = DevExpress.Utils.DefaultBoolean.False;
            this.pivotGridControl1.Size = new System.Drawing.Size(796, 292);
            this.pivotGridControl1.TabIndex = 0;
            this.pivotGridControl1.CustomExportFieldValue += new System.EventHandler<DevExpress.XtraPivotGrid.CustomExportFieldValueEventArgs>(this.pivotGridControl1_CustomExportFieldValue);
            this.pivotGridControl1.FieldValueDisplayText += new DevExpress.XtraPivotGrid.PivotFieldDisplayTextEventHandler(this.pivotGridControl1_FieldValueDisplayText);
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "NewDataSet";
            this.dataSet1.Tables.AddRange(new System.Data.DataTable[] {
            this.dataTable1});
            // 
            // dataTable1
            // 
            this.dataTable1.Columns.AddRange(new System.Data.DataColumn[] {
            this.dataColumn1,
            this.dataColumn2,
            this.dataColumn3});
            this.dataTable1.TableName = "Data";
            // 
            // dataColumn1
            // 
            this.dataColumn1.ColumnName = "Name";
            // 
            // dataColumn2
            // 
            this.dataColumn2.ColumnName = "Date";
            this.dataColumn2.DataType = typeof(System.DateTime);
            // 
            // dataColumn3
            // 
            this.dataColumn3.ColumnName = "Value";
            this.dataColumn3.DataType = typeof(decimal);
            // 
            // fiedName
            // 
            this.fiedName.Area = DevExpress.XtraPivotGrid.PivotArea.RowArea;
            this.fiedName.AreaIndex = 0;
            this.fiedName.FieldName = "Name";
            this.fiedName.Name = "fiedName";
            // 
            // fieldDate
            // 
            this.fieldDate.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea;
            this.fieldDate.AreaIndex = 0;
            this.fieldDate.Caption = "Year";
            this.fieldDate.FieldName = "Date";
            this.fieldDate.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.DateYear;
            this.fieldDate.Name = "fieldDate";
            this.fieldDate.UnboundFieldName = "fieldDate";
            // 
            // fieldValue
            // 
            this.fieldValue.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
            this.fieldValue.AreaIndex = 0;
            this.fieldValue.Caption = "Sum";
            this.fieldValue.FieldName = "Value";
            this.fieldValue.Name = "fieldValue";
            // 
            // fieldValue1
            // 
            this.fieldValue1.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
            this.fieldValue1.AreaIndex = 1;
            this.fieldValue1.Caption = "Count";
            this.fieldValue1.FieldName = "Value";
            this.fieldValue1.Name = "fieldValue1";
            this.fieldValue1.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Count;
            // 
            // fieldValue2
            // 
            this.fieldValue2.Area = DevExpress.XtraPivotGrid.PivotArea.DataArea;
            this.fieldValue2.AreaIndex = 2;
            this.fieldValue2.Caption = "Average";
            this.fieldValue2.FieldName = "Value";
            this.fieldValue2.Name = "fieldValue2";
            this.fieldValue2.SummaryType = DevExpress.Data.PivotGrid.PivotSummaryType.Average;
            // 
            // fieldDate1
            // 
            this.fieldDate1.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea;
            this.fieldDate1.AreaIndex = 1;
            this.fieldDate1.Caption = "Date";
            this.fieldDate1.FieldName = "Date";
            this.fieldDate1.GroupInterval = DevExpress.XtraPivotGrid.PivotGroupInterval.Date;
            this.fieldDate1.Name = "fieldDate1";
            this.fieldDate1.UnboundFieldName = "fieldDate1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(674, 0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 26);
            this.button1.TabIndex = 1;
            this.button1.Text = "Export";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(796, 292);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.pivotGridControl1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pivotGridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataTable1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraPivotGrid.PivotGridControl pivotGridControl1;
        private System.Data.DataSet dataSet1;
        private System.Data.DataTable dataTable1;
        private System.Data.DataColumn dataColumn1;
        private System.Data.DataColumn dataColumn2;
        private System.Data.DataColumn dataColumn3;
        private DevExpress.XtraPivotGrid.PivotGridField fiedName;
        private DevExpress.XtraPivotGrid.PivotGridField fieldDate;
        private DevExpress.XtraPivotGrid.PivotGridField fieldValue;
        private DevExpress.XtraPivotGrid.PivotGridField fieldValue1;
        private DevExpress.XtraPivotGrid.PivotGridField fieldValue2;
        private DevExpress.XtraPivotGrid.PivotGridField fieldDate1;
        private System.Windows.Forms.Button button1;
    }
}

