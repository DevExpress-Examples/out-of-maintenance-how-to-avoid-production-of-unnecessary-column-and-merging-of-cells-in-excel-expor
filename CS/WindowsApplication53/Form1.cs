using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraPrinting;
using DevExpress.XtraPivotGrid;

namespace WindowsApplication53
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            PopulateTable();
            pivotGridControl1.RefreshData();
            pivotGridControl1.BestFit();
        }
        private void PopulateTable()
        {
            DataTable myTable = dataSet1.Tables["Data"];
            myTable.Rows.Add(new object[] {"Aaa", DateTime.Today, 7});
            myTable.Rows.Add(new object[] { "Aaa", DateTime.Today.AddDays(1), 4 });
            myTable.Rows.Add(new object[] { "Bbb", DateTime.Today, 12 });
            myTable.Rows.Add(new object[] { "Bbb", DateTime.Today.AddDays(1), 14 });
            myTable.Rows.Add(new object[] { "Ccc", DateTime.Today, 11 });
            myTable.Rows.Add(new object[] { "Ccc", DateTime.Today.AddDays(1), 10 });

            myTable.Rows.Add(new object[] { "Aaa", DateTime.Today.AddYears(1), 4 });
            myTable.Rows.Add(new object[] { "Aaa", DateTime.Today.AddYears(1).AddDays(1), 2 });
            myTable.Rows.Add(new object[] { "Bbb", DateTime.Today.AddYears(1), 3 });
            myTable.Rows.Add(new object[] { "Bbb", DateTime.Today.AddDays(1).AddYears(1), 1 });
            myTable.Rows.Add(new object[] { "Ccc", DateTime.Today.AddYears(1), 8 });
            myTable.Rows.Add(new object[] { "Ccc", DateTime.Today.AddDays(1).AddYears(1), 22 });
        }

        string fileName = "pivot.xls";
        private void button1_Click(object sender, EventArgs e)
        {
            pivotGridControl1.BeginUpdate();
            pivotGridControl1.OptionsPrint.PrintColumnHeaders = DevExpress.Utils.DefaultBoolean.False;
            pivotGridControl1.OptionsPrint.PrintDataHeaders = DevExpress.Utils.DefaultBoolean.False;
            pivotGridControl1.OptionsPrint.PrintFilterHeaders = DevExpress.Utils.DefaultBoolean.False;

            PivotGridField fieldExportHeader = pivotGridControl1.Fields.Add();
            fieldExportHeader.Caption = "Export";
            fieldExportHeader.Name = "fieldExportHeader";
            fieldExportHeader.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea;
            fieldExportHeader.Visible = true;
            fieldExportHeader.AreaIndex = 0;
            fieldExportHeader.TotalsVisibility = PivotTotalsVisibility.None;
            pivotGridControl1.OptionsView.ShowGrandTotalsForSingleValues = true;
            pivotGridControl1.EndUpdate();

            try
            {
                pivotGridControl1.ExportToXls(fileName);
                System.Diagnostics.Process.Start(fileName);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                pivotGridControl1.Fields.Remove(fieldExportHeader);
            }
        }

        private void pivotGridControl1_CustomExportFieldValue(object sender, DevExpress.XtraPivotGrid.CustomExportFieldValueEventArgs e)
        {
            TextBrick brick = e.Brick as TextBrick;
            if (e.Brick == null) return;

            if (e.Field != null && e.Field.Name == "fieldExportHeader" && e.ValueType == PivotGridValueType.Value )
            {
                brick.Text = GetLayoutDiscription(sender as PivotGridControl, e.Field );
            }
        }

        private void pivotGridControl1_FieldValueDisplayText(object sender, PivotFieldDisplayTextEventArgs e)
        {
            if (e.Field != null && e.Field.Name == "fieldExportHeader" && e.ValueType == PivotGridValueType.Value)
                e.DisplayText = "Export in progress, please wait...";
        } 


        private string GetLayoutDiscription(PivotGridControl pivot, PivotGridField unnecessaryField)
        {
            string text = "ColumnFields: ";
            List<PivotGridField> fields = pivot.GetFieldsByArea(PivotArea.ColumnArea);
            for (int i = 0; i < fields.Count ; i++)
            {
                if (object.ReferenceEquals(fields[i], unnecessaryField)) continue;
                text = text + fields[i].ToString();
                if (i < fields.Count - 1)
                    text = text + " -> ";
            }
            return text;
        }

     
    }
}