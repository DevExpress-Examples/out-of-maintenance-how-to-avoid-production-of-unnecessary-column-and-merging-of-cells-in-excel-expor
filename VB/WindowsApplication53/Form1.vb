Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraPrinting
Imports DevExpress.XtraPivotGrid

Namespace WindowsApplication53
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
			PopulateTable()
			pivotGridControl1.RefreshData()
			pivotGridControl1.BestFit()
		End Sub
		Private Sub PopulateTable()
			Dim myTable As DataTable = dataSet1.Tables("Data")
			myTable.Rows.Add(New Object() {"Aaa", DateTime.Today, 7})
			myTable.Rows.Add(New Object() { "Aaa", DateTime.Today.AddDays(1), 4 })
			myTable.Rows.Add(New Object() { "Bbb", DateTime.Today, 12 })
			myTable.Rows.Add(New Object() { "Bbb", DateTime.Today.AddDays(1), 14 })
			myTable.Rows.Add(New Object() { "Ccc", DateTime.Today, 11 })
			myTable.Rows.Add(New Object() { "Ccc", DateTime.Today.AddDays(1), 10 })

			myTable.Rows.Add(New Object() { "Aaa", DateTime.Today.AddYears(1), 4 })
			myTable.Rows.Add(New Object() { "Aaa", DateTime.Today.AddYears(1).AddDays(1), 2 })
			myTable.Rows.Add(New Object() { "Bbb", DateTime.Today.AddYears(1), 3 })
			myTable.Rows.Add(New Object() { "Bbb", DateTime.Today.AddDays(1).AddYears(1), 1 })
			myTable.Rows.Add(New Object() { "Ccc", DateTime.Today.AddYears(1), 8 })
			myTable.Rows.Add(New Object() { "Ccc", DateTime.Today.AddDays(1).AddYears(1), 22 })
		End Sub

		Private fileName As String = "pivot.xls"
		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			pivotGridControl1.BeginUpdate()
			pivotGridControl1.OptionsPrint.PrintColumnHeaders = DevExpress.Utils.DefaultBoolean.False
			pivotGridControl1.OptionsPrint.PrintDataHeaders = DevExpress.Utils.DefaultBoolean.False
			pivotGridControl1.OptionsPrint.PrintFilterHeaders = DevExpress.Utils.DefaultBoolean.False

			Dim fieldExportHeader As PivotGridField = pivotGridControl1.Fields.Add()
			fieldExportHeader.Caption = "Export"
			fieldExportHeader.Name = "fieldExportHeader"
			fieldExportHeader.Area = DevExpress.XtraPivotGrid.PivotArea.ColumnArea
			fieldExportHeader.Visible = True
			fieldExportHeader.AreaIndex = 0
			fieldExportHeader.TotalsVisibility = PivotTotalsVisibility.None
			pivotGridControl1.OptionsView.ShowGrandTotalsForSingleValues = True
			pivotGridControl1.EndUpdate()

			Try
				pivotGridControl1.ExportToXls(fileName)
				System.Diagnostics.Process.Start(fileName)
			Catch ex As System.Exception
				MessageBox.Show(ex.Message)
			Finally
				pivotGridControl1.Fields.Remove(fieldExportHeader)
			End Try
		End Sub

		Private Sub pivotGridControl1_CustomExportFieldValue(ByVal sender As Object, ByVal e As DevExpress.XtraPivotGrid.CustomExportFieldValueEventArgs) Handles pivotGridControl1.CustomExportFieldValue
			Dim brick As TextBrick = TryCast(e.Brick, TextBrick)
			If e.Brick Is Nothing Then
				Return
			End If

			If e.Field IsNot Nothing AndAlso e.Field.Name = "fieldExportHeader" AndAlso e.ValueType = PivotGridValueType.Value Then
				brick.Text = GetLayoutDiscription(TryCast(sender, PivotGridControl), e.Field)
			End If
		End Sub

		Private Sub pivotGridControl1_FieldValueDisplayText(ByVal sender As Object, ByVal e As PivotFieldDisplayTextEventArgs) Handles pivotGridControl1.FieldValueDisplayText
			If e.Field IsNot Nothing AndAlso e.Field.Name = "fieldExportHeader" AndAlso e.ValueType = PivotGridValueType.Value Then
				e.DisplayText = "Export in progress, please wait..."
			End If
		End Sub


		Private Function GetLayoutDiscription(ByVal pivot As PivotGridControl, ByVal unnecessaryField As PivotGridField) As String
			Dim text As String = "ColumnFields: "
			Dim fields As List(Of PivotGridField) = pivot.GetFieldsByArea(PivotArea.ColumnArea)
			For i As Integer = 0 To fields.Count - 1
				If Object.ReferenceEquals(fields(i), unnecessaryField) Then
					Continue For
				End If
				text = text & fields(i).ToString()
				If i < fields.Count - 1 Then
					text = text & " -> "
				End If
			Next i
			Return text
		End Function


	End Class
End Namespace