<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/WindowsApplication53/Form1.cs) (VB: [Form1.vb](./VB/WindowsApplication53/Form1.vb))
* [Program.cs](./CS/WindowsApplication53/Program.cs) (VB: [Program.vb](./VB/WindowsApplication53/Program.vb))
<!-- default file list end -->
# How to avoid production of unnecessary column and merging of cells in Excel export


<p>The XtraPivotGrid control is exported in a WYSIWYG (What You See Is What You Get) manner. According to this concept, items are exported as is, and this may cause an undesirable result. E.g. If column header width doesn't exactly match the field value width, the additional column will be produced. I'm afraid there is no way to avoid this problem in all cases. Let suppose that there are three <a href="http://documentation.devexpress.com/#WindowsForms/CustomDocument1711">Data Fields</a> and only two <a href="http://documentation.devexpress.com/#WindowsForms/CustomDocument1709">Row Field</a>. In this case, there is no any way to avoid cell merging without corrupting the layout.</p>
<p>Actually, we recommend turning off header export to avoid this problem. This can be done via the <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraPivotGridDataPivotGridOptionsPrint_PrintColumnHeaderstopic">PivotGridOptionsPrint.PrintColumnHeaders</a>, <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraPivotGridDataPivotGridOptionsPrint_PrintDataHeaderstopic">PivotGridOptionsPrint.PrintDataHeaders</a> and <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraPivotGridDataPivotGridOptionsPrint_PrintFilterHeaderstopic">PivotGridOptionsPrint.PrintFilterHeaders</a> properties. However, in some cases, field headers may be necessary to describe the current pivot grid layout.</p>
<p>This example demonstrates how to create an ancillary <a href="http://documentation.devexpress.com/#WindowsForms/CustomDocument1710">Column Field</a> that will have only one <a href="http://documentation.devexpress.com/#WindowsForms/CustomDocument1694">Field Value</a>, which can be used to export a layout description. The <a href="http://documentation.devexpress.com/#WindowsForms/DevExpressXtraPivotGridPivotGridControl_CustomExportFieldValuetopic">PivotGridControl.CustomExportFieldValue</a> event is used to provide a description to this field. Please note that you can generate any necessary string description by customizing the <strong>GetLayoutDiscription</strong> function.<br /><br /><strong>Update:</strong> starting with version 15.1, the new <a href="https://documentation.devexpress.com/#WindowsForms/CustomDocument1800">Data-aware Export</a> mode is available: <a href="https://community.devexpress.com/blogs/thinking/archive/2015/05/27/winforms-amp-asp-net-pivot-grid-controls-new-excel-data-export-engine.aspx">WinForms &amp; ASP.NET Pivot Grid Controls - New Excel Data Export Engine</a></p>

<br/>


