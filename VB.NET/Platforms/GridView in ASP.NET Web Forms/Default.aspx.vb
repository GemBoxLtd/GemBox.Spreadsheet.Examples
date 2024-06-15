Imports GemBox.Spreadsheet
Imports System
Imports System.Data
Imports System.IO
Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class _Default
    Inherits System.Web.UI.Page

    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' To be able to save ExcelFile to PDF format in Medium Trust environment,
        ' you need to specify a font files location that is under your ASP.NET application's control.
        FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/")

        If Not Page.IsPostBack Then

            Dim people As New DataTable()
            people.Columns.Add("ID", Type.GetType("System.Int32"))
            people.Columns.Add("FirstName", Type.GetType("System.String"))
            people.Columns.Add("LastName", Type.GetType("System.String"))

            Session("people") = people

            Me.LoadDataFromFile(Request.PhysicalApplicationPath & "InputData.xlsx")
            Me.SetDataBinding()

        End If
    End Sub

    ''' <summary>
    ''' Export GridView data to Excel file.
    ''' </summary>
    Sub ExportData_Click(ByVal sender As Object, ByVal s As EventArgs)

        Dim people = DirectCast(Session("people"), DataTable)

        ' Create Excel file.
        Dim workbook As New ExcelFile()
        Dim worksheet = workbook.Worksheets.Add("DataSheet")

        ' Export DataTable that's used as GridView data source into Excel sheet.
        worksheet.InsertDataTable(people, New InsertDataTableOptions("A1") With {.ColumnHeaders = True})

        ' Stream Excel file to client's browser.
        workbook.Save(Me.Response, "Report." & Me.RadioButtonList1.SelectedValue)

    End Sub

    ''' <summary>
    ''' Export GridView data and formatting to Excel file.
    ''' </summary>
    Sub ExportDataAndFormatting_Click(ByVal sender As Object, ByVal s As EventArgs)

        Dim stringWriter As New StringWriter()
        Dim htmlWriter As New HtmlTextWriter(stringWriter)

        ' Export GridView control as HTML content.
        Me.GridView1.RenderControl(htmlWriter)

        Dim htmlOptions = LoadOptions.HtmlDefault
        Dim htmlData = htmlOptions.Encoding.GetBytes(stringWriter.ToString())

        Using htmlStream As New MemoryStream(htmlData)
            ' Load HTML into Excel file.
            Dim workbook = ExcelFile.Load(htmlStream, htmlOptions)

            ' Rename Excel sheet.
            Dim worksheet = workbook.Worksheets(0)
            worksheet.Name = "StyledDataSheet"

            ' Delete Excel column that has Delete and Edit buttons.
            worksheet.Columns.Remove(0)

            ' Stream Excel file to client's browser.
            workbook.Save(Me.Response, "Styled Report." & Me.RadioButtonList1.SelectedValue)
        End Using

    End Sub

    ' Override verification to successfully call GridView1.RenderControl method.
    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
    End Sub

    Sub LoadDataFromFile(ByVal fileName As String)

        Dim people = DirectCast(Session("people"), DataTable)

        ' Load Excel file.
        Dim workbook = ExcelFile.Load(fileName)
        Dim worksheet = workbook.Worksheets(0)

        ' Import Excel data into DataTable that's used as GridView data source.
        worksheet.ExtractToDataTable(people, New ExtractToDataTableOptions("A1", worksheet.Rows.Count))

    End Sub

    Sub SetDataBinding()

        Dim people = DirectCast(Session("people"), DataTable)
        Dim peopleDataView = people.DefaultView
        peopleDataView.AllowDelete = True

        Me.GridView1.DataSource = peopleDataView
        Me.GridView1.DataBind()

    End Sub

    Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)

        Dim people = DirectCast(Session("people"), DataTable)
        people.Rows(e.RowIndex).Delete()
        Me.SetDataBinding()

    End Sub

    Protected Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)

        Dim people = CType(Session("people"), DataTable)

        For i As Integer = 1 To people.Columns.Count
            Dim editTextBox = TryCast(Me.GridView1.Rows(e.RowIndex).Cells(i).Controls(0), TextBox)
            If editTextBox IsNot Nothing Then people.Rows(e.RowIndex)(i - 1) = editTextBox.Text
        Next

        Me.GridView1.EditIndex = -1
        Me.SetDataBinding()

    End Sub

    Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)

        Me.GridView1.EditIndex = e.NewEditIndex
        Me.SetDataBinding()

    End Sub

    Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)

        Me.GridView1.EditIndex = -1
        Me.SetDataBinding()

    End Sub

End Class