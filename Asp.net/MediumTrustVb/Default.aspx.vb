Imports System.IO
Imports System.Xml
Imports System.Text
Imports System.Data
Imports GemBox.Spreadsheet

Public Class _Default
    Inherits System.Web.UI.Page

    Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' By specifying a location that is under ASP.NET application's control, 
        ' GemBox.Spreadsheet can use file system operations to retrieve font data even in Medium Trust environment.
        FontSettings.FontsBaseDirectory = Server.MapPath("Fonts/")

        If Not Page.IsPostBack Then
            Dim people As DataTable = New DataTable()

            people.Columns.Add("ID", Type.GetType("System.Int32"))
            people.Columns.Add("FirstName", Type.GetType("System.String"))
            people.Columns.Add("LastName", Type.GetType("System.String"))

            Session("people") = people

            Me.LoadDataFromFile(Request.PhysicalApplicationPath & "InData.xlsx")
            Me.SetDataBinding()
        End If

    End Sub

    Sub Export_Click(ByVal sender As Object, ByVal s As EventArgs)

        Dim people As DataTable = DirectCast(Session("people"), DataTable)

        ' Create excel file.
        Dim ef As ExcelFile = New ExcelFile
        ef.Styles.Normal.Font.Name = "Calibri"
        Dim ws As ExcelWorksheet = ef.Worksheets.Add("DataSheet")
        ws.InsertDataTable(people, New InsertDataTableOptions(0, 0) With {.ColumnHeaders = True})

        ' Stream file to browser
        ef.Save(Me.Response, "Report." & Me.RadioButtonList1.SelectedValue)

    End Sub

    Sub LoadDataFromFile(ByVal fileName As String)

        Dim ef As ExcelFile = ExcelFile.Load(fileName)
        Dim ws As ExcelWorksheet = ef.Worksheets(0)

        Dim people As DataTable = DirectCast(Session("people"), DataTable)
        ws.ExtractToDataTable(people, New ExtractToDataTableOptions("A1", ws.Rows.Count))

    End Sub

    Sub SetDataBinding()

        Dim people As DataTable = DirectCast(Session("people"), DataTable)
        Dim peopleDataView As DataView = people.DefaultView

        Me.GridView1.DataSource = peopleDataView
        peopleDataView.AllowDelete = True
        Me.GridView1.DataBind()

    End Sub

    Sub GridView1_RowDeleting(ByVal sender As Object, ByVal e As GridViewDeleteEventArgs)

        Dim people As DataTable = DirectCast(Session("people"), DataTable)

        people.Rows(e.RowIndex).Delete()
        Me.SetDataBinding()

    End Sub

    Sub GridView1_RowEditing(ByVal sender As Object, ByVal e As GridViewEditEventArgs)

        Me.GridView1.EditIndex = e.NewEditIndex
        Me.SetDataBinding()

    End Sub

    Sub GridView1_RowUpdating(ByVal sender As Object, ByVal e As GridViewUpdateEventArgs)

        Dim i As Integer
        Dim rowIndex As Integer = e.RowIndex
        Dim people As DataTable = Session("people")

        For i = 1 To people.Columns.Count
            Dim editTextBox As TextBox = TryCast(Me.GridView1.Rows(rowIndex).Cells(i).Controls(0), TextBox)

            If Not (editTextBox Is Nothing) Then
                people.Rows(rowIndex)(i - 1) = editTextBox.Text
            End If
        Next

        Me.GridView1.EditIndex = -1
        Me.SetDataBinding()

    End Sub

    Sub GridView1_RowCancelingEdit(ByVal sender As Object, ByVal e As GridViewCancelEditEventArgs)

        Me.GridView1.EditIndex = -1
        Me.SetDataBinding()

    End Sub

End Class