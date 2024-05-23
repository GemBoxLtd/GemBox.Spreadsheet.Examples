Imports GemBox.Spreadsheet
Imports System
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class MainForm

    Public Sub New()
        ' If using the Professional version, put your serial key below.
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")
        InitializeComponent()
    End Sub

    Private Async Sub LoadButton_Click(sender As Object, e As EventArgs) Handles LoadButton.Click
        ' Capture the current context on UI thread.
        Dim context = SynchronizationContext.Current

        ' Create load options.
        Dim loadOptions = New XlsxLoadOptions()
        AddHandler loadOptions.ProgressChanged,
            Sub(eventSender, args)
                Dim percentage = args.ProgressPercentage
                ' Invoke on UI thread.
                context.Post(
                    Sub(progressPercentage)
                        ' Update UI.
                        Me.ProgressBar.Value = CType(progressPercentage, Integer)
                        Me.PercentageLabel.Text = progressPercentage.ToString() & "%"
                    End Sub, percentage)
            End Sub

        Me.PercentageLabel.Text = "0%"
        ' Use tasks to run the load operation in a new thread.
        Await Task.Run(
            Sub()
                ExcelFile.Load("LargeFile.xlsx", loadOptions)
            End Sub)
    End Sub

End Class
