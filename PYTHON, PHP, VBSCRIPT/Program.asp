<%
  ' Create ComHelper object.
  Set comHelper = CreateObject("GemBox.Spreadsheet.ComHelper")
  ' If using the Professional version, put your serial key below.
  comHelper.ComSetLicense("FREE-LIMITED-KEY")

  ''''''''''''''''''''
  ''' Create Excel '''
  ''''''''''''''''''''

  ' Create new ExcelFile object.
  Set workbook = CreateObject("GemBox.Spreadsheet.ExcelFile")
  ' Add new ExcelWorksheet object.
  Set worksheet = workbook.Worksheets.Add("Sheet1")

  ' Set width and format of column "A".
  Set columnA = comHelper.GetColumn(worksheet, 0)
  columnA.Width = 20 * 256
  columnA.Style.Font.Weight = 700

  ' Set values of cells "A1", "A2", "A3" and "A4".
  columnA.Cells.Item(0).Value = "John Doe"
  columnA.Cells.Item(1).Value = "Bob Garvey"
  columnA.Cells.Item(2).Value = "Ben Stilwell"
  columnA.Cells.Item(3).Value = "Peter Pan"
  
  ' Set values of cells "B1", "B2", "B3" and "B4".
  Set columnB = comHelper.GetColumn(worksheet, 1)
  columnB.Cells.Item(0).Value = 1000
  columnB.Cells.Item(1).Value = 2000
  columnB.Cells.Item(2).Value = 3000
  columnB.Cells.Item(3).Value = 4000

 ' Create new Excel file.
  workbook.Save(Server.MapPath(".") & "\New.xlsx")

  ''''''''''''''''''
  ''' Read Excel '''
  ''''''''''''''''''

  ' Read existing Excel file.
  Set book = comHelper.Load(Server.MapPath(".") & "\New.xlsx")
  ' Get first Excel sheet.
  Set sheet = book.Worksheets.Item(0)
  ' Get first Excel row.
  Set row1 = comHelper.GetRow(sheet, 0)

  ' Display values of cells "A1" and "B1".
  Response.Write("Cell A1:" & row1.Cells.Item(0).Value)
  Response.Write("<br>")
  Response.Write("Cell B1:" & row1.Cells.Item(1).Value)

  ''''''''''''''''''''
  ''' Update Excel '''
  ''''''''''''''''''''

  ' Update values of cells "A1" and "B1".
  row1.Cells.Item(0).Value = "Jane Doe"
  row1.Cells.Item(1).Value = 2000

  ' Write the updated Excel file.
  book.Save(Server.MapPath(".") & "\Updated.xlsx")
%>