<?php
  // Create ComHelper object.
  $comHelper = new Com("GemBox.Spreadsheet.ComHelper", null, CP_UTF8);
  // If using the Professional version, put your serial key below.
  $comHelper->ComSetLicense("FREE-LIMITED-KEY");

 /********************
  *** Create Excel ***
  ********************/

  // Create new ExcelFile object.
  $workbook = new Com("GemBox.Spreadsheet.ExcelFile", null, CP_UTF8);
  // Add new ExcelWorksheet object.
  $worksheet = $workbook->Worksheets->Add("Sheet1");

  // Set width and format of column "A".
  $columnA = $comHelper->GetColumn($worksheet, 0);
  $columnA->Width = 20 * 256;
  $columnA->Style->Font->Weight = 700;

  // Set values of cells "A1", "A2", "A3" and "A4".
  $columnA->Cells->Item(0)->Value = "John Doe";
  $columnA->Cells->Item(1)->Value = "Bob Garvey";
  $columnA->Cells->Item(2)->Value = "Ben Stilwell";
  $columnA->Cells->Item(3)->Value = "Peter Pan";
  
  // Set values of cells "B1", "B2", "B3" and "B4".
  $columnB = $comHelper->GetColumn($worksheet, 1);
  $columnB->Cells->Item(0)->Value = 1000;
  $columnB->Cells->Item(1)->Value = 2000;
  $columnB->Cells->Item(2)->Value = 3000;
  $columnB->Cells->Item(3)->Value = 4000;

  // Create new Excel file.
  $workbook->Save(getcwd() . "\\New.xlsx");

 /******************
  *** Read Excel ***
  ******************/

  // Read existing Excel file.
  $book = $comHelper->Load(getcwd() . "\\New.xlsx");
  // Get first Excel sheet.
  $sheet = $book->Worksheets->Item(0);
  // Get first Excel row.
  $row1 = $comHelper->GetRow(sheet, 0);

  // Display values of cells "A1" and "B1".
  echo "Cell A1:" . $row1->Cells->Item(0)->Value;
  echo "<br>";
  echo "Cell B1:" . $row1->Cells->Item(1)->Value;

 /********************
  *** Update Excel ***
  ********************/

  // Update values of cells "A1" and "B1".
  $row1->Cells->Item(0)->Value = "Jane Doe";
  $row1->Cells->Item(1)->Value = 2000;

  // Write the updated Excel file.
  $book.Save(getcwd() . "\\Updated.xlsx");
?>