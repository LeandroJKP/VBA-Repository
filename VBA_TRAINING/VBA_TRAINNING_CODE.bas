Sub PROCESS_ROUTINE()               'This is macro that will automatize the process. It will, by calling other subroutines, to format the data inputted, update their values based on a database, make some calculations, filter data following or criteria, and report it as a PDF.

Call START_SETUP                    'This subroutine set the Excel configurations so that we can start working.

Call DATA_INPUT_FORMAT              'This subroutine will format the data, delete extra columns, fix the header, and the data format.

Call VLOOKUP_DATA_UPDATE            'This subroutine will make a VLOOKUP to find the sector that is responsible for the SKU.

Call DATA_CALCULATION               'This subroutine will calculate the inventory value os each SKU.

Call DATA_SELECTION                 'This subroutine will filter our report based on the value we choose in cells "C14" and "C15".

Call PDF_REPORT_GENERATOR           'This subroutine will save our report as a PDF. Save location and file name can be chosen in cells "C19" and "C20".

Call ENDING_SETUP                   'This subroutine set the Excel configurations back to normal.

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub DATA_INPUT_FORMAT()                                 'This subroutine will format the data, delete extra columns, fix the headset, and the data format.

Dim ws As Worksheet                                     'Defining the variable that will store the Worksheet we will be working.
Set ws = Workbooks("VBA_TRAINING").Sheets("Report")     'Assigning the value to our variable.

If IsEmpty(ws.Range("E1").Value) = True Then            'Cheking if column "E" is empty.
ws.Columns(5).EntireColumn.Delete                       'If "E" is empty, the line of code will delete it.
End If

With ws                                                 'This following lines will apply changes to the format and header of the sheet
    .Range("A1") = "SKU:"
    .Range("B1") = "DATE:"
    .Range("C1") = "AMOUNT:"
    .Range("D1") = "PRICE:"
    .Range("E1") = "DAYS OF SUPPLY:"
    .Range("F1") = "SECTOR:"
    .Range("G1") = "VALUE:"
    .Range("A:G").HorizontalAlignment = xlCenter        'This will centralize the columns that we are working whit.
    .Range("B:B").NumberFormat = "mm/dd/yy"             'This will fix the format so that will show us a date insted of a number.
    .Range("A1:G1").Font.Bold = True                    'This will change the font on the header to bold.

End With

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub VLOOKUP_DATA_UPDATE()                                                       'This subroutine will make a VLOOUP to find the sector that is responsible for the SKU.

Workbooks("VBA_TRAINING").Sheets("Report").Activate                             'This will acrivate the correct sheet.

Dim lastrow As Long                                                             'Defining the variable that will store the number of lines in column A.
lastrow = Range("A" & Rows.Count).End(xlUp).Row                                 'Assigning the value to our variable.

Application.Calculation = xCalculationManual                                    'Disable automatic calculations. If this was turned on, it would recalculate the whole sheet each time we change a cell in the loop resulting in a massive slowdown.

For Counter = 2 To lastrow                                                      'This is the loop that will apply on column "F", from line 2 until the last one, the VLOOKUP to find our information.
    Range("F" & Counter) = _
        "=VLOOKUP(RC[-5],'Database for VLOOKUP'!C[-4]:C[-5],2,0)"               'This is the VLOOKUP formula that we will use.
        
Next Counter                                                                    'Counter of the loop

Application.Calculation = xCalculationAutomatic                                 'Enable back automatic calculations, so all our VLOOKUP will resolve together and find us the values.

Workbooks("VBA_TRAINING").Sheets("Report").Range("F:F") = Range("F:F").Value    'This will kill the formula, by changing it to the value of the VLOOKUP. This way our data is no longer a formula and it is no longer linked to the source.

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub DATA_CALCULATION()                                                          'This subroutine will calculate the inventory value os each SKU.

Workbooks("VBA_TRAINING").Sheets("Report").Activate                             'This will acrivate the correct sheet.

Dim lastrow As Long                                                             'Defining the variable that will store the number of lines in column A.
lastrow = Range("A" & Rows.Count).End(xlUp).Row                                 'Assigning the value to our variable.

Application.Calculation = x1CalculationManual                                   'Disable automatic calculations. If this was turned on, it would recalculate the whole sheet each time we change a cell in the loop resulting in a massive slowdown.

For Counter = 2 To lastrow                                                      'This is the loop that gets the result in column "G", of the multiplication between the cells "C" by "D"
    Range("G" & Counter) = _
        "=RC[-4]*RC[-3]"
        
Next Counter                                                                    'Counter of the loop

Application.Calculation = xlCalculationAutomatic                                'Enable back automatic calculations, so all our VLOOKUP will resolve together and find us the values.

Workbooks("VBA_TRAINING").Sheets("Report").Range("G:G") = Range("G:G").Value    'This will kill the formula so our data is no longer a formula.

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub DATA_SELECTION()                                                            'This subroutine will filter our report based on the value we choose in cells "C14" and "C15".

Workbooks("VBA_TRAINING").Sheets("Report").Activate                             'This will acrivate the correct sheet.

Dim first_criteria As String                                                    'Defining the variable that will store the value of the first criteria.
Dim second_criteria As String                                                   'Defining the variable that will store the value of the second criteria.
first_criteria = Workbooks("VBA_TRAINING").Sheets("Example").Range("C14")       'Assigning the value to our variable based on cell "C14".
second_criteria = Workbooks("VBA_TRAINING").Sheets("Example").Range("C15")      'Assigning the value to our variable based on cell "C15".

Workbooks("VBA_TRAINING").Sheets("Report").Cells.AutoFilter                     'This will remove any filter currently applied
                                                                                'The next two lines will sort our data on descending order of values in column "G"
Range("A1:G1").Sort _
Key1:=Range("G1"), Order1:=xlDescending

With Workbooks("VBA_TRAINING").Sheets("Report").Range("A1:G1")                  'This will define the range in which columns we will filter.

.AutoFilter Field:=7, Criteria1:=">=" & first_criteria                          'Filter values that are greater or equal than the first_criteria.

.AutoFilter Field:=6, Criteria1:=second_criteria                                'Filter values that are qual than the second_criteria

End With

End Sub

Sub PDF_REPORT_GENERATOR()                                                          'This subroutine will save our report as a PDF. Save location and file name can be chosen in cells "C19" and "C20".

Dim saveLocation As String                                                          'Defining the variable that will store the save locarion.
saveLocation = Workbooks("VBA_TRAINING").Sheets("Example").Range("C21")             'Assigning the value to our variable based on cell "C21".
                                                                                    'Saving the file as a PDF.
Workbooks("VBA_TRAINING").Sheets("Report").ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=saveLocation

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub START_SETUP()                                       'This subroutine set the Excel configurations so that we can start working.

Application.ScreenUpdating = False                      'This will disable the visual update on Excel, making the process run faster.
Workbooks("VBA_TRAINING").Sheets("Report").Activate     'This will acrivate the correct sheet.

End Sub
_______________________________________________________________________________________________________________________________________________________________

Sub ENDING_SETUP()                                      'This subroutine set the Excel configurations back to normal.

Application.ScreenUpdating = True                       'This enables back the visual update on Excel.
Application.Calculation = xlCalculationAutomatic        'This enables back the automatic calculations.

End Sub
