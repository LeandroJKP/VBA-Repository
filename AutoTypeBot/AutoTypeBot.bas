Public SKU As Long
Public Quant As Long
Public Price As Double
Public Comment As String
Public TypeOfSolicitation As Byte
Public Motif As Byte
Public Adress As String
Public FirstItem As Integer
Public LastItem As Integer
Public Msg As String
Public Counter As Long
Public Speed As String
_______________________________________________________________________________________________________________________

Sub AutoTypeBot()


Call Parameters                                        'Subrotine to get values to the variable that we will use

Call ErrorCheck                                        'Subrotine to check for erros on the information


AppActivate "Name of Application", True                'Activate the application
     
    
For Counter = FirstItem To LastItem                    'Loop that goes from the first item to the last
    
    Call LineScan                                      'Subroutine to get the values of the current line we want to type

    Call TypyingSCinApp                                'Subroutine that will type the values on selected program
    
    'Call Debugging                                    'Subroutine for Debug
    
 
Next Counter                                           'End of the loop


Call TypingCompleted                                   'Subroutine that send the "Typing Complete" message and set the ending configuration


End Sub
_______________________________________________________________________________________________________________________

Sub Parameters()

Application.ScreenUpdating = False

FirstItem = Range("B3").Value + 2                      'Defines from which line the typing will start
LastItem = Cells(Cells.Rows.Count, "E").End(xlUp).Row  'Identifies the last filled cell

TypeOfSolicitation = Range("B6").Value                 'Defines the type of solicitation
Motif = Range("B7").Value                              'Defines the motif of solicitation
Adress = Range("B8").Value                             'Defines the sector code of solicitation
Speed = Range("B11").Value                             'Defines the speed that the bot will type

End Sub
_______________________________________________________________________________________________________________________

Sub WaitTime()                                          'Set the speed of typing

If StrComp(Speed, "No") = 0 Then
    Application.Wait Now + #12:00:01 AM#
    
Else
    Application.Wait Now + 1 / (24 * 60 * 60# * 2)
    
    End If
    
End Sub
_______________________________________________________________________________________________________________________

Sub LineScan()                                         'Subroutine that grab the value of each line

SKU = Range("E" & Counter).Value                       'Variable that stores the code of the product
Comment = Range("F" & Counter).Value                   'Variable that stores the comment
Quant = Range("G" & Counter).Value                     'Variable that stores the quantity
Price = Range("H" & Counter).Value                     'Variable that stores the price

End Sub
_______________________________________________________________________________________________________________________

Sub Debugging()                                        'Subroutine for Debuging

    Debug.Print (FirstItem)
    
    Debug.Print (LastItem)
    
    Debug.Print (SKU)
 
End Sub
_______________________________________________________________________________________________________________________

Sub ErrorCheck()                                      'Subroutine for error check


If FirstItem > LastItem Then                          'Check if the there is any item to be typed after the selected starting line

Dim Msg As String
Msg = "The last consecutive item is on the line " & LastItem - 2 & " , the typing can not start until this is addressed. " & FirstItem - 2 & "." & vbNewLine & vbNewLine & "Please change the starting line on cell B3"
MsgBox Msg, 0, "Error - Starting line after last line"
Stop

End If

Dim SKUquantityZero As Long

For i = FirstItem To LastItem                         'End of the check


SKUquantityZero = Range("G" & i).Value                'Check to see if there are cells with a value equal to zero or blank

If SKUquantityZero = 0 Or IsEmpty(SKUquantityZero) = True Then
Msg = "The item on line " & i & " is zero/blank." & vbNewLine & vbNewLine & "Please insert a value in column G to proceed."
MsgBox Msg, 0, "Error - Item quantity is equal zero"
Stop

End If

Next i                                                 'End of the check

Range("J20").Select                                    'Select a harmless cell, so if the Appactivatethe fails for some reason the bot will write on empty cells

End Sub
_______________________________________________________________________________________________________________________

Sub CleanProducts()                                     'Subroutine to clean the Sheet


Worksheets("Front").Range("E3:G502").ClearContents
   
    If ActiveSheet.AutoFilterMode Then
 
        ActiveSheet.AutoFilterMode = False
 
    End If

End Sub
_______________________________________________________________________________________________________________________

Sub TypingCompleted()                                  'Subroutine for the final setup and complete message


Application.Wait Now + #12:00:01 AM#

Application.SendKeys "{NUMLOCK}%s"                     'Enable Numpad

SendKeys ("%{TAB}")                                    'Alt+Tab to go back to Excel

Application.ScreenUpdating = True                      'Enable ScreenUpdating

If LastItem - FirstItem + 1 = 1 Then                   'Ending message for one item

Msg = "Typing of solicitation with 1 item finished."
MsgBox Msg, 0, "Task completed "

Else                                                   'Ending message for more than one item

Msg = "Typing of solicitation with " & LastItem - FirstItem + 1 & " items ."
MsgBox Msg, 0, "Task completed"

End If


End Sub
_______________________________________________________________________________________________________________________

Sub TypyingSCinApp()                                    'Subroutine that will send the commands and values

    SendKeys ("{ENTER}")
    
    Application.Wait Now + #12:00:01 AM#
    
    SendKeys (SKU)
    
    Application.Wait Now + #12:00:01 AM#
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys (Comment)

    Call WaitTime
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
      
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys (Quant)
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys (Price)
    
    SendKeys ("{ENTER}")
    
    SendKeys ("{RIGHT}")
      
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys ("0")
    SendKeys (TypeOfSolicitation)
    
 
    SendKeys ("{RIGHT}")
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys (Motif)
          
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    SendKeys ("{RIGHT}")
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys (Adress)
    
    SendKeys ("{ENTER}")
    
    Call WaitTime
    
    SendKeys ("{DOWN}")
    SendKeys ("{RIGHT}")
    
    Application.Wait Now + #12:00:01 AM#


End Sub
