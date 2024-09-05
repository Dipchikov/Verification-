Attribute VB_Name = "Telegram_aspect_inependent"
Option Explicit
Type TelegramData
    'AspectName As String
    Sign As String
    Variable As String
    Packet As String
    value As Integer
    TrackCondition As String
End Type

Sub CheckTelegramConditions()
    Dim folderPath As String
    Dim rs As Worksheet
    Dim ws As Worksheet
    Dim row As Long
    Dim fso As Object

    
'On Error Resume Next
Set rs = ThisWorkbook.Sheets("Results")
    If rs.AutoFilterMode Then
        ' Remove filters
        rs.AutoFilterMode = False
    End If

    ThisWorkbook.Sheets("Results").Rows("2:1048539").Delete
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then ' if OK is pressed
            folderPath = .SelectedItems(1)
        End If
    End With
    
    Set ws = ThisWorkbook.Sheets(1) ' Use the first sheet in the workbook
    row = 1 ' Start writing data from the first row

    ' Create File System Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Start processing the folder and its subfolders
    ProcessFolder fso.GetFolder(folderPath), ws, row
    
    ThisWorkbook.Sheets("Results").Activate
    
    MsgBox "Extraction completed!", vbInformation
End Sub

Sub ProcessFolder(ByVal folder As Object, ByRef ws As Worksheet, ByRef row As Long)

    Dim file As Object
    Dim fileNum As Integer
    Dim fileContent As String
    Dim line As String
    Dim insideTelegram As Boolean
    Dim nidPacketFound As Boolean
    Dim nidPacketIter  As Boolean
    Dim vLoaValue As Long 'String
    Dim rs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim AspectDict As Object
    Dim AspectValue As String
    Dim key As Variant
    Dim bal_group_name As String
    Dim bal_group_id As String
    Dim bal_group_aspect As String
    Dim bal_group_aspect_name As String
    Dim print_row As Long
    Dim lastrow_print As Long
    Dim Q_DIR As String
    Dim infill As String
    Dim BAL_TYPE As String
    Dim M_TRACKCOND As String
    Dim tempVariableName As String

    
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
  ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Configuration")
    Set rs = ThisWorkbook.Sheets("Results")
    
        rs.Cells(1, 1).value = "Aspect [ID](Telegram number)"
        rs.Cells(1, 2).value = "Trans. Point (ID)[Chennel]"
        rs.Cells(1, 3).value = "Direction"
        rs.Cells(1, 4).value = "Infil"
        rs.Cells(1, 5).value = "Packet"
        rs.Cells(1, 6).value = "Variable"
        rs.Cells(1, 7).value = "Extracted Value"
        rs.Cells(1, 8).value = "Result"
        rs.Cells(1, 9).value = "Violation"
         rs.Cells(1, 10).value = "Dir"
    'print_row = 2


    ' Create a new dictionary
    Set AspectDict = CreateObject("Scripting.Dictionary")
    
        ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).row
    print_row = rs.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    
    ' Declare the array with 1-based indexing to match the row numbers
    'Dim TelegramArray(1 To 4) As TelegramData
    ReDim TelegramArray(2 To lastRow) As TelegramData
    
     For j = 2 To lastRow
            'TelegramArray(j).AspectName = ws.Cells(j, 1).Value  '
            If ws.Cells(j, 10).value <> "" Then
                TelegramArray(j).Packet = "NID_PACKET=" & ws.Cells(j, 10).value
                    Else
                TelegramArray(j).Packet = ""
            End If
            If ws.Cells(j, 10).value = "" Then
                TelegramArray(j).Variable = GetValueFromDictionary(ws.Cells(j, 11).value)
                    Else
                TelegramArray(j).Variable = RemoveSpaces(ws.Cells(j, 11).value)
            End If
            TelegramArray(j).Sign = ws.Cells(j, 12).value
            TelegramArray(j).value = ws.Cells(j, 13).value
            TelegramArray(j).TrackCondition = ws.Cells(j, 14).value
    Next j
    
    
    ' Loop through each file in the folder
    For Each file In folder.Files
        If Right(file.Name, 4) = ".sdi" Or Right(file.Name, 4) = ".bdi" Then  ' Adjust file type if necessary
            fileNum = FreeFile
            Open file.Path For Input As #fileNum
            fileContent = ""
            
            ' Read the entire content of the file
            Do While Not EOF(fileNum)
                Line Input #fileNum, line
                fileContent = fileContent & line & vbCrLf
            Loop
            Close #fileNum
            
            ' Initialize flags
            insideTelegram = False
            nidPacketFound = False
            
            ' Split the content into lines for easier processing
            Dim lines() As String
            lines = Split(fileContent, vbCrLf)



            If Right(file.Name, 4) = ".bdi" And lines(7) = "BAL_TYPE=1" Then
            AspectDict.Add "TEL_ASPECT_NAME(1)=Balise default[128]", "(" & ExtractAspectName(lines(4)) & ")"
            bal_group_aspect_name = "Balise default[128]"
            bal_group_id = "(" & ExtractAspectName(lines(4)) & ")"
            End If
            

'------------------------------new code -----------------------------------------------------------------
            ' Loop through each line in the file

            For i = LBound(lines) To UBound(lines)
                        ' Loop through each line in array
            For j = 2 To lastRow
                 If InStr(lines(i), "BAL_GROUP_NAME") > 0 Then
                         bal_group_name = ExtractAspectName(lines(i))
              End If
               If InStr(lines(i), "TEL_ASPECT_NAME") > 0 Then
                    If Not AspectDict.Exists(lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]") Then
                        AspectDict.Add lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]", "(" & ExtractAspectName(lines(i - 3)) & ")[" & ExtractAspectName(lines(i - 2) & "]")
                    End If
                End If
            
               Next j
            Next i


'----------------------------------------------------------------------------------------------



            For j = 2 To lastRow
                    infill = "No"
                    BAL_TYPE = ""
                    vLoaValue = 0
                    
                    
                    If TelegramArray(j).Packet = "" Then
                    nidPacketFound = True
                    End If
                    
               For i = LBound(lines) To UBound(lines)
                        ' Loop through each line in array
                        
                    If InStr(lines(i), "BAL_GROUP_NAME") > 0 Then
                         bal_group_name = ExtractAspectName(lines(i))
                     End If
                     
'                    If InStr(lines(i), "BAL_GROUP_ID") > 0 Then
'                        If InStr(lines(i + 1), "TEL_LEU_CHANNEL") > 0 Then
'                         bal_group_id = "(" & ExtractAspectName(lines(i)) & ")" & "[" & ExtractAspectName(lines(i + 1)) & "]"
'                         Else
'                         bal_group_id = "(" & ExtractAspectName(lines(i)) & ")"
'                         End If
'                     End If
                    
                    Dim myVar As String
                     If InStr(lines(i), "Q_DIR") > 0 Then
                         myVar = ExtractValue(lines(i))
                         Q_DIR = TelegramDirection(myVar)
                     End If
                    '
                     If InStr(lines(i), "BAL_TYPE") > 0 Then
                     BAL_TYPE = BaliseType(ExtractAspectName(lines(i)))
                     End If
                    
                    '----------------begin telegram-------------------------
                   If InStr(lines(i), "BEGIN_TELEGRAM") > 0 And AspectDict.Count > 0 Then
                         bal_group_aspect = "TEL_ASPECT_NAME(" & ExtractTextBetweenParentheses(lines(i)) & ")"
                         'ExtractAspectName(lines(i))
                         For Each key In AspectDict.keys
                            If InStr(key, bal_group_aspect) > 0 Then
                                bal_group_aspect_name = ExtractAspectName(key) & "(" & ExtractTextBetweenParentheses(key) & ")"
                                bal_group_id = AspectDict(key)
                            End If
                         Next key
                         
                         For z = i To UBound(lines)
                         infill = "No"
                             If InStr(lines(z), "136(") > 0 Then
                                infill = "Yes"
                            End If
                            If InStr(lines(z), TelegramArray(j).Packet) > 0 Then
                               Exit For
                            End If
                        Next z
                         
                     End If
                     
                    If InStr(lines(i), "BEGIN_TELEGRAM") > 0 And AspectDict.Count = 0 Then
                         bal_group_aspect_name = BAL_TYPE
                    End If
            
            
                
                ' Check for NID_PACKET=12 and if we are inside BEGIN_TELEGRAM(1)
                If InStr(lines(i), TelegramArray(j).Packet) > 0 Then
                    nidPacketFound = True
                End If
                
                If InStr(lines(i), "D_TRACKCOND=") > 0 Then
                        nidPacketIter = True
                End If
                
                ' Check for V_LOA=0 if NID_PACKET=12 was found
                If nidPacketFound And InStr(lines(i), TelegramArray(j).Variable) > 0 Then
                     tempVariableName = ExtractVariable(lines(i))
                
                If nidPacketIter = True And InStr(TelegramArray(j).Variable, "D_TRACKCOND") > 0 Then    '
                    vLoaValue = CInt(ExtractValue(lines(i))) + vLoaValue
                    M_TRACKCOND = "[" & ExtractValue(lines(i + 2)) & "]"
                    Else
                   vLoaValue = ExtractValue(lines(i))
                End If
                
                 If nidPacketIter = True And InStr(TelegramArray(j).Variable, "L_TRACKCOND") > 0 Then    '
                    M_TRACKCOND = "[" & ExtractValue(lines(i + 1)) & "]"
                End If
                
           If TelegramArray(j).Packet <> "NID_PACKET=68" Then
                 M_TRACKCOND = ""
              End If
                    
                    
                    Select Case TelegramArray(j).Sign
                            Case "="
                                If Trim(vLoaValue) = TelegramArray(j).value Then
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                         rs.Cells(print_row, 8).value = "SUCCESS"
                                         'rs.Cells(print_row, 5).value = "" 'bal_group_aspect_name
                                         rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                         print_row = print_row + 1
                                        Else
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                        rs.Cells(print_row, 8).value = "FAIL"
                                        rs.Cells(print_row, 9).value = TelegramArray(j).Variable & M_TRACKCOND & " value " & vLoaValue & " (expected " & TelegramArray(j).value & ")."    'bal_group_aspect_name & " Direction " & Q_DIR & " - Wrong " &
                                        rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                        print_row = print_row + 1
                                 End If
                        
                            Case "<="
                                If Trim(vLoaValue) <= TelegramArray(j).value Then
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                         rs.Cells(print_row, 8).value = "SUCCESS"
                                         rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                         print_row = print_row + 1
                                          Else
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                        rs.Cells(print_row, 8).value = "FAIL"
                                        rs.Cells(print_row, 9).value = TelegramArray(j).Variable & M_TRACKCOND & " value " & vLoaValue & " (expected <= " & TelegramArray(j).value & ")."   'bal_group_aspect_name & " Direction " & Q_DIR & " - Wrong " &
                                        rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                        print_row = print_row + 1
                                End If
                        
                            Case ">="
                                If Trim(vLoaValue) >= TelegramArray(j).value Then
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                         rs.Cells(print_row, 8).value = "SUCCESS"
                                         rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                         print_row = print_row + 1
                                        Else
                                        rs.Cells(print_row, 1).value = bal_group_aspect_name
                                        rs.Cells(print_row, 2).value = bal_group_name & bal_group_id
                                        rs.Cells(print_row, 3).value = Q_DIR
                                        rs.Cells(print_row, 4).value = infill
                                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet & M_TRACKCOND
                                        rs.Cells(print_row, 6).value = tempVariableName & "=" & TelegramArray(j).value 'TelegramArray(j).Variable & "=" & TelegramArray(j).value
                                        rs.Cells(print_row, 7).value = vLoaValue
                                        rs.Cells(print_row, 8).value = "FAIL"
                                        rs.Cells(print_row, 9).value = TelegramArray(j).Variable & M_TRACKCOND & " value " & vLoaValue & " (expected >= " & TelegramArray(j).value & ")." 'bal_group_aspect_name & " Direction " & Q_DIR & " - Wrong " &
                                        rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                                        print_row = print_row + 1
                                End If
                        
                            ' Add more cases if needed for other operators
                        End Select
                                            
                    
                    

                End If
                
                ' Reset flags when encountering a new BEGIN_TELEGRAM
                If InStr(lines(i), "NID_PACKET") > 0 And InStr(lines(i), TelegramArray(j).Packet) = 0 Then
                    'insideTelegram = False
                    nidPacketFound = False
                    nidPacketIter = False
                    vLoaValue = 0
                    M_TRACKCOND = ""
                End If
            Next i
            nidPacketFound = False
            tempVariableName = ""
            Next j

    End If
    
    
        'Clear the ids
        bal_group_name = ""
        bal_group_id = ""
        bal_group_aspect_name = ""
        'Clear the dictionary
                 
        AspectDict.RemoveAll
    
    Next file
 

    
    ' Loop through each subfolder in the current folder
    Dim subFolder As Object
    For Each subFolder In folder.Subfolders
        ProcessFolder subFolder, ws, row ' Recursive call for subfolders
    Next subFolder
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub



