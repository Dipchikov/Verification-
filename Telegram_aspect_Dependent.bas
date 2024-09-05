Attribute VB_Name = "Telegram_aspect_Dependent"
Option Explicit
Type TelegramData
    AspectName As String
    AspectNumber  As String
    Variable As String
    Packet As String
    value As Integer
    Chennel As String
    row As String
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
    ' Set the folder path where the files are located
    'folderPath = "C:\DPVT_Help\project\rsi-bph_01.00_06_unrel" ' Change this to your folder path
    
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
'------------------------------version 2 ------------------23.08.2024----------------------------------
    Dim file As Object
    Dim fileNum As Integer
    Dim fileContent As String
    Dim line As String
    Dim insideTelegram As Boolean
    Dim nidPacketFound As Boolean
    Dim TelegramMissing As Boolean
    Dim nidPacketMissing As Boolean
    Dim nidPacketVariable As Boolean
    Dim vLoaValue As String
    Dim rs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim z As Long
    Dim AspectDict As Object
    Dim MissingPacketDict As Object
    Dim AspectValue As String
    Dim AspectName As String
    Dim key As Variant
    Dim bal_group_name As String
    Dim bal_group_id As String
    Dim print_row As Long
    Dim Q_DIR As String
    Dim infill As String
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
  ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Configuration")
    Set rs = ThisWorkbook.Sheets("Results")
    
        rs.Cells(1, 1).value = "Aspect[ID](Telegram Number)"
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
    Set MissingPacketDict = CreateObject("Scripting.Dictionary")

        ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).row
    print_row = rs.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ' Declare the array with 1-based indexing to match the row numbers
    'Dim TelegramArray(1 To 4) As TelegramData
    ReDim TelegramArray(1 To lastRow) As TelegramData
    ReDim TelegramArray(1 To lastRow) As TelegramData
    
     For j = 2 To lastRow
            TelegramArray(j).AspectName = ws.Cells(j, 1).value  '
            TelegramArray(j).AspectNumber = ws.Cells(j, 2).value  '
            TelegramArray(j).Packet = "NID_PACKET=" & ws.Cells(j, 3).value
            TelegramArray(j).Variable = RemoveSpaces(ws.Cells(j, 4).value)
            TelegramArray(j).value = ws.Cells(j, 5).value
            TelegramArray(j).Chennel = UCase(ws.Cells(j, 6).value)
       
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
            TelegramMissing = False
            nidPacketMissing = False
            
            ' Split the content into lines for easier processing
            Dim lines() As String
            lines = Split(fileContent, vbCrLf)

            If Right(file.Name, 4) = ".bdi" And lines(7) = "BAL_TYPE=1" Then
            AspectDict.Add "TEL_ASPECT_NAME(1)=Balise default[128]", "(" & ExtractAspectName(lines(4)) & ")"
            insideTelegram = True
            End If
            
              If Right(file.Name, 4) = ".bdi" And lines(7) <> "BAL_TYPE=1" Then
                 GoTo Skip_file '
            End If
            
            ' Loop through each line in the file

            For i = LBound(lines) To UBound(lines)
                        ' Loop through each line in array
                    For j = 2 To lastRow
                             If InStr(lines(i), "BAL_GROUP_NAME") > 0 Then
                                         bal_group_name = ExtractAspectName(lines(i))
                              End If
                               ' If InStr(lines(i), TelegramArray(j).AspectName) > 0 Then
                                AspectName = ExtractAspectName(lines(i))
                
                                
                                If AspectName = TelegramArray(j).AspectName And TelegramArray(j).AspectName <> "" And TelegramArray(j).AspectNumber = "" Then
                                'If TelegramArray(j).Chennel = ExtractAspectName(lines(i - 2)) Then
                                        If Not AspectDict.Exists(lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]") Then
                                            AspectDict.Add lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]", "(" & ExtractAspectName(lines(i - 3)) & ") [" & ExtractAspectName(lines(i - 2)) & "]"
                                        End If
                                   'End If
                                End If
                                
                               If InStr(lines(i), "TEL_ASPECT_NR") > 0 And ExtractAspectName(lines(i)) = TelegramArray(j).AspectNumber And TelegramArray(j).AspectName = "" Then
                                ' If TelegramArray(j).Chennel = ExtractAspectName(lines(i - 1)) Then
                                        If Not AspectDict.Exists(lines(i + 1) & "[" & ExtractAspectName(lines(i)) & "]") Then
                                            AspectDict.Add lines(i + 1) & "[" & ExtractAspectName(lines(i)) & "]", "(" & ExtractAspectName(lines(i - 2)) & ") [" & ExtractAspectName(lines(i - 1)) & "]"
                                        End If
                                    'End If
                                End If
                                
                                 If TelegramArray(j).AspectNumber <> "" And TelegramArray(j).AspectName <> "" And InStr(lines(i), "TEL_ASPECT_NR") > 0 Then
                                     If ExtractAspectName(lines(i + 1)) & "[" & ExtractAspectName(lines(i)) & "]" = TelegramArray(j).AspectName & "[" & TelegramArray(j).AspectNumber & "]" Then
                                    ' If TelegramArray(j).Chennel = ExtractAspectName(lines(i - 1)) Then
                                            If Not AspectDict.Exists(lines(i + 1) & "[" & ExtractAspectName(lines(i)) & "]") Then
                                                AspectDict.Add lines(i + 1) & "[" & ExtractAspectName(lines(i)) & "]", "(" & ExtractAspectName(lines(i - 2)) & ") [" & ExtractAspectName(lines(i - 1)) & "]"
                                            End If
                                        'End If
                                      End If
                                End If

                       Next j
            Next i
            

            
          '-------------------------------loop under aspects------------------------------------------------------
 Dim TempAspectName As String
 For j = 2 To lastRow
            
If TelegramArray(j).AspectName <> "" Or TelegramArray(j).AspectNumber <> "" Then

            For Each key In AspectDict.keys
     
                    bal_group_id = AspectDict(key)
                                 
                    If TelegramArray(j).AspectName = "" And TelegramArray(j).AspectNumber <> "" Then
                        If InStr(key, "[" & TelegramArray(j).AspectNumber & "]") > 0 Then
                            TempAspectName = ExtractAspectName(key)
                            AspectValue = ExtractTextBetweenParentheses(key)
                        End If
                    End If
                    
                    If TelegramArray(j).AspectName <> "" And TelegramArray(j).AspectNumber = "" Then
                        If InStr(key, TelegramArray(j).AspectName) > 0 Then
                            TempAspectName = ExtractAspectName(key)
                            AspectValue = ExtractTextBetweenParentheses(key)
                        End If
                    End If
                    
                If TelegramArray(j).AspectName <> "" And TelegramArray(j).AspectNumber <> "" Then
                        If InStr(key, TelegramArray(j).AspectName & "[" & TelegramArray(j).AspectNumber & "]") > 0 Then
                            TempAspectName = ExtractAspectName(key)
                            AspectValue = ExtractTextBetweenParentheses(key)
                        End If
                    End If
            
            
            
     '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            
            
            
             If InStr(key, TempAspectName) > 0 And TempAspectName <> "" Then
            
                infill = "No"
                
                For i = LBound(lines) To UBound(lines)
                
                
                    If Right(file.Name, 4) = ".bdi" Then
                        insideTelegram = True
                        TelegramMissing = True
                    End If
                        
                     
'                    If InStr(lines(i), "BAL_GROUP_ID") > 0 Then
'                         bal_group_id = ExtractAspectName(lines(i))
'                     End If
                
          ' Loop through each line in array
          



                ' Check for BEGIN_TELEGRAM(1)
                If InStr(lines(i), "BEGIN_TELEGRAM(" & AspectValue & ")") > 0 Then
                    insideTelegram = True
                    TelegramMissing = True
                    
                    For z = i To UBound(lines)
                    
                      If InStr(lines(z), "136(") > 0 Then
                         infill = "Yes"
                     End If
                        If insideTelegram And InStr(lines(z), TelegramArray(j).Packet) > 0 Then
                           Exit For
                        End If
                    Next z
                    
                End If
                
                ' Check for NID_PACKET=12 and if we are inside BEGIN_TELEGRAM(1)
                If insideTelegram And InStr(lines(i), TelegramArray(j).Packet) > 0 Then
                    nidPacketFound = True
                    nidPacketMissing = True
                End If
                
                     Dim myVar As String
                     If InStr(lines(i), "Q_DIR") > 0 Then
                         myVar = ExtractValue(lines(i))
                         Q_DIR = TelegramDirection(myVar)
                     End If

                    If InStr(TelegramArray(j).AspectName, ExtractAspectKeyName(key)) > 0 Or InStr(ExtractAspectName(key), "[" & TelegramArray(j).AspectNumber & "]") > 0 Then
                         nidPacketVariable = True
                    End If


                
                ' Check for V_LOA=0 if NID_PACKET=12 was found
                If nidPacketFound And InStr(lines(i), TelegramArray(j).Variable) > 0 And nidPacketVariable Then
                
                        vLoaValue = ExtractValue(lines(i))
                        If Trim(vLoaValue) = TelegramArray(j).value Then
                            ' Write the result in the Excel sheet
                            rs.Cells(print_row, 1).value = ExtractAspectName(key) & "(" & ExtractTextBetweenParentheses(key) & ")"
                            rs.Cells(print_row, 2).value = bal_group_name & " " & bal_group_id
                            rs.Cells(print_row, 3).value = Q_DIR
                             rs.Cells(print_row, 4).value = infill
                            rs.Cells(print_row, 5).value = TelegramArray(j).Packet
                            rs.Cells(print_row, 6).value = TelegramArray(j).Variable & "=" & TelegramArray(j).value
                            rs.Cells(print_row, 7).value = vLoaValue
                            rs.Cells(print_row, 8).value = "SUCCESS"
                            rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                            print_row = print_row + 1
                            Else
                            rs.Cells(print_row, 1).value = ExtractAspectName(key) & "(" & ExtractTextBetweenParentheses(key) & ")"
                            rs.Cells(print_row, 2).value = bal_group_name & " " & bal_group_id
                            rs.Cells(print_row, 3).value = Q_DIR
                            rs.Cells(print_row, 4).value = infill
                            rs.Cells(print_row, 5).value = TelegramArray(j).Packet
                            rs.Cells(print_row, 6).value = TelegramArray(j).Variable & "=" & TelegramArray(j).value
                            rs.Cells(print_row, 7).value = vLoaValue
                            rs.Cells(print_row, 8).value = "FAIL"
                            rs.Cells(print_row, 9).value = "Balise " & bal_group_name & " " & bal_group_id & "  " & "aspect " & ExtractAspectName(key) & " " & TelegramArray(j).Packet & " - Wrong " & TelegramArray(j).Variable & " value " & vLoaValue & " (expected " & TelegramArray(j).value & ")."
                            rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                            print_row = print_row + 1
                        End If
                    
                End If
                
                        ' Reset flags when encountering a new BEGIN_TELEGRAM
                        If InStr(lines(i), "BEGIN_TELEGRAM") > 0 And InStr(lines(i), "BEGIN_TELEGRAM(" & AspectValue & ")") = 0 And insideTelegram = True Then
                            insideTelegram = False
                            nidPacketFound = False
                            infill = "No"
                            If Right(file.Name, 4) = ".sdi" Then
                            Exit For
                            End If
                        End If
                        
                          If InStr(lines(i), "NID_PACKET") > 0 And InStr(lines(i), TelegramArray(j).Packet) = 0 Then
                            nidPacketFound = False
                            Q_DIR = ""
                        End If
                
                    
                Next i
            
             End If
             
             '-----------------------------Missing Telegram---------------------------------------------------
                If TelegramMissing = True And nidPacketMissing = False Then
                     If Not MissingPacketDict.Exists("Balise " & bal_group_name & "(" & bal_group_id & ") " & "aspect " & ExtractAspectName(key) & " missing " & TelegramArray(j).Packet) Then
                        MissingPacketDict.Add "Balise " & bal_group_name & "(" & bal_group_id & ") " & "aspect " & ExtractAspectName(key) & " missing " & TelegramArray(j).Packet, key
                        rs.Cells(print_row, 1).value = ExtractAspectName(key) & "(" & ExtractTextBetweenParentheses(key) & ")"
                        rs.Cells(print_row, 2).value = bal_group_name & " " & bal_group_id
                        rs.Cells(print_row, 4).value = infill
                        rs.Cells(print_row, 5).value = TelegramArray(j).Packet '"BEGIN_TELEGRAM(" & AspectValue & ")" & " " & TelegramArray(j).Packet
                        rs.Cells(print_row, 8).value = "FAIL"
                        rs.Cells(print_row, 9).value = "Balise " & bal_group_name & " " & bal_group_id & " " & "aspect " & ExtractAspectName(key) & " missing " & TelegramArray(j).Packet
                        rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
                        print_row = print_row + 1
                        
                    End If
                End If
                

            TelegramMissing = False
            nidPacketMissing = False
            AspectValue = ""
            
            
            Next key
            
'            nsideTelegram = False
'            nidPacketFound = False
            TempAspectName = ""
            nidPacketVariable = False
            
            
            End If

        Next j
            


    End If
    
    
         'Clear the dictionary
        AspectDict.RemoveAll
        MissingPacketDict.RemoveAll
         'Clear the ids
        bal_group_name = ""
        bal_group_id = ""

        
Skip_file:
        
    Next file
 

    
    ' Loop through each subfolder in the current folder
    Dim subFolder As Object
    For Each subFolder In folder.Subfolders
        ProcessFolder subFolder, ws, row ' Recursive call for subfolders
    Next subFolder
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.CutCopyMode = False


End Sub


