Attribute VB_Name = "CheckEqualTelegrams"
Option Explicit
Type TelegramData
    AspectName As String
    Aspect As String
    AspectName_2 As String
    Packet As String
    Chennel As String
    Chennel_2 As String
End Type

Sub CheckEqualTelegrams()
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


    'ThisWorkbook.Sheets("Results").ShowAllData
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
    Dim vLoaValue As String
    Dim rs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Dim y As Long
    Dim z As Integer
    Dim AspectDict As Object
    Dim MissingPacketDict As Object
    Dim AspectValue As String
    Dim AspectName As String
    Dim key As Variant
    Dim bal_group_name As String
    Dim bal_group_id As String
    Dim print_row As Long
    Dim AspectName_Chennel  As String
    Dim TelegramContentFisrtAspect() As String
    Dim TelegramContentSecondAspect() As String
    Dim newWorkbook As Workbook

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
    lastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).row
    print_row = rs.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    ' Declare the array with 1-based indexing to match the row numbers
    'Dim TelegramArray(1 To 4) As TelegramData
    ReDim TelegramArray(2 To lastRow) As TelegramData
    
     For j = 2 To lastRow
            TelegramArray(j).AspectName = ws.Cells(j, 17).value  '
            TelegramArray(j).AspectName_2 = ws.Cells(j, 18).value
            TelegramArray(j).Packet = "NID_PACKET=" & ws.Cells(j, 19).value
            TelegramArray(j).Chennel = "[" & UCase(ws.Cells(j, 20).value) & "]"
            TelegramArray(j).Chennel_2 = "[" & UCase(ws.Cells(j, 21).value) & "]"
    Next j
    
    
    ' Loop through each file in the folder
    For Each file In folder.Files
        If Right(file.Name, 4) = ".sdi" Then ' Adjust file type if necessary
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


            
            ' Loop through each line in the file
            
            
        For j = 2 To lastRow
            For i = LBound(lines) To UBound(lines)
                        ' Loop through each line in array
            
                 If InStr(lines(i), "BAL_GROUP_NAME") > 0 Then
                    bal_group_name = ExtractAspectName(lines(i))
                 End If
                 
               ' If InStr(lines(i), TelegramArray(j).AspectName) > 0 Then
                AspectName = ExtractAspectName(lines(i))

                
                    If AspectName = TelegramArray(j).AspectName Then
                    If TelegramArray(j).Chennel = "[" & ExtractAspectName(lines(i - 2)) & "]" Then
                            If Not AspectDict.Exists(lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]") Then
                                AspectDict.Add lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]", bal_group_name & "(" & ExtractAspectName(lines(i - 3)) & ") [" & ExtractAspectName(lines(i - 2)) & "]"
                        End If
                    End If
                    End If
                    
                    If AspectName = TelegramArray(j).AspectName_2 Then
                      If TelegramArray(j).Chennel_2 = "[" & ExtractAspectName(lines(i - 2)) & "]" Then
                                If Not AspectDict.Exists(lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]") Then
                                    AspectDict.Add lines(i) & "[" & ExtractAspectName(lines(i - 1)) & "]", bal_group_name & "(" & ExtractAspectName(lines(i - 3)) & ") [" & ExtractAspectName(lines(i - 2)) & "]"
                                End If
                        End If
                     End If
               Next i
            Next j
            
            
            If AspectDict.Count = 0 Then
                GoTo Skip  ' Jump to the label named Skip
            End If


            
            
            z = 1
          '-------------------------------loop under aspects------------------------------------------------------

                         
For j = 2 To lastRow
            
            ReDim TelegramContentFisrtAspect(0 To 100)
            ReDim TelegramContentSecondAspect(0 To 100)
            
                 For Each key In AspectDict.keys
                            AspectValue = ExtractTextBetweenParentheses(key)
                            x = 1
                            y = 1
                             bal_group_id = AspectDict(key)
            
             
                For i = LBound(lines) To UBound(lines)
                
                
          ' Loop through each line in array
          

                ' Check for BEGIN_TELEGRAM(1)
                If InStr(lines(i), "BEGIN_TELEGRAM(" & AspectValue & ")") > 0 Then
                    insideTelegram = True
                End If
                
                ' Check for NID_PACKET=12 and if we are inside BEGIN_TELEGRAM(1)
                If insideTelegram And InStr(lines(i), TelegramArray(j).Packet) > 0 Then
                    nidPacketFound = True
                End If
                
                
                ' Check for V_LOA=0 if NID_PACKET=12 was found
                If nidPacketFound = True Then
                
                          If InStr(key, TelegramArray(j).AspectName) > 0 And InStr(AspectDict(key), TelegramArray(j).Chennel) > 0 Then
                             TelegramContentFisrtAspect(x) = lines(i)
                             'TelegramContentFisrtAspect(0) = key
                             x = x + 1
                         End If
                         
                        If InStr(key, TelegramArray(j).AspectName_2) > 0 And InStr(AspectDict(key), TelegramArray(j).Chennel_2) > 0 Then
                             TelegramContentSecondAspect(y) = lines(i)
                             'TelegramContentSecondAspect(0) = key
                             y = y + 1
                         End If
                 End If
                
                ' Reset flags when encountering a new BEGIN_TELEGRAM
                If InStr(lines(i), "BEGIN_TELEGRAM") > 0 And InStr(lines(i), "BEGIN_TELEGRAM(" & AspectValue & ")") = 0 Then 'Or InStr(lines(i), "BEGIN_PACKET") > 0
                    insideTelegram = False
                    nidPacketFound = False
                End If
                
                  If InStr(lines(i), "NID_PACKET") > 0 And InStr(lines(i), TelegramArray(j).Packet) = 0 Or InStr(lines(i), "BEGIN_PACKET") > 0 Then
                    nidPacketFound = False
                End If
                
                Next i
            
    Next key
    
    
    
          ' ------------------------------------------Loop under arrays--------------------------------------------
      Dim PacketNameFirstAspect As String
      Dim PacketNameSecondAspect As String
Dim keys As Variant
  ' Extract all keys as an array
    keys = AspectDict.keys

    If AspectDict.Count = 2 Then
      If Join(TelegramContentFisrtAspect, ",") = Join(TelegramContentSecondAspect, ",") Then
         rs.Cells(print_row, 1).value = ExtractAspectName(keys(0)) & "(" & ExtractTextBetweenParentheses(keys(0)) & ")" & "-" & ExtractAspectName(keys(1)) & "(" & ExtractTextBetweenParentheses(keys(1)) & ")"
         rs.Cells(print_row, 2).value = AspectDict.Items()(0)
         rs.Cells(print_row, 5).value = TelegramArray(j).Packet
        rs.Cells(print_row, 8).value = "SUCCESS"
        rs.Cells(print_row, 9).value = AspectDict.Items()(0) & " " & TelegramContentFisrtAspect(0) & " Packets are equal "
        rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
        print_row = print_row + 1
      Else
            rs.Cells(print_row, 1).value = ExtractAspectName(keys(0)) & "(" & ExtractTextBetweenParentheses(keys(0)) & ")" & "-" & ExtractAspectName(keys(1)) & "(" & ExtractTextBetweenParentheses(keys(1)) & ")"
            rs.Cells(print_row, 2).value = AspectDict.Items()(0)
             rs.Cells(print_row, 5).value = TelegramArray(j).Packet
            rs.Cells(print_row, 8).value = "FAIL"
            rs.Cells(print_row, 9).value = ExtractAspectName(keys(0)) & "-" & AspectDict.Items()(0) & " and " & ExtractAspectName(keys(1)) & "-" & AspectDict.Items()(1) & " " & TelegramContentFisrtAspect(0) & " Packets are not equal "
            rs.Cells(print_row, 10).value = Right(file.Path, FindThirdBackslashFromRight(file.Path))
            print_row = print_row + 1

      End If
   End If
     ' Clear the array
    Erase TelegramContentFisrtAspect
    Erase TelegramContentSecondAspect

Next j
            

 'Clear the dictionary
AspectDict.RemoveAll
MissingPacketDict.RemoveAll
         'Clear the ids
bal_group_name = ""
bal_group_id = ""
        
End If
    
Skip:  ' Label indicating where execution should jump
If Not IsEmpty(keys) Then
Erase keys
End If



 Next file


    
    ' Loop through each subfolder in the current folder
    Dim subFolder As Object
    For Each subFolder In folder.Subfolders
        ProcessFolder subFolder, ws, row ' Recursive call for subfolders
    Next subFolder
    
    
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic


End Sub



