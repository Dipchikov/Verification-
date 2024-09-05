Attribute VB_Name = "fuctions"
Option Explicit

Function ExtractValue(inputLine As String) As String
    Dim startPos As Long
    Dim endPos As Long
    On Error Resume Next
    ' Find the position of "=" and "("
    startPos = InStr(inputLine, "=")
    endPos = InStr(startPos, inputLine, "(")
    
    ' Check if both characters are found
    If startPos > 0 And endPos > startPos Then
        ' Extract the value between "=" and "("
        ExtractValue = Trim(Mid(inputLine, startPos + 1, endPos - startPos - 1))
    Else
        ExtractValue = "" ' Return an empty string if not found
    End If
End Function
Function ExtractVariable(inputLine As String) As String
    Dim startPos As Long
    On Error Resume Next
    ' Find the position of "=" and "("
    startPos = InStr(inputLine, "=")

    ' Check if both characters are found
    If startPos > 0 Then
        ' Extract the value between "=" and "("
        ExtractVariable = Left(inputLine, startPos - 1)
    Else
        ExtractVariable = "" ' Return an empty string if not found
    End If
End Function


Function ExtractTextBetweenParentheses(ByVal inputText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedText As String
    
    ' Find the position of the first "("
    startPos = InStr(1, inputText, "(")
    
    ' Find the position of the first ")" after the "("
    If startPos > 0 Then
        endPos = InStr(startPos, inputText, ")")
        
        ' Check if both "(" and ")" are found
        If startPos > 0 And endPos > startPos Then
            ' Extract the text between "(" and ")"
            extractedText = Mid(inputText, startPos + 1, endPos - startPos - 1)
        Else
            ' If no parentheses found, return empty string
            extractedText = ""
        End If
    End If
    ' Return the extracted text
    ExtractTextBetweenParentheses = extractedText
End Function


Function PrintResultsInSheet(val1 As Variant, val2 As Variant, val3 As Variant, val4 As Variant, val5 As Variant)
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Set the worksheet to "Results"
    Set ws = ThisWorkbook.Sheets("Results")

    ' Find the last used row in the sheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1

    ' Fill the values in columns A, B, C, and D in the next available row
    ws.Cells(lastRow, 1).value = val1
    ws.Cells(lastRow, 2).value = val2
    ws.Cells(lastRow, 3).value = val3
    ws.Cells(lastRow, 4).value = val4
    ws.Cells(lastRow, 4).value = val5
End Function



Function ExtractAspectName(ByVal inputText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedText As String
    
    ' Find the position of the first "("
    startPos = InStr(1, inputText, "=")
    
    ' Find the position of the first ")" after the "("
    endPos = Len(inputText)
    
    ' Check if both "(" and ")" are found
    If startPos > 0 And endPos > startPos Then
        ' Extract the text between "(" and ")"
        extractedText = Mid(inputText, startPos + 1, endPos - startPos)
    Else
        ' If no parentheses found, return empty string
        extractedText = ""
    End If
    
    ' Return the extracted text
    ExtractAspectName = extractedText
End Function



Function TelegramDirection(ByVal inputText As String) As String
                    Dim Q_DIR As String
                         Select Case inputText
                            Case "0"
                                ' Code for when myVar is 1
                                Q_DIR = "Reverse"
                            Case "1"
                                ' Code for when myVar is 2
                                Q_DIR = "Nominal"
                            Case "2"
                                ' Code for when myVar is 2
                                Q_DIR = "Both"
                        End Select
    ' Return the extracted text
    TelegramDirection = Q_DIR
End Function


Function RemoveSpaces(ByVal inputString As String) As String
    ' Replace spaces with an empty string
    RemoveSpaces = Replace(inputString, " ", "")
End Function


Function BaliseType(ByVal inputText As String) As String
                    Dim BAL_TYPE As String
                         Select Case inputText
                            Case "0"
                                ' Code for when myVar is 1
                                BAL_TYPE = "Fixed balise"
                            Case "1"
                                ' Code for when myVar is 2
                                BAL_TYPE = "Balise default[128]"
                            Case ""
                                ' Code for when myVar is 2
                                BAL_TYPE = "Switched balise"
                        End Select
    ' Return the extracted text
    BaliseType = BAL_TYPE
End Function


Function FindThirdBackslashFromRight(ByVal inputString As String) As Long
    Dim i As Long
    Dim backslashCount As Long
    
    ' Loop from the end of the string to the beginning
    For i = Len(inputString) To 1 Step -1
        ' Check if the current character is a backslash
        If Mid(inputString, i, 1) = "\" Then
            backslashCount = backslashCount + 1
        End If
        
        ' If the third backslash is found, return the position
        If backslashCount = 3 Then
            FindThirdBackslashFromRight = Len(inputString) - i
            Exit Function
        End If
    Next i
    
    ' If less than 3 backslashes were found, return 0 (not found)
    FindThirdBackslashFromRight = 0
End Function





Function GetValueFromDictionary(inputKey As String) As String
    ' Declare and create a new dictionary
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' Add key-value pairs to the dictionary
    dict.Add "M_VERSION", "M_VERSION="
    dict.Add "Q_LINK", "Q_LINK="

    ' Check if the input key exists in the dictionary
    If dict.Exists(inputKey) Then
        ' If key exists, return the corresponding value
        GetValueFromDictionary = dict(inputKey)
    Else
        ' If key does not exist, return a message
        GetValueFromDictionary = "Key not found"
    End If

    ' Clean up
    Set dict = Nothing
End Function


Function ExtractAspectKeyName(ByVal inputText As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedText As String
    
    ' Find the position of the first "("
    startPos = InStr(1, inputText, "=")
    
    ' Find the position of the first ")" after the "("
    endPos = InStr(startPos, inputText, "[")
    
    ' Check if both "(" and ")" are found
    If startPos > 0 And endPos > startPos Then
        ' Extract the text between "(" and ")"
        extractedText = Mid(inputText, startPos + 1, endPos - startPos - 1)
    Else
        ' If no parentheses found, return empty string
        extractedText = ""
    End If
    
    ' Return the extracted text
    ExtractAspectKeyName = extractedText
End Function
