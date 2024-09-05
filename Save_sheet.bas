Attribute VB_Name = "Save_sheet"
Sub save_sheet()
    
    Dim lr          As Long
    Dim LastColumn  As Long
    Dim developer As String
    Dim InitialFoldr$
    ActiveWorkbook.save
    
    Dim sheetName   As String
    Set wb = ThisWorkbook
   developer = Project_Main_Form.Author.Caption
    On Error Resume Next
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    'ActiveWorkbook.ActiveSheet.Select
    sheetName = wb.ActiveSheet.name
    
    Set wb = Workbooks.Add
    
    Application.CopyObjectsWithCells = False
    ThisWorkbook.Sheets(sheetName).Copy before:=wb.Sheets(1)
    'ActiveSheet.name = sheet_name
    
    '--- remove data ranges-------
    Dim nm          As name
    For Each nm In ActiveWorkbook.Names
        nm.Delete
        'If nm.RefersToRange.Parent.name = “Sheet1” Then nm.Delete
    Next nm
    
    Application.CopyObjectsWithCells = True
    
    '---------Èçòðèâàíå íà Sheet1------------------
    Application.DisplayAlerts = False
    Sheets("Sheet1").Delete
    Application.DisplayAlerts = True
    
    '-------------add user in Footer ---------------
    With ActiveSheet.PageSetup
        .LeftFooter = "&D" & Chr(13) & "&9" & Application.UserName
        .RightFooter = "Page " & "&P" & Chr(13) & "&9" & developer
    End With
    
    Application.CutCopyMode = False        'esp
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Dim sFileSaveName As Variant
    Dim sPath       As String
    sPath = sheetName & "_" & Format(Date, "ddmmyyyy")        'Format(Now, "mm_dd_yy")
    InitialFoldr$ = ""
    sFileSaveName = Application.GetSaveAsFilename(InitialFileName:=sPath, FileFilter:="Excel Files (*.xlsx), *.xlsm")
    If sFileSaveName <> False Then
        Application.DisplayAlerts = False
        ActiveWorkbook.SaveAs sFileSaveName
        
        ActiveWorkbook.Close
        Application.DisplayAlerts = True
    End If
    
End Sub
