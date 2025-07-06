Attribute VB_Name = "ModIndesign"
Dim sRep As String
Public answer As Variant
Sub SwapYellow(ByVal lP1 As Long, ByVal lP2 As Long)
    If IsYellow(lP1) = True And IsYellow(lP2) = True Then
    ElseIf IsYellow(lP1) = True Then
        SetYellow lP2
        ResetYellow lP1
    ElseIf IsYellow(lP2) = True Then
        SetYellow lP1
        ResetYellow lP2
    End If
End Sub
Sub SwapPagesModule(ByVal lP1 As Long, ByVal lP2 As Long)
    Dim r1 As Range, r2 As Range, r3 As Range
'    Set r1 = ThisWorkbook.Names("Page_" & lP1).RefersToRange    ' prva strana
'    Set r2 = ThisWorkbook.Names("Page_" & lP2).RefersToRange    ' druga strana
    
    Set r1 = GetPageRange(lP1)
    Set r2 = GetPageRange(lP2)
    
    Set r3 = ThisWorkbook.Names("swapplace").RefersToRange    ' temp mesto
    r3.ClearContents
    r3.ClearComments
    r3.UnMerge
    r2.Copy
    r3.Cells(1, 1).PasteSpecial
    r2.ClearContents
    r2.ClearComments
    r2.UnMerge
    r1.Copy
    r2.Cells(1, 1).PasteSpecial
    r1.ClearContents
    r1.ClearComments
    r1.UnMerge
    r3.Copy
    r1.Cells(1, 1).PasteSpecial
    r3.ClearContents
    r3.ClearComments
    r3.UnMerge
    Application.CutCopyMode = False
    'bleed, ako ga ima
    If IsBleed(lP1) = True Or IsBleed(lP2) = True Then
        'zameni bleedove
        SwapBleedsModule lP1, lP2
    End If
    'strane markirane ko kolor
    If IsYellow(lP1) = True Or IsYellow(lP2) = True Then
        'zameni bleedove
        SwapYellow lP1, lP2
    End If
End Sub
Sub SwapBleedsModule(ByVal lP1 As Long, ByVal lP2 As Long)

    If IsBleed(lP1) = True And IsBleed(lP2) = True Then
    ElseIf IsBleed(lP1) = True Then
        SetBleed lP2
        ResetBleed lP1
    ElseIf IsBleed(lP2) = True Then
        SetBleed lP1
        ResetBleed lP2
    End If

End Sub
Function IsEmptyPage2(pn As Long) As Boolean
'gleda i format
    Dim r1 As Range, rc As Range
    Dim b As Boolean
    b = True
    'Set r1 = ThisWorkbook.Names("Page_" & pn).RefersToRange
    Set r1 = GetPageRange(pn)
    For Each rc In r1.Cells
        If Len(rc.value) > 0 Then
            b = False
            Exit For
        End If
        If rc.MergeCells = True Then
            b = False
            Exit For
        End If
        If rc.Interior.Pattern <> xlNone Then
            b = False
            Exit For
        End If
        If rc.Interior.TintAndShade <> 0 Then
            b = False
            Exit For
        End If
    Next rc
    IsEmptyPage2 = b
End Function
Function IsEmptyPage(pn As Long) As Boolean
    Dim r1 As Range, rc As Range
    Dim b As Boolean
    b = True
    'Set r1 = ThisWorkbook.Names("Page_" & pn).RefersToRange
    Set r1 = GetPageRange(pn)
    For Each rc In r1.Cells
        If Len(rc.value) > 0 Then
            b = False
            Exit For
        End If
    Next rc
    IsEmptyPage = b
End Function
Sub MyDeletePages()
'brise samo do max pageta
    If IDP Then Exit Sub
'    If InName = False Then
'        MsgBox "Select page where you want to start delete"
'        Exit Sub
'    End If

    Dim x As Long
    Dim lFilledPage As Long
    Dim lPagesToDelete As Long
    Dim lActivePage As Long
    Dim lPageReserve As Long
    Dim lMaxPage As Long
    Dim lMinPage As Long
    Dim bOdZadnjeNapunjeneStrane As Boolean
    Dim ws As Worksheet
'    Dim keysWs As Worksheet
'    Dim matchWs As Worksheet

'On Error Resume Next
'bookPageCount = UBound(PagesArr)
'If Err.Number <> 0 Then
'On Error GoTo 0
'    PopulatePageArray
'End If
'On Error GoTo 0

    If Not MyRefreshLayout Then ' refresh and draw NewLayout before swap
        Exit Sub
    End If

    Set ws1 = ThisWorkbook.Sheets("NewLayout")
    If Application.activeSheet Is ws1 Then
        lActivePage = GetActivePage
    Else
        lActivePage = 0
    End If
    
    NoOfPages = CInt(ThisWorkbook.Sheets("NewLayout").Range("U1"))
    If NoOfPages = 0 Then Exit Sub
    
    frmDelPages.Show (vbModal)
    If frmDelPages.GetRetVal <> 1 Then
        Call Unload(frmDelPages)
        Exit Sub
    End If
    vPageFrom = frmDelPages.GetPageFrom
    vPageTo = frmDelPages.GetPageTo
    vPageNo = frmDelPages.GetPageNo
    Call Unload(frmDelPages)
'    If lActivePage < 2 Or lActivePage > NoOfPages Then
'        lActivePage = Int(val(InputBox("Enter Page No to start deleting from:", , 2)))
'        If lActivePage = 0 Then Exit Sub
'
'    End If
    If vPageNo = 0 Then
        lActivePage = vPageFrom
    Else
        lActivePage = vPageNo
    End If
    If lActivePage < 2 Or lActivePage > NoOfPages Then
        MsgBox "You can not delete pages outside of book" & vbCrLf & "(Page: " & 2 & " to " & NoOfPages & ")"
        Exit Sub
    End If
    'lPagesToDelete = Int(val(InputBox("Enter number of pages to delete:", , 1)))
    If vPageNo = 0 Then
        lPagesToDelete = vPageTo - vPageFrom + 1
    Else
        lPagesToDelete = 1
    End If
    If lPagesToDelete = 0 Then Exit Sub
    If (lActivePage + lPagesToDelete - 1) > NoOfPages Then
        MsgBox "You can not delete pages outside of book" & vbCrLf & "(You want to delete up to " & (lActivePage + lPagesToDelete - 1) & " page in book of " & NoOfPages & " pages)"
        Exit Sub
    End If

''Application.EnableEvents = False
'    enable (False)
'    StoreActiveSheets
'Application.ScreenUpdating = False
'commandOpRunning = True
Set ws = ThisWorkbook.Sheets("Main List")
'Set keysWs = ThisWorkbook.Sheets("KeysSheet")
'Set matchWs = ThisWorkbook.Sheets("match")
Application.ScreenUpdating = False
Application.EnableEvents = False
Dim sIssue As String
'sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
sIssue = GetCurrentIssue
result = MyCheckLockAndReconcile("You can not delete pages in this issue now", True, False)
If Not result Then
    'MyRestoreContext
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
End If
'MySaveContext
LockIssue sIssue


'lRow = lastNonBlankRow(ws, "A")
'lRow = GetNewRow - 1
lRow = GetLastRow(ws)
'Call DeletePagesFromPageArray(lPagesToDelete, lActivePage)
    For r = 2 To lRow
        If ws.Range("G" & r) <> "" Then
'            r1 = findRowInKeysSheet(ws.Range("A" & r) & ws.Range("B" & r) & ws.Range("C" & r))
'            r1 = findRowInKeysSheet(ws.Range("M" & r))
            'r1 = matchWs.Range("C" & r)
            'If (r1 <> 0) Then
                If CInt(ws.Range("G" & r)) >= lActivePage And CInt(ws.Range("G" & r)) < lActivePage + lPagesToDelete Then
                        ws.Range("G" & r & ":H" & r).ClearContents
'                        keysWs.Range("E" & r1 & ":N" & r1).ClearContents
                        ws.Cells(r, AdLastChangeCol + 1) = MakeID
                        gMainListArr(r - 1, 7) = ""
                        gMainListArr(r - 1, 8) = ""
                        gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
                End If
                If CInt(ws.Range("G" & r)) >= lActivePage + lPagesToDelete Then
                    ws.Range("G" & r) = CInt(ws.Range("G" & r)) - lPagesToDelete
                    ws.Cells(r, AdLastChangeCol + 1) = MakeID
                        gMainListArr(r - 1, 7) = ws.Range("G" & r)
                        gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
'                    keysWs.Range("E" & r1) = ws.Range("G" & r)
                End If
            'End If
        End If
    Next r
' ThisWorkbook.Sheets("Settings").Range("B11") = NoOfPages - lPagesToDelete
' Application.StatusBar = "Pages deleted - Refreshing pages now"
'    RestoreActiveSheets
'    DoRefreshLayout ' Added by Hussam
'    enable (True)
'If (False) Then
' If Application.ActiveSheet Is ws1 Then
'    Call DrawPages(lActivePage - 1)
' Else
'    gNeedRedraw = True
' End If
'End If
 maxpage = Application.WorksheetFunction.Max(ws.Range("G2:G" & lRow))
 If ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value <> maxpage Then
    ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = maxpage
 End If
 'Application.StatusBar = "Rechecking Pages ..."
'CheckPages
r = PagesStartRow + ((lActivePage - 2) \ 8) * RowsPerPage
c = startCol + ((lActivePage - 2) Mod 8) * columnsPerPage
'Application.EnableEvents = True
'enable (True)
'Application.ScreenUpdating = True

 'ThisWorkbook.Sheets("Layout").Activate

MySaveIssue (False)
UnlockIssue sIssue
 
 drawPages (2)
 'MyRestoreContext
 If Application.activeSheet Is ws1 Then
    On Error Resume Next
    ThisWorkbook.Sheets("NewLayout").Cells(r, c).Select
    On Error GoTo 0
  End If
 'Application.StatusBar = "Ready"
 
 Application.ScreenUpdating = True
Application.EnableEvents = True

'commandOpRunning = False
End Sub

Sub DeletePages2()
'brise samo do max pageta
    If IDP Then Exit Sub
    If InName = False Then
        MsgBox "Select page where you want to start delete"
        Exit Sub
    End If

    Dim x As Long
    Dim lFilledPage As Long
    Dim lPagesToDelete As Long
    Dim lActivePage As Long
    Dim lPageReserve As Long
    Dim lMaxPage As Long
    Dim lMinPage As Long
    Dim bOdZadnjeNapunjeneStrane As Boolean

    If ThisWorkbook.Worksheets("Settings").Range("FindLastFilledCell").value = "Yes" Then
        bOdZadnjeNapunjeneStrane = True
    End If

    lMaxPage = ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value
    lMinPage = ThisWorkbook.Worksheets("Settings").Range("MinPageNo").value

    lActivePage = AktivnaStrana
    If lActivePage < lMinPage Or lActivePage > lMaxPage Then
        MsgBox "You can not delete pages outside of book" & vbCrLf & "(Page: " & lMinPage & " to " & lMaxPage & ")"
        Exit Sub
    End If


    lPagesToDelete = Int(val(InputBox("Enter number of pages to delete:", , 1)))
    If lPagesToDelete = 0 Then Exit Sub

    If (lActivePage + lPagesToDelete - 1) > lMaxPage Then
        MsgBox "You can not delete pages outside of book" & vbCrLf & "(You want to delete up to " & (lActivePage + lPagesToDelete - 1) & " page in book of " & lMaxPage & " pages)"
        Exit Sub
    End If

    If bOdZadnjeNapunjeneStrane = True Then
        lFilledPage = lMinPage
        'pocni sa krajnjom stranom KNJIGE pa smanjuj
        For x = lMaxPage To lMinPage Step -1
            If IsEmptyPage2(x) = False Then
                lFilledPage = x
                Exit For
            End If
        Next x
    End If

    '    lPageReserve = 849 - lFilledPage
    '    If lBookSize < lPagesToDelete Then
    '        MsgBox "You can not delete " & lPagesToDelete & " in book of " & lBookSize & " pages!"
    '
    '        Exit Sub
    '    End If
    Dim r1 As Range, r2 As Range, r3 As Range
    'obrisi strane
    For x = lActivePage To lActivePage + lPagesToDelete - 1
        'Set r1 = ThisWorkbook.Names("Page_" & x).RefersToRange
        Set r1 = GetPageRange(x)
        r1.ClearContents
        r1.UnMerge
        With r1.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        If IsBleed(x) = True Then
            'obrisi bleedove
            ResetBleed x
        End If
        'strane markirane ko kolor
        If IsYellow(x) = True Then
            ResetYellow x
        End If
    Next x
    If bOdZadnjeNapunjeneStrane = False Then
        For x = (lActivePage + lPagesToDelete) To lMaxPage
            SwapPagesModule x, x - lPagesToDelete
        Next x
    Else
        For x = lActivePage To lFilledPage
            SwapPagesModule x, x + lPagesToDelete
        Next x
    End If
    'ThisWorkbook.Names("Page_" & lActivePage).RefersToRange.Select
    Set r1 = GetPageRange(lActivePage)
    r1.Select
End Sub

Sub DeletePages()
    If IDP Then Exit Sub
    If InName = False Then
        MsgBox "Select page where you want to start delete"
        Exit Sub
    End If

    Dim x As Long
    Dim lFilledPage As Long
    Dim lPagesToAdd As Long
    Dim lActivePage As Long
    Dim lPageReserve As Long
    Dim lBookSize As Long
    lBookSize = ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value
    lActivePage = AktivnaStrana
    If lActivePage < 2 Then Exit Sub
    lPagesToAdd = Int(val(InputBox("Enter number of pages to delete:", , 1)))
    If lPagesToAdd = 0 Then Exit Sub
    lFilledPage = 849    'pocni sa krajnjom stranom pa smanjuj
    For x = 849 To 2 Step -1
        If IsEmptyPage(x) = False Then
            lFilledPage = x
            Exit For
        End If
    Next x

    lPageReserve = 849 - lFilledPage
    If lBookSize < lPagesToAdd Then
        MsgBox "You can not delete " & lPagesToAdd & " in book of " & lBookSize & " pages!"

        Exit Sub
    End If
    Dim r1 As Range, r2 As Range, r3 As Range
    'obrisi strane
    For x = lActivePage To lActivePage + lPagesToAdd - 1
        'Set r1 = ThisWorkbook.Names("Page_" & x).RefersToRange
        Set r1 = GetPageRange(x)
        r1.ClearContents
        r1.UnMerge
    With r1.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        If IsBleed(x) = True Then
            'obrisi bleedove
            ResetBleed x
        End If
        'strane markirane ko kolor
        If IsYellow(x) = True Then
            ResetYellow x
        End If
    Next x

    For x = lActivePage To lFilledPage
        SwapPagesModule x, x + lPagesToAdd
    Next x
    'ThisWorkbook.Names("Page_" & lActivePage).RefersToRange.Select
    Set r1 = GetPageRange(lActivePage)
    r1.Select
End Sub

Sub MyAddPages()
If IDP Then Exit Sub
'    If InName = False Then     'comment by hosam
'        MsgBox "Select the page right before the new pages"
'        Exit Sub
'    End If
    Dim x As Long
    Dim ws As Worksheet
'    Dim keysWs As Worksheet
'    Dim matchWs As Worksheet
    Dim lFilledPage As Long
    Dim lPagesToAdd As Long
    Dim lActivePage As Long
    Dim lPageReserve As Long
    
    If Not MyRefreshLayout Then ' refresh and draw NewLayout before swap
        Exit Sub
    End If
    Set ws1 = ThisWorkbook.Sheets("NewLayout")
    If Application.activeSheet Is ws1 Then
        lActivePage = GetActivePage
    Else
        lActivePage = 0
    End If
    
    NoOfPages = CInt(ThisWorkbook.Sheets("NewLayout").Range("U1"))
    If NoOfPages = 0 Then Exit Sub
    
    load frmAddPages
    frmAddPages.SetPageAfter (IIf(lActivePage = 0, 1, lActivePage))
    frmAddPages.Show (vbModal)
    If frmAddPages.GetRetVal <> 1 Then
        Call Unload(frmAddPages)
        Exit Sub
    End If
    lActivePage = frmAddPages.GetPageAfter
    lPagesToAdd = frmAddPages.GetPageCount
    Call Unload(frmAddPages)
    
    

    'If lActivePage < 2 Or lActivePage > NoOfPages Then Exit Sub
'    If lActivePage < 2 Or lActivePage > NoOfPages Then
'        lActivePage = Int(val(InputBox("Enter Page No to add the pages after:", , 1)))
'        If lActivePage = 0 Then Exit Sub
'
'    End If
'    lPagesToAdd = Int(val(InputBox("Enter number of pages to add:", , 25)))
    If lPagesToAdd = 0 Then Exit Sub
    If lActivePage = 0 Then Exit Sub

'Application.EnableEvents = False
'Application.ScreenUpdating = False
'    enable (False)
'    StoreActiveSheets
'commandOpRunning = True
Set ws = ThisWorkbook.Sheets("Main List")
'Set keysWs = ThisWorkbook.Sheets("KeysSheet")
'Set matchWs = ThisWorkbook.Sheets("match")
'lRow = lastNonBlankRow(ws, "A")

Application.ScreenUpdating = False
Application.EnableEvents = False
Dim sIssue As String
'sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value
sIssue = GetCurrentIssue
result = MyCheckLockAndReconcile("You can not add pages in this issue now", True, False)
If Not result Then
    'MyRestoreContext
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
End If
'MySaveContext
LockIssue sIssue
'lRow = GetNewRow - 1
lRow = GetLastRow(ws)
    For r = 2 To lRow
        If ws.Range("G" & r) <> "" Then
            If CInt(ws.Range("G" & r)) > lActivePage Then
                ws.Range("G" & r) = CInt(ws.Range("G" & r)) + lPagesToAdd
                ws.Cells(r, AdLastChangeCol + 1) = MakeID
                gMainListArr(r - 1, 7) = ws.Range("G" & r)
                gMainListArr(r - 1, 14) = ws.Cells(r, AdLastChangeCol + 1)
'                r1 = findRowInKeysSheet(ws.Range("A" & r) & ws.Range("B" & r) & ws.Range("C" & r))
''                r1 = findRowInKeysSheet(ws.Range("M" & r))
''                'r1 = matchWs.Range("C" & r)
''                If (r1 <> 0) Then
''                    keysWs.Range("E" & r1) = ws.Range("G" & r)
''                End If
            End If
        End If
    Next r
 'ThisWorkbook.Sheets("Settings").Range("B11") = NoOfPages + lPagesToAdd
' Application.StatusBar = "Adding Pages... "
'' Call AddPagesToPageArray(lPagesToAdd, lActivePage)
 'PopulatePageArray
 'Application.StatusBar = "Pages added - Refreshing pages now"
 
    ' RestoreActiveSheets
''    DoRefreshLayout ' Added by Hussam
''    enable (True)

'' If (False) Then
''    If Application.ActiveSheet Is ws1 Then
''       Call DrawPages(lActivePage + 1)
''    Else
''       gNeedRedraw = True
''    End If
'' End If
 maxpage = Application.WorksheetFunction.Max(ws.Range("G2:G" & lRow))
 If ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value <> maxpage Then
    ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = maxpage
 End If
' Application.StatusBar = "Rechecking Pages ..."
'CheckPages
r = PagesStartRow + ((lActivePage - 2) \ 8) * RowsPerPage
c = startCol + ((lActivePage - 2) Mod 8) * columnsPerPage

'Application.EnableEvents = True
'Application.ScreenUpdating = True
'enable (True)
 'ThisWorkbook.Sheets("Layout").Activate

MySaveIssue (False)
UnlockIssue sIssue
 
 drawPages (2)
' MyRestoreContext
 If Application.activeSheet Is ws1 Then
   On Error Resume Next
   ThisWorkbook.Sheets("NewLayout").Cells(r, c).Select
   On Error GoTo 0
 End If
' Application.StatusBar = "Ready"
'commandOpRunning = False
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
Sub AddPages()
If IDP Then Exit Sub
    If InName = False Then
        MsgBox "Select the page right before the new pages"
        Exit Sub
    End If
    Dim x As Long
    Dim lFilledPage As Long
    Dim lPagesToAdd As Long
    Dim lActivePage As Long
    Dim lPageReserve As Long
    lActivePage = AktivnaStrana
    If lActivePage < 2 Then Exit Sub
    lPagesToAdd = Int(val(InputBox("Enter number of pages to add:", , 25)))
    If lPagesToAdd = 0 Then Exit Sub
    lFilledPage = 849    'pocni sa krajnjom stranom pa smanjuj
    For x = 849 To 2 Step -1
        If IsEmptyPage(x) = False Then
            lFilledPage = x
            Exit For
        End If
    Next x
    lPageReserve = 849 - lFilledPage
    If lPageReserve < lPagesToAdd Then
        MsgBox "No place to add " & lPagesToAdd & " pages!"
        Exit Sub
    End If

    Dim r1 As Range, r2 As Range, r3 As Range
    'prvo gurni sve strane od aktivne+1 to poslednje zauzete
    For x = lFilledPage To lActivePage + 1 Step -1
        SwapPagesModule x, x + lPagesToAdd
    Next x

 
    If lFilledPage + lPagesToAdd > ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value Then
        ThisWorkbook.Worksheets("Settings").Range("MaxNoPages").value = lFilledPage + lPagesToAdd
        CheckPages
    End If
End Sub
Sub ExportNumChart()
Const FName As String = "C:\Numbers.jpg"
Dim pic_rng As Range
Dim ShTemp As Worksheet
Dim ChTemp As Chart
Dim PicTemp As Picture
Application.ScreenUpdating = False
Set pic_rng = Worksheets("Layout").Range("Page_2")
Set ShTemp = Worksheets.Add
Charts.Add
ActiveChart.Location Where:=xlLocationAsObject, Name:=ShTemp.Name
Set ChTemp = ActiveChart
pic_rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
ChTemp.Paste
Set PicTemp = Selection
With ChTemp.Parent
.Width = PicTemp.Width + 8
.Height = PicTemp.Height + 8
End With
ChTemp.Export filename:="c:\Users\shonius\Documents\Rad\Elance\2013-06-25  Excel to Indesign4\Numbers.jpg", FilterName:="jpg"
'UserForm1.Image1.Picture = LoadPicture(FName)
'Kill FName
Application.DisplayAlerts = False
ShTemp.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub ListAdsFromLayoutSheet()
If IDP Then Exit Sub
ListAds
 ThisWorkbook.Sheets("Layout").Activate
End Sub
Sub FilterAds()
'remove + sign and text after it from ad list and move results in new Excel file
    Dim r As Range, rResiz As Range, r2 As Range
    Dim x As Long, Y As Long
    Dim SourceAr()
    Dim wbNew As Workbook

    ThisWorkbook.Sheets("Report").Activate
    Set r = Range("g9", Range("h10000").End(xlUp))
    r.Select
    Set rResiz = r.Resize(columnsize:=3)
    If r.Cells.count < 1 Then
        Exit Sub
    End If
    rResiz.Select
    SourceAr = rResiz
    For x = LBound(SourceAr) To UBound(SourceAr)
        Y = InStr(1, SourceAr(x, 1), "+")
        If Y > 0 Then
            SourceAr(x, 1) = left(SourceAr(x, 1), Y - 1)
        End If
    Next x

    Set wbNew = Workbooks.Add
    wbNew.Worksheets(1).Activate
    wbNew.Worksheets(1).Range("a1").value = "Ad:"
    wbNew.Worksheets(1).Range("b1").value = "Page:"
    wbNew.Worksheets(1).Range("c1").value = "Ad file:"
    wbNew.Worksheets(1).Range("a1:c1").Font.Bold = True
    Set r2 = wbNew.Worksheets(1).Range("a2")
    r2.Resize(UBound(SourceAr), 3).value = SourceAr
    wbNew.Worksheets(1).columns("A:C").EntireColumn.AutoFit
End Sub
Sub ListAds(Optional bAllPages As Boolean = False)    'kad se uvoze rezervacije potrebne su sve strane
    If IDP Then Exit Sub
    Dim nName As Name
    'Dim cPages As New Collection
    Dim cPUnits As New Collection
    Dim cPUnitsSizes As New Collection
    Dim cpErrors As New Collection
    Dim cErrors As Collection
    Dim cUnits As Collection
    Dim cUnitsSizes As Collection
    Dim cUnitsColumns As Collection
    Dim cUnitsPositions As Collection

    Dim cAds As New Collection
    Dim cFiles As New Collection
    Dim cOnPage As New Collection

    Dim x As Long, Y As Long, z As Long
    Dim x2 As Long    'brojac za imena
    Dim r As Range, rc As Range
    Dim r2 As Range

    Dim rCaption As Range, rTemp As Range

    Dim bAtLeastOneUnitOnPage As Boolean
    Dim lUnitsTotalPerPage As Long
    Dim lUnitsFilledPerPage As Long
    Dim lFilesPerPage As Long
    Dim lPercent As Long


    Dim aSize(1 To 6)    'row,col,Indesign width,height,filename,adname
    Dim aPos(1 To 2)    'top ,left
    Dim vc As Variant
    Dim lMinPage As Long
    Dim lMaxPages As Long
    Dim sFile As String

    '****vars for report
    Dim lEmptyPages As Long
    Dim lCompletedPages As Long
    Dim lPagesInProgress As Long
    Dim lErrorPages As Long
    Dim lErrorsPerPage As Long
    Dim rRep As Range
    Dim rRep2 As Range    'abecedno sortiranje
    Dim x3 As Long    'counter for report
    '****
    ThisWorkbook.Sheets("Layout").Activate
    If bAllPages = False Then
        lMinPage = Range("MinPageNo").value
        lMaxPages = Range("MaxNoPages").value
    Else
        lMinPage = 2
        lMaxPages = 401
    End If

    For x2 = lMinPage To lMaxPages
'        SetPageRangeAddress (x2)
'        Set NName = ThisWorkbook.Names("Page_" & x2)
        
        Set cUnits = New Collection
        Set cUnitsSizes = New Collection

'            tmpStr = NName.RefersTo
'            vArr = Split(tmpStr, "!")
'            Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))
       ' Set r = NName.RefersToRange
        Set r = GetPageRange(x2)
        For Each rc In r.Cells
            If rc.MergeCells = True Then
                On Error Resume Next
                cUnits.Add rc.MergeArea, rc.MergeArea.Address
                On Error GoTo 0
            Else
                cUnits.Add rc
            End If
        Next rc

        cPUnits.Add cUnits

        For Each vc In cUnits
            Set r2 = vc
            'ad name
            aSize(6) = CStr(r2.Cells(1, 1).value)
            'file name
            aSize(5) = GetPathFromComment(r2.Cells(1, 1))

            If aSize(6) <> "" Then
                cAds.Add aSize(6)
                cFiles.Add aSize(5)
                cOnPage.Add x2
            End If
        Next vc


    Next x2


    With ThisWorkbook.Sheets("Report")
        With .Range("d8:f3200")
            .ClearContents
            .Font.ColorIndex = xlAutomatic
        End With

        With .Range("g8:i3200")
            .ClearContents
            .Font.ColorIndex = xlAutomatic
        End With

        .Range("d7").value = "Ad:"
        .Range("e7").value = "Page:"
        .Range("f7").value = "Ad file:"

        .Range("g7").value = "Ad:"
        .Range("h7").value = "Page:"
        .Range("i7").value = "Ad file:"
    End With

    Set rRep = ThisWorkbook.Sheets("Report").Range("D8")

    For x = 1 To cAds.count
        rRep.Offset(x, 0).value = cAds(x)
        rRep.Offset(x, 1).value = cOnPage(x)
        rRep.Offset(x, 2).value = cFiles(x)
        If cFiles(x) <> "" Then
            If DirU(cFiles(x)) = "" Then
                With rRep.Offset(x, 2).Font
                    .Color = -16776961
                    .TintAndShade = 0
                End With
            End If

        End If
    Next x
    ThisWorkbook.Sheets("Report").Activate
    Range("d7").Select
    If cAds.count < 1 Then Exit Sub
    Range("d9", rRep.Offset(x - 1, 2)).Select

    Range("d9:d3200").NumberFormat = "@"
    Range("d9", rRep.Offset(x - 1, 2)).Copy
    Range("g9").PasteSpecial
    ThisWorkbook.Worksheets("Report").Sort.SortFields.Clear
    ThisWorkbook.Worksheets("Report").Sort.SortFields.Add key:=Range("G9"), _
                                                          SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ThisWorkbook.Worksheets("Report").Sort
        .SetRange Range("G9", rRep.Offset(x - 1, 5))
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    With Range("G9", rRep.Offset(x - 1, 3)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ThemeColor = 6
        '        .TintAndShade = 0.399945066682943
        .Weight = xlThin
    End With
    Range("d7").Select
    Application.CutCopyMode = False
End Sub

Sub MarkYellow()
    If InName Then
        If ActiveCell.Interior.Color = 65535 Then
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Else
            With ActiveCell.Interior
                Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Else
        MsgBox "Select cell in page"
    End If

End Sub
Sub CheckPagesfromReportSheet()
If IDP Then Exit Sub
CheckPages
ThisWorkbook.Sheets("Report").Activate
End Sub
Sub CheckPages()
Attribute CheckPages.VB_ProcData.VB_Invoke_Func = "c\n14"
    If IDP Then Exit Sub
    Dim nName As Name
    'Dim cPages As New Collection
    Dim cPUnits As New Collection
    Dim cPUnitsSizes As New Collection
    Dim cpErrors As New Collection
    Dim cErrors As Collection
    Dim cUnits As Collection
    Dim cUnitsSizes As Collection
    Dim cUnitsColumns As Collection
    Dim cUnitsPositions As Collection
    Dim x As Long, Y As Long, z As Long
    Dim x2 As Long    'brojac za imena
    Dim r As Range, rc As Range
    Dim r2 As Range

    Dim rCaption As Range, rTemp As Range

    Dim bAtLeastOneUnitOnPage As Boolean
    Dim lUnitsTotalPerPage As Long
    Dim lUnitsFilledPerPage As Long
    Dim lFilesPerPage As Long
    Dim lPercent As Long


    Dim aSize(1 To 6)    'row,col,Indesign width,height,filename, ad name
    Dim aPos(1 To 2)    'top ,left
    Dim vc As Variant
    Dim lMaxPages As Long
    Dim lMinPages As Long
    Dim sFile As String

    '****vars for report
    Dim lEmptyPages As Long
    Dim lCompletedPages As Long
    Dim lPagesInProgress As Long
    Dim lErrorPages As Long
    Dim lErrorsPerPage As Long
    Dim rRep As Range
    Dim x3 As Long    'counter for report
    Dim ws As Worksheet
        
'On Error Resume Next
'bookPageCount = UBound(PagesArr)
'If Err.Number <> 0 Then
'On Error GoTo 0
'    PopulatePageArray
'End If
'On Error GoTo 0
'
    '****
    'Set ws = ThisWorkbook.Sheets("NewLayout")
    If Application.ActiveWorkbook.Name <> ThisWorkbook.Name Then
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ThisWorkbook.Sheets("NewLayout").Activate
    lMaxPages = ThisWorkbook.Sheets("Settings").Range("MaxNoPages").value
    lMinPages = ThisWorkbook.Sheets("Settings").Range("MinPageNo").value
    
    fDrawPageOne = LCase(ThisWorkbook.Sheets("additional_Settings").Range("Assign_Ads_in_Page_1")) = "yes"
    lMinPages = IIf(fDrawPageOne, 1, 2)
    lMaxPages = ThisWorkbook.Sheets("NewLayout").Range("U1").value
'    lMinPages = 2
'    lMaxPages = UBound(PagesArr)
'    If lMinPages < 2 Then lMinPages = 2
'    If lMaxPages < 2 Or lMaxPages > UBound(PagesArr) Then lMaxPages = UBound(PagesArr)
    For x2 = lMinPages To lMaxPages

'        Set NName = ThisWorkbook.Names("Page_" & x2)
'        bAtLeastOneUnitOnPage = False
'        For J = 1 To 8
'            If PagesArr(x2).AssignedAdRow(J) <> 0 Then
'                bAtLeastOneUnitOnPage = True
'                Exit For
'            End If
'        Next J
'''        bAtLeastOneUnitOnPage = PageHasAllocations(x2)
'        SetPageRangeAddress (x2)
'        Set NName = ThisWorkbook.Names("Page_" & x2)
'''        If bAtLeastOneUnitOnPage Then
            Set cUnits = New Collection
            Set cUnitsSizes = New Collection
'''            tmpStr = NName.RefersTo
'''            vArr = Split(tmpStr, "!")
'''            Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))
    

'            Set r = NName.RefersToRange
            Set r = GetPageRange(x2)
            For Each rc In r.Cells
                If rc.MergeCells = True Then
                    On Error Resume Next
                    cUnits.Add rc.MergeArea, rc.MergeArea.Address
                    On Error GoTo 0
                Else
                    cUnits.Add rc
                End If
            Next rc
    
            cPUnits.Add cUnits
            lErrorsPerPage = 0
            lUnitsFilledPerPage = 0
            lUnitsTotalPerPage = cUnits.count
            bAtLeastOneUnitOnPage = False
    
            For Each vc In cUnits
                Set r2 = vc
    
                'file name
                aSize(5) = GetPathFromComment(r2.Cells(1))
                aSize(6) = CStr(r2.Cells(1, 1).value)
                isSPREADPage = CheckPageIsSPREAD(x2)
                If isSPREADPage And (x2 Mod 2 = 0) Then
                    pageOneSPREADComment = aSize(5)
                End If
                If isSPREADPage And (x2 Mod 2 = 1) Then
                    aSize(5) = pageOneSPREADComment
                End If
               'determine color
                If aSize(5) <> "" And aSize(6) <> "" Then
                    'normalan slucaj: i celija i komentar postoje
                    
                    aSize(5) = DirU(CStr(aSize(5)))
                    If aSize(5) = "" Then    'file does not exist, altough it is listed in commeny - put patern
                        With r2.Cells(1, 1).Interior
                            .Pattern = xlGray16
                            .PatternColor = 255
                        End With
                        lErrorsPerPage = lErrorsPerPage + 1
                        lErrorPages = lErrorPages + 1
                        bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing to draw gradient
                    Else    'file is here
                        With r2.Cells(1, 1).Interior
                            .Pattern = xlSolid
                            .PatternColor = 255
                        End With
                        lUnitsFilledPerPage = lUnitsFilledPerPage + 1
                        bAtLeastOneUnitOnPage = True
                    End If
    
'                    cUnitsSizes.Add aSize
                ElseIf aSize(5) <> "" And aSize(6) = "" Then
                    With r2.Cells(1, 1).Interior
                        .Pattern = xlGray16
                        .PatternColor = 255
                    End With
                    lErrorsPerPage = lErrorsPerPage + 1
                    lErrorPages = lErrorPages + 1
                    bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing to draw gradient
'                ElseIf r2.Cells(1).Comment Is Nothing = False And aSize(6) = "" Then
'                    With r2.Cells(1, 1).Interior
'                        .Pattern = xlGray16
'                        .PatternColor = 255
'                    End With
'                    lErrorsPerPage = lErrorsPerPage + 1
'                    lErrorPages = lErrorPages + 1
'                    bAtLeastOneUnitOnPage = True    ' mark presence altough file is missing to draw gradient
    
                Else
                    If aSize(6) <> "" Then
                        With r2.Cells(1, 1).Interior
                            .Pattern = xlGray8
                            .PatternColor = 255
                        End With
                        lErrorsPerPage = lErrorsPerPage + 1
                        lErrorPages = lErrorPages + 1
                        bAtLeastOneUnitOnPage = True
    
                    Else
                        'koristi da ovde resetujes celiju ako ima nekog formata a prazna je
                        If r2.Cells(1, 1).Interior.Pattern <> xlNone Then    ' Or r2.Cells(1, 1).Interior.TintAndShade <> 0 'mozda?
                            With r2.Cells(1, 1).Interior
                                .Pattern = xlNone
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                    End If
                End If
            Next vc
            Set rTemp = r.Range("a1")
            Set rCaption = rTemp.Offset(-1, 0)
        'End If
        If bAtLeastOneUnitOnPage = False Then
            lEmptyPages = lEmptyPages + 1
            lPercent = 0
            cpErrors.Add "Empty page"
            'cpErrors.Add NName.Name
            cpErrors.Add "Page_" & x2
            cpErrors.Add "Page " & x2
'            Call FillGradient(rCaption, lPercent)
            Call FillGradient(r, lPercent)

        ElseIf lUnitsFilledPerPage = lUnitsTotalPerPage Then
            lCompletedPages = lCompletedPages + 1
            lPercent = 100
            cpErrors.Add "Completed page"
            'cpErrors.Add NName.Name
            cpErrors.Add "Page_" & x2
            cpErrors.Add "Page " & x2
            Call FillGradient(rCaption, lPercent)
        ElseIf bAtLeastOneUnitOnPage = True Then
            lPagesInProgress = lPagesInProgress + 1
            lPercent = (lUnitsFilledPerPage / lUnitsTotalPerPage) * 100
            If lPercent = 0 Then lPercent = 1    'to draw gradient if page contain only errors
            If lErrorsPerPage < 1 Then
                cpErrors.Add "Page in progress"
                'cpErrors.Add NName.Name
                cpErrors.Add "Page_" & x2
                cpErrors.Add "Page " & x2
            Else
                cpErrors.Add "Page in progress (with " & lErrorsPerPage & " error[s])"
                'cpErrors.Add NName.Name
                cpErrors.Add "Page_" & x2
                cpErrors.Add "Page " & x2
            End If
            Call FillGradient(rCaption, lPercent)
        End If
        'cPUnitsSizes.Add cUnitsSizes
    Next x2
    MarkLastPage
    With ThisWorkbook.Sheets("Report")
        .Range("a1:b1100").ClearContents
        '.Range("a8:b1000").Clear
        .Range("a1").value = "Overview:"
        .Range("a2").value = lCompletedPages & " Completed pages"
        .Range("b2").value = lCompletedPages / (lMaxPages - 1)
        .Range("a3").value = lEmptyPages & " Empty pages"
        .Range("b3").value = lEmptyPages / (lMaxPages - 1)
        .Range("a4").value = lPagesInProgress & " Pages in progress"
        .Range("b4").value = lPagesInProgress / (lMaxPages - 1)
        .Range("a5").value = lErrorPages & " Pages with errors"
        .Range("b5").value = lErrorPages / (lMaxPages - 1)
        .Range("a7").value = "Details:"
    End With

    Set rRep = ThisWorkbook.Sheets("Report").Range("a8")
    For x = 1 To cpErrors.count Step 3
        If cpErrors(x) <> "Completed page" Then
            x3 = x3 + 1
            rRep.Offset(x3, 0).value = cpErrors(x + 2) & " " & cpErrors(x)
            rRep.Offset(x3, 0).Hyperlinks.Add Anchor:=rRep.Offset(x3, 0), Address:="", SubAddress:=cpErrors(x + 1)
            'Debug.Print cpErrors(x + 2), cpErrors(x)
        End If
    Next x
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
Sub FillGradient(r As Range, ByVal lPrecent As Long)
'make range gradient according to percent

    If lPrecent < 100 And lPrecent > 0 Then
        With r.Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 0
            .Gradient.ColorStops.Clear
        End With
        With r.Interior.Gradient.ColorStops.Add(0)
            .Color = vbGreen
            .TintAndShade = 0
        End With
        With r.Interior.Gradient.ColorStops.Add(lPrecent / 100)
            .Color = vbGreen
            .TintAndShade = 0
        End With
        With r.Interior.Gradient.ColorStops.Add(1)
            .Color = 255
            .TintAndShade = 0
        End With
    ElseIf lPrecent = 100 Then
        With r.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 11263438
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    ElseIf lPrecent = 0 Then
        With r.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub

Sub NewBuildLayout()

'    If Unplaced_AdsYesNo() = False Then Exit Sub
    'Version 3.2
    checkTags = False
    Set processedLayers = New Collection
    Set processedTags = New Collection
    
    If IDP Then Exit Sub
    Const StrokeWidth As Double = 1.38888888888889E-02
    Dim nName As Name
    Dim cPages As New Collection
    Dim cPUnits As New Collection
    Dim cPUnitsSizes As New Collection
    Dim cpUnitsPositions As New Collection
    Dim cUnits As Collection
    Dim cUnitsSizes As Collection
    Dim cUnitsColumns As Collection
    Dim cUnitsPositions As Collection
    Dim x As Long, Y As Long, z As Long
    Dim x2 As Long    'brojac za imena
    Dim r As Range, rc As Range
    Dim r2 As Range
    Dim aSize(1 To 7)    'row,col,Indesign width,height,filename,jel bled [onda 1, inace 0], ad name
    Dim aPos(1 To 2)    'top ,left
    Dim vc As Variant
    Dim lMaxPages As Long

    Dim iDA   ' As InDesign.Application
    Dim iDDoc   ' As InDesign.Document
    Dim iDPage    'As InDesign.Page
    Dim idColor    'As InDesign.Color
    Dim iDRectangle    ' As InDesign.Rectangle
    Dim myY1, myX1, myY2, myX2    'geometric bounds for rectangle
    Dim sFile As String
    Dim sTempFile As String    ' ako nema fajla [a upisano je nesto ko ad, ubaci to ko text]


    Dim lMinPage As Long
    Dim lNoOfPages As Long    'number of exported pages

    Dim bOddPage As Boolean    '?
    Dim bCS3fix As Boolean    'do not empty frames if cs3fix is active
    Dim errMsg As String ' Concatanate string to store error messages
    
    
    BuildNames ' added by Hussam
    
    If ThisWorkbook.Sheets("Settings").Range("CS3fix") = "Yes" Then bCS3fix = True

    If ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value = "" Then
        MsgBox "InDesign template is not defined!"
        Exit Sub
    ElseIf Dir(ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value) = "" Then
        MsgBox "InDesign template: " & ThisWorkbook.Sheets("Settings").Range("IndesignTemplate").value & " is not found!"
        Exit Sub
    End If
    If DirExists("C:\regkey") = False Then
        MkDir "C:\regkey"
        errMsg = errMsg & vbCrLf & "info: C:\regkey folder created!"
    End If

    Dim dHbleed As Double, dVbleed As Double    'za bleed
    dHbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Horizontal").value)
    dVbleed = val(ThisWorkbook.Sheets("Settings").Range("Bled_Size_Vertical").value)

    'ThisWorkbook.Sheets("Layout").Activate
    ThisWorkbook.Sheets("NewLayout").Activate
    If LCase(ThisWorkbook.Sheets("Settings").Range("Use_Export_Page_Settings").value) = "yes" Then
        lMinPage = ThisWorkbook.Sheets("Settings").Range("MinPageNo").value
        lMaxPages = ThisWorkbook.Sheets("Settings").Range("BOO").value
    Else
        lMinPage = 2
        lMaxPages = ThisWorkbook.Sheets("NewLayout").Range("U1").value
    End If

    'add dummy pages at 1, to keep counters sinhronized. It was needed when Yaakov asked for removing of first page
    Application.StatusBar = "Preparing..."
    cPages.Add 0&
    cPUnits.Add 0&
    cPUnitsSizes.Add 0&
    cpUnitsPositions.Add 0&

    
    lNoOfPages = lMaxPages - lMinPage + 1
    For x2 = lMinPage To lMaxPages
'        SetPageRangeAddress (x2)
'        Set NName = ThisWorkbook.Names("Page_" & x2)
        
        Set cUnits = New Collection
        Set cUnitsSizes = New Collection
        Set cUnitsPositions = New Collection
'        tmpStr = NName.RefersTo
'        vArr = Split(tmpStr, "!")
'        Set r = ThisWorkbook.Sheets("NewLayout").Range(vArr(1))
        
        'Set r = NName.RefersToRange
        Set r = GetPageRange(x2)

        cPages.Add r
        For Each rc In r.Cells
            If rc.MergeCells = True Then
                On Error Resume Next
                cUnits.Add rc.MergeArea, rc.MergeArea.Address
                On Error GoTo 0
            Else
                cUnits.Add rc
            End If
        Next rc
        cPUnits.Add cUnits
        For Each vc In cUnits
            Set r2 = vc
            aSize(1) = r2.rows.count
            aSize(2) = r2.columns.count
            aSize(3) = GetSizes(aSize(2), aSize(1))(1)
            aSize(4) = GetSizes(aSize(2), aSize(1))(2)
            'pokupi ime reklame
            aSize(7) = r2.Cells(1).value
            If Right(aSize(7), 1) = " " Then ' second page of SPREAD
                aSize(5) = prevPageComment
            Else
                aSize(5) = GetPathFromComment(r2)
                prevPageComment = aSize(5)
            End If
            
         
            If IsBleed(x2) = True Then
                If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                    aSize(3) = GetBleedPageSizes()(1) + dHbleed + dHbleed
                    aSize(4) = GetBleedPageSizes()(2) + dVbleed + dVbleed
                End If
            End If
            cUnitsSizes.Add aSize
            aPos(1) = GetPositions(r2.Cells(1, 1), r)(1)
            aPos(2) = GetPositions(r2.Cells(1, 1), r)(2)
            If aSize(1) = 4 And aSize(2) = 2 Then    'full page
                aPos(1) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Top").value
                aPos(2) = ThisWorkbook.Sheets("Settings").Range("Full_Page_Position_1_Left").value
            End If
            cUnitsPositions.Add aPos
        Next vc
        cPUnitsSizes.Add cUnitsSizes
        cpUnitsPositions.Add cUnitsPositions

    Next x2

    'id fun

    Dim sw As New StopWatch
    sw.StartTimer
    Application.StatusBar = "Starting InDesign"
    Set iDA = CreateObject("InDesign.Application")
    Dim idalerts    'da vratim nivo kukanja na pocetni nivo
    Dim idunitsHOR    'horizontalne jedinice
    Dim idunitsVER    'vertikalne jedinice
    Dim currentPageCount As Integer
    Dim oneFileFeature As Boolean
    Dim FirstPage As Long
    Dim LastPage As Long
    Dim TargetFile As String

    ' Ensure the sheet and range exist
    On Error Resume Next
    TargetFile = ThisWorkbook.Sheets("Additional_settings").Range("TargetFile").value
    On Error GoTo 0
    
    ' Check if the value is valid and not empty
    If Not IsError(TargetFile) Then
        If Len(Trim(CStr(TargetFile))) > 0 Then
            oneFileFeature = True
        Else
            oneFileFeature = False
        End If
    Else
        MsgBox "'TargetFile' range is invalid or doesn't exist.", , "Error"
    End If

    idalerts = iDA.ScriptPreferences.UserInteractionLevel
    iDA.ScriptPreferences.UserInteractionLevel = 1699640946    'idUserInteractionLevels.idNeverInteract
    ' iDA.ScriptPreferences.EnableRedraw = False
    'iDA.Visible = True
    Set iDDoc = iDA.Open(Range("IndesignTemplate").value)

    If oneFileFeature = True Then
         Set iDDoc = iDA.Open(TargetFile) ' destination file
         Set FirstTemplate = iDA.Open(Range("IndesignTemplate").value)

        'iDDestonation = iDA.ActiveDocument
    End If


    '???============= Combining 2 files =================???
    If ThisWorkbook.Sheets("Additional_settings").Range("d2").value > 0 Then

        Application.StatusBar = "Combining InDesign Templates"

        Dim second_template As Variant

        If oneFileFeature = True Then

                ' Open the second document
                Set doc1 = iDA.Open(FirstTemplate)
        
                ' Loop through layers in the First document
            For Each layer In doc1.Layers
                ' Create a new layer in the Main document with the layers name
                Set newLayer = iDDoc.Layers.Add
                newLayer.Name = layer.Name
                
                ' Loop through page items in the current layer
                For Each item In layer.pageItems
                    ' Duplicate item to the new layer in the main document
                    Set newItem = item.Duplicate(newLayer)
                Next item
            Next layer

    End If 'oneFileFeature = True

        ' Get the second template path from the sheet

        second_template = ThisWorkbook.Sheets("Additional_settings").Range("d2").value

        ' Open the second document
        Set doc2 = iDA.Open(second_template)
            
        ' Loop through layers in the Second document
        For Each layer In doc2.Layers
            ' Create a new layer in the Main document with the layers name
            Set newLayer = iDDoc.Layers.Add
            newLayer.Name = layer.Name
            
            ' Loop through page items in the current layer
            For Each item In layer.pageItems
                ' Duplicate item to the new layer in the main document
                Set newItem = item.Duplicate(newLayer)
            Next item
        Next layer
    End If
    '???======================end combining 2 file===============???

    iDA.Windows(1).Minimize    'ako je prozor skriven ovo pravi gresku!
    ThisWorkbook.Activate
    'set jedinice mere
    idunitsHOR = iDDoc.ViewPreferences.HorizontalMeasurementUnits
    iDDoc.ViewPreferences.HorizontalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
    idunitsVER = iDDoc.ViewPreferences.HorizontalMeasurementUnits
    iDDoc.ViewPreferences.VerticalMeasurementUnits = 2053729891    'idMeasurementUnits.idInches
    ThisWorkbook.Activate

    Application.StatusBar = "Adding pages to InDesign"

    If oneFileFeature = True Then
        currentPageCount = iDDoc.Pages.count
 
        ' Add missing pages until lMaxPages is reached
        If currentPageCount < lMaxPages Then
            For x = currentPageCount + 1 To lMaxPages
                iDDoc.Pages.Add 1701733408 ' idAtEnd
                Application.StatusBar = "Adding pages to InDesign: " & x & " of " & lMaxPages
            Next x
        End If

    Else 'if oneFileFeature = False then
        
        For x = lMinPage To lMaxPages
            iDDoc.Pages.Add 1701733408    'idAtEnd
            Application.StatusBar = "Adding pages to InDesign: " & x & " of " & lMaxPages
        Next x

    End If

    ' We'll need to create a color. Check to see if the color already exists.
    On Error Resume Next
    Set idColor = iDDoc.Colors.item("YaakovBlack")
    If Error.Number <> 0 Then
        Set idColor = iDDoc.Colors.Add
        idColor.Name = "YaakovBlack"
        idColor.Model = idColorModel.idProcess
        idColor.ColorValue = Array(0, 0, 0, 100)
        Error.Clear
    End If
    ' Resume normal error handling.
    On Error GoTo 0

    '???=========== Detect and handel new Layer ====================???
    'get articleColor color
    articleColor = ThisWorkbook.Sheets("Additional_settings").Range("b7").Interior.Color

    If articleColor = 0 Then
        errMsg = errMsg & vbCrLf & "-Error: The Article Color have no fill."
    End If

    ' Add a New layer To the document
    LSPLayerName = "LSP"

    layerExists = False

    For i = 1 To iDDoc.Layers.count
        If iDDoc.Layers.item(i).Name = LSPLayerName Then
            layerExists = True
            Exit For
        End If
    Next i

    ' Create the layer If it doesn't exist
    If Not layerExists Then
        Set LSPLayer = iDDoc.Layers.Add
        LSPLayer.Name = LSPLayerName
    Else
        Set LSPLayer = iDDoc.Layers.item(LSPLayerName)
    End If
    '???=========== Detect and handel new Layer ====================???
    If oneFileFeature = True Then

        FirstPage = lMinPage
        LastPage = lMaxPages

    Else 'if oneFileFeature = False then
        FirstPage = 2
        LastPage = lNoOfPages
    End If
        
        Debug.Print "first Page: " & FirstPage & " last page: " & LastPage
    For x = FirstPage To LastPage + 1
        Application.StatusBar = "Creating boxes for page: " & x & " (of " & LastPage & ")"
        Set iDPage = iDDoc.Pages(x)
        z = x - FirstPage + 2
        For Y = 1 To cPUnits(z).count
            Set vc = cPUnits(z).item(Y)
            myY1 = cpUnitsPositions(z).item(Y)(1)
            myY2 = myY1 + cPUnitsSizes(z).item(Y)(4)

            myX1 = cpUnitsPositions(z).item(Y)(2)
            myX2 = myX1 + cPUnitsSizes(z).item(Y)(3)
            sFile = cPUnitsSizes(z).item(Y)(5)
            sTempFile = cPUnitsSizes(z).item(Y)(7)
            
            Debug.Print "indsign and excel page: " & x & " page index: " & z & " ad name: " & sTempFile & " src: " & sFile
            


            ' Get current cell color
            currentColor = vc.Interior.Color

            If currentColor <> articleColor Then 'If is Not an Article
            
                Set iDRectangle = iDPage.Rectangles.Add
                iDRectangle.GeometricBounds = Array(CDbl(myY1), CDbl(myX1), CDbl(myY2), CDbl(myX2))
    '            If ThisWorkbook.Worksheets("Settings").Range("BlackBorder").Value = "Yes" Then
                If ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value > 0 Then
                    iDRectangle.StrokeWeight = ThisWorkbook.Worksheets("Settings").Range("BlackBorder").value
                    iDRectangle.StrokeColor = iDDoc.Swatches.item("YaakovBlack")
                    iDRectangle.StrokeAlignment = 1936998729    ' idInsideAlignment
                Else
                    iDRectangle.StrokeWeight = 0
                End If

                    If Len(sFile) > 0 And DirU(sFile) <> "" Then
                        If getExt(sFile) <> "indd" Then
                            On Error Resume Next ' Allow errors without breaking the program
                                iDRectangle.place sFile ' Try placing the file
                                
                                If Err.Number <> 0 Then
                                    ' Show a message about the corrupted file
                                    errMsg = errMsg & vbCrLf & "-Error: " & sFile & " is a corrupted file."
                                    Err.Clear ' Clear the error to avoid interference later
                                End If
                                
                                ' Continue with the next action, regardless of error
                                iDRectangle.Fit 1668575078 ' idFitOptions.idContentToFrame
                        Else
                        'omoguci ovo kad plati
        '                    iDA.ScriptPreferences.UserInteractionLevel = 1699311169
                            iDRectangle.place sFile, True
                            iDRectangle.Fit 1668575078    'idFitOptions.idContentToFrame
                            'i ovo
        '                    iDA.ScriptPreferences.UserInteractionLevel = 1699640946
                        End If
                    Else
                        'set options for frame to fit when place file manually
                        If bCS3fix = False Then
                            iDRectangle.FrameFittingOptions.AutoFit = True
                            iDRectangle.FrameFittingOptions.FittingOnEmptyFrame = 1668575078    'idEmptyFrameFittingOptions.idContentToFrame
                        End If
                        If Len(sFile) > 0 And DirU(sFile) = "" Then
                            If Range("WriteText").value = "Yes" Then
                                ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                Str2TXT CStr(sFile), "C:\regkey\mfile.txt"
                                'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                iDRectangle.place "C:\regkey\mfile.txt"
                            End If
                        ElseIf Len(sFile) = 0 And Len(sTempFile) > 0 Then
                            If Range("WriteText").value = "Yes" Then
                                ' Str2TXT CStr(sFile), ThisWorkbook.Path & "\mfile.txt"
                                Str2TXT CStr(sTempFile), "C:\regkey\mfile.txt"
                                'iDRectangle.Place ThisWorkbook.Path & "\mfile.txt"
                                iDRectangle.place "C:\regkey\mfile.txt"
                            End If
                        End If
                    End If

                ElseIf currentColor = articleColor Then ' Else If its Not a AD nor an Article
                    '???=========================== LAYER BEGIN =============================???

                    elementName = sTempFile

                    ' Check If layer has already been processed
                    LayerProcessed = False ' Skip the rest of the processing For this layer

                    On Error Resume Next
                    processedLayers.Add elementName, elementName

                    If Err.Number = 457 Then ' Layer already processed (duplicate key)
                        Err.Clear
                        Debug.Print "Skipping layer: " & elementName
                        LayerProcessed = True ' Skip the rest of the processing For this layer
                    End If

                    layerExists = False

                    Debug.Print ("checking If File has Layer " & elementName)

                    'checking For matching layer name
                    For i = 1 To iDDoc.Layers.count
                        If iDDoc.Layers.item(i).Name = elementName Then
                            Set layer = iDDoc.Layers.item(i)
                            layerExists = True
                            Exit For
                        End If
                    Next i

                    If layerExists And LayerProcessed = False Then

                        Application.StatusBar = "Moving layer " & elementName

                        Set layerItems = layer.pageItems

                        layerGroup = False

                        If layerItems.count > 1 Then
                            layerGroup = True
                        End If
                        ' Redefine the array To hold all page items
                        ReDim sortedLayerItems(1 To layerItems.count)

                        ' Populate the array With page items
                        For i = 1 To layerItems.count
                            Set sortedLayerItems(i) = layerItems(i)
                        Next i

                        ' Sort the array by page item id
                        For i = 1 To UBound(sortedLayerItems) - 1
                            For j = i + 1 To UBound(sortedLayerItems)
                                If sortedLayerItems(i).id > sortedLayerItems(j).id Then
                                    ' Swap the items
                                    Set temp = sortedLayerItems(i)
                                    Set sortedLayerItems(i) = sortedLayerItems(j)
                                    Set sortedLayerItems(j) = temp
                                End If
                            Next j
                        Next i

                        ' Move all items in the matched layer To the New location
                        For j = 1 To UBound(sortedLayerItems)
                            Set pageItem = sortedLayerItems(j)

                            ' Calculate the New position based on current position
                            Dim bounds As Variant
                            bounds = pageItem.GeometricBounds
                            Dim currentX As Double, currentY As Double
                            currentX = bounds(1)
                            currentY = bounds(0)

                            ' Calculate the offset To move the items
                            Dim offsetX As Double, offsetY As Double

                            oldPage = pageItem.ParentPage.DocumentOffset ' Get the current page number (1-based in VB

                            If lastelementName = elementName Then ' If its the second With the same layer

                                offsetPage = offsetPage ' Keep offset from previous one
                            Else ' it the first item in layer
                                offsetPage = x - oldPage ' Define the offset , Current page - where it is -oldPage

                                lastelementName = elementName ' Update the lastelementName
                            End If

                            If layerGroup = True Then

                                layerMoveOnPage = False
                            Else
                                layerMoveOnPage = True
                            End If

                            ' Move the page item by the calculated offset
                            newPageNumber = oldPage + offsetPage


                            If newPageNumber < 1 Or newPageNumber > iDDoc.Pages.count Then
                                MsgBox "Can't move further backwards Or beyond the document."
                                Exit Sub
                            End If

                            Set targetPage = iDDoc.Pages(newPageNumber)

                            If layerMoveOnPage = True Then
                                ' Calculate the offset To move the items
                                offsetX = myX1
                                offsetY = myY1

                            Else

                                offsetX = currentX
                                offsetY = currentY

                            End If

                            pageItem.Move targetPage
                            pageItem.Move Array(offsetX, offsetY)
                            Debug.Print "Moved Page Item: " & pageItem.id & " within Layer " & elementName & " from page " & oldPage & " To page " & newPageNumber

                        Next j
                    ElseIf checkTags = True And LayerProcessed = False Then 'If layer dosnt exsist, chck For xml Tag

                        errMsg = errMsg & vbCrLf & "-Error: Layer " & elementName & " was Not found."

                        '???========================= LAYER END | XML TAG BEGIN =====================================???

                        ' Check If Tag has already been processed
                        TagProcessed = False ' on true, Skip the rest of the processing For this layer


                        On Error Resume Next
                        processedTags.Add elementName, elementName
                        If Err.Number = 457 Then ' Tag already processed (duplicate key)
                            Err.Clear
                            Debug.Print "Skipping Tag: " & elementName
                            TagProcessed = True ' Skip the rest of the processing For this Tag
                        End If


                        ' Get all page items in the document
                        Set pageItems = iDDoc.pageItems

                        ' Redefine the array To hold all page items
                        ReDim sortedPageItems(1 To pageItems.count)

                        ' Populate the array With page items
                        For i = 1 To pageItems.count
                            Set sortedPageItems(i) = pageItems(i)
                        Next i

                        ' Sort the array by page item id using Bubble Sort
                        For i = 1 To UBound(sortedPageItems) - 1
                            For j = i + 1 To UBound(sortedPageItems)
                                If sortedPageItems(i).id > sortedPageItems(j).id Then
                                    ' Swap the items
                                    Set temp = sortedPageItems(i)
                                    Set sortedPageItems(i) = sortedPageItems(j)
                                    Set sortedPageItems(j) = temp
                                End If
                            Next j
                        Next i
                        Debug.Print ("checking If Tag " & elementName & " exsist")


                        ' Loop through all page items
                        For i = 1 To UBound(sortedPageItems)
                            'Set pageItem = sortedPageItems(i).Item(i)

                            ' Check If the page item has an associated XML element
                            On Error Resume Next
                            Set xmlTag = sortedPageItems(i).AssociatedXMLElement

                            On Error GoTo 0

                                If Not xmlTag Is Nothing And TagProcessed = False Then 'if xmlYag found a tag and tag processed is False

                                    ' Check If the XML tag name matches Tag Name
                                    If xmlTag.MarkupTag.Name = elementName Then

                                        Application.StatusBar = "Moving tag " & elementName

                                        tagExists = True


                                        ' Check If this is same tag name than the previous one
                                        If lastelementName = elementName Then
                                            offsetPage = offsetPage ' Keep offset from previous one
                                            oldPage = sortedPageItems(i).ParentPage.DocumentOffset
                                            tagMoveOnPage = False

                                        Else
                                            ' Get the current page number (1-based in VBA)
                                            oldPage = sortedPageItems(i).ParentPage.DocumentOffset

                                            ' Define the offset , Current page - where it is -oldPage
                                            offsetPage = x - oldPage

                                            lastelementName = elementName ' Update the lastelementName
                                            tagMoveOnPage = False

                                        End If



                                        newPageNumber = oldPage + offsetPage
                                        'Debug.Print ("x is now on page " & X & " Old Page is " & oldPage & " So x-oldpage = " & offsetPage & ", New page number is " & newPageNumber)
                                        ' Ensure the New page number is valid
                                        If newPageNumber < 1 Or newPageNumber > iDDoc.Pages.count Then
                                            MsgBox "Can't move further backwards Or beyond the document."
                                            Exit Sub
                                        End If

                                        ' Get the target page based on the offset
                                        Set targetPage = iDDoc.Pages(newPageNumber)


                                        ' Calculate the New position based on current rectangle
                                        bounds = sortedPageItems(i).GeometricBounds
                                        currentX = bounds(1)
                                        currentY = bounds(0)

                                        If tagMoveOnPage = True Then
                                            ' Calculate the offset To move the items
                                            offsetX = myX1 - currentX
                                            offsetY = myY1 - currentY

                                        Else

                                            offsetX = currentX
                                            offsetY = currentY

                                        End If

                                        ' Move the page item To the target page And restore its geometric bounds
                                        sortedPageItems(i).Move targetPage
                                        sortedPageItems(i).Move Array(offsetX, offsetY)


                                        Debug.Print "Moved Element ID: " & sortedPageItems(i).id; " With Tag Name " & elementName & " from page " & oldPage & " To page " & newPageNumber & " Page Item is " & i
                                    End If 'XML tag name matches Tag Name

                                End If 'If Not xmlTag Is Nothing


                            Next i

                            '************************** XML TAG END ***********************************

                    End If 'End If layerExists

                If tagExists = True Or layerExists = True Then

                    'Debug.Print "Tag '" & elementName & "' DOES exist."
                Else
                    errMsg = errMsg & vbCrLf & "-Error: Tag Or Layer '" & elementName & "' does Not exist."
                    'Exit Sub
                End If
            End If 'If is Not an Article

        Next Y
                                    'Debug.Print cPUnits(x).Count
    Next x

Dim maxpage As Integer

 maxpage = lNoOfPages + 1
If oneFileFeature = False Then

    For p = iDDoc.Pages.count To 1 Step -1
        Debug.Print "maxpage is "; maxpage & " Checking page: " & p
        ' Only delete pages strictly greater than lNoOfPages
        If p > maxpage Then
            Debug.Print "maxpage is "; maxpage & " Deleting page: " & p
            iDDoc.Pages(p).Delete
        End If
    Next p
        
End If
    'vrati jedinice mere
    iDDoc.ViewPreferences.HorizontalMeasurementUnits = idunitsHOR
    iDDoc.ViewPreferences.VerticalMeasurementUnits = idunitsVER
    'vrati kukanje za script
    iDA.ScriptPreferences.UserInteractionLevel = idalerts
    Application.StatusBar = "Saving InDesign file"
'    sIssue = ThisWorkbook.Worksheets("Main List").Range("CurrentIssue").value

If oneFileFeature = True Then
    If iDDoc.Saved Then
        ' Save directly to the existing file path
        iDDoc.save
    Else
        ' Force save without triggering Save As dialog
        iDDoc.save TargetFile
    End If
    
    errMsg = errMsg & vbCrLf & "-Info: oneFileFeature is on and your file was saved to " & TargetFile

Else

    sIssue = GetCurrentIssue
    iDDoc.save ThisWorkbook.path & "\issue " & sIssue & "-" & Format(Now, "yy-mm-dd-hh-nn") & ".indd"
End If

    ' iDDoc.Windows.Add 'dodaje prozor, tj pokazuje skriveni dokument
    'iDDoc.Close idSaveOptions.idNo
    Application.StatusBar = "Done"
    ThisWorkbook.Activate
    MsgBox "Done for: " & sw.EndTimer / 1000 & " seconds." & vbCrLf & errMsg

End Sub
Function GetPositionsRedni(rSout As Range, rSearchIn As Range) As Long 'vraca od 1 do 8
    Dim ap(1 To 2)
    Dim x As Long
    Dim r As Range, rc As Range
    Set r = Range("PositionForID")
    Dim lPos As Long
    For x = 1 To rSearchIn.Cells.count
        Set rc = rSearchIn.Cells(x)
        If rc.Address = rSout.Address Then
            lPos = x
        End If
    Next x
'    For Each rc In r.Cells
'        If rc.Value = lPos Then
'            ap(1) = rc.Offset(0, 1).Value
'            ap(2) = rc.Offset(0, 2).Value
'            Exit For
'        End If
'    Next rc

    GetPositionsRedni = lPos
End Function
Private Function GetPositions(rSout As Range, rSearchIn As Range) As Variant()
    Dim ap(1 To 2)
    Dim x As Long
    Dim r As Range, rc As Range
    Set r = ThisWorkbook.Worksheets("Settings").Range("PositionForID")
    Dim lPos As Long
    For x = 1 To rSearchIn.Cells.count
        Set rc = rSearchIn.Cells(x)
        If rc.Address = rSout.Address Then
            lPos = x
        End If
    Next x
    For Each rc In r.Cells
        If rc.value = lPos Then
            ap(1) = rc.Offset(0, 1).value
            ap(2) = rc.Offset(0, 2).value
            Exit For
        End If
    Next rc

    GetPositions = ap
End Function
 Function GetSizesNames(lCols, lRows) As String
    Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = ThisWorkbook.Worksheets("Settings").Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 1).value = lCols And rc.Offset(0, 2).value = lRows Then
            ad(1) = rc.Offset(0, 5).value
            
            Exit For
        End If
    Next rc
    GetSizesNames = ad(1)
End Function
 Private Function GetSizes(lCols, lRows) As Variant()
  Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 1).value = lCols And rc.Offset(0, 2).value = lRows Then
            ad(1) = rc.Offset(0, 3).value
            ad(2) = rc.Offset(0, 4).value
            Exit For
        End If
    Next rc
    GetSizes = ad
    
   
End Function
 Private Function GetBleedPageSizes() As Variant()
 Dim ad(1 To 2)
    Dim r As Range, rc As Range
    Set r = Range("Type_Size")
    For Each rc In r.Cells
        If rc.Offset(0, 5).value = "FB" Then
            ad(1) = rc.Offset(0, 3).value
            ad(2) = rc.Offset(0, 4).value
            Exit For
        End If
    Next rc
    GetBleedPageSizes = ad
 
 End Function

Function GetFirstFile(ByVal sName As String, Optional bColor As Boolean = False) As String
    Dim sTemp As String
    Dim sSuffix As String
    Dim sBaseFolder As String
    
    sSuffix = " -col"
    sBaseFolder = Range("Base_folder").value
    sBaseFolder = TrailingSlash(sBaseFolder)
    If Len(Trim(sBaseFolder)) < 3 Then
        GetFirstFile = ""
    End If
    Dim x As Long
    Dim aFN()
    aFN = Array(".jpg", ".tif", ".psd", ".pdf")
    
    If bColor = False Then
        For x = LBound(aFN) To UBound(aFN)
            sTemp = sBaseFolder & sName & aFN(x)
            If DirU(sTemp) <> "" Then
                GetFirstFile = sTemp
                Exit Function
            End If
        Next x
    Else
        For x = LBound(aFN) To UBound(aFN)
            sTemp = sBaseFolder & sName & sSuffix & aFN(x)
            If DirU(sTemp) <> "" Then
                GetFirstFile = sTemp
                Exit Function
            End If
        Next x
    End If

End Function

Public Sub MarkLastPage()
    Dim lPages As Long
    Dim r As Range, r2 As Range, r3 As Range, R4 As Range
'    If Range("MaxNoPages").Value = "" Then
'        Exit Sub
'    End If
'    lPages = Range("MaxNoPages").Value

    lPages = ThisWorkbook.Sheets("NewLayout").Range("U1").value

'    If lPages < 2 Or lPages > UBound(PagesArr) Then
'        lPages = UBound(PagesArr)
'    End If
On Error GoTo skipEOB
    With Range("EndOfBook").Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
skipEOB:
    'Set r = ThisWorkbook.Names("Page_" & lPages).RefersToRange
    On Error GoTo 0
    Set ws = ThisWorkbook.Sheets("NewLayout")
'
'    PageR = PagesStartRow + ((lPages - 2) \ 8) * RowsPerPage
'    PageC = startCol + ((lPages - 2) Mod 8) * columnsPerPage
'    Set r = ws.Range(ws.Cells(PageR + 1, PageC), ws.Cells(PageR + RowsPerPage - 2, PageC + 1))
    
    Set r = GetPageRange(lPages)

    'Debug.Print r.columns.Count
    Set r2 = r.Cells(4, 2).Offset(0, 1)
    Set r3 = r2.Offset(-3, 0)
    Set R4 = Range(r2.Cells(1, 1), r3.Cells(1, 1))
     ThisWorkbook.Names("EndOfBook").RefersTo = "=NewLayout!" & R4.Address
    With R4.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = vbBlue
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

' Sub AddAndOverwritePages()

'         Dim lMinPage As Integer
'         Dim lMaxPages As Integer
'         Dim currentPageCount As Integer
'         Dim x As Integer
    
'         ' Example values for demonstration (replace with your actual variables)
'         lMinPage = 14
'         lMaxPages = 20
    
'         ' Get the current number of pages in the document
'         currentPageCount = iDDoc.Pages.Count
    
'         Application.StatusBar = "Adding pages to InDesign..."
    
'         ' Step 1: Add missing pages until lMaxPages is reached
'         If currentPageCount < lMaxPages Then
'             For x = currentPageCount + 1 To lMaxPages
'                 iDDoc.Pages.Add 1701733408 ' idAtEnd
'                 Application.StatusBar = "Adding pages to InDesign: " & x & " of " & lMaxPages
'             Next x
'         End If
    
'         ' Step 2: Overwrite pages in the specified range
'         For x = lMinPage To lMaxPages
'             ' Overwrite the content of the page (this depends on what "overwrite" means for your case)
'             ' For example, clearing or replacing page content could be done here
'             Application.StatusBar = "Working on page: " & x & " of " & lMaxPages
'             ' Implement your page-specific operations here
'         Next x
    
'         Application.StatusBar = "Operation complete!"
    
' End Sub
    



