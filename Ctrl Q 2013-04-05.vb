Sub SalesQuoteFormatting()
    '
    ' SalesQuoteFormatting Macro
    ' This macro will auto set margins, scaling and other formatting features needed for proper Quote printing/PDFing.
    '
    ' Keyboard Shortcut: Ctrl+q
    '
    ' Copy Platinum, Gold, or Solver from template file SupportTypes.xls and paste to current Sales Quote
    
    Dim isQuote As Boolean, isRequirements As Boolean, isTnC As Boolean, isSnP As Boolean, isSupportPlatinum As Boolean, isSupportGold As Boolean, isSupportSilver As Boolean
    Dim sngStart As Single, sngEnd As Single, sngElapsed As Single
    Dim ActiveCellWidth As Single, PossNewRowHeight As Single, CurrentRowHeight As Single, MergedCellRgWidth As Single
    Dim UsedRange As Range, TnCRange As Range, Cell As Range, innerCell As Range, LastCell As Range
    Dim WS_Count As Integer, FirstSheet_PageCount As Integer, I As Integer, LastRow As Integer, N As Integer, ActiveSht As Integer
    Dim ActiveWB As String
    
    sngStart = Timer
    Application.ScreenUpdating = False
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    For I = 1 To WS_Count ' process all sheets
        With ActiveWorkbook.Worksheets(I)
            LastRow = .Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByRows).Row
            
            isQuote = False
            isRequirements = False
            isTnC = False
            isSnP = False
            isSupportPlatinum = False
            isSupportGold = False
            isSupportSilver = False
            
            Set UsedRange = .Range(.Range("E8"), .Range("AF8"))
            For Each Cell In UsedRange
                With Cell.MergeArea
                    If Not isQuote And Mid(Cell, 1) = "QUOTATION" Then
                        isQuote = True
                    End If
                End With
            Next Cell
            
            Set UsedRange = Nothing
            Set Cell = Nothing
            
            If Not isQuote Then
                Set UsedRange = .Range(.Range("D7"), .Range("I7"))
                For Each Cell In UsedRange
                    With Cell.MergeArea
                        If Not isRequirements And Mid(Cell, 1) = "Customer Requirements" Then
                            isRequirements = True
                        End If
                    End With
                Next Cell
            End If
            
            Set UsedRange = Nothing
            Set Cell = Nothing
            
            If Not isQuote And Not isRequirements Then
                Set UsedRange = .Range(.Range("C6"), .Range("I6"))
                For Each Cell In UsedRange
                    With Cell.MergeArea
                        If Not isSnP And Mid(Cell, 1) = "SUPPORT AND SERVICE LEVEL AGREEMENT" Then
                            isSnP = True
                        End If
                    End With
                Next Cell
            End If
            
            Set UsedRange = Nothing
            Set Cell = Nothing
            
            If Not isQuote And Not isRequirements And Not isSnP Then
                Set UsedRange = .Range(.Range("E7"), .Range("I7"))
                For Each Cell In UsedRange
                    With Cell.MergeArea
                        If Not isTnC And InStr(1, Cell, "TERMS AND CONDITIONS OF AGREEMENT") > 0 Then
                            isTnC = True
                        End If
                    End With
                Next Cell
            End If

            If Not isQuote And Not isRequirements And Not isSnP And Not isTnC Then
                Set UsedRange = .Range(.Range("D6"), .Range("F6"))
                For Each Cell In UsedRange
                    If Not isSupportPlatinum And Not isSupportGold And Not isSupportSilver Then
                        With Cell.MergeArea
                            If InStr(1, Cell, "Platinum") > 0 Then
                                isSupportPlatinum = True
                            ElseIf InStr(1, Cell, "Gold") > 0 Then
                                isSupportGold = True
                            ElseIf InStr(1, Cell, "Silver") > 0 Then
                                isSupportSilver = True
                            End If
                        End With
                        If isSupportPlatinum Or isSupportGold Or isSupportSilver Then
                            ActiveWB = ActiveWorkbook.Name
                            ActiveSht = I
                        End If
                    End If
                Next Cell
            End If
            
            Set UsedRange = Nothing
            Set Cell = Nothing
            
            With .PageSetup
                .LeftMargin = Application.InchesToPoints(0)
                .RightMargin = Application.InchesToPoints(0)
                .TopMargin = Application.InchesToPoints(0)
                .BottomMargin = Application.InchesToPoints(0.5)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0)
                .CenterHorizontally = True
                .CenterVertically = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .CenterFooter = ""
                .LeftFooter = ""
                .AlignMarginsHeaderFooter = True
                
                If isQuote Then
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range(ActiveWorkbook.Worksheets(I).Range("D11"), ActiveWorkbook.Worksheets(I).Range("W11"))
                    Call ResizeCells(UsedRange)
                    Set UsedRange = Nothing
                            
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range(ActiveWorkbook.Worksheets(I).Range("D13"), ActiveWorkbook.Worksheets(I).Range("W19"))
                    Call ResizeCells(UsedRange)
                    Set UsedRange = Nothing
                            
                    Set LastCell = ActiveWorkbook.Worksheets(I).Range(Cells(LastRow, "T").Address)
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range(ActiveWorkbook.Worksheets(I).Range("E27"), LastCell)
                    Call ResizeCells(UsedRange)
                    Set UsedRange = Nothing
                    
                    FirstSheet_PageCount = ActiveWorkbook.Worksheets(I).HPageBreaks.Count + 1
                    
                    .RightFooter = "Page &P of " & FirstSheet_PageCount + WS_Count - 1
                    .LeftFooter = "&7 - Prices exclude taxes, shipping costs, and traveling expenses" & Chr(10) + " - Please email Purchase Orders to: PurchaseOrders@adtechglobal.com or Fax to (678) 679-2347 - Attn: Sales" & Chr(10) + " - Options are not included in total."
                ElseIf isRequirements Then
                    Set LastCell = ActiveWorkbook.Worksheets(I).Range(Cells(LastRow, "K").Address)
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range(ActiveWorkbook.Worksheets(I).Range("D9"), LastCell)
                    Call ResizeCells(UsedRange)
                    Set UsedRange = Nothing
                    Set LastCell = Nothing
                
                    .RightFooter = "Page " & FirstSheet_PageCount + I - 1 & " of " & FirstSheet_PageCount + WS_Count - 1
                    .LeftFooter = "&7 - Prices exclude taxes, shipping costs, and traveling expenses" & Chr(10) + " - Please email Purchase Orders to: PurchaseOrders@adtechglobal.com or Fax to (678) 679-2347 - Attn: Sales" & Chr(10) + " - Options are not included in total."
                ElseIf isTnC Then
                    Set TnCRange = Union(Range("E11:J11"), Range("E13:J13"), Range("E15:J15"), Range("E17:J17"), Range("E19:J19"))
                    TnCRange.Font.Size = 6
                    Set TnCRange = Nothing
                    
                    For N = 12 To 18 Step 2
                        Rows(N).Hidden = True
                    Next N
                    
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("E11")
                    UsedRange.RowHeight = 273.5
                    Set UsedRange = Nothing
                    
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("E13")
                    UsedRange.RowHeight = 249.5
                    Set UsedRange = Nothing
                    
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("E15")
                    UsedRange.RowHeight = 93.25
                    Set UsedRange = Nothing
                    
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("E17")
                    UsedRange.RowHeight = 340.5
                    Set UsedRange = Nothing

                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("E19")
                    UsedRange.RowHeight = 236.25
                    Set UsedRange = Nothing
                    .FitToPagesTall = 2
                
                    .RightFooter = "Page " & FirstSheet_PageCount + I - 1 & " of " & FirstSheet_PageCount + WS_Count - 1
                    .LeftFooter = "&7 - Prices exclude taxes, shipping costs, and traveling expenses" & Chr(10) + " - Please email Purchase Orders to: PurchaseOrders@adtechglobal.com or Fax to (678) 679-2347 - Attn: Sales" & Chr(10) + " - Options are not included in total."
                ElseIf isSnP Then
                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("C48")
                    UsedRange.RowHeight = 97
                    Set UsedRange = Nothing

                    Set UsedRange = ActiveWorkbook.Worksheets(I).Range("C50")
                    UsedRange.RowHeight = 78
                    Set UsedRange = Nothing
                    
                    .FitToPagesTall = 1
                
                    .RightFooter = "Page " & FirstSheet_PageCount + I - 1 & " of " & FirstSheet_PageCount + WS_Count - 1
                    .LeftFooter = "&7 - Prices exclude taxes, shipping costs, and traveling expenses" & Chr(10) + " - Please email Purchase Orders to: PurchaseOrders@adtechglobal.com or Fax to (678) 679-2347 - Attn: Sales" & Chr(10) + " - Options are not included in total."
                ElseIf isSupportPlatinum Or isSupportGold Or isSupportSilver Then
                    Workbooks.Open Filename:="C:\Sales Quote\Template\SupportTypes.xls"
                    If isSupportPlatinum Then
                        Sheets("Platinum").Copy After:=Workbooks(ActiveWB).Sheets(ActiveSht)
                    ElseIf isSupportGold Then
                        Sheets("Gold").Copy After:=Workbooks(ActiveWB).Sheets(ActiveSht)
                    ElseIf isSupportSilver Then
                        Sheets("Silver").Copy After:=Workbooks(ActiveWB).Sheets(ActiveSht)
                    End If
                    Workbooks("SupportTypes.xls").Close
                    Application.DisplayAlerts = False
                    Workbooks(ActiveWB).Sheets(ActiveSht).Delete
                    Application.DisplayAlerts = True
                Else
                    .FitToPagesTall = 1
                End If
            End With
        End With
    Next I
    
    Application.ScreenUpdating = True
    sngEnd = Timer
    sngElapsed = Format(sngEnd - sngStart, "Fixed")
    
    MsgBox "Autoformatting complete (Time elapsed: " & sngElapsed & ")"
End Sub

Function ResizeCells(cellRange As Range)
    For Each Cell In cellRange
        If Cell <> "" And Cell.MergeArea.Count > 1 Then
            With Cell.MergeArea
                If .Rows.Count = 1 Then
                    CurrentRowHeight = .RowHeight
                    ActiveCellWidth = .Cells(1).ColumnWidth
                    MergedCellRgWidth = 0
                    For Each innerCell In Cell.MergeArea
                        MergedCellRgWidth = innerCell.ColumnWidth + MergedCellRgWidth
                    Next
                    .MergeCells = False
                    .Cells(1).ColumnWidth = MergedCellRgWidth + Cell.MergeArea.Count
                    .EntireRow.AutoFit
                    PossNewRowHeight = .RowHeight
                    .Cells(1).ColumnWidth = ActiveCellWidth
                    .MergeCells = True
                    .RowHeight = IIf(CurrentRowHeight > PossNewRowHeight, CurrentRowHeight, PossNewRowHeight)
                End If
            End With
        End If
    Next Cell
    Set Cell = Nothing
    Set innerCell = Nothing
End Function





