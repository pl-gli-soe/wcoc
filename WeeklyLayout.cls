VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeeklyLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2019 FORREST
' Mateusz Milewski mateusz.milewski@mpsa.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

Implements ILayout


Private layoutType As E_TYPE_OF_LAYOUT
Private area As Range

Private coverCols As Integer
Private s As Worksheet


Private dataFromCorail As SuitableData2720


Private LAYOUT_2510_LABEL As String



Private Sub Class_Initialize()

    ' layoutType
    ''
    '
    layoutType = E_TYPE_OF_LAYOUT_WEEKLY_UNDEFINED
    '
    ''
    Set dataFromCorail = Nothing
    
    
    LAYOUT_2510_LABEL = "QTY SELLER      SHIPPER    TYPE  ST  DHEO           DHRQ                        " & chr(10)

End Sub

Private Sub ILayout_finalTouchOnRep(rep As Worksheet)
    

    If rep.Cells(1, 1).Value Like "Weekly Coverage on Corail*" Then
        repFinalTouch rep
    Else
        ordersFinalTouch rep
    End If
End Sub


Private Sub repFinalTouch(rep As Worksheet)
    rep.Columns("A:A").ColumnWidth = 4
    ' rep.Columns("B:B").ColumnWidth = 17
    rep.Columns("C:C").ColumnWidth = 4 ' changed from 8
    rep.Columns("D:D").ColumnWidth = 1
    rep.Columns("J:AC").ColumnWidth = 6
    ''''''''''''''''''''''''''''''''''''
    
    
    ' Paulina
    ''''''''''''''''''''''''''''''''''''
    rep.Columns("B:B").EntireColumn.AutoFit
    rep.Activate
    '''''''''''''''''''''''''''''''''''''
    ActiveWindow.DisplayGridlines = False
    
    '''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''
End Sub

Private Sub ordersFinalTouch(o As Worksheet)

    ' od kiedy tworzenie listy jest opcjonalne nie ma potrzeby juz jej chowac
    ' jednak samego suba sobie zostawie dla potencjalnych przyszlych formowatowan
    ' o.Visible = xlSheetHidden
End Sub

Private Sub ILayout_nowMergeWithDataFromCorail(r As Range, main2720data As SuitableData2720)
    Set dataFromCorail = main2720data
End Sub

Private Sub ILayout_putSomeColorsForThisPartNumber(r As Range)


    If Not dataFromCorail Is Nothing Then
    
    
        Dim rrow As Long


        Set area = getAreaBasedOn(r)
        Set s = area.Parent
        rrow = r.Row
        
        If coverCols = 0 Then
            coverCols = 20
        End If
        
        pFormatCommonDataOnLeft s, CLng(rrow), coverCols
        
        
        
        
        
        
        
    
        Dim cwr As Range
        Dim d1 As Date, d2 As Date, refRqm As Range, refOrder As Range, refShipment As Range
        
        ' Najpierw wpisz daty
        Dim x As Variant
        For x = 0 To coverCols - 1
            
            ' obsolete implementation
            ' s.Cells(rrow, x + 10).Value = IsoWeekNumber(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
            
            ' using instead built-in iso week num foo directly from worksheet functions
            s.Cells(rrow, x + 10).Value = Application.WorksheetFunction.IsoWeekNum(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
            
            Set cwr = s.Cells(rrow, x + 10)
            Set refRqm = cwr.Offset(1, 0)
            Set refOrder = cwr.Offset(2, 0)
            Set refShipment = cwr.Offset(3, 0)
            
            d1 = CDate(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1, "yyyy-mm-dd"))
            d2 = CDate(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 7, "yyyy-mm-dd"))
            
            refRqm = getRqms(refRqm, d1, d2, dataFromCorail)
            refOrder = getOrdersAndPutSomeComment(refOrder, d1, d2, dataFromCorail)
            refShipment = getShipmentsAndPutSomeComment(refShipment, d1, d2, dataFromCorail)
            
            cwr.AddComment CStr(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1, "dd-mm-yyyy") + " - " + Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 7, "dd-mm-yyyy"))
            cwr.Comment.Shape.Width = 150
            cwr.Comment.Shape.Height = 30
        Next x
        ' teraz formu³y - old school comment :D
        ' R[-3]C - order - tego nie liczymy
        

        s.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-2]C+R[-1]C"
        
        s.Cells(rrow + 5, 10).NumberFormat = "0_ ;[Red]-0 "
        s.Cells(rrow + 5, 10).Font.Bold = True
        s.Cells(rrow + 5, 10).Interior.Color = RGB(183, 222, 232)
        

        
        
        For x = 1 To coverCols - 1
            ' R[-3]C - order - tego nie liczymy
            ' S.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
            s.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-2]C+R[-1]C"
            
            s.Cells(rrow + 5, 10 + x).NumberFormat = "0_ ;[Red]-0 "
            s.Cells(rrow + 5, 10 + x).Font.Bold = True
            s.Cells(rrow + 5, 10 + x).Interior.Color = RGB(183, 222, 232)
        Next x
        
        
        ' if TT = zero nothing gonna happen!
        putBordersOnTTWeek s, r
        
    End If
    

End Sub


Private Function getOrdersAndPutSomeComment(rfr As Range, d1 As Date, d2 As Date, data As SuitableData2720) As Double


    Dim titb As TableInTextBox
    Set titb = New TableInTextBox
    titb.defineLables "QTY", "XELLER", "SHIPPER", "TYPE", "ST", "EO", "EF", "RQ"
    titb.defineSizes 8, 10, 10, 10, 2, 10, 10, 10
    titb.setCols 8
    
    Dim labelAdded As Boolean
    labelAdded = False

    rfr.AddComment ""


    getOrdersAndPutSomeComment = 0
    
    
    Dim tmp As Double, ciDate As Date
    tmp = 0
    
    Dim ci As CoverageItem
    Dim ciOrdersColl As Collection
    Dim oi As OrderItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
        ' Set cmnt = Nothing
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmp = tmp + CDbl(ci.getOrder())
            
            
            Set ciOrdersColl = ci.getOrdersCollection()
            
            If Not ciOrdersColl Is Nothing Then
                If ciOrdersColl.Count > 0 Then

                    
                    For Each oi In ciOrdersColl
                    
                        If Not labelAdded Then
                            'rfr.Comment.Text "" & LAYOUT_2510_LABEL
                            titb.tryToAddLineOfTheData "QTY", "SELLER", "SHIPPER", "TYPE", "ST", "EO", "EF", "RQ"
                            labelAdded = True
                        End If
                        
                        'rfr.Comment.Text "" & rfr.Comment.Text & oi.getElement(E_2510_QTY) & "  " & _
                        '    oi.getElement(E_2510_SELLER) & " " & oi.getElement(E_2510_SHIPPER) & " " & oi.getElement(E_2510_TYPE) _
                        '    & " " & oi.getElement(E_2510_S2) & " " & _
                        '    trimToDatesOnly(oi.getElement(E_2510_DHEO)) & " " & _
                        '    trimToDatesOnly(oi.getElement(E_2510_DHRQ)) & chr(10)
                        titb.tryToAddLineOfTheData oi.getElement(E_2510_QTY), _
                            oi.getElement(E_2510_SELLER), _
                            oi.getElement(E_2510_SHIPPER), _
                            oi.getElement(E_2510_TYPE), _
                            oi.getElement(E_2510_S2), _
                            trimToDatesOnly(oi.getElement(E_2510_DHEO)), _
                            trimToDatesOnly(oi.getElement(E_2510_DHEF)), _
                            trimToDatesOnly(oi.getElement(E_2510_DHRQ))
                    Next oi
                End If
            End If
            
        End If
    Next ci
    
    
    rfr.Comment.Text titb.drawThisTable()
    
    
    If Trim(rfr.Comment.Text) <> "" Then
        rfr.Comment.Shape.TextFrame.AutoSize = True
        rfr.Comment.Shape.TextFrame.Characters.Font.Name = "Lucida Sans Typewriter"
        rfr.Comment.Shape.TextFrame.Characters.Font.Size = 8
    Else
        rfr.Comment.Delete
    End If
    
    
    
    
    
    getOrdersAndPutSomeComment = tmp
    
End Function

Private Function getShipmentsAndPutSomeComment(rfr As Range, d1 As Date, d2 As Date, data As SuitableData2720)

    Dim titb As TableInTextBox
    Set titb = New TableInTextBox
    titb.defineLables "QTY", "XELLER", "SHIPPER", "TYPE", "ST", "EO", "EF", "RQ"
    titb.defineSizes 8, 10, 10, 10, 2, 10, 10, 10
    titb.setCols 8
    
    Dim labelAdded As Boolean
    labelAdded = False

    rfr.AddComment ""

    getShipmentsAndPutSomeComment = 0
    
    
    Dim tmp As Double, ciDate As Date, ciOrdersColl As New Collection, oi As OrderItem, theStatus As String
    tmp = 0
    
    Dim ci As CoverageItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
        ' Set cmnt = Nothing ' ?????
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmp = tmp + CDbl(ci.getShip())
            
            Set ciOrdersColl = ci.getOrdersCollection()
            
            If Not ciOrdersColl Is Nothing Then
                If ciOrdersColl.Count > 0 Then
                    
                    For Each oi In ciOrdersColl
                        
                        ' now this is shipment so only EF or more
                        theStatus = oi.getElement(E_2510_S2)
                        
                        
                        ' An ADDED FROM 0.16
                        If theStatus <> "EO" And _
                            theStatus <> "CA" And _
                                theStatus <> "FO" And _
                                    theStatus <> "AN" Then
                        
                        
                            If Not labelAdded Then
                                ' rfr.Comment.Text "" & LAYOUT_2510_LABEL
                                titb.tryToAddLineOfTheData "QTY", "SELLER", "SHIPPER", "TYPE", "ST", "EO", "EF", "RQ"
                                labelAdded = True
                            End If
                            
                            
                            
                        
                            'rfr.Comment.Text "" & rfr.Comment.Text & oi.getElement(E_2510_QTY) & "  " & _
                            '    oi.getElement(E_2510_SELLER) & "  " & oi.getElement(E_2510_SHIPPER) & "  " & oi.getElement(E_2510_TYPE) _
                            '    & " " & oi.getElement(E_2510_S2) & "  " & _
                            '    trimToDatesOnly(oi.getElement(E_2510_DHEO)) & "  " & _
                            '    trimToDatesOnly(oi.getElement(E_2510_DHRQ)) & chr(10)
                            ''''''''''''''''''''''
                            titb.tryToAddLineOfTheData oi.getElement(E_2510_QTY), _
                                oi.getElement(E_2510_SELLER), _
                                oi.getElement(E_2510_SHIPPER), _
                                oi.getElement(E_2510_TYPE), _
                                oi.getElement(E_2510_S2), _
                                trimToDatesOnly(oi.getElement(E_2510_DHEO)), _
                                trimToDatesOnly(oi.getElement(E_2510_DHEF)), _
                                trimToDatesOnly(oi.getElement(E_2510_DHRQ))
                                
                            
                            
                            
                            
                        End If
                    Next oi
                    
                End If
            End If
        End If
    Next ci
    
    rfr.Comment.Text titb.drawThisTable()
    
    
    If Trim(rfr.Comment.Text) <> "" Then
        rfr.Comment.Shape.TextFrame.AutoSize = True
        rfr.Comment.Shape.TextFrame.Characters.Font.Name = "Lucida Sans Typewriter"
        rfr.Comment.Shape.TextFrame.Characters.Font.Size = 8
        rfr.Interior.Color = RGB(255, 255, 160)
    Else
        rfr.Comment.Delete
        rfr.ClearFormats
    End If
    
    
    getShipmentsAndPutSomeComment = tmp
    
End Function

Private Function trimToDatesOnly(strDT As String) As String
    trimToDatesOnly = "" & Split(strDT, " ")(0)
End Function



Private Function getRqms(rfr As Range, d1 As Date, d2 As Date, data As SuitableData2720) As Double

    rfr.AddComment ""
    
    Dim tmpRqm As Double, ciDate As Date
    tmpRqm = 0
    
    Dim ci As CoverageItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmpRqm = tmpRqm + CDbl(ci.getRqms())
            
            rfr.Comment.Text "" & rfr.Comment.Text & ci.getXtraRqmsStr()
        End If
    Next ci
    
    
    If Trim(rfr.Comment.Text) <> "" Then
        rfr.Comment.Shape.TextFrame.AutoSize = True
    Else
        rfr.Comment.Delete
    End If
    
    getRqms = tmpRqm
End Function

Private Sub ILayout_setHowManyWeeks(mhmw As Integer)
    coverCols = Int(mhmw)
End Sub

Private Sub ILayout_setLayout(typeOfLayout As E_TYPE_OF_LAYOUT)
    layoutType = typeOfLayout
End Sub


Private Function getAreaBasedOn(r As Range) As Range
    Set getAreaBasedOn = r
End Function


Private Sub putBordersOnTTWeek(s As Worksheet, rrow As Range)
    
    
    ' tutaj rrow troche przeklamana nazwa , ale nie chce mi sie juz zmieniac
    
    ' putBordersOnTTWeek
    
    If Not dataFromCorail Is Nothing Then
        Dim howManyWeeksForTT As Integer
        howManyWeeksForTT = -1
        On Error Resume Next
        howManyWeeksForTT = Int(dataFromCorail.tt / 7#)
        
        If howManyWeeksForTT > 0 Then
            
            Dim properPlace As Range
            Set properPlace = s.Range(s.Cells(rrow.Row, 10 - 1 + howManyWeeksForTT), s.Cells(rrow.Row + 5, 10 - 1 + howManyWeeksForTT))
            ' properPlace.Interior.Color = RGB(255, 0, 0)
            
            properPlace.Borders(xlDiagonalDown).LineStyle = xlNone
            properPlace.Borders(xlDiagonalUp).LineStyle = xlNone
            With properPlace.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With properPlace.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With properPlace.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With properPlace.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            properPlace.Borders(xlInsideVertical).LineStyle = xlNone
            properPlace.Borders(xlInsideHorizontal).LineStyle = xlNone
        End If
    End If
    
End Sub

Private Sub pFormatCommonDataOnLeft(s As Worksheet, rrow As Long, hmc As Integer)


    Dim tmp As String

    s.Cells(rrow, 1).Value = "PLT"
    s.Cells(rrow, 9).Value = "Past due"
    s.Cells(rrow, 7).Value = "First runout"
    '========================================================

    With s.Range(s.Cells(rrow + 1, 1), s.Cells(rrow + 5, 1))
        .Interior.Color = RGB(140, 140, 140)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Merge
        .Orientation = 90
        .Value = dataFromCorail.plt 'working
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
    End With
    
    ' color req
    With s.Range(s.Cells(rrow + 1, 10), s.Cells(rrow + 1, hmc + 9))
        .Interior.Color = RGB(242, 242, 242)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
   
    '========================================================
    
    With s.Range(s.Cells(rrow, 1), s.Cells(rrow, hmc + 9))
        .Interior.Color = RGB(140, 140, 140)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    tmp = dataFromCorail.supplierName
    tmp = Mid(tmp, 8)
    
    With s.Cells(rrow + 1, 2)
        ' .Value = dataFromCorail.plt ' zatem tutaj wolny slot teraz
        .Value = tmp ' dataFromCorail.supplierName ' wrzucilem tymczasowo - mozna zmienic! :)  '''' OK can stay ''''
        .Interior.Color = RGB(183, 222, 232)
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    With s.Cells(rrow + 2, 2)
        .Value = dataFromCorail.pn
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    
    
    s.Cells(rrow + 3, 2).Value = dataFromCorail.approId
    With s.Cells(rrow + 4, 2)
        .Value = dataFromCorail.approName
        .Font.Size = 9
    End With
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' przykladowo wybralem vendor cofor - mozna teraz wybrac rowniez shipper oraz manufacturer / zmienione na shipper
    ' ====> changed!!!
    'S.Cells(rrow + 5, 2).Value = dataFromCorail.shipperCofor ==> put in comment in favour of Lns 388-399
     With s.Cells(rrow + 5, 2)
        .Value = dataFromCorail.shipperCofor
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = True
    End With
    
    With s.Cells(rrow, 2)
        .Value = "VENDOR / SELLER:     " & dataFromCorail.vendorCofor
        .Interior.Color = RGB(140, 140, 140)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    

    ' tutaj specjalne miesjce pod ruchomy first runout
    ' i chyba powinno to byc opcjonalne :D
    ' ========================================================
    s.Cells(rrow + 1, 7).FormulaR1C1 = "=calcFirstRunOut(R[4]C[3]:R[4]C[22])"
    s.Cells(rrow + 1, 7).Font.Bold = True
    ' ========================================================
    On Error Resume Next
    s.Cells(rrow + 2, 7).Value = "tt:" & dataFromCorail.tt & " days"
    s.Cells(rrow + 2, 7).Font.Bold = True
    ' ========================================================
    
    
    s.Cells(rrow + 1, 8).Value = "RQM"
    s.Cells(rrow + 2, 8).Value = "Ordered"
    s.Cells(rrow + 3, 8).Value = "Shipped"
    s.Cells(rrow + 4, 8).Value = "Manual"
    
    With s.Cells(rrow + 5, 8)
        .Value = "Coverage"
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(183, 222, 232)
    End With
    
    With s.Range(s.Cells(rrow + 1, 8), s.Cells(rrow + 5, 7))
        .Interior.Color = RGB(140, 140, 140)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    
    s.Cells(rrow + 1, 3).Value = "  " '"SLOT"
    s.Cells(rrow + 2, 3).Value = "  " '"SLOT"
    s.Cells(rrow + 3, 3).Value = "  " '"SLOT"
    s.Cells(rrow + 4, 3).Value = "  " '"SLOT"
    s.Cells(rrow + 5, 3).Value = "  " '"SLOT"
    
    With s.Range(s.Cells(rrow + 1, 3), s.Cells(rrow + 5, 3))
        .Interior.Color = RGB(140, 140, 140)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    
    s.Cells(rrow + 1, 5).Value = "BAL"
    s.Cells(rrow + 1, 6).FormulaR1C1 = "=R[1]C+R[2]C"
    s.Cells(rrow + 2, 5).Value = "STOCK"
    s.Cells(rrow + 2, 6).Value = dataFromCorail.stock
    s.Cells(rrow + 3, 5).Value = "RECV"
    s.Cells(rrow + 3, 6).Value = dataFromCorail.recv
    s.Cells(rrow + 4, 5).Value = "Hazards from FJU"
    s.Cells(rrow + 4, 6).Value = dataFromCorail.hazardsFromFJU
    s.Cells(rrow + 5, 5).Value = "Blockages in progress"
    s.Cells(rrow + 5, 6).Value = dataFromCorail.blockagesInProgress
    
   

End Sub
