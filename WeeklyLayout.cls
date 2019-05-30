VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeeklyLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private S As Worksheet


Private dataFromCorail As SuitableData2720


Private Sub Class_Initialize()

    ' layoutType
    ''
    '
    layoutType = E_TYPE_OF_LAYOUT_WEEKLY_UNDEFINED
    '
    ''
    Set dataFromCorail = Nothing

End Sub

Private Sub ILayout_nowMergeWithDataFromCorail(r As Range, main2720data As SuitableData2720)
    Set dataFromCorail = main2720data
End Sub

Private Sub ILayout_putSomeColorsForThisPartNumber(r As Range)


    If Not dataFromCorail Is Nothing Then


        Set area = getAreaBasedOn(r)
        Set S = area.Parent
        rrow = r.Row
        
        coverCols = 20
        
        S.Cells(rrow, 1).Value = "PLT"
        S.Cells(rrow + 1, 2).Value = dataFromCorail.plt
        S.Cells(rrow + 2, 2).Value = dataFromCorail.pn
        S.Cells(rrow + 3, 2).Value = dataFromCorail.approId
        S.Cells(rrow + 4, 2).Value = dataFromCorail.approName
        ' S.Cells(rrow + 5, 2).Value = dataFromCorail.xxx
        
        S.Cells(rrow, 9).Value = "Past due"
        S.Cells(rrow, 7).Value = "First runout"
        ' tutaj specjalne miesjce pod ruchomy first runout
        ' i chyba powinno to byc opcjonalne :D
        ' ========================================================
        S.Cells(rrow, 7).Value = "First runout"
        S.Cells(rrow + 1, 7).FormulaR1C1 = "=calcFirstRunOut(R[4]C[3]:R[4]C[22])"
        S.Cells(rrow + 1, 7).Font.Bold = True
        ' ========================================================
        
        
        ' obsolete
        ' s.Cells(rrow, 8).Value = Component
        
        '    s.Cells(rrow + 1, 1).Value = inPLT
        '    s.Cells(rrow + 1, 2).Value = inPRT
        '    s.Cells(rrow + 4, 2).Value = Comment
        '    s.Cells(rrow + 4, 2).Interior.Color = kolor
        
        S.Cells(rrow + 1, 8).Value = "RQM"
        S.Cells(rrow + 2, 8).Value = "Ordered"
        S.Cells(rrow + 3, 8).Value = "Shipped"
        S.Cells(rrow + 4, 8).Value = "Manual"
        S.Cells(rrow + 5, 8).Value = "Coverage"
        
        S.Cells(rrow + 1, 3).Value = "BANK"
        S.Cells(rrow + 2, 3).Value = "MODE"
        S.Cells(rrow + 3, 3).Value = "PCS TO GO"
        S.Cells(rrow + 4, 5).Value = "TTIME"
        S.Cells(rrow + 5, 3).Value = "MISC"
        
        S.Cells(rrow + 1, 5).Value = "BAL"
        S.Cells(rrow + 1, 6).FormulaR1C1 = "=R[1]C-R[2]C"
        S.Cells(rrow + 2, 5).Value = "STOCK"
        S.Cells(rrow + 2, 6).Value = dataFromCorail.stock
        S.Cells(rrow + 3, 5).Value = "RECV"
        S.Cells(rrow + 3, 6).Value = dataFromCorail.recv
        
        
    
        Dim cwr As Range
        Dim d1 As Date, d2 As Date, refRqm As Range, refOrder As Range, refShipment As Range
        ' Najpierw wpisz daty
        For x = 0 To coverCols - 1
            
            ' obsolete implementation
            ' s.Cells(rrow, x + 10).Value = IsoWeekNumber(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
            
            ' using instead built-in iso week num foo directly from worksheet functions
            S.Cells(rrow, x + 10).Value = Application.WorksheetFunction.IsoWeekNum(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
            
            Set cwr = S.Cells(rrow, x + 10)
            Set refRqm = cwr.Offset(1, 0)
            Set refOrder = cwr.Offset(2, 0)
            Set refShipment = cwr.Offset(3, 0)
            
            d1 = CDate(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1, "yyyy-mm-dd"))
            d2 = CDate(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 7, "yyyy-mm-dd"))
            
            refRqm = getRqms(d1, d2, dataFromCorail)
            refOrder = getOrdersAndPutSomeComment(refOrder, d1, d2, dataFromCorail)
            refShipment = getShipmentsAndPutSomeComment(refShipment, d1, d2, dataFromCorail)
            
            cwr.AddComment CStr(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1, "dd-mm-yyyy") + " - " + Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 7, "dd-mm-yyyy"))
            cwr.Comment.Shape.Width = 150
            cwr.Comment.Shape.Height = 30
        Next x
        ' teraz formu³y - old school comment :D
        ' R[-3]C - order - tego nie liczymy
        S.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-2]C+R[-1]C"
        
        If Comment = "ALC" Then
            ' R[-3]C - order - tego nie liczymy
            S.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-2]C+R[-1]C+RC[-6]"
        End If
        
        For x = 1 To coverCols - 1
            ' R[-3]C - order - tego nie liczymy
            ' S.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
            S.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-2]C+R[-1]C"
        Next x
    End If
    

End Sub


Private Function getOrdersAndPutSomeComment(rfr As Range, d1 As Date, d2 As Date, data As SuitableData2720) As Double


    ' rfr: reference as range - to put some comment
    rfr.AddComment ""

    getOrdersAndPutSomeComment = 0
    
    
    Dim tmp As Double, ciDate As Date
    tmp = 0
    
    Dim ci As CoverageItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmp = tmp + CDbl(ci.getOrder())
        End If
    Next ci
    
    
    getOrdersAndPutSomeComment = tmp
    
End Function

Private Function getShipmentsAndPutSomeComment(rfr As Range, d1 As Date, d2 As Date, data As SuitableData2720)


    ' rfr: reference as range - to put some comment
    rfr.AddComment ""

    getShipmentAndPutSomeComment = 0
    
    
    Dim tmp As Double, ciDate As Date
    tmp = 0
    
    Dim ci As CoverageItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmp = tmp + CDbl(ci.getShip())
            
        End If
    Next ci
    
    
    getShipmentAndPutSomeComment = tmp
    
End Function


Private Function getRqms(d1 As Date, d2 As Date, data As SuitableData2720) As Double

    ' co prawda 3 argument jest zbedny - ale niech zostanie jawnie - ku pamieci :)
    
    Dim tmpRqm As Double, ciDate As Date
    tmpRqm = 0
    
    Dim ci As CoverageItem
    For Each ci In data.getCollectionOfItems()
        
        ciDate = ci.getDate()
    
        If CDate(d1) <= CDate(ciDate) And CDate(d2) >= CDate(ciDate) Then
            tmpRqm = tmpRqm + CDbl(ci.getRqms())
        End If
    Next ci
    
    getRqms = tmpRqm
End Function

Private Sub ILayout_setLayout(typeOfLayout As E_TYPE_OF_LAYOUT)
    layoutType = typeOfLayout
End Sub


Private Function getAreaBasedOn(r As Range) As Range

    Set getAreaBasedOn = r
End Function
