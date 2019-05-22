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
Private s As Worksheet


Private Sub Class_Initialize()

    ' layoutType
    ''
    '
    layoutType = E_TYPE_OF_LAYOUT_WEEKLY_UNDEFINED
    '
    ''

End Sub

Private Sub ILayout_putSomeColorsForThisPartNumber(r As Range)


    Set area = getAreaBasedOn(r)
    Set s = area.Parent
    rrow = r.Row
    
    coverCols = 20
    
    s.Cells(rrow, 1).Value = "PLT"
    s.Cells(rrow, 9).Value = "Past due"
    s.Cells(rrow, 7).Value = "First runout"
    ' tutaj specjalne miesjce pod ruchomy first runout
    ' i chyba powinno to byc opcjonalne :D
    ' ========================================================
    s.Cells(rrow, 7).Value = "First runout"
    s.Cells(rrow + 1, 7).FormulaR1C1 = "=calcFirstRunOut(R[4]C[3]:R[4]C[22])"
    s.Cells(rrow + 1, 7).Font.Bold = True
    ' ========================================================
    
    
    ' obsolete
    ' s.Cells(rrow, 8).Value = Component
    
    '    s.Cells(rrow + 1, 1).Value = inPLT
    '    s.Cells(rrow + 1, 2).Value = inPRT
    '    s.Cells(rrow + 4, 2).Value = Comment
    '    s.Cells(rrow + 4, 2).Interior.Color = kolor
    
    s.Cells(rrow + 1, 8).Value = "req"
    s.Cells(rrow + 2, 8).Value = "Vessel"
    s.Cells(rrow + 3, 8).Value = "AIR"
    s.Cells(rrow + 4, 8).Value = "manual"
    s.Cells(rrow + 5, 8).Value = "coverage"
    
    s.Cells(rrow + 1, 3).Value = "BANK"
    s.Cells(rrow + 2, 3).Value = "MODE"
    s.Cells(rrow + 3, 3).Value = "PCS TO GO"
    s.Cells(rrow + 4, 5).Value = "TTIME"
    s.Cells(rrow + 5, 3).Value = "MISC"
    
    s.Cells(rrow + 1, 5).Value = "CBAL"
    s.Cells(rrow + 2, 5).Value = "DOH"
    s.Cells(rrow + 3, 5).Value = "QHD"
    s.Cells(rrow + 4, 3).Value = "STD PACK"
    s.Cells(rrow + 5, 5).Value = "BACKLOG"
    
    

    Dim cwr As Range
    ' Najpierw wpisz daty
    For x = 0 To coverCols - 1
        
        ' obsolete implementation
        ' s.Cells(rrow, x + 10).Value = IsoWeekNumber(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
        
        ' using instead built-in iso week num foo directly from worksheet functions
        s.Cells(rrow, x + 10).Value = Application.WorksheetFunction.IsoWeekNum(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1)
        
        Set cwr = s.Cells(rrow, x + 10)
        cwr.AddComment CStr(Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 1, "dd-mm-yyyy") + " - " + Format(Round(GlobalFooModule.ostatniaNiedziela(Date)) + x * 7 + 7, "dd-mm-yyyy"))
        cwr.Comment.Shape.Width = 150
        cwr.Comment.Shape.Height = 30
    Next x
    ' teraz formu³y
    s.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
    
    If Comment = "ALC" Then
        s.Cells(rrow + 5, 10).FormulaR1C1 = "=R[-4]C[-4]+R[-3]C[-1]+R[-2]C[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C+RC[-6]"
    End If
    
    For x = 1 To coverCols - 1
        s.Cells(rrow + 5, 10 + x).FormulaR1C1 = "=RC[-1]-R[-4]C+R[-3]C+R[-2]C+R[-1]C"
    Next x
    

End Sub

Private Sub ILayout_setLayout(typeOfLayout As E_TYPE_OF_LAYOUT)

End Sub


Private Function getAreaBasedOn(r As Range) As Range

    
End Function
