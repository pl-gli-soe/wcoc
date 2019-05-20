Attribute VB_Name = "GlobalFooModule"
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



Function ostatniaNiedziela(Data)
    ostatniaNiedziela = Data - Weekday(Data) + 1
End Function


Public Function calcFirstRunOut(r As Range)


    rok = Int(Year(Date)) * 100
    calcFirstRunOut = ""

    Dim rr As Range
    Set rr = r.Offset(-5, 0)
    If rr.Item(1) < rr.Item(r.Count) Then
        For Each i In r
            If i < 0 Then
                calcFirstRunOut = rok + i.Offset(-5, 0)
                Exit Function
            End If
        Next i
    Else
        ' tutaj dodatkowo dochodzi opcja ze mamy przejscie przez nowy rok i mamy zalamanie ciaglosci danych jesli
        ' chodzi tylko i wylacznie o czysty CW
        ' zatem musi algorytm w szybki i prosty sposob umiec to rozpoznac
        For Each i In r
            If i < 0 Then
                If i.Offset(-5, 0) >= rr.Item(1) Then
                    calcFirstRunOut = rok + i.Offset(-5, 0)
                    Exit Function
                Else
                    rok = rok + 100
                    calcFirstRunOut = rok + i.Offset(-5, 0)
                    Exit Function
                End If
            End If
        Next i
    End If

    If rr.Item(1) < rr.Item(r.Count) Then
        calcFirstRunOut = rok + r.Item(r.Count).Offset(-5, 0)
    Else
        rok = rok + 100
        calcFirstRunOut = rok + r.Item(r.Count).Offset(-5, 0)
    End If
    
    
End Function
