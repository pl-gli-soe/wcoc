VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuitableData6510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The MIT License (MIT)
'
'Copyright (c) 2020 FORREST
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



' SUITABLE DATA FROM 6510



Private tt As TTItem
Private tts As Collection

Public plt As String
Public sellerCofor As String
Public shipperCofor As String

Public Function getDominantItem() As TTItem
    
    Set getDominantItem = Nothing
    
    Dim innerttitem As TTItem, counterForWinner As Integer, k As String
    Dim slownikDoPodliczania As New Dictionary
    For Each innerttitem In tts
    
    
        
        k = CStr(CLng(innerttitem.getTT()))
    
        If Not slownikDoPodliczania.Exists(CStr(CLng(innerttitem.getTT()))) Then
            
            slownikDoPodliczania.Add k, 1
        Else
            slownikDoPodliczania(k) = slownikDoPodliczania(k) + 1
        End If
        
    Next innerttitem
    
    counterForWinner = 0
    Dim winner As TTItem
    For Each innerttitem In tts
        k = CStr(CLng(innerttitem.getTT()))
        
        If CLng(counterForWinner) < slownikDoPodliczania(k) Then
            Set winner = innerttitem
            counterForWinner = Int(slownikDoPodliczania(k))
        End If
    Next
    
    If Not winner Is Nothing Then
        Set getDominantItem = winner
    Else
        Set getDominantItem = New TTItem
    End If
    
    
End Function


Public Function getTTs() As Collection
    Set getTTs = tts
End Function


Public Sub addTT(arg As TTItem)
    tts.Add arg
End Sub



Private Sub Class_Initialize()
    Set tts = New Collection
End Sub

Private Sub Class_Terminate()
    Set tts = Nothing
End Sub

