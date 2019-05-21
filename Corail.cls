VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Corail"
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


Implements ICorail


Private scr2720 As ICorailScreen
Private scr2510 As ICorailScreen
Private httpReqRef As HTTPRequestHandler

Private theCorailData As ICorailData



Private Sub Class_Initialize()
    Set scr2720 = New Corail_2720_Screen
    Set scr2510 = New Corail_2510_Screen
    Set httpReqRef = Nothing
    Set theCorailData = New CorailData
End Sub

Private Sub Class_Terminate()
    Set scr2720 = Nothing
    Set scr2510 = Nothing
    
    Set httpReqRef = Nothing
    Set theCorailData = Nothing
End Sub

Private Function ICorail_connectWithHttpReq(httpReq As HTTPRequestHandler) As Boolean

    ' =========================================
    Set httpReqRef = httpReq
    ' =========================================
    
End Function

Private Function ICorail_connectWithScreen(numberOfScreen As Variant) As Boolean
    If numerOfScreen = 2510 Or numberOfScreen = 2720 Then
        ICorail_connectWithScreen = True
    Else
        ICorail_connectWithScreen = False
    End If
End Function

Private Function ICorail_getDataFromScreen(numberOfScreen As Variant, pltCode As Variant, partNumber As Variant) As ICorailData

    If numberOfScreen = 2720 Then
        scr2720.screenSetup numberOfScreen, httpReqRef
        Set ICorail_getDataFromScreen = scr2720.getData(pltCode, partNumber)
    ElseIf numberOfScreen = 2510 Then
        scr2510.screenSetup numberOfScreen, httpReqRef
        Set ICorail_getDataFromScreen = scr2510.getData(pltCode, partNumber)
    End If
End Function
