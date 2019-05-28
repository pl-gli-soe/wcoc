VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Corail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' __        __        _    _          ____
' \ \      / /__  ___| | _| |_   _   / ___|_____   _____ _ __ __ _  __ _  ___
'  \ \ /\ / / _ \/ _ \ |/ / | | | | | |   / _ \ \ / / _ \ '__/ _` |/ _` |/ _ \
'   \ V  V /  __/  __/   <| | |_| | | |__| (_) \ V /  __/ | | (_| | (_| |  __/
'    \_/\_/ \___|\___|_|\_\_|\__, |  \____\___/ \_/ \___|_|  \__,_|\__, |\___|
'   ___  _ __    / ___|___  _|___/_ _(_) |                         |___/
'  / _ \| '_ \  | |   / _ \| '__/ _` | | |
' | (_) | | | | | |__| (_) | | | (_| | | |
'  \___/|_| |_|  \____\___/|_|  \__,_|_|_|
'
'
'01010111 01100101 01100101 01101011 01101100 01111001  01000011 01101111 01110110 01100101 01110010 01100001 01100111 01100101
'01101111 01101110  01000011 01101111 01110010 01100001 01101001 01101100
'
'
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
    Set theCorailData = Nothing
End Sub

Private Sub Class_Terminate()
    Set scr2720 = Nothing
    Set scr2510 = Nothing
    
    Set httpReqRef = Nothing
    Set theCorailData = Nothing
End Sub

Private Function ICorail_connectWithHttpReq(httpReq As HTTPRequestHandler) As Boolean

    ' really simple setter for private variable
    ' ===========================================
    Set httpReqRef = httpReq
    ' ===========================================
    
End Function

Private Function ICorail_connectWithScreen(numberOfScreen As Variant) As Boolean


    ' for now really simple logic, which checking and handling only 2 transactions from Corail
    ' -------------------------------------------------------------------------------------------
    If numerOfScreen = 2510 Or numberOfScreen = 2720 Then
        ICorail_connectWithScreen = True
    Else
        ICorail_connectWithScreen = False
    End If
    ' -------------------------------------------------------------------------------------------
    
    
End Function

Private Function ICorail_getDataFromScreen(numberOfScreen As Variant, pltCode As Variant, partNumber As Variant) As ICorailData


    ' 90% of the implementation of this class
    ' ----------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------

    If numberOfScreen = 2720 Then
        scr2720.screenSetup numberOfScreen, httpReqRef
        Set ICorail_getDataFromScreen = scr2720.getData(pltCode, partNumber) ' As ICorailData
    ElseIf numberOfScreen = 2510 Then
        scr2510.screenSetup numberOfScreen, httpReqRef
        Set ICorail_getDataFromScreen = scr2510.getData(pltCode, partNumber, E_TYPE_OF_DATA_FROM_CORAIL_DEFAULT) ' As ICorailData
    End If
    
    ' ----------------------------------------------------------------------------------
    ' ----------------------------------------------------------------------------------
End Function
