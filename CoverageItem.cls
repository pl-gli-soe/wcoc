VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CoverageItem"
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
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.



' for suitable data (CovaregeItem: Me - is component for SuitableData2720 object) for excel sheet from 2720 Corail Screen

Private d As Date
Private pcDemand As Double
Private incBc As Double
Private ordered As Double
Private shipped As Double


Private ready As Boolean

Public Sub makeReady()
    ready = True
End Sub

Public Function isReady() As Boolean
    isReady = ready
End Function

Public Sub constructor(md As Date, mPcDemand As Double, mIncBc As Double, mo As Double, msh As Double)

    d = CDate(md)
    pcDemand = CDbl(mPcDemand)
    incBc = CDbl(mIncBc)
    ordered = CDbl(mo)
    shipped = CDbl(msh)
End Sub



Public Function getDate() As Date
    getDate = d
End Function

Public Function getStrDate() As String
    tmp = ""
    yyyy = "" & Year(d)
    mm = "" & Month(d)
    
    If Len(mm) = 1 Then
        mm = "0" & m
    End If
    
    dd = "" & Day(d)
    
    If Len(dd) = 1 Then
        dd = "0" & dd
    End If
    
    
    getStrDate = "" & yyyy & "-" & mm & "-" & dd
End Function

Public Function getRqms() As Double
    getRqms = pcDemand + incBc
End Function

Public Function getOrder() As Double
    getOrder = ordered
End Function

Public Function getShip() As Double
    getShip = shipped
End Function

Private Sub Class_Initialize()
    pcDemand = 0
    incBc = 0
    ordered = 0
    shipped = 0
    ready = False
End Sub

