VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Corail_6510_Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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


'page=1&max=45&charset=UTF-8
'&ecwAutoId=false&ecwDivId=insureScheduleCoherence
'&cu=EN&seller=29763M  01&shipper=A009CF  01&type=Departure&milestoneFilterName=RQ

' http://ei.control.erp.corail.inetpsa.com/getInsureScheduleCoherencePager.do


Implements ICorailScreen


Private preUrl As String
Private screenUrl As String
Private url As String
Private num As Integer
Private strona As Integer
Private ileLiniiNaStronie As Integer


Private req As HTTPRequestHandler

Private theCorailData As ICorailData


Private valid As Validator


Private Sub Class_Initialize()


    num = 6510
    
    
    
    
    strona = 1
    ileLiniiNaStronie = 40
    
    screenUrl = "/getInsureScheduleCoherencePager.do"
            
    preUrl = ".control.erp.corail.inetpsa.com"
    url = ""

    Set req = Nothing
    Set theCorailData = Nothing
    
    Set valid = New Validator
End Sub

Private Sub Class_Terminate()

    preUrl = ""
    screenUrl = ""
    url = ""

    Set req = Nothing
    Set theCorailData = Nothing
    
    Set valid = Nothing
End Sub

Private Function ICorailScreen_getData(plt As Variant, partNumber As Variant, Optional e As E_TYPE_OF_DATA_FROM_CORAIL) As ICorailData
    Set ICorailScreen_getData = Nothing
End Function

Private Function ICorailScreen_getDataFrom6510(plt As Variant, vc1 As String, sc1 As String) As ICorailData

    
    ' just in case
    Set ICorailScreen_getDataFrom6510 = Nothing
    
    
    
    Dim answer As String, args As String
    Dim ktoraStrona As Integer
    
    answer = ""
    ktoraStrona = 1
    
    
    Set theCorailData = New CorailDataFrom6510
    
    ' we taking only first page!

    url = "http://" & CStr(plt) & preUrl & screenUrl
    args = makeSomeArgsForMe(vc1, sc1)
    
    Debug.Print "http post: " & url & " " & args
    answer = req.httpPost(url, args)
    
    ' but for this example this set is actually add!!!
    theCorailData.setString answer
    ' Debug.Print answer
    
    ' -----------------------------------------------------------------------------


    ' now parse all
    theCorailData.parseString
    Set ICorailScreen_getDataFrom6510 = theCorailData
    
End Function



Private Function makeSomeArgsForMe(param1 As String, param2 As String) As String


    ' param1 - seller cofor
    ' param2 - shipper cofor
    

    
    Dim tmp As String

    
    tmp = "page=1" & _
        "&max=45" & _
        "&charset=UTF-8" & _
        "&ecwAutoId=false&ecwDivId=insureScheduleCoherence" & _
        "&cu=EN" & _
        "&seller=" & CStr(param1) & _
        "&shipper=" & CStr(param2) & _
        "&type=Departure" & _
        "&milestoneFilterName=RQ"
        
    
    makeSomeArgsForMe = CStr(tmp)
    
End Function

Function parseFromDate(i As Integer)

    Dim yyyy As String, mm As String, dd As String

    
    ' i - stands for start (1) or end (2)
    'yyyy = "" & Year(Date)
    'mm = "" & Month(Date)
    'dd = "" & Day(Date)
    ' parseFromDate = "" & yyyy & "-" & mm & "-" & dd & "T00:00"
    
    ' template
    ' &startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z
    
    
    Dim d1 As Date, d2 As Date
    d1 = Date - 30
    d2 = Date + 140
    
    If i = 1 Then
        
        yyyy = "" & Year(d1)
        
        mm = "" & Month(d1)
        If Len(mm) = 1 Then
            mm = "0" & mm
        End If
        
        dd = "" & Day(d1)
        If Len(dd) = 1 Then
            dd = "0" & dd
        End If
        
        
        parseFromDate = "&startDate=" & yyyy & "-" & mm & "-" & dd & "T22:00:00Z"
        
        
    ElseIf i = 2 Then
    
    
        yyyy = "" & Year(d2)
        
        mm = "" & Month(d2)
        If Len(mm) = 1 Then
            mm = "0" & mm
        End If
        
        dd = "" & Day(d2)
        If Len(dd) = 1 Then
            dd = "0" & dd
        End If
        
        
        parseFromDate = "&endDate=" & yyyy & "-" & mm & "-" & dd & "T22:00:00Z"
        
    End If
    
    
    
    
End Function

Private Function ICorailScreen_openData(plt As Variant, partNumber As Variant) As Boolean
    ICorailScreen_openData = False
End Function

Private Sub ICorailScreen_screenSetup(number As Variant, httpReq As HTTPRequestHandler)

    Set req = httpReq
    
End Sub


