VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Corail_2510_Screen"
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


' getProductSearchList.html
' POST /getOrderViewPager.do?page=1&max=256&charset=UTF-8&location=&productCode=9465048780
' &startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z
' &eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN
' &logisticObjectTypeCode=CALCULATED
' &forecast=false HTTP/1.1
' Host: sx.Control.erp.Corail.inetpsa.com

' http://sx.control.erp.corail.inetpsa.com/getOrderViewPager.do


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


    num = 2510
    
    strona = 1
    ileLiniiNaStronie = 256
    
    screenUrl = "/getOrderViewPager.do"
            
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

    
    ' just in case
    Set ICorailScreen_getData = Nothing
    
    
    
    Dim answer As String
    answer = ""
    
    



    If e = E_TYPE_OF_DATA_FROM_CORAIL_2510_CALC Then
    
    
        ' CALCULATED
        ' -----------------------------------------------------------------------------
    
    
        ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
        url = "http://" & CStr(plt) & preUrl & screenUrl
        args = makeSomeArgsForMe("CALCULATED", partNumber)
        answer = req.httpPost(url, args)
        Set theCorailData = New CorailDataFrom2510
        theCorailData.setString answer
        theCorailData.parseString
        Set ICorailScreen_getData = theCorailData
        ' -----------------------------------------------------------------------------
    
    ElseIf e = E_TYPE_OF_DATA_FROM_CORAIL_2510_UPDATED Then
    
        ' UPDATED
        ' -----------------------------------------------------------------------------
        
        ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
        url = "http://" & CStr(plt) & preUrl & screenUrl
        args = makeSomeArgsForMe("UPDATED", partNumber)
        answer = req.httpPost(url, args)
        Set theCorailData = New CorailDataFrom2510
        theCorailData.setString answer
        theCorailData.parseString
        Set ICorailScreen_getData = theCorailData
        ' -----------------------------------------------------------------------------
        
    Else
        ' sth is wrong
        'Debug.Print "This type of data for this screen is not available: ICorailScreen_getData 2510 class: Corail_2510_Screen line: 94"
        'Set ICorailScreen_getData = Nothing
        
        ' bez rozroznienia jesli chodzi czy updated, czy calculated
        
        ' UPDATED
        ' -----------------------------------------------------------------------------
        
        ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
        url = "http://" & CStr(plt) & preUrl & screenUrl
        args = makeSomeArgsForMe("", partNumber)
        answer = req.httpPost(url, args)
        Set theCorailData = New CorailData
        theCorailData.setString answer
        theCorailData.parseString
        Set ICorailScreen_getData = theCorailData
        ' -----------------------------------------------------------------------------
    End If


    
    ' here for later usage for sum validation:
    ''
    ''
    '
    '
    'If valid.checkAnswer(num, answer) Then
    '    Set theCorailData = New CorailData
    '    theCorailData.setString answer
    '    theCorailData.parseString
    'Else
    '    Set theCorailData = Nothing
    'End If
    ''
    ''
    '
    '
    
End Function



Private Function makeSomeArgsForMe(arg, partNumber) As String
    
    ' arg - updated or calculated - but at the end will have same url from contruction perspective
    ' becuae all rest in baiscally the same dates and other params
    
    
    ' template
    'args = "page=1&max=256&charset=UTF-8&ecwAutoId=false&ecwDivId=orderView&statusList=EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:ST" & _
    '    "&eventType=RQ&forecast=false&logisticObjectTypeCode=CALCULATED&displayLocal=false&productCode=9810867580"
    
    
    thisStartDate = parseFromDate(1)
    thisEndDate = parseFromDate(2)
    
    ' Debug.Print "makeSomeArgsForMe: " & thisStartDate & "  -  " & thisEndDate
    
    
    Dim tmp As String
'    tmp = "page=1&max=500&charset=UTF-8" & _
'        "&ecwAutoId=false&ecwDivId=orderView&statusList=EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ" & _
'        thisStartDate & thisEndDate & _
'        "&eventType=RQ" & _
'        "&forecast=true" & _
'        "&logisticObjectTypeCode=" & arg & _
'        "&displayLocal=false" & _
'        "&productCode=" & partNumber

    tmp = "page=1&max=500&charset=UTF-8" & _
        "&forecast=true" & _
        "&eventType=RQ" & _
        thisStartDate & thisEndDate & _
        "ecwDivId=orderView" & _
        "&productCode=" & partNumber
    
    makeSomeArgsForMe = tmp
    
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
    d1 = Date - 10
    d2 = Date + 100
    
    
    
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
