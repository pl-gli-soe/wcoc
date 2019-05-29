VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Corail_2510_Screen"
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
    ileLiniiNaStronie = 40
    
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
    Dim ktoraStrona As Integer
    
    answer = ""
    ktoraStrona = 1
    
    
    Set theCorailData = New CorailDataFrom2510
    

    Do

        If e = E_TYPE_OF_DATA_FROM_CORAIL_2510_CALC Then
        
        
            ' CALCULATED
            ' -----------------------------------------------------------------------------
        
        
            ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
            url = "http://" & CStr(plt) & preUrl & screenUrl
            args = makeSomeArgsForMe("CALCULATED", partNumber, ktoraStrona, False)
            
            Debug.Print "http post: " & url & " " & args
            answer = req.httpPost(url, args)
            
            ' but for this example this set is actually add!!!
            theCorailData.setString answer
            
            ' -----------------------------------------------------------------------------
            
        ElseIf e = E_TYPE_OF_DATA_FROM_CORAIL_2510_CALC_WITH_FORECAST Then
        
            ' CALCULATED WIth FORECAST
            ' -----------------------------------------------------------------------------
        
        
            ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
            url = "http://" & CStr(plt) & preUrl & screenUrl
            args = makeSomeArgsForMe("CALCULATED", partNumber, ktoraStrona, True)
            
            Debug.Print "http post: " & url & " " & args
            answer = req.httpPost(url, args)
            
            ' but for this example this set is actually add!!!
            theCorailData.setString answer
            
            ' -----------------------------------------------------------------------------
        
        
        ElseIf e = E_TYPE_OF_DATA_FROM_CORAIL_2510_UPDATED Then
        
            ' UPDATED
            ' -----------------------------------------------------------------------------
            
            ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
            url = "http://" & CStr(plt) & preUrl & screenUrl
            args = makeSomeArgsForMe("UPDATED", partNumber, ktoraStrona)
            answer = req.httpPost(url, args)
            
            ' but for this example this set is actually add!!!
            theCorailData.setString answer
            
            ' -----------------------------------------------------------------------------
            
        ElseIf e = E_TYPE_OF_DATA_FROM_CORAIL_2510_UPDATED_WITH_FORECAST Then
        
            ' UPDATED WITH FORECAST
            ' -----------------------------------------------------------------------------
            
            ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
            url = "http://" & CStr(plt) & preUrl & screenUrl
            args = makeSomeArgsForMe("UPDATED", partNumber, ktoraStrona, True)
            answer = req.httpPost(url, args)
            
            ' but for this example this set is actually add!!!
            theCorailData.setString answer
            
            ' -----------------------------------------------------------------------------
            
        Else
            
            ' CALCULATED WITHOUT FORECAST BY DEFAULT
            ' -----------------------------------------------------------------------------
            
            ' not dry, but i want to this - to see exactly what is going on close betweeen the lines
            url = "http://" & CStr(plt) & preUrl & screenUrl
            args = makeSomeArgsForMe("CALCULATED", partNumber, ktoraStrona, False)
            answer = req.httpPost(url, args)
            
            
            ' but for this example this set is actually add!!!
            theCorailData.setString answer
            
            ' -----------------------------------------------------------------------------
        End If
        
        
        ktoraStrona = ktoraStrona + 1
    
    Loop While isThereStillSthToParse(answer)

    ' now parse all
    theCorailData.parseString
    Set ICorailScreen_getData = theCorailData

    
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


Private Function isThereStillSthToParse(answer) As Boolean
    isThereStillSthToParse = False
    
    
    
    firstLine = Split(answer, Chr(10))(0)
    
    arr = Split(firstLine, " ")
    
    
    Dim newVal As Long
    
    For x = LBound(arr) To UBound(arr)
        ' Debug.Print arr(x)
        
        If CStr(arr(x)) Like "numberOfElement=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = CLng(tmp)
            numberOfElements = newVal
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "pageSize=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = Int(tmp)
            currentPageSize = newVal
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "indicePage=*" Then
        
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            newVal = Int(tmp)
            currentPage = newVal
            ' Exit For
        End If
        
        If CStr(arr(x)) Like "numberOfPage=*" Then
            tmp = Split(arr(x), "=")(1)
            tmp = Replace(tmp, Chr(34), "")
            tmp = Replace(tmp, "-->", "")
            newVal = Int(tmp)
            pages = newVal
        End If
    Next x
    
    If currentPage <= pages Then
        isThereStillSthToParse = True
    Else
        isThereStillSthToParse = False
    End If
End Function



Private Function makeSomeArgsForMe(arg, partNumber, ktoraStrona, Optional forecast As Boolean) As String
    
    ' arg - updated or calculated - but at the end will have same url from contruction perspective
    ' becuae all rest in baiscally the same dates and other params
    
    ' arg: UPDATED -> SHIPPED
    ' arg: CALCULATED -> ORDER
    
    ' comment regarding order:
    ' need extra param on status: CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN ' do not know why...
    
    
    ' example template
    'args = "page=1&max=256&charset=UTF-8&ecwAutoId=false&ecwDivId=orderView&statusList=EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:ST" & _
    '    "&eventType=RQ&forecast=false&logisticObjectTypeCode=CALCULATED&displayLocal=false&productCode=9810867580"
    
    ' pod ordery! 209-05-28
    ' &statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN
    ' ? FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN
    
    
    
    thisStartDate = parseFromDate(1)
    thisEndDate = parseFromDate(2)
    
    ' Debug.Print "makeSomeArgsForMe: " & thisStartDate & "  -  " & thisEndDate
    
    
    Dim tmp As String
    
    If forecast Then
    
    
        If arg <> "" Then
    
            tmp = "page=" & ktoraStrona & _
                "&max=" & ileLiniiNaStronie & _
                "&charset=UTF-8" & _
                "&forecast=true" & _
                "&logisticObjectTypeCode=" & arg & _
                "&displayLocal=false" & _
                "&statusList=FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN" & _
                "&eventType=RQ" & _
                thisStartDate & thisEndDate & _
                "&productCode=" & partNumber
        
        Else
        
            tmp = "page=" & ktoraStrona & _
                "&max=" & ileLiniiNaStronie & _
                "&charset=UTF-8" & _
                "&forecast=true" & _
                "&displayLocal=false" & _
                "&statusList=FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN" & _
                "&eventType=RQ" & _
                thisStartDate & thisEndDate & _
                "&productCode=" & partNumber
        
        End If
    
    
    Else


        If arg <> "" Then
    
            tmp = "page=" & ktoraStrona & _
                "&max=" & ileLiniiNaStronie & _
                "&charset=UTF-8" & _
                "&forecast=false" & _
                "&logisticObjectTypeCode=" & arg & _
                "&displayLocal=false" & _
                "&statusList=FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN" & _
                "&eventType=RQ" & _
                thisStartDate & thisEndDate & _
                "&productCode=" & partNumber
        
        Else
        
            tmp = "page=" & ktoraStrona & _
                "&max=" & ileLiniiNaStronie & _
                "&charset=UTF-8" & _
                "&forecast=false" & _
                "&displayLocal=false" & _
                "&statusList=FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN" & _
                "&eventType=RQ" & _
                thisStartDate & thisEndDate & _
                "&productCode=" & partNumber
        
        End If
    End If
    
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
