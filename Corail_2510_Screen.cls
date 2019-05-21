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
' GET /getOrderViewList.do?location=&productCode=9465048780
' &startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z
' &eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN
' &logisticObjectTypeCode=CALCULATED
' &forecast=false HTTP/1.1
' Host: sx.Control.erp.Corail.inetpsa.com


Implements ICorailScreen


Private preUrl As String
Private screenCalcUrl As String
Private screenUpdatedUrl As String
Private url As String


Private req As HTTPRequestHandler

Private theCorailData As ICorailData

Private valid As Validator


Private Sub Class_Initialize()


    num = 2510

    'screenCalcUrl = "/getOrderViewList.do?location=&startDate=&endDate=" & _
    '    "&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN" & _
    '    "&logisticObjectTypeCode=CALCULATED&forecast=false&productCode="
    
    screenCalcUrl = "/getOrderViewList.do?productCode="
        
    screenUpdatedUrl = "/getOrderViewList.do?location=&startDate=&endDate=" & _
        "&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN" & _
        "&logisticObjectTypeCode=UPDATED&forecast=false&productCode="
        
    preUrl = ".control.erp.corail.inetpsa.com"
    url = ""

    Set req = Nothing
    
    Set valid = New Validator
End Sub

Private Sub Class_Terminate()

    preUrl = ""
    screenCalcUrl = ""
    screenUpdatedUrl = ""
    url = ""

    Set req = Nothing
    
    Set valid = Nothing
End Sub

Private Function ICorailScreen_getData(plt As Variant, partNumber As Variant) As ICorailData


    ' CALCULATED
    ' -----------------------------------------------------------------------------
    
    
'    testCalcUrl = "/getOrderViewList.do?location=&productCode=9465048780" & _
'        "&startDate=2019-05-01T22:00:00Z&endDate=2019-06-27T22:00:00Z" & _
'        "&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN" & _
'        "&logisticObjectTypeCode=CALCULATED" & _
'        "&forecast=false"
'
'    testCalcUrl2 = "/getOrderViewList.do?location=&productCode=9465048780" & _
'        "&startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z" & _
'        "&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN" & _
'        "&logisticObjectTypeCode=CALCULATED&forecast=false"
'
'
'
'    url = "http://" & CStr(plt) & preUrl & CStr(testCalcUrl2)
'    'url = "http://" & CStr(plt) & preUrl & screenCalcUrl & CStr(partNumber)
    
    
    
    
    
    'answer = req.httpGet(url)
    'answer = req.httpGet("http://sx.control.erp.corail.inetpsa.com/getOrderViewList.do?location=
    '&productCode=9465048780&startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z&eventType=RQ
    '&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN&logisticObjectTypeCode=CALCULATED&forecast=false")
    
    
    ''
    ''
    'args = "location=&productCode=9465048780&startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z" & _
    '    "&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN" & _
    '    "&logisticObjectTypeCode=CALCULATED&forecast=false"
    ' args = "location=&productCode=9465048780&startDate=2019-05-26T22:00:00Z&endDate=2019-05-27T21:59:00Z&eventType=RQ&statusList=CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:AN&logisticObjectTypeCode=CALCULATED&forecast=false"
        
    ' answer = answer + req.httpPost("http://sx.control.erp.corail.inetpsa.com/getOrderViewPager.do", args)
    
    
    
    Dim answer As String
    uryl = "http://sx.control.erp.corail.inetpsa.com/getOrderViewPager.do"
    'args = "page=1&max=10&charset=UTF-8&ecwAutoId=false&ecwDivId=orderView&productCode=9810867580&statusList=FO:CA:EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:IN_CD:ST:BL:AN" & _
    '    "&eventType=RQ&forecast=false&logisticObjectTypeCode=CALCULATED&displayLocal=false"
    
    args = "page=1&max=256&charset=UTF-8&ecwAutoId=false&ecwDivId=orderView&statusList=EO:EF:XC:RP:XP:DT:AT:EP:AS:RQ:ST" & _
        "&eventType=RQ&forecast=false&logisticObjectTypeCode=CALCULATED&displayLocal=false&productCode=9810867580"
    
    
    answer = req.httpPost(uryl, args)
    Set theCorailData = New CorailData
    theCorailData.setString answer
    theCorailData.parseString
    
    'If valid.checkAnswer(num, answer) Then
    '    Set theCorailData = New CorailData
    '    theCorailData.setString answer
    '    theCorailData.parseString
    'Else
    '    Set theCorailData = Nothing
    'End If
    
    
    Set ICorailScreen_getData = theCorailData
    ' -----------------------------------------------------------------------------
    
    
    ' UPDATED
    ' -----------------------------------------------------------------------------
    ' url = "http://" & CStr(plt) & preUrl & screenUpdatedUrl & CStr(partNumber)
    ' -----------------------------------------------------------------------------
    
End Function

Private Function ICorailScreen_openData(plt As Variant, partNumber As Variant) As Boolean
    ICorailScreen_openData = False
End Function

Private Sub ICorailScreen_screenSetup(number As Variant, httpReq As HTTPRequestHandler)

    Set req = httpReq
    
End Sub
