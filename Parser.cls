VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parser"
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


Implements IParser


Private valid As Validator

Private rawCorailData As ICorailData

' for 2720 and 2510
Private sd As Variant


Private Sub Class_Initialize()

    Set rawCorailData = Nothing
    Set valid = New Validator
    Set sd = Nothing

End Sub

Private Sub Class_Terminate()

    Set rawCorailData = Nothing
    Set valid = Nothing
    Set sd = Nothing
End Sub

Private Function IParser_getConvertedDataSuitableForExcel() As Variant
    Set IParser_getConvertedDataSuitableForExcel = sd
End Function

Private Function IParser_getPackage() As ICorailData
    Set IParser_getPackage = rawCorailData
End Function

Private Sub IParser_importPackageOfData(pckg As ICorailData)
    
    Set rawCorailData = pckg
    
    ' and try parse it!
    ' CorailDataFrom2720 or CorailDataFrom2510
    
    
    If TypeName(pckg) = "CorailDataFrom2720" Then
        innerParse2720
    ElseIf TypeName(pckg) = "CorailDataFrom2510" Then
        innerParse2510
    End If
End Sub













' 222222222222222    555555555555555555   1111111        000000000
'2:::::::::::::::22  5::::::::::::::::5  1::::::1      00:::::::::00
'2::::::222222:::::2 5::::::::::::::::5 1:::::::1    00:::::::::::::00
'2222222     2:::::2 5:::::555555555555 111:::::1   0:::::::000:::::::0
'            2:::::2 5:::::5               1::::1   0::::::0   0::::::0
'            2:::::2 5:::::5               1::::1   0:::::0     0:::::0
'         2222::::2  5:::::5555555555      1::::1   0:::::0     0:::::0
'    22222::::::22   5:::::::::::::::5     1::::l   0:::::0 000 0:::::0
'  22::::::::222     555555555555:::::5    1::::l   0:::::0 000 0:::::0
' 2:::::22222                    5:::::5   1::::l   0:::::0     0:::::0
'2:::::2                         5:::::5   1::::l   0:::::0     0:::::0
'2:::::2             5555555     5:::::5   1::::l   0::::::0   0::::::0
'2:::::2       2222225::::::55555::::::5111::::::1110:::::::000:::::::0
'2::::::2222222:::::2 55:::::::::::::55 1::::::::::1 00:::::::::::::00
'2::::::::::::::::::2   55:::::::::55   1::::::::::1   00:::::::::00
'22222222222222222222     555555555     111111111111     000000000
'
'
' #################################################################################
' #################################################################################
Private Sub innerParse2510()
    
    If valid.makeSomeValidationOnImportedPackage(rawCorailData) Then
        ' it is ok go on we working with this data
        ' -------------------------------------------
        
        Dim d As ICorailData ' CorailDataFrom2510
        Set d = rawCorailData
        
        With d
        
            Set sd = New SuitableData2510
            
        End With
        
        
        ' -------------------------------------------
    Else
        MsgBox "Wrong type of package!"
    End If
End Sub
'
'
' #################################################################################



                                                                               
'
' 222222222222222    77777777777777777777 222222222222222         000000000
'2:::::::::::::::22  7::::::::::::::::::72:::::::::::::::22     00:::::::::00
'2::::::222222:::::2 7::::::::::::::::::72::::::222222:::::2  00:::::::::::::00
'2222222     2:::::2 777777777777:::::::72222222     2:::::2 0:::::::000:::::::0
'            2:::::2            7::::::7             2:::::2 0::::::0   0::::::0
'            2:::::2           7::::::7              2:::::2 0:::::0     0:::::0
'         2222::::2           7::::::7            2222::::2  0:::::0     0:::::0
'    22222::::::22           7::::::7        22222::::::22   0:::::0 000 0:::::0
'  22::::::::222            7::::::7       22::::::::222     0:::::0 000 0:::::0
' 2:::::22222              7::::::7       2:::::22222        0:::::0     0:::::0
'2:::::2                  7::::::7       2:::::2             0:::::0     0:::::0
'2:::::2                 7::::::7        2:::::2             0::::::0   0::::::0
'2:::::2       222222   7::::::7         2:::::2       2222220:::::::000:::::::0
'2::::::2222222:::::2  7::::::7          2::::::2222222:::::2 00:::::::::::::00
'2::::::::::::::::::2 7::::::7           2::::::::::::::::::2   00:::::::::00
'2222222222222222222277777777            22222222222222222222     000000000
'
'
'
' #################################################################################
' #################################################################################
Private Sub innerParse2720()
    
    If valid.makeSomeValidationOnImportedPackage(rawCorailData) Then
        ' it is ok go on we working with this data
        ' -------------------------------------------
        
        
        Dim d As ICorailData ' CorailDataFrom2720
        Set d = rawCorailData
        
        
        ' contructor init also collection for coverage items
        Set sd = New SuitableData2720
        
        Dim inputElement As IHTMLInputElement
        
        
        ' suitable data Part Number
        ' ---------------------------------------------------------------------
        
        Set inputElement = d.getDom.idoc.getElementById("detailBeanId")
        
        sd.pn = inputElement.Value
        ' Debug.Print sd.pn OK!
        ' ---------------------------------------------------------------------
        
        
        ' suitable data for PLT
        ' ---------------------------------------------------------------------
        Set inputElement = d.getDom.idoc.getElementById("stockDepart")
        sd.stock = CDbl(forceStringToBeDouble(CStr(inputElement.Value)))
        ' ---------------------------------------------------------------------
        
        
        
        
        
        ' get full table from top of the 2720
        ' !!!
        ' table#tableauIdentificationProduct
        ''
        '
        '
        
        Dim tb As IHTMLTable
        Set tb = d.getDom.idoc.getElementById("tableauIdentificationProduct")
        
        Dim ec As IHTMLElementCollection
        ' .getDom.doc.getElementsByTagName
        Set ec = tb.getElementsByTagName("input")
        
        iterator = 1
        For Each inputElement In ec
            ' Debug.Print inputElement.Value & " " & inputElement.outerHTML
            
            
            If iterator = 6 Then
                ' (ID appro)
                If Trim(inputElement.Value) <> "" Then
                    sd.approId = Split(inputElement.Value, " - ")(0)
                    sd.approName = Split(inputElement.Value, " - ")(1)
                    
                    ' Debug.Print sd.approId & " 666 " & sd.approName ' OK!
                End If
                
            ElseIf iterator = 7 Then
                ' (Area)
                sd.area = inputElement.Value
                
            ElseIf iterator = 8 Then
                ' (Transfert in product date)
                sd.tipd = inputElement.Value
            ElseIf iterator = 11 Then
                ' (CMJ JL Hebdo)
                sd.hebdo = inputElement.Value
            ElseIf iterator = 12 Then
                '(CMJ JL Périodique)
                sd.periodique = inputElement.Value
            ElseIf iterator = 13 Then
                sd.upstreamFlow = inputElement.Value
            ElseIf iterator = 14 Then
                sd.internals = inputElement.Value
            End If
            
            iterator = iterator + 1
        Next
        
        ' ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
        ' in ec indxs on 23 of may 2019:
        ' 1: <INPUT id=detailBeanId maxLength=20 size=15 value=9810867580>
        ' 2: <INPUT id=hidDetailBeanId type=hidden name=detailBeanId>
        ' 3: <INPUT id=sgrCode type=hidden name=sgrCode>
        ' 4: <INPUT id=lineCode type=hidden name=lineCode>
        ' 5: <INPUT id=lineId type=hidden name=lineId>
        ' 6: P526139 - LAMBERT GEORGES <INPUT class=number readOnly maxLength=20 size=15 value="P526139 - LAMBERT GEORGES"> (ID appro)
        ' 7: E41 <INPUT class=number readOnly maxLength=20 size=15 value=E41> (Area)
        ' 8: 14/06/2016 <INPUT class=number readOnly maxLength=20 size=15 value=14/06/2016> (Transfert in product date)
        ' 9: <INPUT class=number readOnly maxLength=20 size=15>
        ' 10: <INPUT class=number readOnly maxLength=20 size=15>
        ' 11: 2544.0 <INPUT class=number readOnly maxLength=20 size=15 value=2544.0> (CMJ JL Hebdo)
        ' 12: 2632.0 <INPUT class=number readOnly maxLength=20 size=15 value=2632.0> (CMJ JL Périodique)
        ' 13: 9 <INPUT class=number readOnly maxLength=20 size=15 value=9> (Upstream Flow)
        ' 14 743 <INPUT class=number readOnly maxLength=20 size=15 value=743> (Internals)
        ' 15: 2019-05-16T07:08:14Z <INPUT id=internalAlertBeginDate type=hidden value=2019-05-16T07:08:14Z name=internalAlertBeginDate>
        ' 16: 2019-05-23T07:08:14Z <INPUT id=internalAlertEndDate type=hidden value=2019-05-23T07:08:14Z name=internalAlertEndDate>
        ' 17: on <INPUT disabled CHECKED type=checkbox value="">
        ' 18: on <INPUT disabled CHECKED type=checkbox value="">
        ' 19: on <INPUT disabled CHECKED type=checkbox value="">
           
        ''
        '
        '
            
            
        
        
        ' -------------------------------------------
        
        
        ' TEXTAREA MAJOR COMMENT (LAST UPDATE):
        Dim textAreaElement As IHTMLElement
        
        Set textAreaElement = d.getDom.idoc.getElementById("getProductSummaryRead_detailBean_productSummaryNoteList_lastProductSummaryNotes_noteComments")
        sd.majorComment = textAreaElement.innerText
        ' Debug.Print sd.majorComment ' OK!
        
        
        
        
        ' tables of rqms, transit end disruption
        ' ----------------------------------------------------------------------------------------------
        ' ----------------------------------------------------------------------------------------------
        ' ----------------------------------------------------------------------------------------------
        ''
        ''
        '
        
        Dim mainTableDiv As IHTMLElement
        Set mainTableDiv = d.getDom.idoc.getElementById("demandAndSupplyPerDay")
        
        ' now inside this mainTableDiv we have another div.ecwTableHead on top and on bottom div.ecwTableBody
        ' those 2 have direct connection with tables which includes data of our interests :)
        
        Dim maintables As IHTMLElementCollection
        Dim divTableHead As IHTMLElement, divTableBody As IHTMLElement, divTempElement As IHTMLElement
        'Set maintables = d.getDom.idoc.getElementsByClassName("ecwTableHead")
        Set maintables = mainTableDiv.getElementsByTagName("div")
        
        For Each divTempElement In maintables
        
            If divTempElement.className = "ecwTableHead" Then
                Set divTableHead = divTempElement
            End If
            
            If divTempElement.className = "ecwTableBody" Then
                Set divTableBody = divTempElement
            End If
            
        Next
        
        
        
        ' Debug.Print divTableBody.innerHTML ' OK
        ' ecwTableHead should have exactly same data as ecwTableBody - and I do not know why
        ' but at the I'll go with body
        
        
        Dim tds As IHTMLElementCollection, t As IHTMLTable, tc As IHTMLTableCell ' divTempElement - juz zadeklarowany i go wykorzystam tutaj
        Dim r As HTMLTableRow
        Dim c As HTMLTableCell
        Set t = divTableBody.getElementsByTagName("table")(0)
        
        
        
        Dim tmpDate As String, tmpRqm1 As String, tmpRqm2 As String, tmpOrderStr As String, tmpShipStr As String
        Dim tmpDateDate As Date, tmpRqm1Dbl As Double, tmpRqm2Dbl As Double, tmpOrderStrDbl As Double, tmpShipStrDbl As Double
        
        
        Dim covItem As CoverageItem
        
        For Each r In t.Rows
        
            Set covItem = Nothing
            Set covItem = New CoverageItem
        
            For Each c In r.Cells
                'Debug.Print c.outerHTML & ", class: " & c.className & " inner: " & c.innerHTML & " tag: " & c.tagName
                ' <TD ecwKeyName="date"><DIV class=date>22/05/2019</DIV></TD>, class:  inner: <DIV class=date>22/05/2019</DIV> tag: TD
                If c.tagName = "TD" Then
                    Set divTempElement = c.getElementsByTagName("div")(0)
                    
                    ' div.innerHTML: 22/05/2019 class: [date] (DATE)
                    ' div.innerHTML: 113.0 class: [number fbpcDetailByDay] (PC DEMAND)
                    ' div.innerHTML: 0.0 class: [number] (Incl. BC) - not suff -> ?
                    ' div.innerHTML: 0.0 class: [ecwButtonTexteOverOrdered number] ' <TD ecwKeyName="orderedQuantity"
                    ' div.innerHTML: 0.0 class: [ecwButtonTexteOverShipped number] ' <TD ecwKeyName="shippedQuantity"
                    ' div.innerHTML: &nbsp; class: [dateRupture]
                    ' div.innerHTML: &nbsp; class: [dateRupture]
                    ' div.innerHTML: &nbsp; class: [dateRupture] ' <TD ecwKeyName="formattedRelativeDisruptionDate"
                    
                    ' Debug.Print "div.innerHTML: " & divTempElement.innerHTML & " class: [" & divTempElement.className & "]"
                    
                    If divTempElement.className = "date" Then
                        ' date
                        tmpDate = divTempElement.innerText
                        tmpDateDate = parseStringToDate(tmpDate)
                        
                        
                        
                    ElseIf divTempElement.className = "number fbpcDetailByDay" Then
                        ' pc demand
                        tmpRqm1 = divTempElement.innerText
                        tmpRqm1Dbl = CDbl(forceStringToBeDouble(tmpRqm1))
                        
                    ElseIf divTempElement.className = "number" And c.outerHTML Like "*requiredBCQuantity*" Then
                        ' incl bc
                        tmpRqm2 = divTempElement.innerText
                        tmpRqm2Dbl = CDbl(forceStringToBeDouble(tmpRqm2))
                        
                        
                    ElseIf divTempElement.className Like "ecwButtonTexteOverOrdered*" Then
                        ' ordered
                        tmpOrderStr = divTempElement.innerText
                        tmpOrderStrDbl = CDbl(forceStringToBeDouble(tmpOrderStr))
                        
                    ElseIf divTempElement.className Like "ecwButtonTexteOverShipped*" Then
                        ' shipped
                        tmpShipStr = divTempElement.innerText
                        tmpShipStrDbl = CDbl(forceStringToBeDouble(tmpShipStr))
                        
                        covItem.makeReady
                        Exit For
                        
                    End If
                    
                    
                Else
                    Exit For
                End If
            Next c
            
            
            If covItem.isReady() Then
                covItem.constructor CDate(tmpDateDate), CDbl(tmpRqm1Dbl), CDbl(tmpRqm2Dbl), CDbl(tmpOrderStrDbl), CDbl(tmpShipStrDbl)
                sd.addCovLine covItem
            End If
        Next r
        
        
        
        
        ' OK!
        ' For Each covItem In sd.getCollectionOfItems()
        '     Debug.Print "ci : " & covItem.getStrDate & " r " & covItem.getRqms & " t " & " o " & covItem.getOrder & covItem.getShip
        ' Next covItem
        
        '
        ''
        ''
        ' ----------------------------------------------------------------------------------------------
        ' ----------------------------------------------------------------------------------------------
        ' ----------------------------------------------------------------------------------------------
        
        
        
        
    Else
        MsgBox "Wrong type of package!"
    End If
End Sub
'
'
' #################################################################################



Private Function parseStringToDate(strD) As Date
    
    ' format MM/DD/YYYY
    
    Dim d As Integer
    Dim m As Integer
    Dim y As Integer
    
    If checkIfDate(strD) Then
    
        If IsNumeric(Left(strD, 2)) Then
            d = Int(Left(strD, 2))
        End If
        
        If IsNumeric(Mid(strD, 4, 2)) Then
            m = Int(Mid(strD, 4, 2))
        End If
        
        If IsNumeric(Mid(strD, 7, 4)) Then
            y = Int(Right(strD, 4))
        End If
    
    End If
    
    Dim s As String
    s = Format(y & "-" & m & "-" & d, "yyyy-mm-dd")
    
    parseStringToDate = CDate(CStr(s))
    
End Function

Private Function checkIfDate(strD) As Boolean

    checkIfDate = False
    
    If Len(strD) = 10 Then
    
        If Mid(strD, 3, 1) = "/" Then
            If Mid(strD, 6, 1) = "/" Then
                checkIfDate = True
            End If
        End If
    End If
End Function


Private Function forceStringToBeDouble(arg As String) As Double

    ' Application.DecimalSeparator
    Dim convertedValue As String
    convertedValue = CStr(arg)
    ' double sure shot! :D
    convertedValue = Replace(convertedValue, ",", Application.DecimalSeparator)
    convertedValue = Replace(convertedValue, ".", Application.DecimalSeparator)
    
    forceStringToBeDouble = CDbl(convertedValue)
End Function


