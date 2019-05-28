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
        
        
        Set sd = New SuitableData2510
        
        Dim coll As Collection
        Debug.Print "test d.getCollectionOfData().count: " & d.getCollectionOfData().Count
        Set coll = d.getCollectionOfData()
        
        
        ' c is an collection of pages
        
        ' !!!
        '  -> change into dom
        Dim stronaOrderow As Variant
        Dim maintables As IHTMLElementCollection
        Dim divElement As IHTMLElement
        Dim divTempElement As IHTMLElement
        Dim parentDivTable As IHTMLElement
        
        
        Dim tds As IHTMLElementCollection, t As IHTMLTable, tc As IHTMLTableCell ' divTempElement - juz zadeklarowany i go wykorzystam tutaj
        Dim r As HTMLTableRow
        Dim c As HTMLTableCell
        
        Dim oItem As OrderItem
        
        
        ' temps vars
        ' --------------------------------------------------------------------------------------------------
        ' --------------------------------------------------------------------------------------------------
        
        Dim tmpStrSgrLine As String, tmpStrPn As String, tmpStrQty As String, tmpStrOrderNum As String
        Dim tmpStrDn As String, tmpStrDHEO As String, tmpStrDHEF As String, tmpStrDHXC As String
        Dim tmpStrDHAS As String, tmpStrDHRQ As String, tmpStrS1 As String, tmpStrRouting As String
        Dim tmpStrD As String, tmpStrSeller As String, tmpStrShipper As String, tmpStrS2 As String, tmpStrUM As String
        
        Dim tmpDH_Handler As DH_Handler
        
        ' --------------------------------------------------------------------------------------------------
        ' --------------------------------------------------------------------------------------------------
        
        
        
        For Each stronaOrderow In coll
        
        
            ' every page will be treat as seperate web-page
            d.getDom.idoc.body.innerHTML = stronaOrderow
            d.getDom.doc.body.innerHTML = stronaOrderow
            
            
            
            Set maintables = d.getDom.idoc.getElementsByTagName("div")
            
            
            ' just get first one!
            ' ===========================================
            For Each divElement In maintables
            
                If divElement.className = "ecwTableBody" Then
                
                    Set parentDivTable = divElement
                    Exit For
                End If
            Next divElement
            ' ===========================================
            
            

            Set t = parentDivTable.getElementsByTagName("table")(0)
        
        
            For Each r In t.Rows
            
                Set oItem = Nothing
                Set oItem = New OrderItem
                
                
                ' init by empty
                ' ----------------------------------------------
                tmpStrSgrLine = ""
                tmpStrPn = ""
                tmpStrQty = ""
                tmpStrOrderNum = ""
                tmpStrDn = ""
                tmpStrDHEO = ""
                tmpStrDHEF = ""
                tmpStrDHXC = ""
                tmpStrDHAS = ""
                tmpStrDHRQ = ""
                tmpStrS1 = ""
                tmpStrRouting = ""
                tmpStrD = ""
                tmpStrSeller = ""
                tmpStrShipper = ""
                tmpStrS2 = ""
                tmpStrUM = ""
                
                Set tmpDH_Handler = New DH_Handler
                ' ----------------------------------------------
                
                i = 1
            
                For Each c In r.Cells
                    ' Debug.Print c.outerHTML & ", class: " & c.className & " inner: " & c.innerHTML & " tag: " & c.tagName
                    ' <TH class=ecwTableSortable>SGR/Line</TH>, class: ecwTableSortable inner: SGR/Line tag: TH
                    ' ...
                    ' <TH>Qty</TH>, class:  inner: Qty tag: TH
                    ' <TH class=ecwTableSortable>DHEO</TH>, class: ecwTableSortable inner: DHEO tag: TH
                    ' ...
                    ' <TH>Seller </TH>, class:  inner: Seller  tag: TH
                    ' <TH>Shipper </TH>, class:  inner: Shipper  tag: TH
                    If c.tagName = "TD" Then
                        Set divTempElement = c.getElementsByTagName("div")(0)
                        
                        ' Debug.Print "c.getElementsByTagName(div)(0) > o: " & divTempElement.outerHTML & " i: " & divTempElement.innerHTML
                        
                        ' 1 &nbsp (consult button)
                        
                        ' 2 <DIV class=text>891/9</DIV> (SGR LIne)
                        If i = 2 Then tmpStrSgrLine = CStr(divTempElement.innerHTML)
                        
                        ' 3 <DIV class="ecwButtonTexteOver text">9810867580</DIV> i: 9810867580 (PN)
                        If i = 3 Then tmpStrPn = CStr(divTempElement.innerHTML)
                        
                        ' 4 <DIV class=number>720.0</DIV> i: 720.0 (Qty)
                        If i = 4 Then
                            tmpStrQty = CStr(divTempElement.innerHTML)
                        End If
                        
                        ' 5 <DIV class=text>004010222</DIV> i: 004010222 (order num)
                        If i = 5 Then tmpStrOrderNum = CStr(divTempElement.innerHTML)
                        
                        ' 6 <DIV class=text>69206418</DIV> i: 69206418 (dn)
                        If i = 6 Then tmpStrDn = CStr(divTempElement.innerHTML)
                        
                        ' 7 <DIV class=date>19/05/21 08:50</DIV> i: 19/05/21 08:50 (DHEO)
                        ' 8 <DIV class=date>19/05/23 14:00</DIV> i: 19/05/23 14:00 (DHEF)
                        ' 9 <DIV class=date>&nbsp;</DIV> i: &nbsp; (DHXC)
                        ' 10 <DIV class=date>19/05/28 10:20</DIV> i: 19/05/28 10:20 (DHAS)
                        ' 11 <DIV class=date>19/05/28 11:30</DIV> i: 19/05/28 11:30 (DHRQ)
                        If i >= 7 And i <= 11 Then
                            ' all DHXX here
                            ' --------------------------------------------------------------------------
                            
                            If i = 7 Then tmpStrDHEO = CStr(divTempElement.innerHTML)
                            If i = 8 Then tmpStrDHEF = CStr(divTempElement.innerHTML)
                            If i = 9 Then tmpStrDHXC = CStr(divTempElement.innerHTML)
                            If i = 10 Then tmpStrDHAS = CStr(divTempElement.innerHTML)
                            If i = 11 Then tmpStrDHRQ = CStr(divTempElement.innerHTML)
                            ' --------------------------------------------------------------------------
                        End If
                        
                        ' 12 <DIV class=text>3383</DIV> i: 3383 (S (1))
                        If i = 12 Then tmpStrS1 = CStr(divTempElement.innerHTML)
                        
                        ' 13 <DIV class=text>L9 FBC</DIV> i: L9 FBC (routing)
                        If i = 13 Then tmpStrRouting = CStr(divTempElement.innerHTML)
                        
                        ' 14 <DIV class=text>TC2</DIV> i: TC2 (D)
                        If i = 14 Then tmpStrD = CStr(divTempElement.innerHTML)
                        
                        ' 15 <DIV class=text>94656L 01</DIV> i: 94656L 01 (SELLER)
                        If i = 15 Then tmpStrSeller = CStr(divTempElement.innerHTML)
                        
                        ' 16 <DIV class=text>94656L 02</DIV> i: 94656L 02 (SHIPPER)
                        If i = 16 Then tmpStrShipper = CStr(divTempElement.innerHTML)
                        
                        ' 17 <DIV class="text RED">EF</DIV> i: EF (s (2) status)
                        If i = 17 Then tmpStrS2 = CStr(divTempElement.innerHTML)
                        
                        ' 18 <DIV class=text>20308</DIV> i: 20308 (UM)
                        If i = 18 Then
                            tmpStrUM = CStr(divTempElement.innerHTML)
                            
                            
                            
                            ' and this is the last one - here we can make some inital setup
                            ' on data...
                            'setup all data from i = 1 to 18
                            ' ==========================================================================
                            ' ==========================================================================
                            ''
                            '
                            Set tmpDH_Handler = Nothing
                            Set tmpDH_Handler = New DH_Handler
                            
                            If checkIfDateAndTime(tmpStrDHEO) Then
                                tmpDH_Handler.DHEO = parseStringToDateTime(tmpStrDHEO)
                            End If
                            
                            If checkIfDateAndTime(tmpStrDHEF) Then
                                tmpDH_Handler.DHEF = parseStringToDateTime(tmpStrDHEF)
                            End If
                            
                            If checkIfDateAndTime(tmpStrDHXC) Then
                                tmpDH_Handler.DHXC = parseStringToDateTime(tmpStrDHXC)
                            End If
                            
                            If checkIfDateAndTime(tmpStrDHAS) Then
                                tmpDH_Handler.DHAS = parseStringToDateTime(tmpStrDHAS)
                            End If
                            
                            If checkIfDateAndTime(tmpStrDHRQ) Then
                                tmpDH_Handler.DHRQ = parseStringToDateTime(tmpStrDHRQ)
                            End If
                            
                            
                            Dim qtyAsDouble As Double
                            qtyAsDouble = 0
                            On Error Resume Next
                            qtyAsDouble = CDbl(forceStringToBeDouble(tmpStrQty))
                            
                            With oItem
                                .coforConstructor tmpStrSeller, tmpStrShipper
                                .dhConstructor tmpDH_Handler
                                .commonConstructor tmpStrSgrLine, tmpStrPn, qtyAsDouble, tmpStrOrderNum, tmpStrDn, tmpStrS1, tmpStrRouting, tmpStrD, tmpStrS2, tmpStrUM
                            End With
                            
                            
                            '
                            ''
                            ' ==========================================================================
                            ' ==========================================================================
                        End If
                    End If
                    
                    i = i + 1
                Next c
                
                
                
                If oItem.checkIsObjectReady() Then
                    sd.addOrder oItem
                End If
                
                
            Next r
        

        Next stronaOrderow
        
        
        
        
            

        'For Each oItem In sd.getOrders()
        '     Debug.Print "oi pn: " & oItem.getElement(E_2510_PRODUCT) & " qty: " & oItem.getElement(E_2510_QTY) & " DHRQ: " & oItem.getElement(E_2510_DH).DHRQ
        'Next oItem
        
        
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
    
    Dim S As String
    S = Format(y & "-" & m & "-" & d, "yyyy-mm-dd")
    
    parseStringToDate = CDate(CStr(S))
    
End Function

Private Function parseStringToDateTime(strDT) As Date
    
    ' format 19/05/29 10:20
    ' format yy/mm/dd hh:mm
    
    Dim d As Integer
    Dim m As Integer
    Dim y As Integer
    
    Dim hh As Integer
    Dim mm As Integer
    

    If IsNumeric(Left(strDT, 2)) Then
        y = Int(Left(strDT, 2))
    End If
    
    If IsNumeric(Mid(strDT, 4, 2)) Then
        m = Int(Mid(strDT, 4, 2))
    End If
    
    If IsNumeric(Mid(strDT, 7, 2)) Then
        d = Int(Mid(strDT, 7, 2))
    End If
    
    If IsNumeric(Mid(strDT, 10, 2)) Then
        hh = Int(Mid(strDT, 10, 2))
    End If
    
    If IsNumeric(Mid(strDT, 13, 2)) Then
        mm = Int(Mid(strDT, 13, 2))
    End If
    
    
    Dim S As String
    S = Format(y & "-" & m & "-" & d & " " & hh & ":" & mm, "yy-mm-dd hh:mm")
    
    parseStringToDateTime = CDate(CStr(S))
    
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


Private Function checkIfDateAndTime(strDT) As Boolean

    checkIfDateAndTime = False
    
    If Len(strDT) = 14 Then
    
        If Mid(strDT, 3, 1) = "/" Then
            If Mid(strDT, 6, 1) = "/" Then
                If Mid(strDT, 9, 1) = " " Then
                    If Mid(strDT, 12, 1) = ":" Then
                        checkIfDateAndTime = True
                    End If
                End If
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


