VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeeklyCoverage"
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


Implements ICoverage


Private valid As Validator


Private theCorail As ICorail
Private theCorailData As ICorailData
Private theLayout As ILayout


Private repWorksheet As Worksheet
Private repCell As Range

Private ordersWorksheet As Worksheet
Private orderCell As Range



Private httpReq As HTTPRequestHandler




Private Sub Class_Initialize()
    Set valid = New Validator
    Set httpReq = New HTTPRequestHandler
    Set theCorail = New Corail
    
    Set theCorailData = Nothing
End Sub

Private Sub Class_Terminate()
    Set valid = Nothing
    Set httpReq = Nothing
    Set theCorail = Nothing
    
    Set theCorailData = Nothing
End Sub


Private Sub ICoverage_runCoverage(login As Variant, pass As Variant)



    ' init http req
    ' -------------------------------------
    httpReq.setup login, pass
    ' -------------------------------------

    ' new rep
    Set repWorksheet = ThisWorkbook.Sheets.Add
    repWorksheet.Cells(1, 1).Value = "Weekly Coverage on Corail, new rep: " & CStr(Now)
    repWorksheet.Cells(1, 2).Value = "BIND: "
    
    Set repCell = repWorksheet.Cells(2, 1)
    
    ' new order rep
    Set ordersWorksheet = ThisWorkbook.Sheets.Add
    ordersWorksheet.Cells(1, 1).Value = "Order list on Corail, new rep: " & CStr(Now)
    Set orderCell = ordersWorksheet.Cells(2, 1)
    fillWithLabels ordersWorksheet, orderCell
    Set orderCell = orderCell.Offset(1, 0)
    
    ordersWorksheet.Cells(1, 2).Value = "BIND: "
    ordersWorksheet.Cells(1, 3).Value = repWorksheet.Name
    
    ' repWorksheet.Cells(1, 2).Value = "BIND: "
    repWorksheet.Cells(1, 3).Value = ordersWorksheet.Name

    initialSetupForInputList ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM)
    
    Dim inputRef As Range
    ' starting point
    Set inputRef = ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM).Range("A2")
    
    Do
        
        makeInteration inputRef
    
        Set inputRef = inputRef.Offset(1, 0)
    Loop Until Trim(inputRef) = ""
    
    
    theLayout.finalTouchOnRep repWorksheet
    
End Sub


' ===========================================================================
' ===========================================================================
' MAIN LOGIC FOR ONE ITERATION - ONE PART NUMBER
Private Sub makeInteration(ByRef inputRef As Range)


    ' ----------------------------------------------------------------------------------------------------------
    ' iteracja pracuje na jednej linii danych
    ' kolejne etapy (pamietamy ze juz jestesmy w petli i raczej chcemy uniknac dodatkowej petli w
    ' tymze subie juz):
    '
    ' 3.
    ' poniewaz juz mamy gotowe http request handler - zakladam, ze pracowac bede pracowac synchronicznie
    ' i czekac kolejne linie bez sciagania sie - do zmiany (moze) w pozniejszym czasie
    '
    ' 5.
    ' nowy obiekt Corail o interfejsie ICorail wystepowac bedzie w postaci pojedynczej instancji
    ' sprawdzy wstepnie czy httReq jest uzyteczny i jesli tak to robimy connection
    '
    '
    ' 7. robimy kolejne standardowe polaczenie z danymi z ekranow 2720 oraz 2510 + opcjonalnie w kodzie widac
    ' linijki wyrzucajace cale pliki na lokalny dysk C do katalogu Temp - bardzo uzyteczne na poczatku mojej drogi
    ' powiazanej z mozliwoscia wykorzystania http post :)
    '
    '
    ' 11. theCorail.getDataFromScreen(2720, plt, pn) - to jest bardzo kompleksowa metoda, ktora pobiera dane z ekranu
    ' za pomoca zapytania http get lub post w zaleznosci od screenu jakiego pobieramy
    '
    ' jej interfejs wymuszony jest przez Interfejs ICorail:
    ' Private Function ICorail_getDataFromScreen(numberOfScreen As Variant, pltCode As Variant, partNumber As Variant) As ICorailData
    '
    ' wnetrznosci klasy Corail same ustalaja w jaki sposob dane maja zostac pobrane bez jego jakiej kolwiek ingerencji usera,
    ' nastepnie jest poprawnie parsowana i wyrzucona do postaci ustandaryzowanego obiekty interfejsu ICorailData
    ' i teraz bardzo wazne, poniewaz ICorailData moze byc wlasciwie wszystkim i niczym, jedne co maja ze soba wspolnego to,
    ' ze pierwotnym elementem wyjsciowym jest surowy string w standardzie html, ktory de facto moze zostac wyrzucony jako
    ' lokalny plik html
    ' nalezy pamietac ze klasy interfejsy ICorailData jest komponentem klasy Corail (oprocz ICorailData, jest jeszcze HTTP Request Handler oraz klasy ekranow Corail)
    '
    '
    ' schemat: Corail uruchamia obiekt klasy: HTTPRequestHandler, ktory usluguje swoim klientom: klasom ekranow Corail, ktore z kolei usuluguja jednemu
    ' standardowi: ICorailData, ktory jest ostatnia instancja przed wyrzuceniem danych do arkusza Excela
    '
    '
    '
    ' ICorailData priv components:
    ' Private rawString As String ( to ten surowy string )
    ' Private theDomHandler As DOMHandler ( surowy string uzbrojony w DOM - dzieki czemu mozemy odnosic sie do danych funkcjami typu getElementById itd itp)
    ' Private theParser As IParser (parser, ktory pozwala wyluskac bardzo szczegolowe i konkretne dane - duzo wewnetrznej wspolpracy pomiedzy tym obiektem a DOMem)
    '
    '
    '
    ' ----------------------------------------------------------------------------------------------------------

    Dim plt As String
    Dim pn As String
    Dim fso As New FSOHandler
    
    plt = CStr(inputRef.Value)
    pn = CStr(inputRef.Offset(0, 1).Value)
    
    If httpReq.usable() Then
        
        theCorail.connectWithHttpReq httpReq
        ' force english usage - basically for the date format ymd h:m
        ' ale jest NOK aktualnie - nie dziala poprawnie to zapytanie HTTP - widocznie potrzebne sa
        ' dodatkowe zapytania
        ' httpReq.httpGet "http://" & plt & ".commons.erp.corail.inetpsa.com/rest/sts-i18n/locales/en-US"
    
        If theCorail.connectWithScreen(2720) Then
        
            ' 2720
            ' --------------------------------------------------------------
            Set theCorailData = theCorail.getDataFromScreen(plt, pn, E_TYPE_OF_DATA_FROM_CORAIL_2720_COMMON)
            ' Debug.Print "theCorailData: " & TypeName(theCorailData) ' OK
            
            Dim suitableDataFrom2720 As SuitableData2720, ci As CoverageItem
            Set suitableDataFrom2720 = theCorailData.getSuitableData()
            
            ' OK
            ' Debug.Print "suitableDataFrom2720.stock: " & suitableDataFrom2720.stock
            
            
            ' 2720 xtra rqms
            ' --------------------------------------------------------------
            Set theCorailData = theCorail.getDataFromScreen(plt, pn, E_TYPE_OF_DATA_FROM_CORAIL_2720_XTRA_RQMS)
            Dim suitableXtraDataFrom2720 As SuitableXtraData2720
            Set suitableXtraDataFrom2720 = theCorailData.getSuitableData()
            
            
            ' --------------------------------------------------------------
            
            ' tylko na potrzeby testu wrzucamy odpowiedz http do lokalnego pliku
            'fso.makeHtmlFile "t2720_" & plt & "_" & pn, theCorailData.getString()
            
            ' 2510
            ' --------------------------------------------------------------
            Set theCorailData = theCorail.getDataFromScreen(plt, pn, E_TYPE_OF_DATA_FROM_CORAIL_2510_CALC)
            ' Debug.Print "theCorailData: " & TypeName(theCorailData) ' OK
            
            Dim suitableDataCalcFrom2510 As SuitableData2510, oi As OrderItem
            Set suitableDataCalcFrom2510 = theCorailData.getSuitableData()
            
            
            Dim r As Range
            
            
            Set r = orderCell
            
            ' Debug.Print "suit.getOrders() count: " & suitableDataCalcFrom2510.getOrders().Count
            
            
            
            For Each oi In suitableDataCalcFrom2510.getOrders()
            
                r.Offset(0, WCOC.E_ORDER_LIST_REP_PLT) = plt
                r.Offset(0, WCOC.E_ORDER_LIST_REP_PN) = oi.getElement(E_2510_PRODUCT)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_SGR) = oi.getElement(E_2510_SGR_LINE)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_QTY) = oi.getElement(E_2510_QTY)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_ORDER_NUM) = oi.getElement(E_2510_ORDER_NUMBER)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_DN) = oi.getElement(E_2510_DN)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_DHEO) = oi.getElement(E_2510_DHEO)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_DHEF) = oi.getElement(E_2510_DHEF)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_DHAS) = oi.getElement(E_2510_DHAS)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_DHRQ) = oi.getElement(E_2510_DHRQ)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_SELLER) = oi.getElement(E_2510_SELLER)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_SHIPPER) = oi.getElement(E_2510_SHIPPER)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_TYPE) = oi.getElement(E_2510_TYPE)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_UM) = oi.getElement(E_2510_UM)
                r.Offset(0, WCOC.E_ORDER_LIST_REP_STATUS) = oi.getElement(E_2510_S2)
                
                
                Set r = r.Offset(1, 0)
            Next oi
            
            ' set next empty position for next part number
            Set orderCell = r
            
            
            
            ' more like moving data from 2510 into 2720 into coverage item object
            ' ---------------------------------------------------------------------
            dataMerge suitableDataFrom2720, suitableXtraDataFrom2720, suitableDataCalcFrom2510
            ' ---------------------------------------------------------------------
            
            makeLineInCoverageRep suitableDataFrom2720, repCell
            
            ' --------------------------------------------------------------
            
            
            
            ' but this is really raw data without any setup at all
            'i = 1
            'For Each item In theCorailData.getCollectionOfData()
            '    ' Debug.Print TypeName(item)  ' String OK
            '    fso.makeHtmlFile "t2510_" & "_" & i & "_" & plt & "_" & pn, CStr(item)
            '    i = i + 1
            'Next item
            
        End If
    Else
        MsgBox "Fatal error during http connection!", vbExclamation
        End
    End If

End Sub
' ===========================================================================
' ===========================================================================


Private Sub makeLineInCoverageRep(main2720data As SuitableData2720, r As Range)



    
    ' args:
    ' r - starting point for this "line" of data
    ' main2720data - is exactly much more because already has also all important data collected from screen 2510
    
    Set theLayout = New WeeklyLayout
    
    With theLayout
        .setLayout E_TYPE_OF_LAYOUT_WEEKLY_DEFAULT
        .nowMergeWithDataFromCorail r, main2720data
        ' "colors" stands as symbol for how data and layout processing
        .putSomeColorsForThisPartNumber r
    End With
    
    Set r = r.Offset(7, 0)
    
End Sub



Private Sub initialSetupForInputList(ish As Worksheet)
    
    ' initial setup for my input worksheet
    ' ================================================================================================
    ''
    '
    
    If valid.checkIfThereIsAInputList(ish) Then
    
        ' no operation req!
        
    Else
        MsgBox "Wrong input list!", vbExclamation
        End
    End If
    '
    ''
    ' ================================================================================================
End Sub


Private Sub fillWithLabels(osh As Worksheet, refRange As Range)


    Dim r As Range
    Set r = refRange

    r.Offset(0, E_ORDER_LIST_PLT).Value = "PLT"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_PN) = "PN"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_SGR) = "SGR LINE"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_QTY) = "QTY"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_ORDER_NUM) = "Order#"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_DN) = "DN"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_DHEO) = "DHEO"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_DHEF) = "DHEF"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_DHAS) = "DHAS"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_DHRQ) = "DHRQ"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_SELLER) = "SELLER"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_SHIPPER) = "SHIPPER"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_TYPE) = "TYPE"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_UM) = "UM"
    r.Offset(0, WCOC.E_ORDER_LIST_REP_STATUS) = "STATUS"
End Sub




Private Sub dataMerge(data2720 As SuitableData2720, xtraData2720 As SuitableXtraData2720, data2510 As SuitableData2510)

    Dim orderItemDetails As IOrderDetails, covItem As CoverageItem, xtraRqms As XtraRqmItem
    
    ' OK!
    'Debug.Print "count collection for data merger (2720): " & data2720.getCollectionOfItems().Count
    'Debug.Print "count collection for data merger (2510): " & data2510.getOrders().Count
    
    ' Debug.Print xtraData2720.xtraRqms.Count ' OK
    
    
    ' add xtra rqms
    For Each xtraRqms In xtraData2720.xtraRqms
    
    
        
        
            Set covItem = findCovItem(xtraRqms.d, data2720.getCollectionOfItems())
            
            If covItem Is Nothing Then
                ' need a new covItem
                ' --------------------------------------------------------------------------
                If Len(Trim(xtraRqms.sgrLine)) = 2 Then
                
                    Set covItem = New CoverageItem
                    covItem.constructor xtraRqms.d, xtraRqms.Total, 0, 0, 0
                    covItem.setXtraRqms xtraRqms
                    data2720.addCovLine covItem
                End If
                ' --------------------------------------------------------------------------
            Else
                ' there is sth - check if data is the same
                ' --------------------------------------------------------------------------
                If Len(Trim(xtraRqms.sgrLine)) = 2 Then
                    covItem.setXtraRqms xtraRqms
                End If
                ' --------------------------------------------------------------------------
            End If
        
        
    Next xtraRqms
    
    
    
    
    For Each covItem In data2720.getCollectionOfItems()
        Set orderItemDetails = New OrderDetails
        ' inside this sub we asign orders into covItem already
        orderItemDetails.findYourOrders data2510, covItem
    Next covItem
End Sub


Private Function findCovItem(d As Date, c As Collection) As CoverageItem

    Set findCovItem = Nothing
    
    Dim ci As CoverageItem
    For Each ci In c
        If CDate(ci.getDate()) = CDate(d) Then
            Set findCovItem = ci
            Exit Function
        End If
    Next ci
End Function
