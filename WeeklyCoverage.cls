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
    Set repCell = repWorksheet.Cells(2, 1)
    
    ' new order rep
    Set ordersWorksheet = ThisWorkbook.Sheets.Add
    ordersWorksheet.Cells(1, 1).Value = "Order list on Corail, new rep: " & CStr(Now)
    Set orderCell = ordersWorksheet.Cells(2, 1)

    initialSetupForInputList ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM)
    
    Dim inputRef As Range
    ' starting point
    Set inputRef = ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM).Range("A2")
    
    Do
        
        makeInteration inputRef
    
        Set inputRef = inputRef.Offset(1, 0)
    Loop Until Trim(inputRef) = ""
    
    
End Sub


' ===========================================================================
' ===========================================================================
' MAIN LOGIC FOR ONE ITERATION
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
    
        If theCorail.connectWithScreen(2720) Then
        
            ' 2720
            ' --------------------------------------------------------------
            ' Set theCorailData = theCorail.getDataFromScreen(2720, plt, pn)
            ' Debug.Print "theCorailData: " & TypeName(theCorailData) ' OK
            ' --------------------------------------------------------------
            
            ' tylko na potrzeby testu wrzucamy odpowiedz http do lokalnego pliku
            'fso.makeHtmlFile "t2720_" & plt & "_" & pn, theCorailData.getString()
            
            ' 2510
            ' --------------------------------------------------------------
            Set theCorailData = theCorail.getDataFromScreen(2510, plt, pn)
            ' Debug.Print "theCorailData: " & TypeName(theCorailData) ' OK
            
            Dim suit As SuitableData2510, oi As OrderItem
            Set suit = theCorailData.getSuitableData()
            
            
            Dim r As Range
            
            
            Set r = orderCell.Offset(1, 0)
            
            Debug.Print "suit.getOrders() count: " & suit.getOrders().Count
            
            For Each oi In suit.getOrders()
            
                r.Value = plt
                r.Offset(0, 1) = pn
                r.Offset(0, 2) = oi.getElement(E_2510_PRODUCT)
                r.Offset(0, 3) = oi.getElement(E_2510_SGR_LINE)
                r.Offset(0, 4) = oi.getElement(E_2510_QTY)
                r.Offset(0, 5) = oi.getElement(E_2510_ORDER_NUMBER)
                r.Offset(0, 6) = oi.getElement(E_2510_DN)
                r.Offset(0, 7) = oi.getElement(E_2510_DH).DHEO
                r.Offset(0, 8) = oi.getElement(E_2510_DH).DHEF
                r.Offset(0, 9) = oi.getElement(E_2510_DH).DHXC
                r.Offset(0, 10) = oi.getElement(E_2510_DH).DHAS
                r.Offset(0, 11) = oi.getElement(E_2510_DH).DHRQ
                r.Offset(0, 12) = oi.getElement(E_2510_SELLER)
                r.Offset(0, 13) = oi.getElement(E_2510_SHIPPER)
                r.Offset(0, 14) = oi.getElement(E_2510_S2)
                
                
                Set r = r.Offset(1, 0)
            Next oi
            
            
            Set orderCell = r
            
            ' --------------------------------------------------------------
            
            ' tylko na potrzeby testu wrzucamy odpowiedz http do lokalnego pliku
            'fso.makeHtmlFile "t2510_" & plt & "_" & pn, theCorailData.getString()
            'Dim DH As DOMHandler
            'i = 1
            'For Each txt In theCorailData.getCollectionOfData()
            '    fso.makeHtmlFile "t2510_" & "_" & i & "_" & plt & "_" & pn, CStr(txt)
            '    i = i + 1
            'Next txt
            ' Debug.Print theCorailData.getCollectionOfData().Count
            
            
            ' but this is really raw data without any setup at all
            i = 1
            For Each item In theCorailData.getCollectionOfData()
                ' Debug.Print TypeName(item)  ' String OK
                fso.makeHtmlFile "t2510_" & "_" & i & "_" & plt & "_" & pn, CStr(item)
                i = i + 1
            Next item
        End If
    Else
        MsgBox "Fatal error during http connection!", vbExclamation
        End
    End If

End Sub
' ===========================================================================
' ===========================================================================


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
