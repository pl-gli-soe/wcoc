VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailDataFrom2510"
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

Implements ICorailData



Private rawString As String
Private theDomHandler As DOMHandler
Private theParser As IParser
Private sd As SuitableData2510


' collection for collecting rawString and theDomHandler
Private collectionOfDataFrom2510 As Collection

Private numberOfElements As Long
Private currentPage As Integer
Private currentPageSize As Integer
Private pages As Integer

Private Sub Class_Initialize()
    rawString = ""
    Set theDomHandler = New DOMHandler
    Set theParser = New Parser
    Set collectionOfDataFrom2510 = New Collection
    numberOfElements = 0
    currentPage = 0
    currentPageSize = 0
    pages = 0
End Sub

Private Sub Class_Terminate()
    rawString = ""
    Set theDomHandler = Nothing
    Set theParser = Nothing
    Set collectionOfDataFrom2510 = Nothing
    numberOfElements = 0
    currentPage = 0
    currentPageSize = 0
    pages = 0
End Sub

Private Function ICorailData_getCollectionOfData() As Collection
    Set ICorailData_getCollectionOfData = collectionOfDataFrom2510
End Function

Private Function ICorailData_getString() As String
    ICorailData_getString = CStr(rawString)
End Function

Private Function ICorailData_getDom() As DOMHandler
    Set ICorailData_getDom = theDomHandler
End Function

Private Function ICorailData_getSuitableData() As Variant
    Set ICorailData_getSuitableData = sd
End Function

Private Sub ICorailData_parseString()


    ' we're inside parsing
    ' ========================================
    ' Debug.Print "parseString... "
    ' Debug.Print rawString
    '
    ' now from 2510 perspective we need to parse many strings from
    ' collection so we need here some loop
    
    Dim buffColl As Collection
    Set buffColl = New Collection
    
    For Each txt In collectionOfDataFrom2510
        txt = "<style> table tr td { padding: 4px; }</style>" & Chr(10) & txt
        buffColl.Add txt
    Next txt
    
    Set collectionOfDataFrom2510 = Nothing
    Set collectionOfDataFrom2510 = New Collection
    
    For Each txt In buffColl
        collectionOfDataFrom2510.Add txt
    Next txt
    
    Set buffColl = Nothing
    
    
    
    ' teraz parsujemy typowo dane z ekranu 2510
    ' -----------------------------------------------------------------------------------------------
    ' -----------------------------------------------------------------------------------------------
    

    theParser.importPackageOfData Me
    ' sd As SuitableData2510 and foo as Variant but VBA allows
    Set sd = theParser.getConvertedDataSuitableForExcel()
    
    ' Debug.Print "ICorailData_parseString " & TypeName(sd) ' OK!
    
    
    ' -----------------------------------------------------------------------------------------------
    ' -----------------------------------------------------------------------------------------------
    
    ' ========================================

End Sub

Private Sub ICorailData_setString(arg As String)
    rawString = arg
    
    With theDomHandler
        Set .doc = New HTMLDocument
        Set .idoc = New HTMLDocument
        .doc.body.innerHTML = arg
        .idoc.body.innerHTML = .doc.body.innerHTML
    End With
    
    ' check first line for validation
    ' ----------------------------------------------
    firstLine = Split(rawString, Chr(10))(0)
    
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
    
    
    Debug.Print "ICorailData_setString: num " & numberOfElements & " curr p: " & currentPage & " pages: " & pages
    
    If CLng(currentPage) <= CLng(pages) Then
        collectionOfDataFrom2510.Add arg
    End If
    
End Sub

