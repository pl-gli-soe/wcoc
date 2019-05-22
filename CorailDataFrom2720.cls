VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CorailDataFrom2720"
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
Private sd As ISuitableData

Private Sub Class_Initialize()
    rawString = ""
    Set theDomHandler = New DOMHandler
    Set theParser = New Parser
End Sub

Private Sub Class_Terminate()
    rawString = ""
    Set theDomHandler = Nothing
    Set theParser = Nothing
End Sub


Private Function ICorailData_getDom() As DOMHandler
    Set ICorailData_getDom = theDomHandler
End Function


Private Function ICorailData_getString() As String
    ICorailData_getString = CStr(rawString)
End Function

Private Sub ICorailData_parseString()


    ' we're inside parsing
    ' ========================================
    ' Debug.Print "parseString... "
    ' Debug.Print rawString
    
    
    rawString = "<style> table tr td { padding: 4px; }</style>" & Chr(10) & rawString
    
    Set theDomHandler = Nothing
    Set theDomHandler = New DOMHandler

    With theDomHandler
        Set .doc = New HTMLDocument
        Set .idoc = New HTMLDocument
        .doc.body.innerHTML = rawString
        .idoc.body.innerHTML = .doc.body.innerHTML
    End With
    
    
    ' teraz parsujemy typowo dane z ekranu 2720
    ' -----------------------------------------------------------------------------------------------
    ' -----------------------------------------------------------------------------------------------
    

    theParser.importPackageOfData Me
    Set sd = theParser.getConvertedDataSuitableForExcel()
    
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
    
End Sub

