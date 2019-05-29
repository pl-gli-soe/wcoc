VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validator"
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



Public Function makeSomeValidationOnImportedPackage(pckg As ICorailData) As Boolean
    makeSomeValidationOnImportedPackage = False
    
    
    If pckg Is Nothing Then
        makeSomeValidationOnImportedPackage = False
    Else
        makeSomeValidationOnImportedPackage = True
    End If
End Function


Public Function checkIfThereIsAInputList(sh As Worksheet) As Boolean
    checkIfThereIsAInputList = False
    

    ' for now - really simple!
    If Trim(sh.Range("A2")) <> "" And Trim(sh.Range("B2")) <> "" Then
        ' at least one record is ok
        checkIfThereIsAInputList = True
    End If
End Function


Public Function checkAnswer(scrNum, answer) As Boolean
    
    checkAnswer = False
    
    
    ' narazie bez validacji - przepusc bezproblemowo
    checkAnswer = True
End Function
