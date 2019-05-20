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
Private theLayout As ILayout


Private repWorksheet As Worksheet
Private repCell As Range


Private httpReq As HTTPRequestHandler




Private Sub Class_Initialize()
    Set valid = New Validator
    Set httpReq = Nothing
    Set theCorail = New Corail
End Sub

Private Sub Class_Terminate()
    Set valid = Nothing
    Set httpReq = Nothing
    Set theCorail = Nothing
End Sub

Private Sub ICoverage_runCoverage(login As Variant, pass As Variant)



    ' init http req
    ' -------------------------------------
    httpReq.setup login, pass
    ' -------------------------------------

    ' new rep
    Set repWorksheet = ThisWorkbook.Sheets.Add
    repworkhsheet.Cells(1, 1).Value = "Weekly Coverage on Corail, new rep: " & CStr(Now)
    Set repCell = repWorksheet.Cells(2, 1)


    initialSetupForInputList ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM)
    
    Dim inputRef As Range
    ' starting point
    Set inputRef = ThisWorkbook.Sheets(WCOC.G_INPUT_SH_NM).Range("A2")
    
    Do
        
        makeInteration inputRef
    
        Set inputRef = inputRef.Offset(1, 0)
    Loop While Trim(inputRef) = ""
    
    
End Sub


' ===========================================================================
' ===========================================================================
' MAIN LOGIC FOR ONE ITERATION
Private Sub makeInteration(ByRef inputRef As Range)

    Dim plt As String
    Dim pn As String
    
    plt = CStr(inputRef.Value)
    pn = CStr(inputRef.Offset(0, 1).Value)
    
    If theCorail.connectWithScreen() Then
        ' theCorail.getDataFromScreen( 2720, plt, pn )

End Sub
' ===========================================================================
' ===========================================================================


Private Sub initialSetupForInputList(ish As Worksheet)
    
    ' initial setup for my input worksheet
    ' ================================================================================================
    ''
    '
    
    If valid.checkIfThereIsAInputList(ish) Then
    Else
        MsgBox "Wrong input list!", vbExclamation
        End
    End If
    '
    ''
    ' ================================================================================================
End Sub
