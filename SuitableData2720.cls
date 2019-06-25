VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SuitableData2720"
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



' SUITABLE DATA FROM 2720

Public plt As String
Public pn As String
Public stock As Double
Public recv As Double
Public approId As String
Public approName As String
Public area As String
Public tipd As String
Public hebdo As String
Public periodique As String
Public upstreamFlow As String
Public internals As String
Public majorComment As String


' cofors added
Public vendorCofor As String
Public shipperCofor As String
Public manufacturerCofor As String

' supplier name
Public supplierName As String


Private item As CoverageItem
Private collectionOfCovLines As Collection


Public Sub addCovLine(i As CoverageItem)
    
    collectionOfCovLines.Add i
End Sub

Public Function getCollectionOfItems() As Collection
    Set getCollectionOfItems = collectionOfCovLines
End Function



Private Sub Class_Initialize()
    Set collectionOfCovLines = New Collection
End Sub

Private Sub Class_Terminate()
    Set collectionOfCovLines = Nothing
End Sub
