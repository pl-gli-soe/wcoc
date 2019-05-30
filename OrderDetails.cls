VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderDetails"
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


Implements IOrderDetails




Private matchedOrders As Collection


Private Sub Class_Initialize()
    Set matchedOrders = New Collection
End Sub

Private Sub Class_Terminate()
    Set matchedOrders = Nothing
End Sub

Private Sub IOrderDetails_findYourOrders(dataFrom2510 As SuitableData2510, covItem As CoverageItem)


    Dim d As Date
    d = CDate(covItem.getDate())

    ' o1 - order as double from coverage item
    ' s1 - shipment as double from coverage item
    Dim oi As OrderItem, o1 As Double, s1 As Double, qty As Double, status As String
    Dim dateFromDHRQ As Date
    For Each oi In dataFrom2510.getOrders()
    
    
        ' Debug.Print "try to match: " & CDate(oi.getElement(E_2510_DHRQ)) & " =? " & CDate(d)
        'If CDate(oi.getElement(E_2510_DHRQ)) = CDate(d) Then
        'End If
        
        dateFromDHRQ = CDate(Format(oi.getElement(E_2510_DHRQ), "yyyy-mm-dd"))
        
        ' Debug.Print dateFromDHRQ & " = ? " & d
        If CDate(dateFromDHRQ) = CDate(d) Then
        
            o1 = covItem.getOrder()
            s1 = covItem.getShip()
            qty = oi.getElement(E_2510_QTY)
            status = oi.getElement(E_2510_S2)
            
            Debug.Print "dates matched: " & d & " now: o1, s1: " & o1 & ", " & s1 & " and order it self (qty): " & qty & " s " & status
            
            matchedOrders.Add oi
        End If
    Next oi
    
    
    ' all oi found now set to covItem
    covItem.setOrderDetails Me

End Sub

Private Function IOrderDetails_getCollectionOfOrders() As Collection
    Set IOrderDetails_getCollectionOfOrders = matchedOrders
End Function
