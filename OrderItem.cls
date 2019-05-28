VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' __        __        _    _          ____
' \ \      / /__  ___| | _| |_   _   / ___|_____   _____ _ __ __ _  __ _  ___
'  \ \ /\ / / _ \/ _ \ |/ / | | | | | |   / _ \ \ / / _ \ '__/ _` |/ _` |/ _ \
'   \ V  V /  __/  __/   <| | |_| | | |__| (_) \ V /  __/ | | (_| | (_| |  __/
'    \_/\_/ \___|\___|_|\_\_|\__, |  \____\___/ \_/ \___|_|  \__,_|\__, |\___|
'   ___  _ __    / ___|___  _|___/_ _(_) |                         |___/
'  / _ \| '_ \  | |   / _ \| '__/ _` | | |
' | (_) | | | | | |__| (_) | | | (_| | | |
'  \___/|_| |_|  \____\___/|_|  \__,_|_|_|
'
'
'01010111 01100101 01100101 01101011 01101100 01111001  01000011 01101111 01110110 01100101 01110010 01100001 01100111 01100101
'01101111 01101110  01000011 01101111 01110010 01100001 01101001 01101100
' FORREST SOFTWARE
' Copyright (c) 2016 Mateusz Forrest Milewski
'
' Permission is hereby granted, free of charge,
' to any person obtaining a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation the rights to
' use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software,
' and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
' INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
' IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
' WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' privates for controlling splitted constructor
Private dhIsReady As Boolean
Private coforIsReady As Boolean
Private commonDataIsReady As Boolean

' for suitable data (Item: Me - is component for SuitableData2510)

Private SgrLine As String
Private pn As String
Private qty As Double
Private orderNum As String
Private DN As String

Private S1 As String
Private routing As String
Private d As String

Private status As String
Private UM As String

''
' milestones dates
Private DH As DH_Handler

' cofors
Private sellerCofor As String
Private shipperCofor As String




Private Sub Class_Initialize()
    Set DH = New DH_Handler
    dhIsReady = False
    coforIsReady = False
    commonDataIsReady = False
End Sub

Private Sub Class_Terminate()
    Set DH = Nothing
End Sub

Public Sub dhConstructor(mDH As DH_Handler)


    Set DH = mDH
    dhIsReady = True

End Sub

Public Sub coforConstructor(mSellerCofor As String, mShipperCofor As String)

    sellerCofor = mSellerCofor
    shipperCofor = mShipperCofor
    
    
    coforIsReady = True
 End Sub

Public Sub commonConstructor(mSgrLine As String, mPn As String, mQty As Double, mOrderNum As String, mDn As String, mS1 As String, mRouting As String, mD As String, mS2 As String, mUm As String)

    SgrLine = CStr(mSgrLine)
    pn = CStr(mPn)
    qty = CDbl(mQty)
    orderNum = CStr(mOrderNum)
    DN = CStr(mDn)
    S1 = CStr(mS1)
    routing = CStr(mRouting)
    d = CStr(mD)
    status = CStr(mS2)
    UM = CStr(mUm)
    
    
    commonDataIsReady = True
End Sub

Public Function checkIsObjectReady() As Boolean

    ' this funciton provides yes/no anwer
    ' becuase there is a lot of params from screen 2510
    ' so i split the constructor into a few subs
    ' and every of kind have it s own weight
    checkIsObjectReady = commonDataIsReady And coforIsReady And dhIsReady
    
End Function


Public Function getElement(whichElement As E_2510_ELEMENT) As Variant

    
    
    Select Case whichElement
        Case E_2510_ELEMENT.E_2510_D
            getElement = CStr(d)
        Case E_2510_ELEMENT.E_2510_DH
            Set getElement = DH
        Case E_2510_ELEMENT.E_2510_DN
            getElement = CStr(DN)
        Case E_2510_ELEMENT.E_2510_ORDER_NUMBER
            getElement = CStr(orderNum)
        Case E_2510_ELEMENT.E_2510_PRODUCT
            getElement = CStr(pn)
        Case E_2510_ELEMENT.E_2510_QTY
            getElement = CDbl(qty)
        Case E_2510_ELEMENT.E_2510_ROUTING
            getElement = CStr(routing)
        Case E_2510_ELEMENT.E_2510_S1
            getElement = CStr(S1)
        Case E_2510_ELEMENT.E_2510_S2
            getElement = CStr(status)
        Case E_2510_ELEMENT.E_2510_SELLER
            getElement = CStr(sellerCofor)
        Case E_2510_ELEMENT.E_2510_SGR_LINE
            getElement = CStr(SgrLine)
        Case E_2510_ELEMENT.E_2510_SHIPPER
            getElement = CStr(shipperCofor)
        Case E_2510_ELEMENT.E_2510_UM
            getElement = CStr(UM)
        Case Else
            MsgBox "This element is not exist!"
            End
            
    End Select
    
    
End Function




