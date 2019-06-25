Attribute VB_Name = "TestModule"
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
' OK
'Function parseFromDate(i As Integer)
'
'    ' i - stands for start (1) or end (2)
'    yyyy = "" & Year(Date)
'    mm = "" & Month(Date)
'    dd = "" & Day(Date)
'
'    parseFromDate = "" & yyyy & "-" & mm & "-" & dd & "T00:00"
'
'End Function


Private Sub finalTouchTest()
'
' finalTouchTest Macro
'

'
    Columns("J:AC").Select
    Selection.ColumnWidth = 5.43
    
    
End Sub


Private Sub testOnEmptyDate()
    
    Dim d As Date
    Debug.Print d ' 00:00:00
End Sub


' public sub testOnOrderItemAndTry to get
' OK!
Private Sub testOnOrderItem()
    
    Dim oi As OrderItem
    Set oi = New OrderItem
    
    oi.coforConstructor "123", "456"
    
    Debug.Print oi.checkIsObjectReady() ' should be false
    
    Debug.Print oi.getElement(E_2510_SHIPPER) ' should be 456
End Sub
