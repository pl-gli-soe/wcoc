VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "PSA Login"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5025
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SubmitBtn_Click()
    Hide
    main CStr(Me.TextBoxLogin), CStr(Me.TextBoxPass), CBool(Me.CheckBoxOrderList)
End Sub
