VERSION 5.00
Begin VB.Form frm_ini 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INI Test"
   ClientHeight    =   2250
   ClientLeft      =   3225
   ClientTop       =   4530
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2250
   ScaleWidth      =   3060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_ContactName 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txt_ContactEmail 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmd_Load 
      Caption         =   "Load Defaults"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmd_Save 
      Caption         =   "Save Defaults"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Name :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Contact Email :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KeySection As String
Dim KeyKey As String
Dim KeyValue As String

Private Sub loadini()

Dim lngResult As Long
Dim strFileName
Dim strResult As String * 50
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = GetPrivateProfileString(KeySection, _
KeyKey, strFileName, strResult, Len(strResult), _
strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
Else
KeyValue = Trim(strResult)
End If

End Sub

Private Sub saveini()

Dim lngResult As Long
Dim strFileName
strFileName = App.Path & "\Myini.ini" 'Declare your ini file !
lngResult = WritePrivateProfileString(KeySection, _
KeyKey, KeyValue, strFileName)
If lngResult = 0 Then
'An error has occurred
Call MsgBox("An error has occurred while calling the API function", vbExclamation)
End If

End Sub
' unfortunately you have to include ALL the above on any form
' on which you want to use this, but for my needs this is
' the most simple ini manipulation thing i found :)

' Load your ini parameters
Private Sub cmd_Load_Click()

'load ContactEmail
KeySection = "Email"
KeyKey = "ContactEmail"
loadini
txt_ContactEmail.Text = KeyValue

'load ContactName
KeySection = "Name"
KeyKey = "ContactName"
loadini
txt_ContactName.Text = KeyValue

End Sub

' Save your ini parameters
Private Sub cmd_Save_Click()

'Save ContactEmail
KeySection = "Email"
KeyKey = "ContactEmail"
KeyValue = txt_ContactEmail.Text
saveini

'Save  ContactName
KeySection = "Name"
KeyKey = "ContactName"
KeyValue = txt_ContactName.Text
saveini

End Sub
