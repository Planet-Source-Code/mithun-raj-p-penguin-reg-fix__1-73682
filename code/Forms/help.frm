VERSION 5.00
Begin VB.Form help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "help"
   ClientHeight    =   1605
   ClientLeft      =   5100
   ClientTop       =   5235
   ClientWidth     =   5220
   Icon            =   "help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5220
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   $"help.frx":0ECA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

