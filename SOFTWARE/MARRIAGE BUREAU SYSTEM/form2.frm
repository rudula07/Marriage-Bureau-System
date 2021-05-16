VERSION 5.00
Begin VB.Form F2_descrip 
   Caption         =   "Form2"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   Picture         =   "form2.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   4
      Top             =   9240
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "AGREE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   3
      Top             =   8160
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DISAGREE"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4680
      TabIndex        =   2
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DO NOTE THAT ALL THE DETAILS PROVIDED WILL BE ORIGINAL AND AUTHERISED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   855
      Left            =   4680
      TabIndex        =   1
      Top             =   7200
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   11295
      Left            =   -960
      Picture         =   "form2.frx":40EAF
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   21735
   End
End
Attribute VB_Name = "F2_descrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option2 = True Then
F3_login.Show
Else
Option1 = True
Unload Me
End If
End Sub



Private Sub Option2_Click()
If Option2 = True Then
MsgBox "FALSE DETAILS ARE CRIMINAL OFFENCE", vbInformation, "NOTE"
End If
End Sub
