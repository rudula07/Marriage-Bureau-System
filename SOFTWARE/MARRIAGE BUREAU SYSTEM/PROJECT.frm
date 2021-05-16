VERSION 5.00
Begin VB.Form FORM2 
   Caption         =   "Form2"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   Picture         =   "PROJECT.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   15510
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   8520
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "AGREE"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "DISAGREE"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "DO NOTE THAT ALL THE DETAILS PROVIDED WILL BE ORIGINAL AND AUTHERISED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   7440
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   5640
      TabIndex        =   0
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   5640
      Left            =   5640
      Picture         =   "PROJECT.frx":40EAF
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   9105
   End
End
Attribute VB_Name = "FORM2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Option2 = True Then
Form3.Show
Else
Option1 = True
Unload Me
End If
End Sub

Private Sub Option2_Click()
If Option2 = True Then
MsgBox ("FALSE DETAILS ARE CRIMINAL OFFENCE")
End If
End Sub
