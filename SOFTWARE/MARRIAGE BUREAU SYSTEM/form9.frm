VERSION 5.00
Begin VB.Form F9_search 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "form9.frx":0000
      Left            =   3000
      List            =   "form9.frx":0002
      TabIndex        =   1
      Text            =   "SELECT"
      Top             =   8760
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Page"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   13560
      TabIndex        =   0
      Top             =   8640
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT YOUR PARTNERS GENDER"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   2055
      Left            =   600
      TabIndex        =   2
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   11040
      Left            =   0
      Picture         =   "form9.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20640
   End
End
Attribute VB_Name = "F9_search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
If Combo1.Text = "MALE" Then
Image1.Picture = LoadPicture("C:\Users\KANS\Desktop\project pics\traditional-4367785.jpg")
Command1.Enabled = True
Else
Image1.Picture = LoadPicture("C:\Users\KANS\Desktop\project pics\jewellery-3971048.jpg")
Command1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
F91_selected.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem "MALE"
Combo1.AddItem "FEMALE"
End Sub
