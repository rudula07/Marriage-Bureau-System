VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00404040&
      Height          =   10695
      Left            =   0
      ScaleHeight     =   10635
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WELCOME"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   1815
         Left            =   6120
         TabIndex        =   1
         Top             =   6240
         Width           =   7695
      End
      Begin VB.Image Image1 
         Height          =   6960
         Left            =   4440
         Picture         =   "MDIForm1.frx":0000
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   10920
      End
   End
   Begin VB.Menu mnustart 
      Caption         =   "START"
   End
   Begin VB.Menu mnureport 
      Caption         =   "REPORTS"
   End
   Begin VB.Menu mnuabu 
      Caption         =   "ABOUT US"
   End
   Begin VB.Menu mnuclose 
      Caption         =   "CLOSE"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuabu_Click()
F93_aboutus.Show
End Sub

Private Sub mnuclose_Click()
End
End Sub

Private Sub mnureport_Click()
DataReport1.Show
End Sub

Private Sub mnustart_Click()
F1_welcome.Show
End Sub

