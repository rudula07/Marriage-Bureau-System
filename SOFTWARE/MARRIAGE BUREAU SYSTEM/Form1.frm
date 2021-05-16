VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form F1_welcome 
   Caption         =   "Form1"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   19320
      Top             =   9120
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   9960
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Caption         =   "START"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      MaskColor       =   &H00C0FFC0&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9330
      Left            =   6960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9300
      ScaleWidth      =   6780
      TabIndex        =   0
      Top             =   240
      Width           =   6810
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "MARRIAGE BUREAU MANAGMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   360
         TabIndex        =   2
         Top             =   5640
         Width           =   6015
      End
   End
   Begin VB.Image Image1 
      Height          =   11175
      Left            =   -120
      Picture         =   "Form1.frx":45E7
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   20655
   End
End
Attribute VB_Name = "F1_welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Timer1.Enabled = True
ProgressBar1.Visible = True
ProgressBar1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value = ProgressBar1.Max Then
MsgBox "WELCOME", vbInformation + vbDefaultButton3, "MESSAGE"
F2_descrip.Show
Unload Me
Else
ProgressBar1.Value = ProgressBar1.Value + 2
End If
End Sub
