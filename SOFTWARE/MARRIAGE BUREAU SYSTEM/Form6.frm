VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form F6_perinfo 
   Caption         =   "Form6"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form6"
   Palette         =   "Form6.frx":0000
   Picture         =   "Form6.frx":631B3
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11880
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400040&
      Caption         =   "FATHERS INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2415
      Left            =   10560
      TabIndex        =   45
      Top             =   5640
      Width           =   8055
      Begin VB.TextBox Text20 
         DataField       =   "Fname"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   48
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox Text19 
         DataField       =   "Fage"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   47
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text18 
         DataField       =   "focc"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   46
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "FATHER NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   51
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "AGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   50
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   49
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00400040&
      Caption         =   "MOTHERS INFORMATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   2295
      Left            =   10560
      TabIndex        =   38
      Top             =   8280
      Width           =   8055
      Begin VB.TextBox Text17 
         DataField       =   "Mocc"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   41
         Top             =   1560
         Width           =   4575
      End
      Begin VB.TextBox Text16 
         DataField       =   "Mage"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text15 
         DataField       =   "Mname"
         DataSource      =   "Personalinfo"
         Height          =   495
         Left            =   2400
         TabIndex        =   39
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "MOTHERS NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   44
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "AGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   42
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.CommandButton addnewcmd 
      Caption         =   "ADD NEW"
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
      Left            =   18840
      TabIndex        =   34
      Top             =   8640
      Width           =   1335
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   15
      Left            =   0
      TabIndex        =   33
      Top             =   10935
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   26
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Personalinfo 
      Height          =   855
      Left            =   18120
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\KANS\Documents\Personalinfo.accdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\KANS\Documents\Personalinfo.accdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdnext 
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
      Left            =   18840
      TabIndex        =   30
      Top             =   9600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400040&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9615
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   9975
      Begin VB.TextBox Text14 
         DataField       =   "Phno"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   37
         Top             =   8760
         Width           =   6015
      End
      Begin VB.TextBox Text13 
         DataField       =   "OCCUPATION"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   32
         Top             =   8160
         Width           =   6135
      End
      Begin VB.TextBox Text12 
         DataField       =   "QUALIFICATION"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   29
         Top             =   7560
         Width           =   6135
      End
      Begin VB.TextBox Text11 
         DataField       =   "STATE"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   28
         Top             =   6960
         Width           =   6135
      End
      Begin VB.TextBox Text10 
         DataField       =   "NATIONALITY"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   27
         Top             =   6360
         Width           =   6135
      End
      Begin VB.TextBox Text9 
         DataField       =   "COUNTRY"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2760
         TabIndex        =   26
         Top             =   5760
         Width           =   6135
      End
      Begin VB.TextBox Text8 
         DataField       =   "CASTE"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   25
         Top             =   5160
         Width           =   6135
      End
      Begin VB.TextBox Text7 
         DataField       =   "RELIGION"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   24
         Top             =   4560
         Width           =   6135
      End
      Begin VB.TextBox Text6 
         DataField       =   "HEIGHT"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   23
         Top             =   3960
         Width           =   6135
      End
      Begin VB.TextBox Text5 
         DataField       =   "WEIGHT"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   22
         Top             =   3360
         Width           =   6135
      End
      Begin VB.TextBox Text4 
         DataField       =   "AGE"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   21
         Top             =   2760
         Width           =   6135
      End
      Begin VB.TextBox Text3 
         DataField       =   "DOB"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   20
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "GENDER"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   19
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         DataField       =   "NAME"
         DataSource      =   "Personalinfo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   18
         Top             =   960
         Width           =   6135
      End
      Begin VB.OptionButton Optfemale 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   13
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton Optmale 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "PHONE NUMBER"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   36
         Top             =   8880
         Width           =   2535
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "COUNTRY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "RELIGION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "AGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "DATE OF BIRTH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "WEIGHT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "GENDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "NAME"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "STATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "NATIONALITY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "CASTE"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "HEIGHT"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "QUALIFIATION"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "OCCUPATION"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   8280
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400040&
      Caption         =   "UPLOAD PHOTO"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4815
      Left            =   12960
      TabIndex        =   1
      Top             =   360
      Width           =   5295
      Begin VB.CommandButton upcmd 
         Caption         =   "UPLOAD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         MaskColor       =   &H00400040&
         TabIndex        =   2
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Left            =   1080
         Stretch         =   -1  'True
         Top             =   960
         Width           =   3135
      End
   End
   Begin VB.CommandButton baccmd 
      Caption         =   "BACK"
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
      Left            =   18840
      TabIndex        =   0
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label14 
      DataField       =   "PHOTO"
      DataSource      =   "Personalinfo"
      Height          =   1095
      Left            =   18480
      TabIndex        =   35
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   16200
      Left            =   0
      Picture         =   "Form6.frx":A4062
      Top             =   -120
      Width           =   28800
   End
End
Attribute VB_Name = "F6_perinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Private Sub Command3_Click()
F6_flyinfo.Show
End Sub

Private Sub addnewcmd_Click()
personalinfo.Recordset.AddNew
Image1.Picture = LoadPicture(Clear)
End Sub

Private Sub baccmd_Click()
Unload Me
End Sub

Private Sub Cmdclear_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
Text7.Text = " "
Text8.Text = " "
Text9.Text = " "
Text10.Text = " "
Text11.Text = " "
Text12.Text = " "
Text13.Text = " "
End Sub

Private Sub nextcmd_Click()
personalinfo.Recordset.MoveNext
Image1.Picture = LoadPicture(Label14.Caption)
End Sub


Private Sub Cmdmove_Click()
personalinfo.Recordset.MoveNext
Image1.Picture = LoadPicture(Label14.Caption)
If personalinfo.Recordset.EOF Then
personalinfo.Recordset.MoveFirst
End If
End Sub

Private Sub cmdnext_Click()
F8_details.Show
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(Label14.Caption)
End Sub

Private Sub Optfemale_Click()
Text2.Text = "FEMALE"
End Sub
Private Sub Optmale_Click()
Text2.Text = "MALE"
End Sub

Private Sub upcmd_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg*.jpg"
str = CommonDialog1.FileName
Image1.Picture = LoadPicture(str)
Label14.Caption = str
End Sub
