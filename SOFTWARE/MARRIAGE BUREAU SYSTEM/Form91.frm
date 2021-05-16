VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form F91_selected 
   BackColor       =   &H00004040&
   Caption         =   "Form1"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10635
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   7215
      Left            =   6000
      TabIndex        =   42
      Top             =   480
      Width           =   8535
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   6615
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   8055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Frame2"
      Height          =   7335
      Left            =   14760
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label Label54 
         Caption         =   " NAME"
         Height          =   375
         Left            =   360
         TabIndex        =   68
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label53 
         Caption         =   " DOB"
         Height          =   375
         Left            =   360
         TabIndex        =   67
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label52 
         Caption         =   " GENDER"
         Height          =   375
         Left            =   360
         TabIndex        =   66
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label51 
         Caption         =   " COUNTRY"
         Height          =   375
         Left            =   360
         TabIndex        =   65
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label50 
         Caption         =   " FNAME"
         Height          =   375
         Left            =   360
         TabIndex        =   64
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label Label49 
         Caption         =   " MNAME"
         Height          =   375
         Left            =   360
         TabIndex        =   63
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label48 
         Caption         =   "OCCUPATION"
         Height          =   375
         Left            =   360
         TabIndex        =   62
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label47 
         Caption         =   " PHONE NUMBER"
         Height          =   375
         Left            =   360
         TabIndex        =   61
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label Label46 
         Caption         =   "AGE"
         Height          =   375
         Left            =   360
         TabIndex        =   60
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label45 
         Caption         =   "WEIGHT"
         Height          =   375
         Left            =   360
         TabIndex        =   59
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label44 
         Caption         =   " HEIGHT"
         Height          =   375
         Left            =   360
         TabIndex        =   58
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label43 
         Caption         =   " RELIGION"
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label25 
         Caption         =   " CASTE"
         Height          =   375
         Left            =   360
         TabIndex        =   56
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Phno"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   41
         Top             =   6480
         Width           =   2055
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "NAME"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "GENDER"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   25
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "DOB"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   24
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "AGE"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   23
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "WEIGHT"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   22
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "HEIGHT"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   21
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "RELIGION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   20
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "CASTE"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   19
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "COUNTRY"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Fname"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Mname"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "OCCUPATION"
         DataSource      =   "Adodc1"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   6000
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc personalinfo 
      Height          =   855
      Left            =   3360
      Top             =   9600
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from Table1 where Gender='MALE' "
      Caption         =   "personalinfo"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label Label24 
         Caption         =   " NAME"
         Height          =   375
         Left            =   360
         TabIndex        =   55
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   " DOB"
         Height          =   375
         Left            =   360
         TabIndex        =   54
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label23 
         Caption         =   " GENDER"
         Height          =   375
         Left            =   360
         TabIndex        =   53
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label42 
         Caption         =   " COUNTRY"
         Height          =   375
         Left            =   360
         TabIndex        =   52
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label41 
         Caption         =   " FNAME"
         Height          =   375
         Left            =   360
         TabIndex        =   51
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label40 
         Caption         =   " MNAME"
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label39 
         Caption         =   "OCCUPATION"
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label38 
         Caption         =   " PHONE NUMBER"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "AGE"
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "WEIGHT"
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label20 
         Caption         =   " HEIGHT"
         Height          =   375
         Left            =   360
         TabIndex        =   45
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label19 
         Caption         =   " RELIGION"
         Height          =   375
         Left            =   360
         TabIndex        =   44
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   " CASTE"
         Height          =   375
         Left            =   360
         TabIndex        =   43
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Phno"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   40
         Top             =   6360
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "OCCUPATION"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Mname"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "Fname"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "COUNTRY"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "CASTE"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   9
         Top             =   3960
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "RELIGION"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "HEIGHT"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "WEIGHT"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "AGE"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "DOB"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "GENDER"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DataField       =   "NAME"
         DataSource      =   "personalinfo"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   3360
      Top             =   8160
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select *  from Table1 where Gender='FEMALE';"
      Caption         =   "adodc1"
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
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   5640
      TabIndex        =   27
      Top             =   9240
      Width           =   9255
      Begin VB.CommandButton Command1 
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
         Height          =   855
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2040
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
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
         Height          =   855
         Left            =   3960
         TabIndex        =   32
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5760
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Show"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7440
         TabIndex        =   30
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   5640
      TabIndex        =   28
      Top             =   7800
      Width           =   9255
      Begin VB.CommandButton Command6 
         Caption         =   " SHOW"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7320
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5640
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
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
         Height          =   855
         Left            =   3840
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         Caption         =   "PREVIOUS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1920
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
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
         Height          =   855
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label14 
      DataField       =   "PHOTO"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   1800
      TabIndex        =   29
      Top             =   8280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      DataField       =   "PHOTO"
      DataSource      =   "personalinfo"
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   22200
      Left            =   0
      Picture         =   "Form91.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   20655
   End
End
Attribute VB_Name = "F91_selected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
personalinfo.Recordset.MovePrevious
Image1.Picture = LoadPicture(Label1.Caption)
If personalinfo.Recordset.BOF Then
personalinfo.Recordset.MoveLast
End If
End Sub

Private Sub Command3_Click()
Frame1.Visible = False
personalinfo.Recordset.MoveNext
Image1.Picture = LoadPicture(Label1.Caption)
If personalinfo.Recordset.EOF Then
personalinfo.Recordset.MoveFirst
End If
End Sub

Private Sub Command4_Click()
F92_thankyou.Show
End Sub


Private Sub Command5_Click()
Frame1.Visible = True
End Sub

Private Sub Command6_Click()
Frame2.Visible = True
End Sub

Private Sub Command7_Click()
F92_thankyou.Show
End Sub

Private Sub Command8_Click()
Frame2.Visible = False
Adodc1.Recordset.MoveNext
Image1.Picture = LoadPicture(Label14.Caption)
If Adodc1.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command9_Click()
Frame2.Visible = False
Adodc1.Recordset.MovePrevious
Image1.Picture = LoadPicture(Label14.Caption)
If Adodc1.Recordset.BOF Then
Adodc1.Recordset.MoveLast
End If
End Sub

Private Sub Form_Load()
If F9_search.Combo1.Text = "MALE" Then
Image1.Picture = LoadPicture(Label1.Caption)
Frame4.Visible = False
Command6.Visible = False
Command7.Visible = False
Command8.Visible = False
Command9.Visible = False
Command10.Visible = False
End If

If F9_search.Combo1.Text = "FEMALE" Then
Image1.Picture = LoadPicture(Label14.Caption)
Frame3.Visible = False
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
End If
Frame1.Visible = False

End Sub

Private Sub Image1_Click()
personalinfo.RecordSource = "select * from Table1 where Photo='" + Label1.Caption + "'"
End Sub

