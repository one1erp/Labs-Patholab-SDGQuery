VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4016B910-CCE8-4B27-95FA-006C7152BC93}#2.16#0"; "MacabiShared.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.UserControl SDGQueryCtrl 
   ClientHeight    =   9615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15120
   KeyPreview      =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15120
   Begin TabDlg.SSTab SSTab 
      Height          =   9375
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "דרישות שנשלפו"
      TabPicture(0)   =   "SDGQueryCtrl.ctx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblRecords"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCount"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblCountSDGTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LstSDG"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "תנאי השליפה"
      TabPicture(1)   =   "SDGQueryCtrl.ctx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FramePhysician"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FramePatient"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "FrameCollectingStation"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "FrameStatus"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "FrameDates"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "FrameExternalReference"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "FrameSdgID"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "FrameAuthorisedBy"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "FrameMalignant"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "frameSnomed"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "frameFreeText"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "CheckCountSDG"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "frmLetterControl"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkShowOldVersions"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Text1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Frame4"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      Begin VB.Frame Frame4 
         Caption         =   "תוצאות חריגות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   7680
         Width           =   2895
         Begin VB.OptionButton optMalignant 
            Alignment       =   1  'Right Justify
            Caption         =   "לא  חריג (פורטל)"
            Height          =   240
            Index           =   2
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton optMalignant 
            Alignment       =   1  'Right Justify
            Caption         =   "סימון למשרד הבריאות"
            Height          =   240
            Index           =   1
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton optMalignant 
            Alignment       =   1  'Right Justify
            Caption         =   "סימון לפורטל"
            Height          =   240
            Index           =   0
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   137
         Text            =   "SDGQueryCtrl.ctx":0038
         Top             =   8040
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox chkShowOldVersions 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "הצג גרסאות ישנות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   135
         Top             =   8880
         Width           =   2175
      End
      Begin VB.Frame frmLetterControl 
         Caption         =   "מכתבי תשובה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   8040
         Width           =   3495
         Begin VB.ComboBox cmbLetterCode 
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optLetterSent 
            Alignment       =   1  'Right Justify
            Caption         =   "לא נשלח"
            Height          =   195
            Index           =   1
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optLetterSent 
            Alignment       =   1  'Right Justify
            Caption         =   "נשלח"
            Height          =   195
            Index           =   0
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLetterCode 
            Caption         =   "קוד מכתב"
            Height          =   255
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CheckBox CheckCountSDG 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "הצגת פרטי הדרישות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   130
         Top             =   8520
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   -74880
         TabIndex        =   124
         Top             =   840
         Width           =   6615
         Begin VB.CommandButton RemoveSelected 
            Height          =   300
            Left            =   120
            Picture         =   "SDGQueryCtrl.ctx":003E
            Style           =   1  'Graphical
            TabIndex        =   128
            ToolTipText     =   "Remove Selected Request(s)"
            Top             =   270
            Width           =   350
         End
         Begin VB.TextBox TxtRequestBarCode 
            Height          =   300
            Left            =   4800
            TabIndex        =   127
            Top             =   255
            Width           =   1695
         End
         Begin VB.CommandButton InvertSelection 
            Height          =   300
            Left            =   600
            Picture         =   "SDGQueryCtrl.ctx":0188
            Style           =   1  'Graphical
            TabIndex        =   126
            ToolTipText     =   "Invert Selection"
            Top             =   270
            Width           =   350
         End
         Begin VB.CommandButton SelectAll 
            Height          =   300
            Left            =   1080
            Picture         =   "SDGQueryCtrl.ctx":050F
            Style           =   1  'Graphical
            TabIndex        =   125
            ToolTipText     =   "Select All Request(s)"
            Top             =   270
            Width           =   350
         End
         Begin VB.Label LblSampleBarCode 
            AutoSize        =   -1  'True
            Caption         =   "Insert With Barcode:"
            Height          =   240
            Left            =   2895
            TabIndex        =   129
            Top             =   285
            Width           =   1785
         End
      End
      Begin VB.Frame frameFreeText 
         Caption         =   "טקסט חופשי במכתב תשובה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   4200
         Width           =   7100
         Begin VB.CommandButton cmdFreeOr 
            Caption         =   "או"
            Height          =   375
            Left            =   6000
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   945
            Width           =   495
         End
         Begin VB.CommandButton cmdFreeAnd 
            Caption         =   "וגם"
            Height          =   375
            Left            =   5400
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   945
            Width           =   495
         End
         Begin VB.TextBox txtFreeText 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   960
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   5175
         End
         Begin VB.Label lblFreeText 
            Alignment       =   1  'Right Justify
            Caption         =   "טקסט לחיפוש:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   360
            Width           =   1425
         End
      End
      Begin VB.Frame frameSnomed 
         Caption         =   "Snomed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   6000
         Width           =   7100
         Begin VB.CheckBox chkSearchSnomedByPatient 
            Alignment       =   1  'Right Justify
            Caption         =   "חפש לפי פציינט"
            Height          =   255
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   1440
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtLastSnomedMdigit 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            TabIndex        =   32
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox txtSnomedT 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   1200
            Width           =   5175
         End
         Begin VB.TextBox txtSnomedM 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   33
            Top             =   780
            Width           =   5175
         End
         Begin VB.Label lblLastSomedMDigit 
            Alignment       =   1  'Right Justify
            Caption         =   "ספרה אחרונה של Snomed M"
            Height          =   360
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   ":Snomed T"
            Height          =   360
            Left            =   5520
            TabIndex        =   107
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label lblSnomedM 
            Alignment       =   1  'Right Justify
            Caption         =   ":Snomed M"
            Height          =   360
            Left            =   5520
            TabIndex        =   106
            Top             =   780
            Width           =   1305
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -68160
         TabIndex        =   94
         Top             =   840
         Width           =   2415
         Begin VB.CommandButton CmdClose 
            Caption         =   "סגור"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdConds 
            Caption         =   "תנאים"
            Height          =   375
            Left            =   1200
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame FrameMalignant 
         Caption         =   "אחר"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   9120
         Visible         =   0   'False
         Width           =   7100
         Begin VB.ComboBox CmbMalignant 
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   300
            Width           =   1695
         End
         Begin VB.ComboBox CmbConsult 
            Height          =   315
            Left            =   3720
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   750
            Width           =   1695
         End
         Begin VB.ComboBox CmbPositive 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CmbQC 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   690
            Width           =   1695
         End
         Begin VB.Label LblMalignant 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ממאיר:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label LblConsult 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "התייעצות:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label LblPositive 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "חיובי:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label LblQC 
            Alignment       =   1  'Right Justify
            Caption         =   "בקרת איכות:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   720
            Width           =   1305
         End
      End
      Begin VB.Frame FrameAuthorisedBy 
         Caption         =   "מאושר ע""י"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   4680
         Width           =   7100
         Begin VB.ComboBox cmbRoles 
            Height          =   360
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   300
            Width           =   2295
         End
         Begin VB.ComboBox CmbAuthorisedBy 
            Height          =   360
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   900
            Width           =   2295
         End
         Begin VB.TextBox TxtAuthOperID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            TabIndex        =   86
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "תפקיד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   300
            Width           =   705
         End
         Begin VB.Label LblAuthorisedBy 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מאשר:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   900
            Width           =   675
         End
         Begin VB.Label LblAuthorisedName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   900
            Width           =   3615
         End
      End
      Begin VB.Frame FrameSdgID 
         Caption         =   "מס. פנימי"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   1500
         Width           =   7100
         Begin VB.TextBox txtSdgYear 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1080
            TabIndex        =   4
            Top             =   300
            Width           =   1215
         End
         Begin VB.ComboBox cmbSdgType 
            Height          =   360
            ItemData        =   "SDGQueryCtrl.ctx":0951
            Left            =   3360
            List            =   "SDGQueryCtrl.ctx":0953
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   300
            Width           =   1815
         End
         Begin VB.TextBox TxtFromSdgID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3360
            TabIndex        =   5
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox TxtToSdgID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1080
            TabIndex        =   6
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label lblSdgYear 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "שנה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   300
            Width           =   465
         End
         Begin VB.Label lblSdgType 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "סוג דגימה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   300
            Width           =   1065
         End
         Begin VB.Label LblSdgID 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מספר פנימי:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   900
            Width           =   1185
         End
         Begin VB.Label LblFromSdgID 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   900
            Width           =   225
         End
         Begin VB.Label LblToSdgID 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   900
            Width           =   375
         End
      End
      Begin VB.Frame FrameExternalReference 
         Caption         =   "מס. דרישה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   480
         Width           =   4215
         Begin VB.TextBox txtExternalReference 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   600
            TabIndex        =   1
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label LblExternalReference 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "מספר דרישה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2450
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame FrameDates 
         Caption         =   "תאריכים"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2480
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   480
         Width           =   7100
         Begin MSMask.MaskEdBox MaskEdBoxFrom 
            Height          =   360
            Index           =   0
            Left            =   2760
            TabIndex        =   21
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxTo 
            Height          =   360
            Index           =   0
            Left            =   480
            TabIndex        =   22
            Top             =   300
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxFrom 
            Height          =   360
            Index           =   1
            Left            =   2760
            TabIndex        =   23
            Top             =   900
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxTo 
            Height          =   360
            Index           =   1
            Left            =   480
            TabIndex        =   24
            Top             =   900
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxFrom 
            Height          =   360
            Index           =   2
            Left            =   2760
            TabIndex        =   25
            Top             =   1500
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxTo 
            Height          =   360
            Index           =   2
            Left            =   480
            TabIndex        =   26
            Top             =   1500
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxFrom 
            Height          =   360
            Index           =   3
            Left            =   2760
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBoxTo 
            Height          =   360
            Index           =   3
            Left            =   480
            TabIndex        =   28
            Top             =   2040
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   390
            Index           =   0
            Left            =   2760
            TabIndex        =   115
            Top             =   300
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   390
            Index           =   1
            Left            =   2760
            TabIndex        =   116
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   390
            Index           =   2
            Left            =   2760
            TabIndex        =   117
            Top             =   1500
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   390
            Index           =   3
            Left            =   2760
            TabIndex        =   118
            Top             =   2040
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   390
            Index           =   0
            Left            =   480
            TabIndex        =   119
            Top             =   300
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   390
            Index           =   1
            Left            =   480
            TabIndex        =   120
            Top             =   900
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   390
            Index           =   2
            Left            =   480
            TabIndex        =   121
            Top             =   1500
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   390
            Index           =   3
            Left            =   480
            TabIndex        =   122
            Top             =   2040
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   688
            _Version        =   393216
            Format          =   116785153
            CurrentDate     =   38616
         End
         Begin VB.Label lblDateTo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblDateFrom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   2040
            Width           =   225
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "תאריך אישור:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   2040
            Width           =   1500
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "תאריך השלמה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1500
            Width           =   1500
         End
         Begin VB.Label lblDateFrom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1500
            Width           =   225
         End
         Begin VB.Label lblDateTo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1500
            Width           =   375
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "תאריך אישור קבלה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   900
            Width           =   2100
         End
         Begin VB.Label lblDateFrom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   900
            Width           =   225
         End
         Begin VB.Label lblDateTo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   900
            Width           =   375
         End
         Begin VB.Label lblDate 
            Alignment       =   1  'Right Justify
            Caption         =   "תאריך קלט:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   300
            Width           =   1500
         End
         Begin VB.Label lblDateFrom 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   300
            Width           =   225
         End
         Begin VB.Label lblDateTo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   2040
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.Frame FrameStatus 
         Caption         =   "סטטוס"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   480
         Width           =   2535
         Begin VB.ComboBox CmbFromStatus 
            Height          =   360
            ItemData        =   "SDGQueryCtrl.ctx":0955
            Left            =   360
            List            =   "SDGQueryCtrl.ctx":0957
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label LblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "סטטוס:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -65640
         TabIndex        =   63
         Top             =   480
         Width           =   5295
         Begin VB.ComboBox cmbReportLetter 
            Height          =   315
            Left            =   120
            TabIndex        =   134
            ToolTipText     =   "יש לבחור את קוד המכתב מהרשימה וללחוץ לדיווח"
            Top             =   660
            Width           =   1455
         End
         Begin VB.CommandButton cmdReportLetter 
            Caption         =   "דיווח מכתב תשובה לפי קוד"
            Height          =   495
            Left            =   120
            TabIndex        =   133
            ToolTipText     =   "יש לבחור את קוד המכתב מהרשימה וללחוץ לדיווח"
            Top             =   180
            Width           =   1455
         End
         Begin VB.CommandButton CmdCopyToClipboard 
            Caption         =   "Copy To Clipboard"
            Height          =   375
            Left            =   1680
            TabIndex        =   123
            Top             =   600
            Width           =   1455
         End
         Begin VB.CommandButton CmdPrintFinalLetter 
            Caption         =   "מכתב תשובה"
            Height          =   375
            Left            =   3960
            TabIndex        =   43
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton CmdPrintReport 
            Caption         =   "דוח"
            Height          =   375
            Left            =   3240
            TabIndex        =   44
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   240
         TabIndex        =   62
         Top             =   8400
         Width           =   3375
         Begin VB.CommandButton CmdClose 
            Caption         =   "סגור"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "נקה"
            Height          =   375
            Left            =   1200
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton CmdFinished 
            Caption         =   "בצע"
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame FrameCollectingStation 
         Caption         =   "תחנות איסוף"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   3120
         Width           =   7100
         Begin VB.TextBox TxtToCollectingStationID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   240
            TabIndex        =   99
            Top             =   540
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TxtToCollectingStation 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   4920
            TabIndex        =   30
            Top             =   540
            Visible         =   0   'False
            Width           =   700
         End
         Begin VB.CommandButton cmdFindToColStat 
            Caption         =   "..."
            Height          =   375
            Left            =   5640
            TabIndex        =   39
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton cmdFindFromColStat 
            Caption         =   "..."
            Height          =   375
            Left            =   5640
            TabIndex        =   38
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox TxtFromCollectingStation 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   4920
            TabIndex        =   29
            Top             =   300
            Width           =   700
         End
         Begin VB.TextBox TxtFromCollectingStationID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   240
            TabIndex        =   59
            Top             =   300
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label LblToColStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "עד:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   540
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label LblToCollectingStationName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   540
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label LblFromColStat 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "מ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   540
            Visible         =   0   'False
            Width           =   225
         End
         Begin VB.Label LblFromCollectingStationName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   300
            Width           =   4695
         End
         Begin VB.Label LblCollectingStation 
            Alignment       =   1  'Right Justify
            Caption         =   "תחנה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame FramePatient 
         Caption         =   "פציאנטים"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   6240
         Width           =   7100
         Begin VB.TextBox txtPatientMaxAge 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   16
            Top             =   770
            Width           =   735
         End
         Begin VB.TextBox txtPatientMinAge 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5040
            TabIndex        =   15
            Top             =   770
            Width           =   735
         End
         Begin VB.CommandButton CmdFindPatient 
            Caption         =   "..."
            Height          =   375
            Left            =   5760
            TabIndex        =   13
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox TxtPatientID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   4440
            TabIndex        =   14
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox TxtPatientIDNum 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            TabIndex        =   55
            Top             =   300
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "גיל מקסימום"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3315
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   790
            Width           =   1245
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "גיל מינימום"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   790
            Width           =   1155
         End
         Begin VB.Label LblPatientName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   300
            Width           =   4215
         End
         Begin VB.Label LblPatient 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "פציאנט:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.Frame FramePhysician 
         Caption         =   "רופאים"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   3120
         Width           =   7100
         Begin VB.TextBox TxtSubmittingPhysician 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   10
            Top             =   900
            Width           =   1215
         End
         Begin VB.TextBox TxtReferringPhysician 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   8
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton CmdFindSub 
            Caption         =   "..."
            Height          =   375
            Left            =   5085
            TabIndex        =   9
            Top             =   900
            Width           =   255
         End
         Begin VB.CommandButton CmdFindRef 
            Caption         =   "..."
            Height          =   375
            Left            =   5085
            TabIndex        =   7
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox TxtRefID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            TabIndex        =   49
            Top             =   300
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox TxtSubID 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   120
            TabIndex        =   48
            Top             =   900
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label LblReferringPhysician 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "רופא מפנה:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label LblSubmittingPhysician 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "רופא מבצע:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   900
            Width           =   1185
         End
         Begin VB.Label LblRefName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   300
            Width           =   3615
         End
         Begin VB.Label LblSubName 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   900
            Width           =   3615
         End
      End
      Begin MSComctlLib.ListView LstSDG 
         Height          =   7455
         Left            =   -74880
         TabIndex        =   42
         Top             =   1680
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   13150
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglstIcons"
         ColHdrIcons     =   "imglstIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label LblCountSDGTitle 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "נבחרו"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -70560
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   480
         Width           =   585
      End
      Begin VB.Label LblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -71520
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   480
         Width           =   675
      End
      Begin VB.Label LblRecords 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "דרישות"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -72480
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imglstIcons 
      Left            =   0
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":0959
            Key             =   "X"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":0DAB
            Key             =   "A"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":11FD
            Key             =   "C"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":164F
            Key             =   "I"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":1AA1
            Key             =   "P"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":1EF3
            Key             =   "R"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":2345
            Key             =   "S"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":2797
            Key             =   "U"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":2BE9
            Key             =   "V"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":303B
            Key             =   "W"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":348D
            Key             =   "UP"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SDGQueryCtrl.ctx":35E7
            Key             =   "DOWN"
         EndProperty
      EndProperty
   End
   Begin MacabiShared.FindPhysicianDlg FindPhysicianDlg 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin MacabiShared.FindPatientDlg FindPatientDlg 
      Height          =   615
      Left            =   0
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.Menu mnuRequest 
      Caption         =   "דרישה"
      Visible         =   0   'False
      Begin VB.Menu mnuResultEntry 
         Caption         =   "הזנת תוצאות"
      End
      Begin VB.Menu mnuModifySDG 
         Caption         =   "עידכון פרטי דרישה"
      End
      Begin VB.Menu mnuSdgCopy 
         Caption         =   "רויזיה"
      End
      Begin VB.Menu mnuendsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserEvents 
         Caption         =   "ארוע משתמש"
         Begin VB.Menu mnuEvent 
            Caption         =   "(Empty)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuUserEventsALL 
         Caption         =   "ארוע לכל הדרישות"
         Begin VB.Menu mnuEventALL 
            Caption         =   "(Empty)"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "SDGQueryCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

'------------------------------------------------------------------------
'
'
' 02/08/2005 Barak : Add mnuEventAll for running events on all sdg's
'    08/2005 Barak : Change to 2 tabs - Query and results
'
'------------------------------------------------------------------------

Option Explicit
 
'Private Declare Function GetTickCount Lib "kernel32" () As Long
 
'הגדרת צבעים גלובליים
Private Const RED = &HFF&
Private Const WHITE = &HFFFFFF

Private Const PRINT_EVENT = "Print PDF Letter"
'the older value for this constant:
'Private Const PRINT_EVENT ="Print Final Letter"

'to hold the selected sdg names
'to be used in reporting the control letter
'Dim SdgNames As New scripting.Dictionary


Private COLUMN_WIDTH As Integer
Private Const TABLE_LEFT = 600 '460
Private TABLE_TOP As Long
 
Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
Private con As ADODB.connection
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
Private WorkFolder As String
Private strColumnHeaders As String
'---------------------------
' GET QUERY SQL
Private DontRunQuery As Boolean
'---------------------------

Private Enum DateField
    dfCreatedOn = 0
    dfReceivedOn = 1
    dfCompletedOn = 2
    dfAuthorisedOn = 3
End Enum

Private Enum SdgType
    stAll = 0
    stB = 1
    stC = 2
    stP = 3
End Enum

Private sql As String

Private SelectClause As String
Private FromClause As String
Private WhereClause As String
Private StatusCodes As scripting.Dictionary
Private Const MaxRecordToDisplay = 15000
Private RClickli As ListItem
Private SdgCopyParameter As LSEXT.LSExtensionParameters
Private NumColumnHeader As Integer
Private PageNumber As Integer

'ranslates the text describing the letter sent
'to the code in the field u_grp_code:
Private dicGrpCode As New Dictionary

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long

Private Declare Function CloseClipboard Lib "User32" () As Long

Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long

Private Declare Function EmptyClipboard Lib "User32" () As Long

Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long

Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MAXSIZE = 4096

'Private Sub chkMalignancy_Click()
'On Error GoTo ERR_chkMalignancy_Click
'
'    chkMalignancyPatient.Visible = (chkMalignancy.Value = 1)
'
'    Exit Sub
'ERR_chkMalignancy_Click:
'MsgBox "ERR_chkMalignancy_Click" & vbCrLf & Err.Description
'End Sub

Private Sub cmdClose_Click(Index As Integer)
  If Not NtlsSite Is Nothing Then Call NtlsSite.CloseWindow
End Sub

Private Sub cmdConds_Click()
    SSTab.Tab = 1
End Sub

Private Sub CmdCopyToClipboard_Click()
    Dim strFileds As String
    Dim i As Integer
    Dim j As Integer
    Dim li As ListItem
    Dim TmpStr As String

    strFileds = ""
    For i = 1 To LstSDG.ListItems.Count
        strFileds = strFileds & vbTab & Trim(LstSDG.ListItems(i).Text)
        Set li = LstSDG.ListItems(i)
        For j = 1 To LstSDG.ColumnHeaders.Count - 1
            TmpStr = Trim(li.SubItems(j))
            strFileds = strFileds & vbTab & TmpStr
        Next j
        strFileds = strFileds & vbCrLf
    Next i

    If Trim(strFileds) <> "" Then
        strFileds = strColumnHeaders & _
                    vbCrLf & _
                    strFileds
        Call ClipBoard_SetData(strFileds)
        MsgBox "The entity information has been successfuly copied to the clipboard.", _
                vbInformation + vbYes, "Nautilus - Copy To Clipboard"
    End If
End Sub

Function ClipBoard_SetData(MyString As String)
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, X As Long

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Function
   End If

   ' Clear the Clipboard.
   X = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If
End Function

Private Sub cmdFindFromColStat_Click()
    'Set frmFindCollectionStation.connection = con
    frmFindCollectionStation.Initialize con
    'frmFindCollectionStation.SetConnection (con)
    frmFindCollectionStation.Show vbModal
    
    If frmFindCollectionStation.GetClinicId <> "" Then
        TxtFromCollectingStation.Text = frmFindCollectionStation.GetClinicId
    '    LblFromCollectingStationName.Caption = frmFindCollectionStation.GetClinicName
    End If
     
    Call RefreshCollectingStation(TxtFromCollectingStation.Text, LblFromCollectingStationName)
End Sub

'Private Sub cmdFindToColStat_Click()
'    Set frmFindCollectionStation.connection = con
'    frmFindCollectionStation.Initialize con
'    'frmFindCollectionStation.SetConnection (con)
'    frmFindCollectionStation.Show vbModal
'
'    If frmFindCollectionStation.GetClinicId <> "" Then
'        TxtToCollectingStation.Text = frmFindCollectionStation.GetClinicId
'        LblToCollectingStationName.Caption = frmFindCollectionStation.GetClinicName
'    End If
'End Sub

Private Sub CmdFinished_Click()
'    SdgNames.RemoveAll
    FillList
'    SSTab.Tab = 0
'    LstSDG.SetFocus
End Sub

Private Sub cmdFreeAnd_Click()
    txtFreeText = txtFreeText & " & "
    txtFreeText.SetFocus
End Sub

Private Sub cmdFreeOr_Click()
    txtFreeText = txtFreeText & " | "
    txtFreeText.SetFocus
End Sub

Private Sub CmdPrintFinalLetter_Click()
    If LstSDG.ListItems.Count = 0 Then
        MsgBox "אין אפשרות להפיק מכתב תשובה - חסרות דרישות", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - שאילתה לדרישה"
        Exit Sub
    End If
    Dim i As Integer

    For i = 1 To LstSDG.ListItems.Count
        Call RunSdgEvent(LstSDG.ListItems(i).Text, PRINT_EVENT)
    Next i
End Sub

Private Sub CmdPrintReport_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strTitle As String
    Dim strLine As String
    Dim li As ListItem
    Dim NextCol As Long
    Dim CurrY As Long
    Dim TmpStr As String

    If LstSDG.ListItems.Count = 0 Then
        MsgBox "אין אפשרות להפיק דוח נתוני דרישה - חסרים דרישות", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - שאילתה לדרישה"
        Exit Sub
    End If

    PageNumber = 1
    TABLE_TOP = 1000
    Call PrintHeader("נ ת ו נ י    ד ר י ש ה")

    'For Detail
    Printer.FontUnderline = False
    Printer.FontSize = 11

    For i = 1 To LstSDG.ListItems.Count
        CurrY = Printer.CurrentY
        NextCol = TABLE_LEFT
        Printer.CurrentX = NextCol + 20
'        Printer.CurrentY = CurrY
        'Printer.Print Left(Trim(LstSDG.ListItems(i).Text), COLUMN_WIDTH)
        PrintStr Left(Trim(LstSDG.ListItems(i).Text), COLUMN_WIDTH)
        If CurrY > Printer.CurrentY Then CurrY = Printer.CurrentY - Printer.TextHeight("C")

        Set li = LstSDG.ListItems(i)

        For j = 1 To NumColumnHeader - 1

            TmpStr = Trim(li.SubItems(j))
            TmpStr = Left(TmpStr, COLUMN_WIDTH)
            If IsHEB(TmpStr) Then
                NextCol = TABLE_LEFT + (j + 1) * Printer.TextWidth(Space(COLUMN_WIDTH + 1)) - IIf(i = NumColumnHeader - 1, 50, 0)
                NextCol = NextCol - Printer.TextWidth(TmpStr) - 20
            Else
                NextCol = TABLE_LEFT + j * Printer.TextWidth(Space(COLUMN_WIDTH + 1))
                NextCol = NextCol + 20
            End If

            Printer.CurrentX = NextCol
            Printer.CurrentY = CurrY
            PrintStr TmpStr
        Next j
    Next i
    Call PrintFooter(True, True)
End Sub

Private Function IsHEB(str As String) As Boolean
    Dim IsRTL As Boolean
    Dim SearchChars As String
    Dim X As Integer

    SearchChars = "אבגדהוזחטיכלמנסעפצקרשת"
    IsRTL = False
    For X = 1 To Len(SearchChars)
        If InStr(1, str, Mid(SearchChars, X, 1)) > 0 Then
            IsRTL = True
        End If
    Next X
    IsHEB = IsRTL
End Function

Private Function SpecialPad(str As String, PadChar As String, PadLen As Integer) As String
    Dim IsRTL As Boolean
    Dim SearchChars As String
    Dim X As Integer

    SearchChars = "אבגדהוזחטיכלמנסעפצקרשת"
    IsRTL = False
    For X = 1 To Len(SearchChars)
        If InStr(1, str, Mid(SearchChars, X, 1)) > 0 Then
            IsRTL = True
        End If
    Next X
    SpecialPad = PadStr(str, PadChar, PadLen, IsRTL)
End Function

Private Function PadStr(str As String, PadChar As String, PadLen As Integer, Optional PadLeft As Boolean = False) As String
    If Not PadLeft Then
        PadStr = Left(str + String(PadLen, PadChar), PadLen)
    Else
        PadStr = Right(String(PadLen, PadChar) + str, PadLen)
    End If
End Function

Private Sub PrintHeader(Str2Print As String)
    Dim i As Integer
    Dim LogoPic As IPictureDisp
    Dim NextCol As Long
    Dim CurrY As Long
    Dim TmpStr As String

    Printer.RightToLeft = False
    Printer.FontName = "Miriam Fixed"
    Printer.FontUnderline = False

    If PageNumber = 1 Then
        'For Title
        Set LogoPic = LoadPicture("C:\Program Files\LabSystems\Nautilus\Resource\macabi.jpg")
        Printer.PaintPicture LogoPic, Printer.Width - 3800, 0, 3000, 600

        Printer.FontBold = True
        Printer.FontSize = 14
        Printer.CurrentY = 500
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(Str2Print)) / 2
        Printer.Print Str2Print
        Printer.Print vbCrLf
    End If

    'For Header
    Printer.FontBold = False
    Printer.FontSize = 11

    NumColumnHeader = 8
    If LstSDG.ColumnHeaders.Count < NumColumnHeader Then
        NumColumnHeader = LstSDG.ColumnHeaders.Count
    End If
    COLUMN_WIDTH = CInt((83 - 1) / NumColumnHeader) - 1

    CurrY = TABLE_TOP + 20
    For i = 0 To NumColumnHeader - 1

        TmpStr = Trim(LstSDG.ColumnHeaders.Item(i + 1).Text)
        TmpStr = Left(TmpStr, COLUMN_WIDTH)
        If IsHEB(TmpStr) Then
            NextCol = TABLE_LEFT + (i + 1) * Printer.TextWidth(Space(COLUMN_WIDTH + 1)) - IIf(i = NumColumnHeader - 1, 50, 0)
            NextCol = NextCol - Printer.TextWidth(TmpStr) - 20
        Else
            NextCol = TABLE_LEFT + i * Printer.TextWidth(Space(COLUMN_WIDTH + 1))
            NextCol = NextCol + 20
        End If

        Printer.CurrentX = NextCol
        Printer.CurrentY = CurrY
        Printer.Print TmpStr
    Next i

    Printer.Line (TABLE_LEFT, TABLE_TOP)-(TABLE_LEFT + NumColumnHeader * (Printer.TextWidth(Space(COLUMN_WIDTH + 1))) - 50, TABLE_TOP)
    Printer.Line (TABLE_LEFT, TABLE_TOP + Printer.TextHeight(" ") + 20)-(TABLE_LEFT + NumColumnHeader * (Printer.TextWidth(Space(COLUMN_WIDTH + 1))) - 50, TABLE_TOP + Printer.TextHeight(" ") + 40), , B
End Sub

Private Sub PrintFooter(Optional PageNo As Boolean = True, Optional EndDocument As Boolean = True)
    Dim StrFooter As String
    Dim ButtomLine As Long
    Dim NextCol As Long
    Dim i As Integer

    ButtomLine = Printer.CurrentY

    For i = 0 To NumColumnHeader
        NextCol = TABLE_LEFT + i * Printer.TextWidth(Space(COLUMN_WIDTH + 1)) - IIf(i = NumColumnHeader, 50, 0)
        Printer.Line (NextCol, TABLE_TOP)-(NextCol, ButtomLine)
    Next i

    Printer.Line (TABLE_LEFT, ButtomLine)-(NextCol, ButtomLine)

    If PageNo Then
        StrFooter = "Page No. " & PageNumber
        Printer.CurrentY = Printer.Height - 1000
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(StrFooter)) / 2
        Printer.Print StrFooter
    End If

    If EndDocument Then
        StrFooter = " סך-הכל: " & LstSDG.ListItems.Count & " רשומות "
        Printer.CurrentY = Printer.Height - 1500
        Printer.CurrentX = (Printer.Width - Printer.TextWidth(StrFooter)) / 2
        Printer.Print StrFooter
        Printer.EndDoc
    Else
        Printer.NewPage
        PageNumber = PageNumber + 1
    End If
End Sub

Private Sub PrintStr(Str2Print As String)
    Dim NextCol As Long
    
    If Printer.CurrentY > Printer.Height - 2000 Then
        Call PrintFooter(True, False)
        TABLE_TOP = 400
        Call PrintHeader("")
        If IsHEB(Str2Print) Then
            NextCol = TABLE_LEFT + Printer.TextWidth(Space(COLUMN_WIDTH + 1))
            NextCol = NextCol - Printer.TextWidth(Str2Print) - 20
        Else
            NextCol = TABLE_LEFT
            NextCol = NextCol + 20
        End If
        Printer.CurrentX = NextCol
    End If
    Printer.Print Str2Print
End Sub

'reports that a letter of some (grp_code) kind
'was sent for this sdg
Private Sub cmdReportLetter_Click()

On Error GoTo ERR_cmdReportLetter_Click
    Dim iUpdatedRecords As Integer
    Dim i As Integer
    Dim li As ListItem
    Dim rstid As ADODB.Recordset
    Dim sdgrec As ADODB.Recordset
    Dim sql As String
    Dim strSdgId As String
    Dim strSdgName As String
    Dim strGrpCode As String

    If cmbReportLetter.Text = "" Then
        MsgBox " נא לבחור תחילה קוד מכתב "
        Exit Sub
    End If
    
    If LstSDG.ListItems.Count = 0 Then
        MsgBox " לא נבחרו רשומות לעדכון "
        Exit Sub
    End If
    
    
    
    strGrpCode = dicGrpCode(CStr(cmbReportLetter.Text))
    
    iUpdatedRecords = 0
    
    For i = 1 To LstSDG.ListItems.Count
        Set li = LstSDG.ListItems(i)
        
        strSdgName = LstSDG.ListItems(i).Text

        'get the sdg id from the list:
        'strSdgName = Trim(CStr(li.SubItems(0)))

        'get the relevant data to insert into
        'u_letter_control, u_letter_control_user
        sql = "select d.sdg_id, d.name dn, r.name rn, i.name imn, p.name pn, " & _
            "du.u_clinic_code rc, du.u_implementing_clinic ic " & _
            "from lims_sys.sdg d, lims_sys.sdg_user du, " & _
            "lims_sys.client p, lims_sys.supplier r, lims_sys.supplier i " & _
            "where d.sdg_id = du.sdg_id and " & _
            "p.client_id(+) = du.u_patient and " & _
            "i.supplier_id(+) = du.u_implementing_physician and " & _
            "r.supplier_id(+) = du.u_referring_physician and " & _
            "d.name = '" & strSdgName & "'"
            
' "d.name = '" & li.ToolTipText & "'"


            
            '"d.name = '" & SdgNames.Item(i) & "'"

        Set sdgrec = con.Execute(sql)

        If sdgrec.EOF Then
            MsgBox "Code: " & strSdgId & " does not exist."
        Else
            'get the sdg name from the recordset:
            strSdgName = Trim(CStr(sdgrec("dn")))

            'get next u_letter_control_id:
            Set rstid = con.Execute("select lims.sq_u_letter_control.nextval from dual")

            'enter the new record into the u_letter_control table:
            sql = ("insert into lims_sys.u_letter_control " & _
                "(u_letter_control_id, name, version, version_status) " & _
                "values (" & rstid(0) & _
                ",'" & strSdgName & " - ' || to_char(sysdate,'dd/mm/yyyy hh24:mi:ss')," & _
                "'1','A')")
            Call con.Execute(sql)

            'enter the new record into the u_letter_control_user table:
            sql = ("insert into lims_sys.u_letter_control_user (u_letter_control_id, " & _
                "u_client, u_refering, u_implimenting, u_refering_clinic, " & _
                "u_implimenting_clinic, u_grp_code, u_sdg_id, u_send_on) " & _
                "values (" & rstid(0) & ", '" & nte(sdgrec("PN")) & "', '" & _
                nte(sdgrec("RN")) & "', '" & nte(sdgrec("IMN")) & "', '" & _
                nte(sdgrec("RC")) & "', '" & nte(sdgrec("IC")) & "', '" & _
                strGrpCode & "', " & sdgrec("SDG_ID") & ", sysdate)")
            Call con.Execute(sql)
             
            iUpdatedRecords = iUpdatedRecords + 1
        
        End If
    Next i
    
    MsgBox " עודכנו " & iUpdatedRecords & " רשומות "
 
    Exit Sub
ERR_cmdReportLetter_Click:
    MsgBox "ERR_cmdReportLetter_Click" & vbCrLf & Err.Description
End Sub

Private Sub dtFrom_CloseUp(Index As Integer)
    MaskEdBoxFrom(Index) = Format(dtFrom(Index).Value, "dd/mm/yyyy")
End Sub

Private Sub dtFrom_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    dtFrom(Index).Value = Now
    If Trim(MaskEdBoxFrom(Index).Text) <> "__/__/____" Then
        dtFrom(Index).Value = Trim(MaskEdBoxFrom(Index).Text)
    End If
End Sub

Private Sub dtTo_CloseUp(Index As Integer)
    MaskEdBoxTo(Index) = Format(dtTo(Index).Value, "dd/mm/yyyy")
End Sub

Private Sub dtTo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    dtTo(Index).Value = Now
    If Trim(MaskEdBoxTo(Index).Text) <> "__/__/____" Then
        dtTo(Index).Value = Trim(MaskEdBoxTo(Index).Text)
    End If
End Sub

Private Function IExtensionWindow_CloseQuery() As Boolean
    'Happens when the user close the window
    IExtensionWindow_CloseQuery = True
End Function

Private Function IExtensionWindow_DataChange() As LSExtensionWindowLib.WindowRefreshType
    IExtensionWindow_DataChange = windowRefreshNone
End Function

Private Function IExtensionWindow_GetButtons() As LSExtensionWindowLib.WindowButtonsType
    IExtensionWindow_GetButtons = windowButtonsNone
End Function

Private Sub IExtensionWindow_Internationalise()
End Sub

Private Sub IExtensionWindow_PreDisplay()
    Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
    Set FindPhysicianDlg.con = con
    Set FindPatientDlg.con = con

    LstSDG.MultiSelect = True
      
    ' GET QUERY SQL
    DontRunQuery = False

    strColumnHeaders = ""
    SetColumnHeader

    SSTab.Tab = 1
End Sub

Private Sub IExtensionWindow_refresh()
    'Code for refreshing the window
End Sub

Private Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)
End Sub

Private Function IExtensionWindow_SaveData() As Boolean
End Function

Private Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)
End Sub

Private Sub IExtensionWindow_SetParameters(ByVal parameters As String)
    Set con = New ADODB.connection
    Dim constr As String
    constr = "Provider=OraOLEDB.Oracle" & _
        ";Data Source=" & NtlsCon.GetServerDetails & _
        ";User ID=" & NtlsCon.GetUsername & _
        ";Password=" & NtlsCon.GetPassword
  
          If NtlsCon.GetServerIsProxy Then
            constr = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If

        
        
        
    con.Open constr
    con.CursorLocation = adUseClient
    con.Execute "SET ROLE LIMS_USER"
    
End Sub

Private Sub IExtensionWindow_SetServiceProvider(ByVal serviceProvider As Object)

    Set sp = serviceProvider
    Set NtlsCon = sp.QueryServiceProvider("DBConnection")
    Set NtlsUser = sp.QueryServiceProvider("User")
    Set ProcessXML = sp.QueryServiceProvider("ProcessXML")

End Sub

Private Sub IExtensionWindow_SetSite(ByVal Site As Object)
    Set NtlsSite = Site
    NtlsSite.SetWindowInternalName ("MacabiSDGQuery")
    NtlsSite.SetWindowRegistryName ("MacabiSDGQuery")
    Call NtlsSite.SetWindowTitle("Request Query")
End Sub

Private Sub IExtensionWindow_Setup()
    Dim rsGrpCode As Recordset
    Dim Status As ADODB.Recordset
    Dim Operators As ADODB.Recordset
    Dim Roles As ADODB.Recordset

    WorkFolder = ""
    WorkFolder = xmlManager.GetDefaultFolderFromWorkStation(NtlsUser.GetWorkstationId, con)
    If Trim(WorkFolder) <> "" Then
        xmlManager.XmlFolder = WorkFolder & "\SDGQuery\"
    End If
    
    'Call imglstIcons.ListImages.Add(, "L1", LoadPicture("Resource\Sdg.ico"))

    'Init the Status combo
    Set Status = con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
        "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
        "name = 'SDG Status') " & _
        "order by order_number")

    CmbFromStatus.List(0) = "All"
    CmbFromStatus.List(1) = "במעבדה"
    Set StatusCodes = New scripting.Dictionary
    Status.MoveFirst
    Do Until Status.EOF
        CmbFromStatus.List(CmbFromStatus.ListCount) = Status("PHRASE_DESCRIPTION")
        Call StatusCodes.Add(CStr(Status("PHRASE_DESCRIPTION").Value), CStr(Status("PHRASE_NAME").Value))
        Status.MoveNext
    Loop

    'Init the Operator combo
        
    sql = "select NAME from lims_sys.OPERATOR order by FULL_NAME"
    
    Set Operators = con.Execute(sql)

    CmbAuthorisedBy.List(0) = "הכל"
    Operators.MoveFirst
    Do Until Operators.EOF
        CmbAuthorisedBy.List(CmbAuthorisedBy.ListCount) = Operators("NAME")
        Operators.MoveNext
    Loop
    Operators.Close
    Set Operators = Nothing
    Set Roles = con.Execute("select NAME from lims_sys.LIMS_ROLE order by name")
    
    cmbRoles.List(0) = "הכל"
    Roles.MoveFirst
    Do Until Roles.EOF
        cmbRoles.List(cmbRoles.ListCount) = Roles("NAME")
        Roles.MoveNext
    Loop
    Roles.Close
    Set Roles = Nothing
    cmbRoles.Text = "הכל"
    
     
    'fill the dictionary of the group code belonging the letters:
    Set rsGrpCode = con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
        "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
        "name = 'Letter Control Grp') " & _
        "order by phrase_name")
        
        
    While Not rsGrpCode.EOF
        Call dicGrpCode.Add(CStr(rsGrpCode("phrase_description").Value), _
                            CStr(rsGrpCode("phrase_name").Value))
        
        rsGrpCode.MoveNext
    Wend
    
    Call InitGrpCodeList

    CmbMalignant.List(0) = "הכל"
    CmbMalignant.List(1) = "כן"
    CmbMalignant.List(2) = "לא"
    CmbMalignant.Text = "הכל"

    CmbConsult.List(0) = "הכל"
    CmbConsult.List(1) = "כן"
    CmbConsult.List(2) = "לא"
    CmbConsult.Text = "הכל"

    CmbPositive.List(0) = "הכל"
    CmbPositive.List(1) = "כן"
    CmbPositive.List(2) = "לא"
    CmbPositive.Text = "הכל"

    CmbQC.List(0) = "הכל"
    CmbQC.List(1) = "כן"
    CmbQC.List(2) = "לא"
    CmbQC.Text = "הכל"

    cmbSdgType.List(stAll) = "הכל"
    cmbSdgType.List(stB) = "היסטולוגיה (B)"
    cmbSdgType.List(stC) = "ציטולוגיה (C)"
    cmbSdgType.List(stP) = "פאפ (P)"
    cmbSdgType.Text = "הכל"

    CmbFromStatus.Text = "All"

    CmbAuthorisedBy.Text = "הכל"
    LblCount.Caption = ""
        
    SSTab.Tab = 1
    
    GetSelectClause
    GetFromClause
    GetWhereClause
End Sub

Private Function IExtensionWindow_ViewRefresh() As LSExtensionWindowLib.WindowRefreshType
    IExtensionWindow_ViewRefresh = windowRefreshNone
End Function

Private Sub ConnectSameSession(ByVal aSessionID)
    Dim aProc As New ADODB.Command
    Dim aSession As New ADODB.Parameter
    
    aProc.ActiveConnection = con
    aProc.CommandText = "lims.lims_env.connect_same_session"
    aProc.CommandType = adCmdStoredProc

    aSession.Type = adDouble
    aSession.Direction = adParamInput
    aSession.Value = aSessionID
    aProc.parameters.Append aSession

    aProc.Execute
    Set aSession = Nothing
    Set aProc = Nothing
End Sub

Private Sub IExtensionWindow2_Close()
'MsgBox "IExtensionWindow2_Close"
End Sub

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

Private Sub InvertSelection_Click()
    On Error GoTo ErrEnd
    Dim i As Long

    If LstSDG.ListItems.Count = 0 Then Exit Sub

    For i = 1 To LstSDG.ListItems.Count
        LstSDG.ListItems(i).Selected = Not LstSDG.ListItems(i).Selected
    Next i
    LstSDG.SetFocus
    Exit Sub

ErrEnd:
    MsgBox "InvertSelection... " & vbCrLf & _
            Err.Description
End Sub





Private Sub LstSDG_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Static sPrevColHdr As MSComctlLib.ColumnHeader

    LstSDG.Sorted = True
    LstSDG.SortKey = ColumnHeader.Index - 1

    If sPrevColHdr Is Nothing Then
        LstSDG.SortOrder = lvwAscending
        ColumnHeader.Icon = "UP"
    ElseIf sPrevColHdr = ColumnHeader Then
        If ColumnHeader.Icon = "UP" Then
            LstSDG.SortOrder = lvwDescending
            ColumnHeader.Icon = "DOWN"
        Else
            LstSDG.SortOrder = lvwAscending
            ColumnHeader.Icon = "UP"
        End If
    Else
        LstSDG.SortOrder = lvwAscending
        ColumnHeader.Icon = "UP"
        sPrevColHdr.Icon = Empty
    End If
    Set sPrevColHdr = ColumnHeader
End Sub

Private Sub LstSDG_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> vbRightButton Then Exit Sub

    Set RClickli = LstSDG.HitTest(X, Y)
    If RClickli Is Nothing Then Exit Sub

    Call InitEventsMenu
    mnuSdgCopy.Visible = False
    PopupMenu mnuRequest
    mnuRequest.Visible = True
    Call UnloadEventMenu
End Sub

Private Sub MaskEdBoxFrom_LostFocus(Index As Integer)
    If CheckDate(MaskEdBoxFrom(Index)) And MaskEdBoxFrom(Index) <> "__/__/____" Then
        dtFrom(Index).Value = MaskEdBoxFrom(Index).Text
    End If
End Sub

Private Sub MaskEdBoxTo_LostFocus(Index As Integer)
    If CheckDate(MaskEdBoxTo(Index)) And MaskEdBoxTo(Index) <> "__/__/____" Then
        dtTo(Index).Value = MaskEdBoxTo(Index).Text
        If MaskEdBoxFrom(Index).Text > MaskEdBoxTo(Index).Text Then
            MaskEdBoxFrom(Index).BackColor = RED
            MsgBox "מתאריך קלט חייב להיות קטן או שווה ל עד תאריך קלט !", vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical + vbOKOnly, "קלט שגוי"
            MaskEdBoxFrom(Index).BackColor = WHITE
        End If
    End If
End Sub

Private Sub mnuEvent_Click(Index As Integer)
    Call RunSdgEvent(RClickli.Text, mnuEvent(Index).Caption)
End Sub

Private Sub mnuEventALL_Click(Index As Integer)
    Dim i As Integer

    For i = 1 To LstSDG.ListItems.Count
        Call RunSdgEvent(LstSDG.ListItems(i).Text, mnuEventALL(Index).Caption)
    Next i
End Sub

Private Sub RunSdgEvent(SdgName As String, EventName As String)
    Dim doc As New DOMDocument
    Dim res As New DOMDocument
    Dim xmlLogin As IXMLDOMElement
    Dim xmlRec As IXMLDOMElement
    Dim e As IXMLDOMElement
    Dim element As IXMLDOMElement
    Dim FileName As String

    Set e = doc.createElement("lims-request")
    Call doc.appendChild(e)
    Set xmlLogin = doc.createElement("login-request")
    Call e.appendChild(xmlLogin)
    Set xmlRec = doc.createElement("SDG")
    Call xmlLogin.appendChild(xmlRec)
    Set element = doc.createElement("find-by-name")
    element.Text = Trim(SdgName)
    Call xmlRec.appendChild(element)
    Set element = doc.createElement("fire-event")
    element.Text = Trim(EventName)
    Call xmlRec.appendChild(element)
    
    If Trim(WorkFolder) <> "" Then
        FileName = "SDGQuery_" & EventName & "_" & SdgName & "_DOC"
        Call xmlManager.SaveXmlFile(doc, FileName)
    End If
    
    Call ProcessXML.ProcessXMLWithResponse(doc, res)
    
    If Trim(WorkFolder) <> "" Then
        FileName = "SDGQuery_" & EventName & "_" & SdgName & "_RES"
        Call xmlManager.SaveXmlFile(res, FileName)
    End If

End Sub

Private Sub mnuModifySDG_Click()
    Dim RequestNBR As String
    Dim li As ListItem
    Dim strParameters As String
    Dim ParamRec As ADODB.Recordset
    Dim frs As New FrmModifySDG

    RequestNBR = RClickli.Text
    If RequestNBR <> "" Then

        Load frs
        frs.ModifySDGCtrl.RunFromWindow = True
        Call frs.ModifySDGCtrl.IExtensionWindow_SetSite(NtlsSite)

        Call frs.ModifySDGCtrl.IExtensionWindow_SetServiceProvider(sp)
        frs.ModifySDGCtrl.IExtensionWindow_Internationalise

        Set ParamRec = con.Execute("select PARAMETER2 from lims_sys.command " & _
            "where name = 'Modify Sdg'")
    
        If Not ParamRec.EOF Then
            strParameters = Trim(nte(ParamRec("PARAMETER2")))
            Call frs.ModifySDGCtrl.IExtensionWindow_SetParameters(strParameters)
        End If

        ParamRec.Close

'        Call FrmModifySDG.ModifySDGCtrl.IExtensionWindow_RestoreSettings(hKey)

        frs.ModifySDGCtrl.IExtensionWindow_PreDisplay
        frs.ModifySDGCtrl.IExtensionWindow_GetButtons
        frs.ModifySDGCtrl.IExtensionWindow_Setup

'        FrmModifySDG.ModifySDGCtrl.IExtensionWindow_DataChange
'        FrmModifySDG.ModifySDGCtrl.IExtensionWindow_refresh
'        FrmModifySDG.ModifySDGCtrl.IExtensionWindow_SaveData
'        FrmModifySDG.ModifySDGCtrl.IExtensionWindow_ViewRefresh

        frs.ModifySDGCtrl.InitiateSdg (RequestNBR)
        frs.Show vbModal

        frs.ModifySDGCtrl.IExtensionWindow_CloseQuery
        Unload frs
        Set frs = Nothing

    End If
End Sub

Private Sub mnuResultEntry_Click()
    Dim RequestNBR As String
    Dim strParameters As String
    Dim ParamRec As ADODB.Recordset
    Dim frs As New FrmResultEntry

    RequestNBR = RClickli.Text
    If RequestNBR <> "" Then

        Load frs
        frs.ResultEntryCtrl.RunFromWindow = True
        Call frs.ResultEntryCtrl.IExtensionWindow_SetSite(NtlsSite)

        Call frs.ResultEntryCtrl.IExtensionWindow_SetServiceProvider(sp)
        frs.ResultEntryCtrl.IExtensionWindow_Internationalise

        Set ParamRec = con.Execute("select PARAMETER2 from lims_sys.command " & _
            "where name = 'Result Entry'")

        If Not ParamRec.EOF Then
            strParameters = Trim(nte(ParamRec("PARAMETER2")))
            Call frs.ResultEntryCtrl.IExtensionWindow_SetParameters(strParameters)
        End If
        ParamRec.Close

'        Call FrmResultEntry.ResultEntryCtrl.IExtensionWindow_RestoreSettings(hKey)

        frs.ResultEntryCtrl.IExtensionWindow_PreDisplay
        frs.ResultEntryCtrl.IExtensionWindow_GetButtons
        frs.ResultEntryCtrl.IExtensionWindow_Setup

'        FrmResultEntry.ResultEntryCtrl.IExtensionWindow_DataChange
'        FrmResultEntry.ResultEntryCtrl.IExtensionWindow_refresh
'        FrmResultEntry.ResultEntryCtrl.IExtensionWindow_SaveData
'        FrmResultEntry.ResultEntryCtrl.IExtensionWindow_ViewRefresh

        frs.ResultEntryCtrl.InitiateSdg (RequestNBR)
        frs.Show vbModal

        frs.ResultEntryCtrl.IExtensionWindow_CloseQuery
        Unload frs
        Set frs = Nothing

    End If
End Sub

Private Sub InitEventsMenu()
    Dim RequestNBR As String
    Dim EventsRec As ADODB.Recordset
    Dim strSql As String
    Dim i As Integer

    RequestNBR = RClickli.Text
    If RequestNBR <> "" Then
        strSql = "select wn.NAME " & _
                 "from lims_sys.sdg d, lims_sys.workflow_node wn " & _
                 "where d.name = '" & RequestNBR & "' and " & _
                 "d.WORKFLOW_NODE_ID = wn.PARENT_ID and wn.EVENTS = 'Q' " & _
                 "order by wn.ORDER_NUMBER"

        Set EventsRec = con.Execute(strSql)
        If Not EventsRec.EOF Then
            EventsRec.MoveFirst
        End If

        i = 1
        While Not EventsRec.EOF
            If i = 1 Then
                mnuEvent(1).Visible = True
                mnuEvent(1).Caption = Trim(nte(EventsRec("NAME")))
                mnuEventALL(1).Visible = True
                mnuEventALL(1).Caption = Trim(nte(EventsRec("NAME")))
            Else
                Load mnuEvent(i)
                mnuEvent(i).Caption = Trim(nte(EventsRec("NAME")))
                Load mnuEventALL(i)
                mnuEventALL(i).Caption = Trim(nte(EventsRec("NAME")))
            End If
            i = i + 1
            EventsRec.MoveNext
        Wend
    End If
End Sub

Private Sub UnloadEventMenu()
    Dim i As Integer

    For i = 2 To mnuEvent.Count
        Unload mnuEvent(i)
        Unload mnuEventALL(i)
    Next i
End Sub

Private Sub mnuSdgCopy_Click()
    Dim RequestNBR As String

    RequestNBR = RClickli.Text
    If RequestNBR <> "" Then
    End If
End Sub

Private Sub optLetterSent_Click(Index As Integer)
'    If Index = 1 Then
'        lblLetterCode.Visible = False
'        cmbLetterCode.Visible = False
'    Else
'        lblLetterCode.Visible = True
'        cmbLetterCode.Visible = True
'    End If
End Sub


Private Sub RemoveSelected_Click()
    On Error GoTo ErrEnd
    Dim oneElement As Boolean
    Dim i As Long
    Dim RequestCount As Long

    If LstSDG.ListItems.Count = 0 Then Exit Sub

    oneElement = True
    For i = LstSDG.ListItems.Count To 1 Step -1
        If LstSDG.ListItems(i).Selected Then
            LstSDG.ListItems.Remove (i)
            oneElement = False
        End If
    Next i
    If oneElement And Not (LstSDG.SelectedItem Is Nothing) Then
        LstSDG.ListItems.Remove (LstSDG.SelectedItem.Index)
    End If
    RequestCount = LstSDG.ListItems.Count
    LblCount.Caption = RequestCount
    LblCount.Refresh
    LstSDG.SetFocus
    Exit Sub

ErrEnd:
    MsgBox "RemoveSelected... " & vbCrLf & _
            Err.Description
End Sub

Private Sub SelectAll_Click()
    On Error GoTo ErrEnd
    Dim i As Long

    If LstSDG.ListItems.Count = 0 Then Exit Sub

    For i = 1 To LstSDG.ListItems.Count
        LstSDG.ListItems(i).Selected = True
    Next i
    LstSDG.SetFocus
    Exit Sub

ErrEnd:
    MsgBox "SelectAll... " & vbCrLf & _
            Err.Description
End Sub

Private Sub txtExternalReference_LostFocus()
    txtExternalReference = UCase(txtExternalReference)
End Sub

Private Sub txtFreeText_GotFocus()
    SendKeys "{END}"
End Sub

Private Sub TxtFromSdgID_LostFocus()
    TxtFromSdgID = Trim(TxtFromSdgID)
    
    If TxtFromSdgID = "" Then Exit Sub
    
    Select Case Left(TxtFromSdgID, 1)
    Case "b", "B"
        cmbSdgType.ListIndex = stB
        TxtFromSdgID = Mid(TxtFromSdgID, 2)
    Case "c", "C"
        cmbSdgType.ListIndex = stC
        TxtFromSdgID = Mid(TxtFromSdgID, 2)
    Case "p", "P"
        cmbSdgType.ListIndex = stP
        TxtFromSdgID = Mid(TxtFromSdgID, 2)
    Case Else
    End Select
    
    If Left(Right(TxtFromSdgID, 3), 1) = "/" Then
        txtSdgYear = Right(TxtFromSdgID, 2)
        TxtFromSdgID = Left(TxtFromSdgID, Len(TxtFromSdgID) - 3)
    End If
    TxtFromSdgID = PadStr(TxtFromSdgID, "0", 6, True)

End Sub

Private Sub TxtPatientID_KeyUp(KeyCode As Integer, Shift As Integer)
    TxtPatientIDNum = GetPatientID(TxtPatientID.Text)
    Call RefreshPatient(TxtPatientID, TxtPatientIDNum.Text, LblPatientName)
End Sub

Private Sub TxtPatientID_LostFocus()
    If Trim(TxtPatientID.Text) <> "" Then
        If TxtPatientIDNum = -1 Then
            LblPatientName = "לתשומת לבך: פציאנט לא במאגר."
        End If
    Else
        LblPatientName.Caption = ""
    End If
End Sub

Private Sub TxtReferringPhysician_KeyUp(KeyCode As Integer, Shift As Integer)
    TxtRefID = GetPhysicianID(TxtReferringPhysician.Text)
    Call RefreshPhysician(TxtReferringPhysician, TxtRefID.Text, LblRefName)
End Sub

Private Sub TxtReferringPhysician_LostFocus()
    If Trim(TxtReferringPhysician.Text) <> "" Then
        If TxtRefID = -1 Then
            LblRefName = "לתשומת לבך: רופא לא במאגר."
        End If
    Else
        LblRefName.Caption = ""
    End If
End Sub

Private Sub TxtRequestBarCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyCode = vbKeyReturn Then Exit Sub

    If Trim(TxtRequestBarCode.Text) = "" Then
        TxtRequestBarCode.BackColor = RED
        MsgBox "You must barcode request !", vbCritical, "Nautilus - Request Query"
        TxtRequestBarCode.BackColor = WHITE
        Call TxtRequestBarCode.SetFocus
        Exit Sub
    End If
    FillOneRequest
    Call TxtRequestBarCode.SetFocus
End Sub

Private Sub txtSdgYear_LostFocus()
    If Len(txtSdgYear) > 2 Then
        txtSdgYear = Right(txtSdgYear, 2)
    End If
End Sub

Private Sub TxtSubmittingPhysician_KeyUp(KeyCode As Integer, Shift As Integer)
    TxtSubID = GetPhysicianID(TxtSubmittingPhysician.Text)
    Call RefreshPhysician(TxtSubmittingPhysician, TxtSubID.Text, LblSubName)
End Sub

Private Sub TxtSubmittingPhysician_LostFocus()
    If Trim(TxtSubmittingPhysician.Text) <> "" Then
        If TxtSubID = -1 Then
            LblSubName = "לתשומת לבך: רופא לא במאגר."
        End If
    Else
        LblSubName.Caption = ""
    End If
End Sub

Private Sub TxtFromCollectingStation_Change()
    Call RefreshCollectingStation(TxtFromCollectingStation.Text, LblFromCollectingStationName)
End Sub

Private Sub TxtRefID_Change()
    Call RefreshPhysician(TxtReferringPhysician, TxtRefID.Text, LblRefName)
End Sub

Private Sub TxtSubID_Change()
    Call RefreshPhysician(TxtSubmittingPhysician, TxtSubID.Text, LblSubName)
End Sub

Private Function GetPhysicianID(pstLicenseNbr As String) As Long

    Dim lRet As Long
    Dim rsPhys As ADODB.Recordset
    
    On Error GoTo NoPhysician
    
    sql = "SELECT supplier_id FROM lims_sys.supplier_user su" & vbCrLf & _
          "WHERE su.U_LICENSE_NBR = '" & pstLicenseNbr & "' "
    
    Set rsPhys = con.Execute(sql)
    
    If Not rsPhys.EOF Then
        lRet = rsPhys("supplier_id")
    Else
        lRet = -1
    End If
    
    rsPhys.Close
    Set rsPhys = Nothing
    
    GetPhysicianID = lRet
    
Exit Function
NoPhysician:
    GetPhysicianID = -1
End Function

Private Sub RefreshPhysician(TxtRef As TextBox, TxtID As String, LblObj As Label)
    On Error GoTo NoPhys
    Dim Referring As ADODB.Recordset
    Dim NameStr As String

    Set Referring = con.Execute("select U_FIRST_NAME, U_LAST_NAME, U_LICENSE_NBR from lims_sys.supplier_user " & _
            "where supplier_user.supplier_id = " & TxtID)
    If Not Referring.EOF Then
        NameStr = nte(Referring("U_FIRST_NAME")) & " " & nte(Referring("U_LAST_NAME"))
        LblObj.Caption = NameStr
        TxtRef.Text = nte(Referring("U_LICENSE_NBR"))
    Else
        LblObj.Caption = ""
    End If
NoPhys:
End Sub

Private Sub RefreshCollectingStation(TxtColeectingSt As String, LabelObj As Label)
    On Error GoTo NoCollectingSt
    Dim Collection As ADODB.Recordset
    Dim CollectingNameStr As String

    Set Collection = con.Execute("select * from lims_sys.u_clinic, lims_sys.u_clinic_user, lims_sys.address " & _
        "where u_clinic.u_clinic_id = u_clinic_user.u_clinic_id and " & _
        "ADDRESS_TABLE_NAME(+) = 'U_CLINIC' and " & _
        "ADDRESS_ITEM_ID(+) = u_clinic.u_clinic_id and " & _
        "ADDRESS_LINE_1(+) = u_clinic.name and " & _
        "u_clinic.name = '" & TxtColeectingSt & "'")
    If Not Collection.EOF Then
'        CollectingNameStr = Trim(nte(Collection("NAME"))) & " - " & Trim(nte(Collection("U_CLINIC_NAME"))) & " - " & Trim(nte(Collection("PHONE")))
        CollectingNameStr = Trim(nte(Collection("U_CLINIC_NAME"))) & " - " & Trim(nte(Collection("PHONE")))
        LabelObj.Caption = CollectingNameStr
        TxtFromCollectingStationID.Text = Trim(nte(Collection("U_CLINIC_ID")))
    Else
        LabelObj.Caption = ""
        TxtFromCollectingStationID.Text = ""
    End If
NoCollectingSt:
End Sub

Private Sub RefreshClinic(TxtClinic, LblObj As Label)
    On Error GoTo NoClinic
    Dim clinic As ADODB.Recordset
    Dim ClinicStr As String

    Set clinic = con.Execute("select * from lims_sys.address " & _
            "where address_table_name = 'SUPPLIER' and " & _
            "address_line_1 <> '0' order by ADDRESS_LINE_1")
    If Not clinic.EOF Then
        ClinicStr = Trim(nte(clinic("ADDRESS_LINE_1"))) & " - " & _
        Trim(clinic("ADDRESS_LINE_2")) & " - " & _
        Trim(clinic("ADDRESS_LINE_3"))
        LblObj.Caption = ClinicStr
    Else
        LblObj.Caption = ""
    End If
NoClinic:
End Sub

Private Function GetPatientID(pstPatientName As String) As Long
    
    Dim lRet As Long
    Dim rsClt As ADODB.Recordset
    
    On Error GoTo NoPatient
    
    sql = "SELECT client_id FROM lims_sys.client " & vbCrLf & _
          "WHERE name = '" & pstPatientName & "' "
    Set rsClt = con.Execute(sql)
    
    If Not rsClt.EOF Then
        lRet = rsClt("client_id")
    Else
        lRet = -1
    End If
    
    rsClt.Close
    Set rsClt = Nothing
    
    GetPatientID = lRet
    
Exit Function
NoPatient:
    GetPatientID = -1
End Function

Private Sub RefreshPatient(TxtPatientID As TextBox, TxtPatient As String, LblObj As Label)
    On Error GoTo NoPatient
    Dim client As ADODB.Recordset
    Dim ClientStr As String

    Set client = con.Execute("select * from lims_sys.client_user, lims_sys.client " & _
            "where client_user.client_id = client.client_id and client.client_id = '" & TxtPatient & "'")
    If Not client.EOF Then
        ClientStr = Trim(client("U_LAST_NAME")) & " - " & Trim(client("U_FIRST_NAME"))
        LblObj.Caption = ClientStr
        TxtPatientID.Text = Trim(client("NAME"))
    Else
        LblObj.Caption = ""
    End If
NoPatient:
End Sub

Private Sub cmdClear_Click()
    LstSDG.ListItems.Clear

    ' Dates
    MaskEdBoxFrom(dfCreatedOn).Text = "__/__/____"
    MaskEdBoxTo(dfCreatedOn).Text = "__/__/____"
    MaskEdBoxFrom(dfReceivedOn).Text = "__/__/____"
    MaskEdBoxTo(dfReceivedOn).Text = "__/__/____"
    MaskEdBoxFrom(dfCompletedOn).Text = "__/__/____"
    MaskEdBoxTo(dfCompletedOn).Text = "__/__/____"
    MaskEdBoxFrom(dfAuthorisedOn).Text = "__/__/____"
    MaskEdBoxTo(dfAuthorisedOn).Text = "__/__/____"

    ' External Reference
    txtExternalReference = ""
    
    ' Status
    CmbFromStatus.ListIndex = 0
    
    ' Sdg
    cmbSdgType.ListIndex = 0
    txtSdgYear = ""
    TxtFromSdgID = ""
    TxtToSdgID = ""

    ' Doctors
    TxtReferringPhysician = ""
    LblRefName.Caption = ""
    TxtRefID = ""

    TxtSubmittingPhysician = ""
    LblSubName.Caption = ""
    TxtSubID = ""
    
    ' Collection stations
    TxtFromCollectingStation = ""
    LblFromCollectingStationName.Caption = ""
    TxtFromCollectingStationID = ""
    
    TxtToCollectingStation = ""
    LblToCollectingStationName.Caption = ""
    TxtToCollectingStationID = ""

    ' Autorization
    cmbRoles.ListIndex = 0
    CmbAuthorisedBy.ListIndex = 0
    LblAuthorisedName.Caption = ""
    TxtAuthOperID = ""

    ' Patient
    TxtPatientID = ""
    LblPatientName.Caption = ""
    TxtPatientIDNum = ""

    ' showing old versions
    chkShowOldVersions.Value = False


    'malignancy
    optMalignant(0).Value = False
    optMalignant(1).Value = False
    optMalignant(2).Value = False
'    chkMalignancy.Value = 0
'    chkMalignancyPatient.Value = 0

    ' Snomed
    txtSnomedM = ""
    txtSnomedT = ""
    txtLastSnomedMdigit = ""

    ' Free Text
    txtFreeText = ""
    
    'letter control
    optLetterSent(0).Value = False
    optLetterSent(1).Value = False
    InitGrpCodeList
    
    ' Other
    CmbMalignant.ListIndex = 0
    CmbConsult.ListIndex = 0
    CmbPositive.ListIndex = 0
    CmbQC.ListIndex = 0

    LblCount.Caption = ""
    strColumnHeaders = ""
End Sub


Private Sub CmdFindPatient_Click()
    FindPatientDlg.ShowDlg
    TxtPatientIDNum.Text = FindPatientDlg.PatientID
    Call RefreshPatient(TxtPatientID, TxtPatientIDNum.Text, LblPatientName)
End Sub

Private Sub CmdFindRef_Click()
    FindPhysicianDlg.ShowDlg
    TxtRefID.Text = FindPhysicianDlg.SupplierID
    Call RefreshPhysician(TxtReferringPhysician, TxtRefID.Text, LblRefName)
End Sub

Private Sub CmdFindSub_Click()
    FindPhysicianDlg.ShowDlg
    TxtSubID.Text = FindPhysicianDlg.SupplierID
    Call RefreshPhysician(TxtSubmittingPhysician, TxtSubID.Text, LblSubName)
End Sub

'Private Sub CmdNext_Click()
'    SSTab.Tab = SSTab.Tab + 1
'    If SSTab.Tab = 9 Then
'        CmdNext.Enabled = False
'    Else
'        CmdNext.Enabled = True
'    End If
'End Sub

'Private Sub SSTab_Click(PreviousTab As Integer)
'    If SSTab.Tab = 0 Then
'        CmdBack.Enabled = False
'        CmdNext.Enabled = True
'    ElseIf SSTab.Tab = 9 Then
'        CmdBack.Enabled = True
'        CmdNext.Enabled = False
'    Else
'        CmdBack.Enabled = True
'        CmdNext.Enabled = True
'    End If
'End Sub

Private Sub cmbRoles_Change()
    RefreshOperators
End Sub

Private Sub cmbRoles_Click()
    RefreshOperators
End Sub

Private Sub cmbRoles_KeyDown(KeyCode As Integer, Shift As Integer)
    RefreshOperators
End Sub

Private Sub cmbRoles_KeyUp(KeyCode As Integer, Shift As Integer)
    RefreshOperators
End Sub

Private Sub cmbRoles_Scroll()
'    RefreshOperators
End Sub

Private Sub CmbAuthorisedBy_Change()
    RefreshOperatorName
End Sub

Private Sub CmbAuthorisedBy_Click()
    RefreshOperatorName
End Sub

Private Sub CmbAuthorisedBy_KeyDown(KeyCode As Integer, Shift As Integer)
    RefreshOperatorName
End Sub

Private Sub CmbAuthorisedBy_KeyUp(KeyCode As Integer, Shift As Integer)
    RefreshOperatorName
End Sub

Private Sub CmbAuthorisedBy_Scroll()
'    RefreshOperatorName
End Sub

Private Sub RefreshOperatorName()
    On Error GoTo NoOperator
    Dim Operator As ADODB.Recordset

    If CmbAuthorisedBy.Text = "הכל" Then
        LblAuthorisedName.Caption = ""
        TxtAuthOperID.Text = ""
        Exit Sub
    End If

    Set Operator = con.Execute("select operator.OPERATOR_ID, operator_user.U_HEBREW_NAME from lims_sys.operator, lims_sys.operator_user " & _
            "where operator.operator_id = operator_user.operator_id and name = '" & CmbAuthorisedBy.Text & "'")
    If (Not Operator.EOF) And (Trim(Operator("U_HEBREW_NAME")) <> "") Then
        LblAuthorisedName.Caption = Trim(Operator("U_HEBREW_NAME"))
        TxtAuthOperID.Text = Trim(Operator("OPERATOR_ID"))
    Else
        LblAuthorisedName.Caption = ""
        TxtAuthOperID.Text = ""
    End If
NoOperator:
End Sub

Private Sub RefreshOperators()
    On Error GoTo NoOperator
    
    Dim Operators As ADODB.Recordset
    
    Dim lOper As Variant

    CmbAuthorisedBy.Clear
    CmbAuthorisedBy.List(0) = "הכל"
    
    If cmbRoles.Text = "הכל" Then
        sql = "select NAME from lims_sys.OPERATOR order by FULL_NAME"
    Else
        sql = "select O.NAME " & vbCrLf & _
              "from lims_sys.OPERATOR O, lims_sys.OPERATOR_ROLE O_R, lims_sys.LIMS_ROLE R " & vbCrLf & _
              "where O.OPERATOR_ID (+) = O_R.OPERATOR_ID " & vbCrLf & _
              "  and O_R.ROLE_ID = R.ROLE_ID " & vbCrLf & _
              "  and R.NAME = '" & cmbRoles.Text & "'" & vbCrLf & _
              "order by O.FULL_NAME "
    End If
        
    Set Operators = con.Execute(sql)
    Do Until Operators.EOF
        CmbAuthorisedBy.List(CmbAuthorisedBy.ListCount) = Operators("NAME")
        Operators.MoveNext
    Loop
    Operators.Close
    Set Operators = Nothing
        
'    Else
'        For Each lOper In OperatorsRoles
'            If cmbRoles.Text = lOper.Role Then
'                CmbAuthorisedBy.List(CmbAuthorisedBy.ListCount) = lOper.Name
'            End If
'        Next
'
'    End If
    
    CmbAuthorisedBy.Text = "הכל"

NoOperator:
End Sub

Private Sub SetColumnHeader()
    'select the column headers for sdg
    Dim SqlStr  As String
    Dim Column As ADODB.Recordset

    strColumnHeaders = ""
    SqlStr = "select sf.name || decode(ec.LOOKUP_SCHEMA_FIELD_ID, null, '', ' - ' || " & _
                "(select sf1.name from lims_sys.schema_field sf1 where sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)), " & _
                    "ec.column_width " & _
                    "from lims_sys.schema_field sf, " & _
                    "lims_sys.entity_column ec " & _
                    "where ec.SCHEMA_ENTITY_ID = (select se1.SCHEMA_ENTITY_ID from lims_sys.schema_entity se1 " & _
                "where upper(se1.name) = 'SDG') and " & _
                "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
                "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " " & _
                "order by ec.ORDER_NUMBER"
    Set Column = con.Execute(SqlStr)
    If Not Column.EOF Then
        Column.MoveFirst
    End If
    
    strColumnHeaders = strColumnHeaders & vbTab & "NAME"
    Call LstSDG.ColumnHeaders.Add(, , "NAME", 97 * 15)

    While Not Column.EOF
        If UCase(Column(0)) <> "NAME" Then
            strColumnHeaders = strColumnHeaders & vbTab & Trim(nte(Column(0)))
            Call LstSDG.ColumnHeaders.Add(, , Column(0), Column(1) * 15)
        End If
        Column.MoveNext
    Wend
End Sub

'פונקציה כללית לבדיקת תאריך
Private Function CheckDate(MaskEdBox1 As MaskEdBox) As Boolean
    On Error GoTo errDate

    Dim dt As Date  'משתנה שיכול לקבל ערכי תאריך חוקיים בלבד
    
    If MaskEdBox1.Text <> "__/__/____" Then
        
        If Mid(MaskEdBox1.Text, 4, 2) > 12 Then    'האם תוי החודש מתאימים לפורמט הארופאי
            GoTo errDate    'שגוי - ועל כן עבור לקטע קוד להודעת שגיאה
        End If
    
        dt = Replace(MaskEdBox1.Text, "_", "")
        MaskEdBox1.Text = Format(dt, "dd/mm/yyyy")
    End If
    
    CheckDate = True
    Exit Function

errDate:
    'תאריך שגוי
    '**********
    MaskEdBox1.BackColor = RED
    MsgBox (MaskEdBox1.Text & " - " & "תאריך לא תקין"), vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical + vbOKOnly, "קלט שגוי"
    MaskEdBox1.BackColor = WHITE
'    MaskEdBox1.Text = "__/__/____"
    MaskEdBox1.SetFocus
    CheckDate = False
End Function

Private Sub GetSelectClause()
    Dim SqlStr As String
    Dim Column As ADODB.Recordset
    'select clause (column database names for sdg)
    SelectClause = "SDG.NAME, "

    SqlStr = "select decode(ec.LOOKUP_SCHEMA_FIELD_ID, null, st.DATABASE_NAME || '.' ||sf.DATABASE_NAME, " & _
            "(select st1.DATABASE_NAME || ec.schema_field_id || '.' ||sf1.DATABASE_NAME " & _
            " from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
            "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
            "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)) " & _
            "from lims_sys.schema_field sf, " & _
            "lims_sys.entity_column ec, " & _
            "lims_sys.schema_table st " & _
            "where ec.SCHEMA_ENTITY_ID = (select se1.SCHEMA_ENTITY_ID from lims_sys.schema_entity se1 " & _
            "where upper(se1.name) = 'SDG') and " & _
            "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
            "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
            "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID " & _
            "order by ec.ORDER_NUMBER"

    Set Column = con.Execute(SqlStr)
    While Not Column.EOF
        If Column(0) <> "SDG.NAME" Then
            SelectClause = SelectClause & Column(0) & ", "
        End If
        Column.MoveNext
    Wend
    SelectClause = SelectClause & "sdg.status"
End Sub

Private Sub GetFromClause()
    Dim SqlStr As String
    Dim Column As ADODB.Recordset
    'from clause (database table names)
    FromClause = "lims_sys.sdg, lims_sys.sdg_user, "

    SqlStr = "select distinct('lims_sys.' || " & _
        "(select st1.DATABASE_NAME || ' ' || st1.DATABASE_NAME || ec.schema_field_id " & _
        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID)) " & _
        "from lims_sys.schema_field sf, " & _
        "lims_sys.entity_column ec, " & _
        "lims_sys.schema_table st " & _
        "where ec.SCHEMA_ENTITY_ID = (select se1.SCHEMA_ENTITY_ID from lims_sys.schema_entity se1 " & _
        "where upper(se1.name) = 'SDG') and " & _
        "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
        "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
        "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID and " & _
        "ec.LOOKUP_SCHEMA_FIELD_ID is not null"

    Set Column = con.Execute(SqlStr)
    While Not Column.EOF
        FromClause = FromClause & Column(0) & ", "
        Column.MoveNext
    Wend
    FromClause = Left(FromClause, Len(FromClause) - 2)
End Sub

Private Sub GetWhereClause()
    Dim SqlStr As String
    Dim Column As ADODB.Recordset
    'where clause (add hard coded the 'sdg.sdg_id = sdg_user.sdg_id')
    WhereClause = "sdg.sdg_id = sdg_user.sdg_id and "

    SqlStr = "select distinct(st.DATABASE_NAME || '.' || sf.DATABASE_NAME || ' = ' || " & _
        "(select st1.DATABASE_NAME || ec.schema_field_id " & _
        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID) || '.' || " & _
        "lims.lims_function.get_primary_key( " & _
        "(select st1.DATABASE_NAME " & _
        "from lims_sys.schema_table st1, lims_sys.schema_field sf1 " & _
        "where st1.SCHEMA_TABLE_ID = sf1.SCHEMA_TABLE_ID and " & _
        "sf1.SCHEMA_FIELD_ID = ec.LOOKUP_SCHEMA_FIELD_ID))) || '(+)' " & _
        "from lims_sys.schema_field sf, " & _
        "lims_sys.entity_column ec, " & _
        "lims_sys.schema_table st " & _
        "where ec.SCHEMA_ENTITY_ID = (select se1.SCHEMA_ENTITY_ID from lims_sys.schema_entity se1 " & _
        "where upper(se1.name) = 'SDG') and " & _
        "sf.SCHEMA_FIELD_ID = ec.SCHEMA_FIELD_ID and " & _
        "ec.OPERATOR_ID = " & NtlsUser.GetOperatorId & " and " & _
        "st.SCHEMA_TABLE_ID = sf.SCHEMA_TABLE_ID and " & _
        "ec.LOOKUP_SCHEMA_FIELD_ID is not null"

    Set Column = con.Execute(SqlStr)
    While Not Column.EOF
        WhereClause = WhereClause & Column(0) & " and "
        Column.MoveNext
    Wend
    WhereClause = Left(WhereClause, Len(WhereClause) - 5)
End Sub

Private Sub FillList()

    Dim RstRequest As ADODB.Recordset
    Dim WhereStr As String
    Dim li As ListItem
    Dim MalignantFlag As String
    Dim ConsultFlag As String
    Dim PositiveFlag As String
    Dim QCFlag As String
    Dim SqlStatement As String
'    Dim SqlCount As String
    Dim DateFlag As Boolean
    Dim i
    Dim lstSdgType As String
    Dim lstSdgFrom As String
    Dim lstSdgTo As String
    Dim lstFreeText As String
    Dim lstMsg As String
    Dim lTmpFromClause As String
    Dim lTmpWhereClause As String
    
    
    On Error GoTo ErrHnd
    
    LstSDG.ListItems.Clear
    
    MalignantFlag = ""
    ConsultFlag = ""
    PositiveFlag = ""
    QCFlag = ""

    lTmpFromClause = ""
    lTmpWhereClause = ""
    WhereStr = ""

    ' Dates
    If Trim(MaskEdBoxFrom(dfCreatedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.CREATED_ON,'ddd') >= to_date('" & MaskEdBoxFrom(dfCreatedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxTo(dfCreatedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.CREATED_ON,'ddd') <= to_date('" & MaskEdBoxTo(dfCreatedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxFrom(dfReceivedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.RECEIVED_ON,'ddd') >= to_date('" & MaskEdBoxFrom(dfReceivedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxTo(dfReceivedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.RECEIVED_ON,'ddd') <= to_date('" & MaskEdBoxTo(dfReceivedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxFrom(dfCompletedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.COMPLETED_ON,'ddd') >= to_date('" & MaskEdBoxFrom(dfCompletedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxTo(dfCompletedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.COMPLETED_ON,'ddd') <= to_date('" & MaskEdBoxTo(dfCompletedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxFrom(dfAuthorisedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.AUTHORISED_ON,'ddd') >= to_date('" & MaskEdBoxFrom(dfAuthorisedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    If Trim(MaskEdBoxTo(dfAuthorisedOn).Text) <> "__/__/____" Then
        WhereStr = WhereStr & " and trunc(sdg.AUTHORISED_ON,'ddd') <= to_date('" & MaskEdBoxTo(dfAuthorisedOn).Text & "', 'dd/mm/yyyy') " & vbCrLf
    End If

    ' External refferance
    ' check by the 1st 8 digits for a match:
    If Trim(txtExternalReference) <> "" Then
        WhereStr = WhereStr & " and substr(sdg.EXTERNAL_REFERENCE,1,8) = '" & Left(txtExternalReference, 8) & "' " & vbCrLf
        'WhereStr = WhereStr & " and sdg.EXTERNAL_REFERENCE = '" & txtExternalReference & "' " & vbCrLf
    End If
    
    ' Status
    Select Case Trim(CmbFromStatus.Text)
    Case "All"
        WhereStr = WhereStr & " and sdg.STATUS <> 'X' " & vbCrLf
    Case "במעבדה"
        WhereStr = WhereStr & " and sdg.STATUS in ('V','P','C','I') " & vbCrLf
    Case Else
        WhereStr = WhereStr & " and sdg.STATUS = '" & StatusCodes(CmbFromStatus.Text) & "' " & vbCrLf
    End Select

    ' Sdg Type
    If Trim(CmbFromStatus.Text) = "Unreceived" Then
        Select Case cmbSdgType.ListIndex
            Case stB
                WhereStr = WhereStr & " and sdg.NAME LIKE '%B' " & vbCrLf
            Case stC
                WhereStr = WhereStr & " and sdg.NAME LIKE '%C' " & vbCrLf
            Case stP
                WhereStr = WhereStr & " and sdg.NAME LIKE '%P' " & vbCrLf
            Case stAll
        End Select
    Else
        Select Case cmbSdgType.ListIndex
            Case stB
                WhereStr = WhereStr & " and sdg.NAME LIKE 'B%' " & vbCrLf
            Case stC
                WhereStr = WhereStr & " and sdg.NAME LIKE 'C%' " & vbCrLf
            Case stP
                WhereStr = WhereStr & " and sdg.NAME LIKE 'P%' " & vbCrLf
            Case stAll
        End Select
    End If

    If Trim(txtSdgYear.Text) <> "" Then
        WhereStr = WhereStr & " and " & _
            " TO_CHAR(sdg.RECEIVED_ON ,'YY') = '" & _
            Trim(txtSdgYear.Text) & "' " & vbCrLf
    End If

    ' Sdg
    If Trim(TxtFromSdgID) <> "" Then
        Select Case cmbSdgType.ListIndex
        Case stB
            lstSdgType = "B"
        Case stC
            lstSdgType = "C"
        Case stP
            lstSdgType = "P"
        Case stAll
            lstSdgType = ""
        End Select

        lstSdgFrom = lstSdgType & TxtFromSdgID & "/" & txtSdgYear

        If cmbSdgType.ListIndex = stAll Then
            If Trim(TxtToSdgID) = "" Then
                WhereStr = WhereStr & _
                    " and (sdg.NAME = 'B" & lstSdgFrom & "' " & vbCrLf & _
                    "   or sdg.NAME = 'C" & lstSdgFrom & "' " & vbCrLf & _
                    "   or sdg.NAME = 'P" & lstSdgFrom & "') " & vbCrLf
            Else
                lstSdgTo = lstSdgType & TxtToSdgID & "/" & txtSdgYear
                WhereStr = WhereStr & _
                    " and ((sdg.NAME >= 'B" & lstSdgFrom & "' and sdg.NAME <= 'B" & lstSdgTo & "') " & vbCrLf & _
                    "   or (sdg.NAME >= 'C" & lstSdgFrom & "' and sdg.NAME <= 'C" & lstSdgTo & "') " & vbCrLf & _
                    "   or (sdg.NAME >= 'P" & lstSdgFrom & "' and sdg.NAME <= 'P" & lstSdgTo & "')) " & vbCrLf
            End If
        Else
            If Trim(TxtToSdgID) = "" Then
                WhereStr = WhereStr & " and sdg.NAME = '" & lstSdgFrom & "' " & vbCrLf
            Else
                lstSdgTo = lstSdgType & TxtToSdgID & "/" & txtSdgYear
                WhereStr = WhereStr & " and sdg.NAME >= '" & lstSdgFrom & "' " & vbCrLf & _
                                      " and sdg.NAME <= '" & lstSdgTo & "' " & vbCrLf
            End If
        End If
    End If

    ' Doctors
    If Trim(TxtReferringPhysician) <> "" Then
        WhereStr = WhereStr & " and sdg_user.U_REFERRING_PHYSICIAN = " & TxtRefID & vbCrLf
    End If

    If Trim(TxtSubmittingPhysician) <> "" Then
        WhereStr = WhereStr & " and sdg_user.U_IMPLEMENTING_PHYSICIAN = " & TxtSubID & vbCrLf
    End If

    ' Collection stations
    If Trim(TxtFromCollectingStation.Text) <> "" Then
        'lTmpFromClause = lTmpFromClause & ", lims.u_clinic u_clinic "
        'lTmpWhereClause = lTmpWhereClause & " and sdg_user.U_COLLECTION_STATION=u_clinic.u_clinic_id "
         
        'WhereStr = WhereStr & _
          " and u_clinic.name = '" & Trim(TxtFromCollectingStation.Text) & "' " & vbCrLf
    
        WhereStr = WhereStr & " and sdg_user.U_COLLECTION_STATION = " & TxtFromCollectingStationID
    End If
     
'    If Trim(TxtFromCollectingStation) <> "" Then
'        If Trim(TxtToCollectingStation) = "" And TxtFromCollectingStationID <> "" Then
'            WhereStr = WhereStr & " and sdg_user.U_COLLECTION_STATION = " & TxtFromCollectingStationID & vbCrLf
'        Else
'            lTmpFromClause = ", lims_sys.u_clinic u_clinic "
'            lTmpWhereClause = " and sdg_user.U_COLLECTION_STATION=u_clinic.u_clinic_id "
'            WhereStr = WhereStr & _
'                " and to_number(u_clinic.name) >= " & TxtFromCollectingStation & vbCrLf & _
'                " and to_number(u_clinic.name) <= " & TxtToCollectingStation & vbCrLf
'        End If
'    End If
 
    ' Authorized by
    If Trim(CmbAuthorisedBy.Text) <> "הכל" Then
        WhereStr = WhereStr & _
                " and (lims.authorization.signed_by(sdg.sdg_id,1) = " & TxtAuthOperID & vbCrLf & _
                "      or lims.authorization.signed_by(sdg.sdg_id,2) = " & TxtAuthOperID & ") " & vbCrLf
    End If

    ' Patients
    If Trim(TxtPatientID) <> "" Then
        WhereStr = WhereStr & " and sdg_user.U_PATIENT = " & TxtPatientIDNum.Text & vbCrLf
    End If

    'selection by the range of patient age:
'    If Trim(txtPatientMinAge) <> "" And Trim(txtPatientMaxAge) <> "" Then
'        WhereStr = WhereStr & " and sdg_user.u_patient in "
'        WhereStr = WhereStr & " ( "
'        WhereStr = WhereStr & " select c.client_id "
'        WhereStr = WhereStr & " from lims_sys.client c, lims_sys.client_user cu "
'        WhereStr = WhereStr & " where sdg_user.u_patient = c.client_id "
'        WhereStr = WhereStr & " and c.client_id = cu.client_id "
'        WhereStr = WhereStr & " and to_number(to_char(sdg.CREATED_ON, 'yyyy' )) - to_number(to_char(cu.U_DATE_OF_BIRTH, 'yyyy')) >= " & txtPatientMinAge & " "
'        WhereStr = WhereStr & " and to_number(to_char(sdg.CREATED_ON, 'yyyy' )) - to_number(to_char(cu.U_DATE_OF_BIRTH, 'yyyy')) <= " & txtPatientMaxAge & " "
'        WhereStr = WhereStr & " ) "
'    End If
 
    'search by minimum age:
    If Trim(txtPatientMinAge) <> "" Then
        
        If Not IsNumeric(txtPatientMinAge) Then
            MsgBox " יש להכניס ערך מספרי עבור גיל הפציינט "
            SSTab.Tab = 1
            txtPatientMinAge.SetFocus
            Exit Sub
        End If
        
        WhereStr = WhereStr & " and sdg_user.u_patient in "
        WhereStr = WhereStr & " ( "
        WhereStr = WhereStr & " select c.client_id "
        WhereStr = WhereStr & " from lims_sys.client c, lims_sys.client_user cu "
        WhereStr = WhereStr & " where sdg_user.u_patient = c.client_id "
        WhereStr = WhereStr & " and c.client_id = cu.client_id "
        WhereStr = WhereStr & " and to_number(to_char(sdg.CREATED_ON, 'yyyy' )) - to_number(to_char(cu.U_DATE_OF_BIRTH, 'yyyy')) >= " & txtPatientMinAge & " "
        WhereStr = WhereStr & " ) "
    End If
            
    'search by maximum age:
    If Trim(txtPatientMaxAge) <> "" Then
 
        If Not IsNumeric(txtPatientMaxAge) Then
            MsgBox " יש להכניס ערך מספרי עבור גיל הפציינט "
            SSTab.Tab = 1
            txtPatientMaxAge.SetFocus
            Exit Sub
        End If
        
        WhereStr = WhereStr & " and sdg_user.u_patient in "
        WhereStr = WhereStr & " ( "
        WhereStr = WhereStr & " select c.client_id "
        WhereStr = WhereStr & " from lims_sys.client c, lims_sys.client_user cu "
        WhereStr = WhereStr & " where sdg_user.u_patient = c.client_id "
        WhereStr = WhereStr & " and c.client_id = cu.client_id "
        WhereStr = WhereStr & " and to_number(to_char(sdg.CREATED_ON, 'yyyy' )) - to_number(to_char(cu.U_DATE_OF_BIRTH, 'yyyy')) <= " & txtPatientMaxAge & " "
        WhereStr = WhereStr & " ) "
    End If
            
            
            

    ' Add the Documents table to the query
    If Trim(txtFreeText) <> "" Then
    'If (Trim(txtSnomedM) <> "") Or (Trim(txtSnomedT) <> "") Or (Trim(txtFreeText) <> "") Then
        lTmpFromClause = lTmpFromClause & " , lims.letter_doc letter_doc "
        lTmpWhereClause = lTmpWhereClause & " and sdg.sdg_id = letter_doc.sdg_id "
    End If
    ' FreeText
    If Trim(txtFreeText) <> "" Then
        lstFreeText = Replace(Replace(Trim(txtFreeText), "'", "''"), vbCrLf, "")
        WhereStr = WhereStr & " and contains(letter_doc.doc_file,'" & lstFreeText & "') > 0 " & vbCrLf
    End If


    'show old versions or not:
    'if the check box is false don't show them:
    If chkShowOldVersions.Value = False Then
        WhereStr = WhereStr & " and instr(sdg.name,'V',2)=0 " & vbCrLf
    End If


    ' Snomed:
    'must be an sdg that has a snomed-m like one of the entered velues
    'and likewise for the snomed-t:
'    If Trim(txtSnomedM) <> "" And Trim(txtSnomedT) <> "" Then
'        Dim strSnomedM As String
'        Dim strSnomedT As String
'
'        strSnomedM = Trim(txtSnomedM)
'        strSnomedT = Trim(txtSnomedT)
'
'        strSnomedM = GetSnomedText(strSnomedM, "M")
'        strSnomedT = GetSnomedText(strSnomedT, "T")
'
'
'        WhereStr = WhereStr & " and sdg.sdg_id in " & _
'                   "(select sdg_id from lims.sdg_snomed where snomed in " & _
'                   strSnomedM & " )" & " " & vbCrLf
'
'        WhereStr = WhereStr & " and sdg.sdg_id in " & _
'                   "(select sdg_id from lims.sdg_snomed where snomed in " & _
'                   strSnomedT & " )" & " " & vbCrLf
'    End If
    
    
    'Snomed:
    'must be an sdg that has a snomed-m like one of the entered velues
    'and likewise for the snomed-t:
        
    If Trim(txtSnomedM) <> "" Then
        Dim strSnomedM As String
        
        strSnomedM = Trim(txtSnomedM)
        strSnomedM = GetSnomedText(strSnomedM, "M")
         
        'search by sdg_id:
        If chkSearchSnomedByPatient.Value = False Then
            WhereStr = WhereStr & " and sdg.sdg_id in " & _
                       "(select sdg_id from lims.sdg_snomed where snomed in " & _
                       strSnomedM & " )" & " " & vbCrLf
        'search by patient:
        Else
            WhereStr = WhereStr & " and sdg_user.u_patient in " & _
                "(select patient_id from lims.sdg_snomed where snomed in " & _
                strSnomedM & " )" & " " & vbCrLf
        End If
    End If
    
    If Trim(txtSnomedT) <> "" Then
        Dim strSnomedT As String
        
        strSnomedT = Trim(txtSnomedT)
        strSnomedT = GetSnomedText(strSnomedT, "T")
        
        'search by sdg_id:
        If chkSearchSnomedByPatient.Value = False Then
            WhereStr = WhereStr & " and sdg.sdg_id in " & _
                       "(select sdg_id from lims.sdg_snomed where snomed in " & _
                       strSnomedT & " )" & " " & vbCrLf
        'search by patient:
        Else
            WhereStr = WhereStr & " and sdg_user.u_patient in " & _
                "(select patient_id from lims.sdg_snomed where snomed in " & _
                strSnomedT & " )" & " " & vbCrLf
        End If
    End If
    
    
    'snomed:
    'search by the last digit of snomed-m:
    If Trim(txtLastSnomedMdigit) <> "" Then
        Dim strSnomedMLastDigit As String
        
        strSnomedMLastDigit = Trim(txtLastSnomedMdigit)
        strSnomedMLastDigit = GetSnomedLastDigits(strSnomedMLastDigit)
        
        'search by sdg_id:
        If chkSearchSnomedByPatient.Value = False Then
            WhereStr = WhereStr & " and sdg.sdg_id in " & _
                       "(select sdg_id from lims.sdg_snomed " & _
                       "where substr(snomed, 6, 1) in " & _
                       strSnomedMLastDigit & _
                       " and snomed_type='M' " & _
                       ") " & vbCrLf
        'search by patient:
        Else
            WhereStr = WhereStr & " and sdg_user.u_patient in " & _
                       "(select patient_id from lims.sdg_snomed " & _
                       "where substr(snomed, 6, 1) in " & _
                       strSnomedMLastDigit & _
                       " and snomed_type='M' " & _
                       ") " & vbCrLf
        End If
'
'        WhereStr = WhereStr & " and sdg.sdg_id in " & _
'                   "(select sdg_id from lims.sdg_snomed " & _
'                   "where (substr(snomed, 6, 1) in " & _
'                   strSnomedMLastDigit & _
'                   " and substr(snomed, 1, 1)='M') " & _
'                   "or substr(snomed, 1, 1)='T' " & _
'                   ") " & vbCrLf
                   
'MsgBox WhereStr
'
'    (substr(ds.snomed, 6, 1) in ('1','0')
'    and substr(ds.snomed, 1, 1) ='M'
'    or substr(ds.snomed, 1, 1) ='T');
    End If
    
    
    'search also by a response letter:
    'sent / not sent
    'group code of the letter
    If optLetterSent(0).Value = True Or optLetterSent(1).Value = True Then
        Dim strCommonSql As String
     
        strCommonSql = " exists " & _
                       "(select 1 from lims_sys.u_letter_control_user lc " & _
                       "where lc.u_sdg_id = sdg.sdg_id"
        
        If optLetterSent(1).Value = True Then
            strCommonSql = " and not " & strCommonSql
        Else
            strCommonSql = " and " & strCommonSql
        End If
        
        'If txtLetterCode <> "" Then
        If cmbLetterCode.Text <> "" Then
            Dim strGrpCode As String
            strGrpCode = dicGrpCode(CStr(cmbLetterCode.Text))
            
            strCommonSql = strCommonSql & _
                  " and  lc.u_grp_code = '" & strGrpCode & "' "
        End If
    
        strCommonSql = strCommonSql & ") "
        WhereStr = WhereStr & strCommonSql & " " & vbCrLf


    End If
    
    
    'malignancy:
    If optMalignant(0).Value = True Then
        WhereStr = WhereStr & " and  lims.IS_MALIGNANT(sdg.sdg_id)='T' "
    End If
    
    If optMalignant(1).Value = True Then
        WhereStr = WhereStr & " and  lims.IS_MALIGNANT_ADMIN(sdg.sdg_id)='T' "
    End If
    
    If optMalignant(2).Value = True Then
        WhereStr = WhereStr & " and  lims.IS_MALIGNANT(sdg.sdg_id)='F' "
    End If
    
'    If optMalignant(0).Value = True Then
'        WhereStr = WhereStr & " and ( lims.IS_MALIGNANT(sdg.sdg_id)='T' "
'
'        If chkMalignancyPatient.Value = 1 Then
'            WhereStr = WhereStr & " or exists " & _
'                " ( " & _
'                "   select 1 " & _
'                "   from lims_sys.sdg_user duMal " & _
'                "   where duMal.u_patient = sdg_user.u_patient " & _
'                "   and   lims.IS_MALIGNANT(duMal.sdg_id)='T' " & _
'                " ) "
'        End If
'
'        WhereStr = WhereStr & " ) "
'    End If
'
'    If optMalignant(1).Value = True Then
'        WhereStr = WhereStr & " and ( lims.IS_MALIGNANT(sdg.sdg_id)='F' "
'
'        If chkMalignancyPatient.Value = 1 Then
'            WhereStr = WhereStr & " or exists " & _
'                " ( " & _
'                "   select 1 " & _
'                "   from lims_sys.sdg_user duMal " & _
'                "   where duMal.u_patient = sdg_user.u_patient " & _
'                "   and   lims.IS_MALIGNANT(duMal.sdg_id)='F' " & _
'                " ) "
'        End If
'
'        WhereStr = WhereStr & " ) "
'    End If
    
'    If chkMalignancy.Value = 1 Then
'        WhereStr = WhereStr & " and ( lims.IS_MALIGNANT(sdg.sdg_id)='T' "
'
'        If chkMalignancyPatient.Value = 1 Then
'            WhereStr = WhereStr & " or exists " & _
'                " ( " & _
'                "   select 1 " & _
'                "   from lims_sys.sdg_user duMal " & _
'                "   where duMal.u_patient = sdg_user.u_patient " & _
'                "   and   lims.IS_MALIGNANT(duMal.sdg_id)='T' " & _
'                " ) "
'        End If
'
'        WhereStr = WhereStr & " ) "
'    End If
    
    
'    If Trim(txtSnomedM) <> "" Then
'        lstFreeText = Replace(Replace(Trim(txtSnomedM), "'", "''"), vbCrLf, "")
'        WhereStr = WhereStr & " and contains(letter_doc.doc_file,'" & lstFreeText & "') > 0 " & vbCrLf
'    End If
'
'    If Trim(txtSnomedT) <> "" Then
'        lstFreeText = Replace(Replace(Trim(txtSnomedT), "'", "''"), vbCrLf, "")
'        WhereStr = WhereStr & " and contains(letter_doc.doc_file,'" & lstFreeText & "') > 0 " & vbCrLf
'    End If
    

    ' Other
    ' Malignant
    If Trim(CmbMalignant.Text) = "כן" Then
        MalignantFlag = "sdg_user.U_MALIGNANT = 'T'" & vbCrLf
    ElseIf Trim(CmbMalignant.Text) = "לא" Then
        MalignantFlag = "nvl(sdg_user.U_MALIGNANT, ' ') <> 'T'" & vbCrLf
    End If
    If MalignantFlag <> "" Then
        WhereStr = WhereStr & " and " & MalignantFlag & vbCrLf
    End If

    ' Consult
    If Trim(CmbConsult.Text) = "כן" Then
        ConsultFlag = "sdg_user.U_CONSULT = 'T'" & vbCrLf
    ElseIf Trim(CmbConsult.Text) = "לא" Then
        ConsultFlag = "nvl(sdg_user.U_CONSULT, ' ') <> 'T'" & vbCrLf
    End If
    If ConsultFlag <> "" Then
        WhereStr = WhereStr & " and " & ConsultFlag & vbCrLf
    End If

    ' Positive
    If Trim(CmbPositive.Text) = "כן" Then
        PositiveFlag = "sdg_user.U_POSITIVE = 'T'" & vbCrLf
    ElseIf Trim(CmbPositive.Text) = "לא" Then
        PositiveFlag = "nvl(sdg_user.U_POSITIVE, ' ') <> 'T'" & vbCrLf
    End If
    If PositiveFlag <> "" Then
        WhereStr = WhereStr & " and " & PositiveFlag & vbCrLf
    End If

    ' QC
    If Trim(CmbQC.Text) = "כן" Then
        QCFlag = "sdg_user.U_QC is not NULL" & vbCrLf
    ElseIf Trim(CmbQC.Text) = "לא" Then
        QCFlag = "sdg_user.U_QC is NULL" & vbCrLf
    End If
    If QCFlag <> "" Then
        WhereStr = WhereStr & " and " & QCFlag & vbCrLf
    End If

    SqlStatement = "select " & SelectClause & " from " & FromClause & lTmpFromClause & _
        " where " & WhereClause & lTmpWhereClause & WhereStr & " order by sdg.name"


    ' GET QUERY SQL
    If DontRunQuery Then
        SqlStatement = Replace(SqlStatement, " and ", vbCrLf & "        and ", , , vbTextCompare)
        SqlStatement = Replace(SqlStatement, " from ", vbCrLf & " from ", , , vbTextCompare)
        SqlStatement = Replace(SqlStatement, " where ", vbCrLf & " where ", , , vbTextCompare)
        SqlStatement = Replace(SqlStatement, " order by ", vbCrLf & " order by ", , , vbTextCompare)
        SqlStatement = Replace(SqlStatement, vbCrLf & vbCrLf, vbCrLf)
        MousePointer = vbDefault
        
        MousePointer = vbDefault
        Call ClipBoard_SetData(SqlStatement)
        If Clipboard.GetText = SqlStatement Then
            MsgBox "SQL Query Copied to clipboard, " & vbCrLf & "Query will not run!!!", vbInformation
        Else
            MsgBox "Error Copying SQL query.", vbCritical
        End If
        
    Else
        'GO ON AS USUAL
        'Get the sample & Result list
        
        MousePointer = vbHourglass
        Set RstRequest = con.Execute(SqlStatement)
    
        MousePointer = vbDefault
    
       
    
        If CheckCountSDG.Value = Unchecked Then Exit Sub
    
        If RstRequest.RecordCount > MaxRecordToDisplay Then
            MsgBox " מספר הרשומות שנשלפו " & RstRequest.RecordCount & " עולה על המקסימום המוגדר להצגה " & MaxRecordToDisplay, , "Nautilus - Request Quer"
            Exit Sub
        End If
    
    
        While Not RstRequest.EOF
            Set li = LstSDG.ListItems.Add(, , nte(RstRequest(0)), , CStr(RstRequest("STATUS").Value))
            For i = 1 To RstRequest.Fields.Count - 2
                li.SubItems(i) = nte(RstRequest(i))
            Next i
            
            
    '        li.Key = RstRequest("name")
            
           ' li.ToolTipText = RstRequest("name")
            
    
    
    'SdgNames.Item(SdgNames.Count + 1) = RstRequest("name")
    'MsgBox RstRequest("name")
            
            
            RstRequest.MoveNext
    
        Wend
        RstRequest.Close
        
        LblCount.Caption = LstSDG.ListItems.Count
    
    
        SSTab.Tab = 0
        LstSDG.SetFocus
    End If 'get error Sql


Exit Sub
ErrHnd:
    MousePointer = vbDefault
    Dim lFileNumber As Integer
    lFileNumber = FreeFile
    lstMsg = "Error from SDG Query: " & Date & " " & Time & vbCrLf & vbCrLf & _
             "SqlStatement :" & vbCrLf & _
             SqlStatement & vbCrLf & vbCrLf & _
             "Error Number : " & Err.Number & vbCrLf & _
             "Error Description : " & Err.Description
    Open "C:\SDGQuery_Err.txt" For Output As #lFileNumber
    Print #lFileNumber, String(80, "-")
    Print #lFileNumber, lstMsg
    Close #lFileNumber
    lstMsg = "שגיאה נגרמה בשליפה," & vbCrLf & _
             "נרשם קובץ שגיאה במחיצה הראשית" & vbCrLf & _
             "פנה לאחראי או לאחראית."
    MsgBox lstMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "אוף, קרתה תקלה ..."
End Sub

Private Sub TxtToCollectingStation_Change()
    Call RefreshCollectingStation(TxtToCollectingStation.Text, LblToCollectingStationName)
End Sub

Private Sub TxtToCollectingStation_LostFocus()
    
    If TxtToCollectingStation = "" Then Exit Sub
    
    On Error GoTo ErrHnd
    If TxtFromCollectingStation = "" Then
        TxtFromCollectingStation.BackColor = RED
        MsgBox "מתחנת איסוף חייב להיות קטן או שווה ל עד תחנת איסוף !", vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical + vbOKOnly, "קלט שגוי"
        TxtFromCollectingStation.BackColor = WHITE
        TxtToCollectingStation.SetFocus
    ElseIf (CInt(TxtFromCollectingStation) > CInt(TxtToCollectingStation)) Then
        TxtFromCollectingStation.BackColor = RED
        MsgBox "מתחנת איסוף חייב להיות קטן או שווה ל עד תחנת איסוף !", vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical + vbOKOnly, "קלט שגוי"
        TxtFromCollectingStation.BackColor = WHITE
        TxtToCollectingStation.SetFocus
    End If
Exit Sub
ErrHnd:

End Sub

Private Sub TxtToSdgID_LostFocus()
    
    TxtToSdgID = Trim(TxtToSdgID)
    
    If TxtToSdgID = "" Then Exit Sub
    
    Select Case Left(TxtToSdgID, 1)
    Case "b", "B"
        cmbSdgType.ListIndex = stB
        TxtToSdgID = Mid(TxtToSdgID, 2)
    Case "c", "C"
        cmbSdgType.ListIndex = stC
        TxtToSdgID = Mid(TxtToSdgID, 2)
    Case "p", "P"
        cmbSdgType.ListIndex = stP
        TxtToSdgID = Mid(TxtToSdgID, 2)
    Case Else
    End Select
    
    If Left(Right(TxtToSdgID, 3), 1) = "/" Then
        txtSdgYear = Right(TxtToSdgID, 2)
        TxtToSdgID = Left(TxtToSdgID, Len(TxtToSdgID) - 3)
    End If
    TxtToSdgID = PadStr(TxtToSdgID, "0", 6, True)
    
    If (TxtFromSdgID > TxtToSdgID) Or (TxtFromSdgID = "") Then
        TxtFromSdgID.BackColor = RED
        MsgBox "ממספר פנימי חייב להיות קטן או שווה ל עד מספר פנימי !", vbMsgBoxRight + vbMsgBoxRtlReading + vbCritical + vbOKOnly, "קלט שגוי"
        TxtFromSdgID.BackColor = WHITE
    End If

End Sub

Private Sub FillOneRequest()
    On Error GoTo ErrEnd
    Dim RstRequest As ADODB.Recordset
    Dim WhereStr As String
    Dim SqlStatement As String
    Dim li As ListItem
    Dim ReqName As String
    Dim i As Integer
    Dim RequestCount As Long

    WhereStr = " and sdg.name = '" & UCase(Trim(TxtRequestBarCode.Text)) & "' "
    SqlStatement = "select " & SelectClause & " from " & FromClause & _
        " where " & WhereClause & WhereStr & " order by sdg.name"

    'Get the Request
    Set RstRequest = con.Execute(SqlStatement)

    If RstRequest.RecordCount = 0 Or RstRequest.EOF Then
        TxtRequestBarCode.BackColor = RED
        MsgBox "Request: " & Trim(TxtRequestBarCode.Text) & " not found !", vbInformation, "Nautilus - Sample Query"
        TxtRequestBarCode.BackColor = WHITE
        Call TxtRequestBarCode.SetFocus
        Exit Sub
    End If

    ReqName = nte(RstRequest("NAME"))
    If Not (LstSDG.FindItem(ReqName) Is Nothing) Then
        TxtRequestBarCode.BackColor = RED
        MsgBox "Request: " & Trim(ReqName) & " already exists in list !", vbCritical, "Nautilus - Sample Query"
        TxtRequestBarCode.BackColor = WHITE
        Call TxtRequestBarCode.SetFocus
        Exit Sub
    End If

    If Not RstRequest.EOF Then
        Set li = LstSDG.ListItems.Add(, , nte(RstRequest(0)), , CStr(RstRequest("STATUS").Value))
        For i = 1 To RstRequest.Fields.Count - 2
            li.SubItems(i) = nte(RstRequest(i))
        Next i
        RstRequest.MoveNext
    End If
    RstRequest.Close

    TxtRequestBarCode.Text = ""

    RequestCount = LstSDG.ListItems.Count
    LblCount.Caption = RequestCount
    LblCount.Refresh

    LstSDG.SortKey = 0
    LstSDG.SortOrder = lvwAscending
    LstSDG.Sorted = True

    Exit Sub

ErrEnd:
    MsgBox "Error in FillOneRequest... " & vbCrLf & _
            Err.Description, vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "Nautilus - Request Query"
End Sub



'change a text like 123,456,789
'to                 'M123','M456','M789'
Private Function GetSnomedText(str As String, strPrefix As String)
    Dim str2 As String
    Dim strToReplace As String
    strToReplace = "','" & strPrefix
    
    str2 = Replace(str, ",", strToReplace, , , vbTextCompare)
    
    str2 = "('" & strPrefix & str2 & "')"
    
    GetSnomedText = str2
End Function


'change a text like 1,2,3,4
'to                 ('1','2','3','4')
Private Function GetSnomedLastDigits(str As String)
    Dim str2 As String
    Dim strToReplace As String
    strToReplace = "','"
    
    str2 = Replace(str, ",", strToReplace, , , vbTextCompare)
    
    str2 = "('" & str2 & "')"
    
    GetSnomedLastDigits = str2
End Function

Sub InitGrpCodeList()
    Dim i As Integer
    
    cmbLetterCode.Clear
    
    cmbLetterCode.AddItem ""
    
    For i = 0 To dicGrpCode.Count - 1
         cmbLetterCode.AddItem CStr(dicGrpCode.Keys(i))
         cmbReportLetter.AddItem CStr(dicGrpCode.Keys(i))
    Next i
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strVer As String
    On Error GoTo Err_UserControl_KeyDown

    If KeyCode = vbKeyF10 And Shift = 1 Then
           strVer = "Name: " & App.EXEName & vbCrLf & _
                 "Path: " & App.Path & vbCrLf & _
                 "Version: " & "[" & App.Major & "." & App.Minor & "." & App.Revision & "]" & vbCrLf & _
                 "Company: One Software Technologies (O.S.T) Ltd."
                 
        MsgBox "-----GET SQL QUERY-----" & vbCrLf & vbCrLf & _
                strVer, vbInformation, "Nautilus - Project Properties & Get SQL Query"
       'for debugging
        ' GET QUERY SQL
        DontRunQuery = True
        CmdFinished.SetFocus
        CmdFinished_Click
        DontRunQuery = False
        
    End If
       Exit Sub
Err_UserControl_KeyDown:
MsgBox "Err_UserControl_KeyDown : " & Err.Description
DontRunQuery = False
End Sub

