VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFindCollectionStation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Collection Station"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10125
   DrawMode        =   14  'Copy Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10125
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmCollectionStationDetails 
      Caption         =   "Collection Station Details"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton cmdClose 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   11
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   10
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find Now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtDistrict 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtHeadClinic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.Image ImageFind 
         Height          =   480
         Left            =   5520
         Picture         =   "frmFindCollectionStation.frx":0000
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label lblDistrict 
         Caption         =   "District:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblHeadClinic 
         Caption         =   "Head Clinic:"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblCode 
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lstCollectionStation 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imglstIcons"
      ColHdrIcons     =   "imglstHeaderIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Head Clinic"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "District"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblNumOfRecords 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "frmFindCollectionStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strClinicId As String
Private strClinicName As String
Private connection As ADODB.connection
Private rs As ADODB.Recordset

Public Function GetClinicId() As String
    If lstCollectionStation.ListItems.Count > 0 Then
        GetClinicId = strClinicId
    Else
        GetClinicId = ""
    End If
End Function

Public Function GetClinicName() As String
    If lstCollectionStation.ListItems.Count > 0 Then
        GetClinicName = strClinicName
    Else
        GetClinicName = ""
    End If
End Function

'Public Sub SetConnection(con_ As ADODB.connection)
'    Set connection = con_
'End Sub


Private Sub cmdClear_Click()
    Call Resetform
End Sub

Private Sub Resetform()
    lblNumOfRecords.Caption = ""
    txtName = ""
    txtCode = ""
    txtHeadClinic = ""
    txtDistrict = ""
    lstCollectionStation.ListItems.Clear
End Sub

Private Sub cmdClose_Click()
'    Resetform
    Me.Hide
End Sub

Private Sub cmdFind_Click()
    
On Error GoTo ERR_cmdFind_Click
    
    FindNow

Exit Sub
ERR_cmdFind_Click:
MsgBox "ERR_cmdFind_Click " & vbCrLf & Err.Description

End Sub


Private Sub FindNow()
    If txtName.Text = "" And _
       txtCode.Text = "" And _
       txtHeadClinic.Text = "" And _
       txtDistrict.Text = "" Then
       
        MsgBox " לא הוכנסו קריטריונים לחיפוש "
        Exit Sub
    End If
    
    FillList
End Sub


Private Sub FillList()

On Error GoTo ERR_FillList
    
    Dim strSql As String
    Dim li As ListItem
    Dim i As Integer
    
    'show the hourglass mous pointer
    MousePointer = 11
    
    
    strSql = "select * from lims_sys.u_clinic, lims_sys.u_clinic_user, lims_sys.address " & _
             "where u_clinic.u_clinic_id = u_clinic_user.u_clinic_id and " & _
             "ADDRESS_TABLE_NAME(+) = 'U_CLINIC' and " & _
             "ADDRESS_ITEM_ID(+) = u_clinic.u_clinic_id and " & _
             "ADDRESS_LINE_1(+) = u_clinic.name "

    If txtCode.Text <> "" Then
        strSql = strSql & " and u_clinic.name = " & Trim(txtCode.Text) & " "
    End If
    
    If txtName.Text <> "" Then
        strSql = strSql & " and u_clinic_user.u_clinic_name like '%" & _
                         Trim(txtName.Text) & "%' "
    End If
    
    If txtHeadClinic.Text <> "" Then
        strSql = strSql & " and u_clinic_user.u_head_clinic = '" & _
                         Trim(txtHeadClinic.Text) & "' "
    End If
    
    
    If txtDistrict.Text <> "" Then
        strSql = strSql & " and u_clinic_user.u_district = '" & _
                         Trim(txtDistrict.Text) & "' "
    End If
    
'WhereStr = WhereStr & "U_FIRST_NAME like '" & Replace(TxtFirstName.Text, "'", "''") & "%' "
     
    
    Set rs = connection.Execute(strSql)


    lstCollectionStation.ListItems.Clear
    If Not rs.EOF Then
        rs.MoveFirst
        i = 0
    
        While Not rs.EOF
            Set li = lstCollectionStation.ListItems.Add(, , nte(rs("U_CLINIC_NAME")), , 1)
            li.Tag = nte(rs("U_CLINIC_NAME"))
            li.SubItems(1) = nte(rs("NAME"))
            li.SubItems(2) = nte(rs("U_HEAD_CLINIC"))
            li.SubItems(3) = nte(rs("U_DISTRICT"))
            rs.MoveNext
            i = i + 1
        Wend
                                
        lblNumOfRecords.ForeColor = vbBlack
        lblNumOfRecords.Caption = " נמצאו " & i & " רשומות "
    
    Else
        lblNumOfRecords.RightToLeft = True
        lblNumOfRecords.ForeColor = vbRed
        lblNumOfRecords.Caption = " לא נמצאו רשומות "
    End If
    rs.Close

    'show the regular mouse pointer
    MousePointer = 0


Exit Sub
ERR_FillList:
MsgBox "ERR_FillList " & vbCrLf & Err.Description

End Sub

Public Sub Initialize(con_ As ADODB.connection)
    Set connection = con_
'    Call imglstIcons.ListImages.Add(, "L1", LoadPicture("Resource\Client.ico"))

    Call zLang.Hebrew
    txtName.Alignment = vbRightJustify
    txtName.RightToLeft = True
    Resetform
End Sub

Private Sub lstCollectionStation_DblClick()
    CloseForm
End Sub

Private Sub CloseForm()
    If lstCollectionStation.ListItems.Count > 0 Then
        strClinicName = lstCollectionStation.SelectedItem.Tag
        strClinicId = lstCollectionStation.SelectedItem.SubItems(1)
        
    '    Description = LsPatient.SelectedItem.SubItems(1) & " " & _
    '                  LsPatient.SelectedItem.SubItems(2)
    '    ID = PatientID
    End If
    Call zLang.SetOrigLang
    
    Me.Hide
  '  Resetform
End Sub


Public Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function

Private Sub Form_Load()
'    Initialize
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FindNow
    End If
End Sub

Private Sub txtDistrict_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FindNow
    End If
End Sub

Private Sub txtHeadClinic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FindNow
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        FindNow
    End If
End Sub

