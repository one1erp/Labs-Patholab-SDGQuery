VERSION 5.00
Object = "{87B91421-0018-46C6-80FF-DB0FEA100277}#1.5#0"; "ModifySDG.ocx"
Begin VB.Form FrmModifySDG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nautilus - Modify SDG"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   15900
   StartUpPosition =   2  'CenterScreen
   Begin ModifySDG.ModifySDGCtrl ModifySDGCtrl 
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   16325
   End
End
Attribute VB_Name = "FrmModifySDG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Width = ModifySDGCtrl.Width
    Me.Height = ModifySDGCtrl.Height
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub ModifySDGCtrl_CloseClicked()
    Me.Hide
End Sub



