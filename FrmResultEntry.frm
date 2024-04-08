VERSION 5.00
Object = "{4C253C86-D2A4-459F-BD28-88177A11AE4F}#5.0#0"; "ResultEntry.ocx"
Begin VB.Form FrmResultEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nautilus - Result Entry"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   15960
   StartUpPosition =   2  'CenterScreen
   Begin ResultEntry.ResultEntryCtrl ResultEntryCtrl 
      Height          =   10000
      Left            =   5
      TabIndex        =   0
      Top             =   0
      Width           =   16000
      _ExtentX        =   28231
      _ExtentY        =   17648
   End
End
Attribute VB_Name = "FrmResultEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Width = ResultEntryCtrl.Width
    Me.Height = ResultEntryCtrl.Height
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub ResultEntryCtrl_CloseClicked()
    Me.Hide
End Sub


