VERSION 5.00
Begin VB.UserControl ctlSnapShot 
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ToolboxBitmap   =   "SnapShot.ctx":0000
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   2985
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "SnapShot.ctx":0312
      Top             =   180
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.CommandButton cmdSnapShot 
      Caption         =   "Get Snap Shot"
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "ctlSnapShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Used As Boolean

Private Sub cmdSnapShot_Click()

    Load frmSnap
    Used = True

End Sub

Private Sub UserControl_Resize()

    UserControl.Width = cmdSnapShot.Width * 15
    UserControl.Height = cmdSnapShot.Height * 15
    cmdSnapShot.Top = 0
    cmdSnapShot.Left = 0
    
End Sub

Private Sub UserControl_Terminate()

    If Used Then
        Unload frmSnap
        Unload frmShots
        Set frmSnap = Nothing
        Set frmShots = Nothing
    End If
    
End Sub
