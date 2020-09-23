VERSION 5.00
Begin VB.Form frmSnapShot 
   Caption         =   "Snap Shot"
   ClientHeight    =   5955
   ClientLeft      =   6780
   ClientTop       =   6750
   ClientWidth     =   5175
   Icon            =   "SnapShot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin SnapShot.ctlSnapShot SnapShot 
      Height          =   510
      Left            =   90
      TabIndex        =   2
      Top             =   30
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   900
   End
   Begin VB.TextBox txtfrmShapShot 
      Enabled         =   0   'False
      Height          =   5325
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "SnapShot.frx":058A
      Top             =   570
      Width           =   5025
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   510
      Left            =   1335
      TabIndex        =   0
      Top             =   30
      Width           =   1140
   End
End
Attribute VB_Name = "frmSnapShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

    Unload Me
    End
    
End Sub

