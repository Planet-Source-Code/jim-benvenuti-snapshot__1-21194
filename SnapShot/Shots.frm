VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShots 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   Caption         =   "Shots"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   675
   ClientWidth     =   8235
   Icon            =   "Shots.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtfrmShots 
      Height          =   5190
      Left            =   4410
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Shots.frx":058A
      Top             =   195
      Width           =   3390
   End
   Begin VB.Frame frSelShot 
      BackColor       =   &H80000009&
      Caption         =   "Selected Shot"
      Height          =   2025
      Left            =   60
      TabIndex        =   9
      Top             =   2250
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   330
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".bmp"
      Filter          =   "BitMaps (.bmp)|*.bmp|"
   End
   Begin VB.PictureBox picZoom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2445
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   35
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame frShotInfo 
      BackColor       =   &H80000009&
      Caption         =   "Selected Shot Info"
      Height          =   2085
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   2310
      Begin VB.TextBox txtFileName 
         Height          =   330
         Left            =   45
         TabIndex        =   10
         Top             =   1665
         Width           =   2160
      End
      Begin VB.TextBox txtHeight 
         Height          =   330
         Left            =   975
         TabIndex        =   6
         Top             =   1110
         Width           =   1155
      End
      Begin VB.TextBox txtWidth 
         Height          =   330
         Left            =   975
         TabIndex        =   4
         Top             =   675
         Width           =   1155
      End
      Begin VB.TextBox txtFrom 
         Height          =   330
         Left            =   975
         TabIndex        =   2
         Top             =   255
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "FileName"
         Height          =   270
         Left            =   45
         TabIndex        =   11
         Top             =   1470
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         Height          =   270
         Left            =   435
         TabIndex        =   7
         Top             =   1155
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   270
         Left            =   435
         TabIndex        =   5
         Top             =   705
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   270
         Left            =   435
         TabIndex        =   3
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.PictureBox picShot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   2520
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   105
      Width           =   300
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Save 
         Caption         =   "Save Capture"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete Capture"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Zoom 
      Caption         =   "Zoom"
      Begin VB.Menu x4 
         Caption         =   "x4"
      End
      Begin VB.Menu x8 
         Caption         =   "x8"
      End
      Begin VB.Menu x16 
         Caption         =   "x16"
      End
   End
End
Attribute VB_Name = "frmShots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mNewIndex As Integer
Dim CapInfo As cCapInfo
Dim CapCol As Collection
Dim mDefaultDir As String
Dim mLastSelected As Integer

Public Function NewShot(x As Long, y As Long, CapWidth As Long, CapHeight As Long) As PictureBox
    
    'frmSnap calls here to get a new container (picShot(n)) for a new Shot.
    'Once he copies the new Shot into the container he loads this form and Form_Activate gets
    'run which calls picShot_DblClick who makes the new Shot the Selected Shot.
    
    mNewIndex = mNewIndex + 1               'PicShot(0), the origional will not be used (I hate base 0)
    Load picShot(mNewIndex)              'Create a new PictureBox
    picShot(mNewIndex).Width = CapWidth
    picShot(mNewIndex).Height = CapHeight
    picShot(mNewIndex).Visible = True
    Set CapInfo = New cCapInfo
    CapInfo.CapX = x
    CapInfo.CapY = y
    CapInfo.CapWidth = CapWidth
    CapInfo.CapHeight = CapHeight
    CapInfo.CapFileName = ""
    CapCol.Add CapInfo, Str(mNewIndex)      'Make it a string so it is the key not the item #, Item #s can change
                                            'Keys can't
                                            
    Set NewShot = picShot(mNewIndex)  'This passes a reference to frmSnap so it can BitBlt the latest
                                            'Shot into it. When the form is Activated after the BitBlt then it will
                                            'be put into the Selected frame and the image into the clipboard. The is
                                            'done in picShot_DblClick which is run from Form_Activate and whenever you
                                            'DblClick on a Shot.
    
End Function


Private Sub Delete_Click()

    picShot(mLastSelected).Visible = False
    
End Sub

Private Sub Exit_Click()

    Dim i As Integer
    Set CapInfo = Nothing
    Set CapCol = Nothing
    mLastSelected = 0
    mNewIndex = 0
    Unload Me

End Sub

Private Sub Form_Activate()

    picShot_DblClick (mNewIndex)    'puts new Shot into Selected frame and the image onto the clipboard
    
End Sub

Private Sub Form_Load()

    Set CapCol = New Collection         'This is a Collection that holds Shot information.
    
End Sub

Private Sub picShot_DblClick(Index As Integer)
    
    Clipboard.Clear
    Clipboard.SetData picShot(Index).Image, vbCFBitmap                           'Save Selected Shot to clipboard
    txtFrom = CapCol.Item(Str(Index)).CapX & ", " & CapCol.Item(Str(Index)).CapY    'Display Shot Info
    txtWidth = CapCol.Item(Str(Index)).CapWidth
    txtHeight = CapCol.Item(Str(Index)).CapHeight
    txtFileName = CapCol.Item(Str(Index)).CapFileName
    Set picShot(mLastSelected).Container = frmShots       'Take the last selected Shot out of the
    picShot(mLastSelected).Top = 0                           'Selected frame and place it on the form
    picShot(mLastSelected).Left = 180
    frSelShot.Width = picShot(Index).Width + 40           'SetUp frSelShot to receive new selection
    If frSelShot.Width < 101 Then frSelShot.Width = 101
    frSelShot.Height = picShot(Index).Height + 40
    frSelShot.Top = 150
    frSelShot.Left = 4
    Set picShot(Index).Container = frSelShot
    picShot(Index).Left = 300    'The frame works in twips so the 20 must be multiplied by 15
    picShot(Index).Top = 300
    mLastSelected = Index           'Make this selection the last selection and put the frame on top of zorder
    frSelShot.ZOrder 0

End Sub

Private Sub picShot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

            'Visual Basic calls SetShot when the left mouse
            'button is pressed so call ReleaseShot so mouse
            'messages will be sent to Windows
            ReleaseCapture
            'Tell Windows the mouse was pressed in the caption
            'area (initiates dragging)
            SendMessage picShot(Index).hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub picZoom_DblClick()

    picZoom.Visible = False         'Just to get it out of the way
    
End Sub

Private Sub picZoom_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

            'Visual Basic calls SetShot when the left mouse
            'button is pressed so call ReleaseShot so mouse
            'messages will be sent to Windows
            ReleaseCapture
            'Tell Windows the mouse was pressed in the caption
            'area (initiates dragging)
            SendMessage picZoom.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&

End Sub

Private Sub Save_Click()
        
    Dim FileName As String
    On Error Resume Next
    CommonDialog.FileName = ""
    CommonDialog.ShowSave
    FileName = CommonDialog.FileName
    If Err.Number <> 32755 Then                                  'Cancel Button was clicked
        SavePicture picShot(mLastSelected).Image, FileName       'Write out the Shot to a .bmp file
    End If
    txtFileName = FileName
    On Error GoTo 0
    
End Sub


Private Sub x4_Click()

    DoZoom 4
    
End Sub

Private Sub x8_Click()

    DoZoom 8

End Sub

Private Sub x16_Click()

    DoZoom 16
    
End Sub

Private Sub DoZoom(ZoomFactor As Integer)

    Dim rtn As Long
    Dim DestWidth As Single
    Dim DestHeight As Single
    Dim SrcWidth As Single
    Dim srcHeight As Single
    picZoom.ZOrder 0
    picZoom.Picture = LoadPicture()
    picZoom.Visible = True
    SrcWidth = picShot(mLastSelected).Width / 15         'The frame works in twips
    srcHeight = picShot(mLastSelected).Height / 15
    picZoom.Width = SrcWidth * ZoomFactor
    picZoom.Height = srcHeight * ZoomFactor
    DestWidth = picZoom.Width
    DestHeight = picZoom.Height
    picZoom.PaintPicture picShot(mLastSelected).Image, 0, 0, DestWidth, _
                            DestHeight, 0, 0, SrcWidth, srcHeight, SRCCOPY
   
    picZoom.Refresh         'AutoReDraw is true so you must refresh.
    
End Sub


Private Sub txtfrmShots_Click()

    txtfrmShots.Visible = False
    
End Sub

