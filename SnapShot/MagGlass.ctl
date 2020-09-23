VERSION 5.00
Begin VB.UserControl ctlMagGlass 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   Enabled         =   0   'False
   FillColor       =   &H80000000&
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ToolboxBitmap   =   "MagGlass.ctx":0000
   Begin VB.TextBox txtDescription 
      Enabled         =   0   'False
      Height          =   3540
      Left            =   3855
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "MagGlass.ctx":0312
      Top             =   225
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.PictureBox picMagGlass 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      FillColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   0
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   41
         X2              =   11
         Y1              =   271
         Y2              =   271
      End
   End
End
Attribute VB_Name = "ctlMagGlass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mMyHost As Object
Private hDcMemHost As Long
Private hDcMemMag As Long
Private mMemBmpMag As cMemoryBmp
Private mMyWidth As Long
Private mMyHeight As Long
Private mLeftTop As Long
Private mLeftLeft As Long
Private mRightTop As Long
Private mRightLeft As Long
Private mAmLeft As Boolean
Private mMemBmpText As cMemoryBmp
Private hDcMemText As Long
Private mTextBackColor As Long

Public Sub MyHost(Host As Object, hDcMem As Long)
    
    Set mMyHost = Host                              'Host object
    hDcMemHost = hDcMem                             'Handle to Copy of Host's surface in memory
    mMyWidth = UserControl.Extender.Width
    mMyHeight = UserControl.Extender.Height
    mLeftTop = 10
    mLeftLeft = 10
    mRightTop = 10
    mRightLeft = mMyHost.Width / 15 - mMyWidth - 10
    UserControl.Extender.Top = mLeftTop
    UserControl.Extender.Left = mLeftLeft
    mAmLeft = True
    DoEvents
    
End Sub

Public Sub MouseMove(x As Integer, y As Integer)

    Dim StartX As Long
    Dim StartY As Long
    Dim rtn As Long
    Dim i As Long
    Dim UsePen As Long
    Dim Point As POINTAPI
    Dim Start As Long
    'Do we have to move
    If mAmLeft Then
        If x < mMyWidth + 12 And y < mMyHeight + 12 Then
            MoveMe
        End If
    Else
        If x > mRightLeft - 2 And y < mMyHeight + 12 Then
            MoveMe
        End If
    End If
    'Ok lets do the magnification
    StartX = x - 9                  'the MagGlass will display 19x19 pixels magnified x11
    StartY = y - 9
    rtn = StretchBlt(hDcMemMag, 12, 12, 210, 210, _
                     hDcMemHost, StartX, StartY, 19, 19, SRCCOPY)           'Copy Host Memory to MagGlass Memory
    For i = 0 To 210 Step 11                                                'with x11 stretch
        Start = i + 12
        rtn = BitBlt(hDcMemMag, Start, 12, 1, 209, hDcMemMag, Start, 12, DSTINVERT) 'put in grid
        rtn = BitBlt(hDcMemMag, 12, Start, 209, 1, hDcMemMag, 12, Start, DSTINVERT)
    Next
    rtn = BitBlt(hDcMemMag, 113, 116, 7, 1, hDcMemMag, 113, 116, DSTINVERT)         'put in crosshairs
    rtn = BitBlt(hDcMemMag, 116, 113, 1, 7, hDcMemMag, 116, 113, DSTINVERT)
    rtn = BitBlt(picMagGlass.hDC, 12, 12, 210, 210, hDcMemMag, 12, 12, SRCCOPY)     'copy to PictureBox
    DoText x, y
    picMagGlass.Refresh     'AutoReDraw is true therefore must be refreshed
    
End Sub

Private Sub DoText(x As Integer, y As Integer)

    Dim FillArea As RECT
    Dim PointColor As Long
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    Dim sText As String
    Dim rtn As Long
    Dim UseBrush As Long
    Dim UsePen As Long
    
    PointColor = GetPixel(hDcMemHost, CLng(x), CLng(y))         'Get Color of Current Point
    Red = PointColor And &HFF&                                  'Get RGB values
    Green = (PointColor And &HFF00&) / &H100&
    Blue = (PointColor And &HFF0000) / &H10000
    
    'Print out the x,y co-ordinates
    UseBrush = mMemBmpText.GetBrush(mTextBackColor)
    FillArea.Left = 76
    FillArea.Top = 227
    FillArea.Right = 160
    FillArea.Bottom = 260
    rtn = FillRect(picMagGlass.hDC, FillArea, UseBrush)
    sText = "x: " & Format(x, "0000") & " , y: " & Format(y, "0000")
    rtn = DrawText(picMagGlass.hDC, sText, Len(sText), FillArea, DT_LEFT)
    
    'print out the color info
    FillArea.Left = 46
    FillArea.Top = 243
    FillArea.Right = 230
    FillArea.Bottom = 258
    rtn = FillRect(picMagGlass.hDC, FillArea, UseBrush)
    mMemBmpText.DeleteBrush
    sText = "rgb(" & Format(Red, "000") & ", " & Format(Green, "000") & ", " & _
                     Format(Blue, "000") & ") - " & Format(PointColor, "#,##0")
    rtn = DrawText(picMagGlass.hDC, sText, Len(sText), FillArea, DT_LEFT)
    
    'print out color box
    UseBrush = mMemBmpText.GetBrush(PointColor)
    UsePen = mMemBmpText.GetPen(PS_SOLID, 1, vbBlack)
    rtn = Rectangle(hDcMemText, 0, 0, 15, 15)
    mMemBmpText.DeleteBrush
    mMemBmpText.DeletePen
    rtn = BitBlt(picMagGlass.hDC, 26, 243, 15, 15, hDcMemText, 0, 0, SRCCOPY)

End Sub
Private Sub MoveMe()
    
    If mAmLeft Then
        mAmLeft = False
        UserControl.Extender.Top = mRightTop
        UserControl.Extender.Left = mRightLeft
    Else
        mAmLeft = True
        UserControl.Extender.Top = mLeftTop
        UserControl.Extender.Left = mLeftLeft
    End If
    
End Sub

Private Sub UserControl_Resize()

    UserControl.Width = 3495
    UserControl.Height = 3945
    DrawMagGlass
    
End Sub

Private Sub DrawMagGlass()
    
    Dim clrFace As Long
    Dim clrShadow(2) As Long
    Dim clrHiLite(2) As Long
    Dim FaceRect As RECT
    
    clrFace = GetSysColor(COLOR_BTNFACE)
    clrShadow(1) = GetSysColor(COLOR_BTNSHADOW)
    clrShadow(2) = vbBlack
    clrHiLite(1) = RGB(195, 190, 226)
    clrHiLite(2) = GetSysColor(COLOR_BTNHIGHLIGHT)
    mMemBmpMag.Fill clrFace
        'draw outside borders raised
    mMemBmpMag.Fill clrHiLite(2), 0, 0, 233, 1
    mMemBmpMag.Fill clrHiLite(2), 0, 0, 1, 263
    mMemBmpMag.Fill clrHiLite(1), 1, 1, 232, 1
    mMemBmpMag.Fill clrHiLite(1), 1, 1, 1, 262
    mMemBmpMag.Fill clrShadow(2), 232, 0, 1, 263
    mMemBmpMag.Fill clrShadow(2), 0, 262, 232, 1
    mMemBmpMag.Fill clrShadow(1), 231, 1, 1, 261
    mMemBmpMag.Fill clrShadow(1), 1, 261, 231, 1
    mMemBmpMag.Fill vbWhite, 7, 7, 219, 219
        'draw inside borders recessed
    mMemBmpMag.Fill clrHiLite(2), 225, 7, 1, 219
    mMemBmpMag.Fill clrHiLite(2), 7, 225, 218, 1
    mMemBmpMag.Fill clrHiLite(1), 224, 8, 1, 217
    mMemBmpMag.Fill clrHiLite(1), 8, 224, 217, 1
    mMemBmpMag.Fill clrShadow(1), 7, 7, 218, 1
    mMemBmpMag.Fill clrShadow(1), 7, 7, 1, 218
    mMemBmpMag.Fill clrShadow(2), 8, 8, 216, 1
    mMemBmpMag.Fill clrShadow(2), 8, 8, 1, 216
        'draw black strokes to mark center pixel
    mMemBmpMag.Fill vbBlack, 9, 116, 3, 1
    mMemBmpMag.Fill vbBlack, 221, 116, 3, 1
    mMemBmpMag.Fill vbBlack, 116, 9, 1, 3
    mMemBmpMag.Fill vbBlack, 116, 221, 1, 3
    mMemBmpMag.Copy picMagGlass
    picMagGlass.Refresh
    
End Sub

Private Sub UserControl_Initialize()

    Set mMemBmpMag = New cMemoryBmp
    hDcMemMag = mMemBmpMag.Create(233, 263)                                     'Create MemBmp for MagGlass image
    Set mMemBmpText = New cMemoryBmp
    hDcMemText = mMemBmpText.Create(150, 30)                                    'Create MemBmp for Point Information
    mMemBmpText.SetFont picMagGlass
    mTextBackColor = GetSysColor(COLOR_BTNFACE)
                                                                                
End Sub

Private Sub UserControl_Terminate()

    Set mMemBmpMag = Nothing
    Set mMemBmpText = Nothing
    
End Sub
