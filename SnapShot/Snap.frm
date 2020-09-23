VERSION 5.00
Begin VB.Form frmSnap 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5775
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7230
   Icon            =   "Snap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin SnapShot.ctlMagGlass MagGlass 
      Height          =   3945
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6959
   End
   Begin VB.TextBox txtfrmSnap 
      Enabled         =   0   'False
      Height          =   5520
      Left            =   3795
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Snap.frx":058A
      Top             =   105
      Visible         =   0   'False
      Width           =   3165
   End
End
Attribute VB_Name = "frmSnap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private fMouseDown As Boolean
Private mOldX As Long
Private mOldY As Long
Private mStartX As Integer
Private mStartY As Integer
Private mEndX As Integer
Private mEndY As Integer
Private mBeginX As Long
Private mBeginY As Long
Private mHeight As Long
Private mWidth As Long
Private mhDcMem As Long
Private mMemBmp As cMemoryBmp


Private Sub Form_Activate()

    MousePointer = 2        'Turn the cursor into a cross.

End Sub

Private Sub Form_Load()
    
    'The form is loaded with WindowState Maximized (2). Visible is false.
    'The first task is to get the DeskTop DC. Once we have that we can BitBlt a copy of the screen
    'onto the form. The forms AutoRedraw must be true since it is not visible yet.
    'Once we have the image on the full screen on the form we can make it visible. This happens with
    'not a flicker. At the end when we turn the visibility off we get a flicker because windows wants
    'to repaint the screen.
    'Once the form is visible then you can move the mouse to the spot you wish to capture, hold down the
    'left button and drag it to the end of your selection. When you let the button up the capture is
    'made (see Form_MouseUp)
    
    Dim Point As POINTAPI
    Dim rtn As Long
    Dim DeskhWnd As Long
    Dim DeskDC As Long
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Me.Left = 0
    Me.Top = 0
    mHeight = Me.Height / 15
    mWidth = Me.Width / 15
    Set mMemBmp = New cMemoryBmp
    mhDcMem = mMemBmp.Create(mWidth, mHeight)
    DeskhWnd = GetDesktopWindow()           'Get the hWnd of the desktop
    DeskDC = GetDC(DeskhWnd)                'Get Desktop Dc
    rtn = BitBlt(mhDcMem, 0, 0, mWidth, mHeight, DeskDC, 0, 0, SRCCOPY)  'Copy Screen image into memory
    rtn = ReleaseDC(DeskhWnd, DeskDC)
    rtn = BitBlt(Me.hDC, 0, 0, mWidth, mHeight, mhDcMem, 0, 0, SRCCOPY)    'Copy Screen image onto the Form
    MagGlass.MyHost Me, mhDcMem
    Me.Visible = True
    MagGlass.Visible = True
    rtn = GetCursorPos(Point)
    MagGlass.MouseMove CInt(Point.x), CInt(Point.y)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mMemBmp = Nothing
   
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        Unload Me
        Unload frmShots
        Exit Sub
    End If
    
    
    Dim rtn As Long
    Me.AutoRedraw = False       'This is necessary or else the BitBlts run quite jerky
    mStartX = x                  'Save the start co-ordinates
    mStartY = y
    mOldX = x
    mOldY = y
    fMouseDown = True
    
End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    Dim rtn As Long
    MagGlass.MouseMove Int(x), Int(y)                    'Send the MagGlass news of the move
    If fMouseDown Then
        If mOldX = mStartX And mOldY = mStartY Then     'This is run only once the first time and draws the
            If x > mStartX Then                         'initial Inverted rectangle.
                mBeginX = mStartX                       'The rest of the code is finding the top left co-ordinate
                mWidth = x - mStartX                    'of the rectangle along with the width and height.
            Else
                mBeginX = x                             'I know that simple subtraction sometimes giving negative values
                mWidth = mStartX - x                    'and offsets would probably work, call me old fashion but
            End If                                      'I like to work with positive numbers, it makes debugging
            If y > mStartY Then                         'soooo much easier
                mBeginY = mStartY
                mHeight = y - mStartY
            Else
                mBeginY = y
                mHeight = mStartY - y
            End If
            rtn = BitBlt(Me.hDC, mBeginX, mBeginY, mWidth, mHeight, Me.hDC, mBeginX, mBeginY, DSTINVERT) 'Destination Inverted
        Else
            If x > mOldX Then               'This code is run after the first time and there after. The first section
                mBeginX = mOldX             'finds the top left rectangle on the side, that is the rectangle caused by the
                mWidth = x - mOldX          'change in x. The rectangle formed from the start co-ordinated (mBeginX and
            Else                            'mBeginY, always the top left co-ordinate) and the width (change in x) and
                mBeginX = x                 'the height (the change in y from the start of the selection. This rectangle
                mWidth = mOldX - x          'includes the portion where the two rectangles overlap.
            End If
            If y > mStartY Then
                mBeginY = mStartY
                mHeight = y - mStartY
            Else
                mBeginY = y
                mHeight = mStartY - y
            End If
            rtn = BitBlt(Me.hDC, mBeginX, mBeginY, mWidth, mHeight, Me.hDC, mBeginX, mBeginY, DSTINVERT)
            If mOldX > mStartX Then
                mBeginX = mStartX           'This second part is the rectangle formed at the top or bottom by the change
                mWidth = mOldX - mStartX    'in y. The co-ordinate (mBeginX,mBeginY is again the top left corner of the
            Else                            'rectangle. the height is the change in y and the width is the change in x
                mBeginX = mOldX             'from the last move (mOldX). The area where the two rectangles overlap
                mWidth = mStartX - mOldX    'was handle above in the change in x rectangle.
            End If
            If y > mOldY Then               'When you move the mouse back over terain covered it just does an inversion
                mBeginY = mOldY             'of the invert which leaves you with the original.
                mHeight = y - mOldY
            Else
                mBeginY = y
                mHeight = mOldY - y
            End If
            rtn = BitBlt(Me.hDC, mBeginX, mBeginY, mWidth, mHeight, Me.hDC, mBeginX, mBeginY, DSTINVERT)
        End If
        
        mOldX = x
        mOldY = y
    End If

End Sub

Public Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim rtn As Long
    Dim NewShot As PictureBox
    MousePointer = 0
    mEndX = x
    mEndY = y
    fMouseDown = False
                                                                                                
    If mEndX > mStartX Then               'Get the start co-ordinate (top left) and the width and height
        mBeginX = mStartX                 'of the selection
        mWidth = mEndX - mStartX + 1
    Else
        mBeginX = mEndX
        mWidth = mStartX - mEndX + 1
    End If
    If mEndY > mStartY Then
        mBeginY = mStartY
        mHeight = mEndY - mStartY + 1
    Else
        mBeginY = mEndY
        mHeight = mStartY - mEndY + 1
    End If
    Set NewShot = frmShots.NewShot(mBeginX, mBeginY, mWidth, mHeight)  'Go over to the form that saves
                                                                                'the shots (frmShots) and get a
                                                                                'container (PictureBox). Use the
                                                                                'opertunity to pass the info about
                                                                                'the selection.
    rtn = BitBlt(NewShot.hDC, 0, 0, mWidth, mHeight, _
                 mhDcMem, mBeginX, mBeginY, SRCCOPY)                            'Do the copy into the PictureBox
    Unload Me                                                                   'Unload me and show the shots
                                                                                'to date.
    frmShots.WindowState = 0                                                 'Put frmShots into normal WindowState
    frmShots.Show                                                            'and show him
        
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim x As Long
    Dim y As Long
    Dim rtn As Long
    Dim Point As POINTAPI
    rtn = GetCursorPos(Point)
    x = Point.x
    y = Point.y
    If KeyCode = 37 Then    'left
        x = x - 1
        rtn = SetCursorPos(x, y)
    End If
    If KeyCode = 39 Then    'right
        x = x + 1
        rtn = SetCursorPos(x, y)
    End If
    If KeyCode = 38 Then    'up
        y = y - 1
        rtn = SetCursorPos(x, y)
    End If
    If KeyCode = 40 Then    'down
        y = y + 1
        rtn = SetCursorPos(x, y)
    End If

End Sub
