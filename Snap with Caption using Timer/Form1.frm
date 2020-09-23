VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   149
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   463
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   1170
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const SM_CXSMICON     As Long = 49
Private Const SM_CYCAPTION    As Long = 4
Private Const SM_CYFRAME      As Long = 33
Private Const SPI_GETWORKAREA As Long = 48
Private Const VK_LBUTTON      As Long = &H1

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinini As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

' Number of pixels from edge before snap occurs...
Private Const SNAP_GAP As Long = 30

Private m_bMouseDown As Boolean
Private m_rcSnapArea As RECT
Private m_ptOffset   As POINTAPI
Private m_cyCaption  As Long
Private m_cyBorder   As Long
Private m_cxIcon     As Long

Private Sub Form_Load()

    ' Ensure that the 'Movable' property of the form is set to False.  This property can only be set at
    ' design-time.  We'll handle all of the moving responsibilities for the form.

    ' Need these metrics to calculate the proper offset of our mouse click...
    m_cyCaption = GetSystemMetrics(SM_CYCAPTION)
    m_cyBorder = GetSystemMetrics(SM_CYFRAME)
    m_cxIcon = GetSystemMetrics(SM_CXSMICON)
    
    ' Set our snap area to be the current work area of the desktop...
    SystemParametersInfo SPI_GETWORKAREA, 0, m_rcSnapArea, 0

    ' Start listening for mousedown events on the caption bar...
    Timer1.Interval = 10
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    ' Ignore minimized or maximized windows...
    If Me.WindowState <> vbNormal Then Exit Sub

    ' The high bit determines if the button is down...
    Dim bDown As Boolean
    bDown = ((GetAsyncKeyState(VK_LBUTTON) And &H8000) = &H8000)
    
    If Not bDown Then
    
        ' If the mouse is not down now but previously was, reset our module vars...
        If m_bMouseDown Then
            
            m_bMouseDown = False
            m_ptOffset.x = 0
            m_ptOffset.y = 0
            ReleaseCapture
        
        End If
        
        ' If the mouse isn't down, just exit...
        Exit Sub
    
    End If
    
    ' The mouse is down.  If it was down previously, just keep sending MouseMove messages...
    If m_bMouseDown Then
    
        Form_MouseMove vbLeftButton, 0, 0, 0
        Exit Sub
        
    End If
    
    ' If we've made it to this point, the mouse is down but wasn't previously.  Check to see if it's over the caption bar...
    Dim pt As POINTAPI
    GetCursorPos pt
    ScreenToClient hwnd, pt
    
    ' If not, just exit...
    If pt.y < -m_cyCaption Or pt.y > 0 Or pt.x < m_cxIcon + m_cyBorder Or pt.x > ScaleWidth Then Exit Sub
        
    ' Otherwise, init our module vars...
    m_ptOffset = pt
    m_bMouseDown = True
    SetCapture hwnd
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Allow dragging with the left mouse button only...
    If Not m_bMouseDown Then Exit Sub
    
    ' Move the form relative to the current cursor position...
    Dim ptCursor As POINTAPI
    GetCursorPos ptCursor
    
    ' Calculate the new form position...
    Dim pt As POINTAPI
    pt.x = ptCursor.x - (m_cyBorder + m_ptOffset.x)
    pt.y = ptCursor.y - (m_cyCaption + m_cyBorder + m_ptOffset.y)
    
    ' Test the edges...
    If pt.x < m_rcSnapArea.Left + SNAP_GAP Then
        pt.x = m_rcSnapArea.Left
    ElseIf pt.x + ScaleX(Width, vbTwips, vbPixels) > m_rcSnapArea.Right - SNAP_GAP Then
        pt.x = m_rcSnapArea.Right - ScaleX(Width, vbTwips, vbPixels)
    End If
    
    If pt.y < m_rcSnapArea.Top + SNAP_GAP Then
        pt.y = m_rcSnapArea.Top
    ElseIf pt.y + ScaleY(Height, vbTwips, vbPixels) > m_rcSnapArea.Bottom - SNAP_GAP Then
        pt.y = m_rcSnapArea.Bottom - ScaleY(Height, vbTwips, vbPixels)
    End If
    
    ' Set the new form position...
    Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY

End Sub
