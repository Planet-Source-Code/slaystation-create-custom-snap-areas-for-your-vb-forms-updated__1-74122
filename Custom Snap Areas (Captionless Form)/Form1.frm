VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C25B19&
   BorderStyle     =   0  'None
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6180
   ControlBox      =   0   'False
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustom 
      BackColor       =   &H00C25B19&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3840
      TabIndex        =   6
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtCustom 
      BackColor       =   &H00C25B19&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtCustom 
      BackColor       =   &H00C25B19&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtCustom 
      BackColor       =   &H00C25B19&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C25B19&
      Caption         =   "Custom:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C25B19&
      Caption         =   "Use work area"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C25B19&
      Caption         =   "Use full screen"
      ForeColor       =   &H8000000E&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag me toward an edge.  Hit ESC to close."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1530
      TabIndex        =   7
      Top             =   240
      Width           =   3120
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

Private Const SPI_GETWORKAREA As Long = 48

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinini As Long) As Long

' Number of pixels from edge before snap occurs...
Private Const SNAP_GAP As Long = 30

Private m_bMouseDown As Boolean
Private m_rcSnapArea As RECT
Private m_ptOffset   As POINTAPI

Private Sub Form_Load()
    
    ' Default some values for the custom area...
    txtCustom(0).Text = Int(ScaleX(Screen.Width, vbTwips, vbPixels) / 6)
    txtCustom(1).Text = Int(ScaleY(Screen.Height, vbTwips, vbPixels) / 4)
    txtCustom(2).Text = Int(txtCustom(0).Text * 5)
    txtCustom(3).Text = Int(txtCustom(1).Text * 3)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Allow dragging with the left mouse button only...
    If Button <> vbLeftButton Then Exit Sub

    ' Update the bounding rect...
    If Option1(0) Then
        SetArea 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
    ElseIf Option1(1) Then
        SystemParametersInfo SPI_GETWORKAREA, 0, m_rcSnapArea, 0
    Else
        SetArea txtCustom(0), txtCustom(1), txtCustom(2), txtCustom(3)
    End If
    
    ' Save the cursor's position within our form...
    GetCursorPos m_ptOffset
    ScreenToClient Me.hwnd, m_ptOffset

    m_bMouseDown = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Allow dragging with the left mouse button only...
    If Button <> vbLeftButton Or Not m_bMouseDown Then Exit Sub
    
    ' Move the form relative to the current cursor position...
    Dim ptCursor As POINTAPI
    GetCursorPos ptCursor
    
    ' Calculate the new form position...
    Dim pt As POINTAPI
    pt.x = ptCursor.x - m_ptOffset.x
    pt.y = ptCursor.y - m_ptOffset.y
    
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

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> vbLeftButton Then Exit Sub
    m_bMouseDown = False

End Sub

Private Sub Option1_Click(Index As Integer)

    ' Enable/Disable the custom coord section as needed...
    Dim i As Long
    For i = 0 To txtCustom.UBound
    
        txtCustom(i).Enabled = (Index = 2)
        txtCustom(i).BackColor = IIf(Index = 2, vbWhite, &HC25B19)
    
    Next

End Sub

Private Sub SetArea(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)

    m_rcSnapArea.Left = x1
    m_rcSnapArea.Top = y1
    m_rcSnapArea.Right = x2
    m_rcSnapArea.Bottom = y2

End Sub

Private Sub txtCustom_GotFocus(Index As Integer)

    ' Highlight on entry...
    txtCustom(Index).SelStart = 0
    txtCustom(Index).SelLength = 99

End Sub

Private Sub txtCustom_LostFocus(Index As Integer)

    ' Ensure that the right/bottom coords are at least large enough to hold the size of the form...
    txtCustom(2).Text = Max(txtCustom(2).Text, txtCustom(0).Text + ScaleX(Width, vbTwips, vbPixels))
    txtCustom(3).Text = Max(txtCustom(3).Text, txtCustom(1).Text + ScaleY(Height, vbTwips, vbPixels))

End Sub

Private Function Max(ByVal i As Long, ByVal j As Long) As Long

    If i > j Then Max = i Else Max = j

End Function

