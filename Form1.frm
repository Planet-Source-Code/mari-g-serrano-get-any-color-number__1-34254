VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Number"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picCrossHair 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   375
      MouseIcon       =   "Form1.frx":08CA
      Picture         =   "Form1.frx":1194
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   15
      Width           =   495
   End
   Begin VB.TextBox txtColor 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   75
      Width           =   1560
   End
   Begin VB.CommandButton cmdDlg 
      Caption         =   "Color Dialog"
      Height          =   510
      Left            =   3435
      TabIndex        =   0
      Top             =   45
      Width           =   1185
   End
   Begin MSComDlg.CommonDialog c 
      Left            =   2565
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the Icon to get a Color Number"
      Height          =   525
      Left            =   0
      TabIndex        =   3
      Top             =   540
      Width           =   1680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' This is a very simple Program to get any color Number
' i think it is usefull. MaRiØ 30/4/02

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private lDesktopDC As Long
Private m_bDragging As Boolean

Private Sub cmdDlg_Click()
    On Error Resume Next
    c.Color = CLng(txtColor.Text)
    c.ShowColor
    txtColor.Text = CLng(c.Color)
    txtColor.BackColor = txtColor.Text
End Sub

Private Sub Form_Load()
    lDesktopDC = GetDC(&H0)
    c.Flags = &H2 Or &H1 'cdCCFullOpen Or cdlCCRGBInit
End Sub
Private Sub picCrossHair_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are not dragging
    If Button = vbLeftButton And Not m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = True
        ' Set mouse pointer
        Me.MouseIcon = picCrossHair.MouseIcon
        Me.MousePointer = 99
        ' Erase picture from picCrossHair
        picCrossHair.Picture = Nothing
    End If
End Sub

Private Sub picCrossHair_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        Dim tPA As POINTAPI
        ' Get cursor cordinates
        GetCursorPos tPA
        txtColor.BackColor = GetPixel(lDesktopDC, tPA.X, tPA.Y)
        txtColor.Text = txtColor.BackColor
    End If
End Sub

Private Sub picCrossHair_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If user pressed left mouse button and we are dragging
    If Button = vbLeftButton And m_bDragging Then
        ' Set dragging flag to true
        m_bDragging = False
        ' Restore mouse pointer to normal (arrow)
        Me.MousePointer = vbNormal
        ' Load picture into picCrossHair
        picCrossHair.Picture = picCrossHair.MouseIcon
    End If
End Sub
