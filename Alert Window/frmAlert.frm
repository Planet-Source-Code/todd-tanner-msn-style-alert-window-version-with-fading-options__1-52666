VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1755
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2355
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   30
      ScaleHeight     =   1635
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   30
      Width           =   2295
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1980
         Picture         =   "frmAlert.frx":0000
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   2
         Top             =   60
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   30
         TabIndex        =   1
         Top             =   270
         Width           =   2175
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   1560
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim mCounter As Long
Dim mWasClicked As Boolean
Dim MoveThis As Integer
Dim mOpenAlerts As Integer
Dim mFadeOut As Boolean
Dim mFadeIn As Boolean
Dim mCurTranLvl As Integer

Private Sub Form_Load()
    MoveThis = 50 ' Changes the windows moving speed
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Visible = False
End Sub
Private Sub Label1_Click()
    mWasClicked = True
End Sub
Private Sub Picture2_Click()
    Timer1.Enabled = False
    Timer2.Enabled = True
    Me.Visible = False
End Sub
Private Sub Timer1_Timer() ' Grow Timer
    Dim WindowRect As RECT
    If Not mFadeIn Then
        SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
        TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX
        If (Me.Top + (Me.Height * mOpenAlerts) + TaskBar) > Screen.Height Then
            Me.Top = Me.Top - MoveThis
        Else
            mCounter = mCounter - 1
            If mCounter = 0 Then
                Timer1.Enabled = False
                Timer2.Enabled = True
            End If
        End If
    Else
        If mCurTranLvl > 0 Then
            mCurTranLvl = mCurTranLvl - 1
            modMakeTransparent.MakeTransparent Me.hWnd, mCurTranLvl
        Else
            mCounter = mCounter - 1
            If mCounter = 0 Then
                Timer1.Enabled = False
                Timer2.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub Timer2_Timer() ' Shrink Timer / fadeout Timer
    Dim WindowRect As RECT
    If Not mFadeOut Then
        SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
        TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX
        Me.Top = Me.Top + MoveThis
        If Me.Top < Screen.Height And Me.Visible Then
            Me.Top = Me.Top + MoveThis
        Else
            Timer2.Enabled = False
            setTop Me.hWnd, False
            OpenAlerts = OpenAlerts - 1
            Me.Visible = False
        End If
    Else
        If mCurTranLvl < 100 Then
            mCurTranLvl = mCurTranLvl + 1
            modMakeTransparent.MakeTransparent Me.hWnd, mCurTranLvl
        Else
            Timer2.Enabled = False
            setTop Me.hWnd, False
            OpenAlerts = OpenAlerts - 1
            Me.Visible = False
        End If
    End If
End Sub
Public Function DisplayAlert(ByVal mMsg As String, Optional ByVal Duration As Long = 300, Optional ByVal FadeIn As Boolean = False, Optional ByVal FadeOut As Boolean = False) As Boolean
    ' FadeIn and FadeOut only tested on Windows XP/2000
    Dim ClsGradient As cGradient
    Dim TaskBar As Long
    Dim WindowRect As RECT
    Set ClsGradient = New cGradient
    OpenAlerts = OpenAlerts + 1
    mOpenAlerts = OpenAlerts
    Me.Left = Screen.Width - Me.Width
    mCounter = Duration
    mWasClicked = False
    mFadeOut = FadeOut
    mFadeIn = FadeIn
    If mFadeIn Then 'Set up for Fade In if needed
        mCurTranLvl = 100
        SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
        TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.Bottom) * Screen.TwipsPerPixelX
        Me.Top = Screen.Height - TaskBar - (Me.Height * mOpenAlerts)
        modMakeTransparent.MakeTransparent Me.hWnd, mCurTranLvl
    Else
        Me.Top = Screen.Height
    End If
    Label1.Caption = mMsg
    ' Draw the gradient background
    With ClsGradient
        .Angle = -100
        .Color1 = RGB(61, 149, 255)
        .Color2 = RGB(255, 255, 255)
        .Draw Picture1
    End With
    Picture1.Refresh
    Me.Visible = True
    setTop Me.hWnd, True
    Timer1.Enabled = True
    While Me.Visible And Not mWasClicked
        DoEvents
    Wend
    DisplayAlert = mWasClicked
End Function

