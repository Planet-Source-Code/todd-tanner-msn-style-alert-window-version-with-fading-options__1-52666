VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Alert"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   2415
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   3000
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Fade Out"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Fade In"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   1605
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Alert"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    Dim Alert As frmAlert
    Set Alert = New frmAlert
    Caption = Alert.DisplayAlert(Text1.Text, , CBool(Check1.Value), CBool(Check2.Value))
    Set Alert = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
