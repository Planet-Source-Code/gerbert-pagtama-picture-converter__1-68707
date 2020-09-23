VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Picture File Converter Editor"
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   2160
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   3135
      Left            =   120
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Picture File Converter-Editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   435
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5025
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   570
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program and Design by GERBERT PAGTAMA
'for more information Email : gerbert_p@yahoo.com
Private m_cDIB As cDIBSection
Private mGrad As New Gradients
Public mWhatGradient As Integer
Public mWhatColor As Integer

Dim Time_Open As Byte

Private Sub Form_Load()
  mWhatColor = WhatColor.Blue
  mGrad.Gradient02 Me, mWhatColor
  Timer1.Interval = 1
  Time_Open = 0
End Sub

Private Sub FrmAbout_Click()
 Unload Me
End Sub

Private Sub Timer1_Timer()
  Label1.Caption = Time_Open & "%"
 If Time_Open < 100 Then
    Time_Open = Time_Open + 1
    ElseIf Time_Open = 100 Then
           Unload Me
           FrmMain.Show
           Timer1.Interval = 0
           Timer1.Enabled = False
 End If
 Progress.Value = Time_Open
End Sub

