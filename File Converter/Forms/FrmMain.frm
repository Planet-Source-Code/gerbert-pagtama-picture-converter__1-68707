VERSION 5.00
Begin VB.Form FrmMain 
   BorderStyle     =   0  'None
   Caption         =   "Picture File Converter"
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   360
      Width           =   5025
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00E0E0E0&
      Height          =   3975
      Left            =   120
      Top             =   120
      Width           =   5295
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   1320
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label LblBttn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   435
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Exit Program"
      Top             =   3240
      Width           =   585
   End
   Begin VB.Label LblBttn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   435
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      ToolTipText     =   "Edit Picture File"
      Top             =   2280
      Width           =   1845
   End
   Begin VB.Label LblBttn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Convert Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   435
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "Convert Picture File"
      Top             =   1320
      Width           =   2460
   End
End
Attribute VB_Name = "FrmMain"
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

Private Sub Form_Load()
  mWhatColor = WhatColor.Blue
  mGrad.Gradient02 Me, mWhatColor
End Sub

Private Sub Image4_Click()

End Sub

Private Sub LblBttn_Click(Index As Integer)
  Select Case Index
         Case 0
              FileConverter.Show
              Unload Me
         Case 1
              FrmPictureEditor.Show
              Unload Me
         Case 2
              End
  End Select
 
End Sub

Private Sub LblBttn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case Index
         Case 0
              For i = 0 To 2
                  LblBttn.Item(i).ForeColor = &HFFFFC0
              Next i

              LblBttn.Item(0).ForeColor = &HC0FFFF
         Case 1
              For i = 0 To 2
                  LblBttn.Item(i).ForeColor = &HFFFFC0
              Next i
              LblBttn.Item(1).ForeColor = &HC0FFFF
         Case 2
              For i = 0 To 2
                  LblBttn.Item(i).ForeColor = &HFFFFC0
              Next i
              LblBttn.Item(2).ForeColor = &HC0FFFF
  End Select
End Sub
