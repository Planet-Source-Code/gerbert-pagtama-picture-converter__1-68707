VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About the Program"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton FrmAbout 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   4455
      Left            =   120
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_cDIB As cDIBSection
Private mGrad As New Gradients
Public mWhatGradient As Integer
Public mWhatColor As Integer

Private Sub Form_Load()
  mWhatColor = WhatColor.Blue
  mGrad.Gradient02 Me, mWhatColor
End Sub

Private Sub FrmAbout_Click()
 Unload Me
End Sub
