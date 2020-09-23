VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmPrintPreview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Preview"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   13740
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   13515
      TabIndex        =   4
      Top             =   120
      Width           =   13575
      Begin VB.CommandButton Command6 
         Caption         =   "Original Size"
         Height          =   495
         Left            =   5400
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Fit to Window"
         Height          =   495
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   495
         Left            =   12240
         TabIndex        =   8
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Zoom IN +"
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Zoom Out -"
         Height          =   495
         Left            =   2760
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3240
      SmallChange     =   100
      TabIndex        =   3
      Top             =   8880
      Width           =   6735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7815
      Left            =   9960
      Max             =   255
      SmallChange     =   100
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      Height          =   8055
      Left            =   3240
      ScaleHeight     =   7980
      ScaleMode       =   0  'User
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   960
      Width           =   6735
      Begin VB.PictureBox picCurrent1 
         AutoSize        =   -1  'True
         Height          =   8475
         Left            =   -120
         ScaleHeight     =   8415
         ScaleWidth      =   6795
         TabIndex        =   1
         Top             =   -120
         Width           =   6855
         Begin VB.Image picCurrent 
            Height          =   7815
            Left            =   120
            Top             =   120
            Width           =   6615
         End
      End
   End
End
Attribute VB_Name = "FrmPrintPreview"
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
Dim Original_Width, Original_Height

Private Sub Form_Load()
  mWhatColor = WhatColor.Blue
  mGrad.Gradient02 Me, mWhatColor
  picCurrent = LoadPicture(Path_Data_PrintEditor)
  Original_Height = picCurrent.Height
  Original_Width = picCurrent.Width
  VScroll1.Max = picCurrent.Height
  HScroll1.Value = 0 'picCurrent.Left
  Call Scroll_down
End Sub

Private Sub Command1_Click()
  picCurrent.Stretch = True
  picCurrent.Height = picCurrent.Height - 50
  picCurrent.Width = picCurrent.Width - 50
  Call Scroll_down
End Sub

Private Sub Command2_Click() '+
  picCurrent.Stretch = True
  picCurrent.Height = picCurrent.Height + 50
  picCurrent.Width = picCurrent.Width + 50
  Call Scroll_down
End Sub

Private Sub Command3_Click()
  Dim strImgTmpFile
  CommonDialog1.ShowPrinter
  strImgTmpFile = "temp.bmp"
  SavePicture picCurrent.Picture, strImgTmpFile
  picCurrent.Picture = LoadPicture(strImgTmpFile)
  Kill strImgTmpFile
  Printer.PaintPicture picCurrent, 0, 0
  Printer.EndDoc
End Sub

Private Sub Command4_Click()
  Unload Me
End Sub

Private Sub Command5_Click()
 picCurrent.Stretch = True
 picCurrent.Height = 7815
 picCurrent.Width = 6615
 Call Scroll_down
End Sub

Private Sub Command6_Click()
   picCurrent.Stretch = True
   picCurrent.Height = Original_Height
   picCurrent.Width = Original_Width
   Call Scroll_down
End Sub

Private Sub HScroll1_Change()
      picCurrent.Left = -HScroll1.Value
End Sub

Private Sub VScroll1_Change()
      picCurrent.Top = -(VScroll1.Value)
End Sub

Function Scroll_down()
  If CDbl(picCurrent.Height) > 7815 Then
     VScroll1.Enabled = True
     Else
        VScroll1.Enabled = False
  End If
  If CDbl(picCurrent.Width) > 6615 Then
      HScroll1.Enabled = True
      Else
         HScroll1.Enabled = False
  End If
End Function
