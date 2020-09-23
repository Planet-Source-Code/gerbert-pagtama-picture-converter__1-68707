VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FileConverter 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Converter"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14565
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9285
   ScaleWidth      =   14565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Select Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   240
      TabIndex        =   44
      Top             =   6720
      Width           =   3015
      Begin VB.PictureBox picColourReductionOptions 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1455
         ScaleWidth      =   2475
         TabIndex        =   45
         Top             =   120
         Width           =   2475
         Begin VB.OptionButton optReduceMethod 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Floyd-Stucci"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   51
            Top             =   360
            Width           =   2835
         End
         Begin VB.OptionButton optReduceMethod 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Default"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   50
            Top             =   120
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.OptionButton optReduceMethod 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Optimal Palette"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   49
            Top             =   1080
            Width           =   2835
         End
         Begin VB.PictureBox picFloydStucciOptions 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   240
            ScaleHeight     =   495
            ScaleWidth      =   2715
            TabIndex        =   46
            Top             =   600
            Width           =   2715
            Begin VB.OptionButton optFloydStucciType 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "&Halftone"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFC0&
               Height          =   255
               Index           =   0
               Left            =   60
               TabIndex        =   48
               Top             =   0
               Width           =   2355
            End
            Begin VB.OptionButton optFloydStucciType 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               Caption         =   "&Web Safe"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFC0&
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   47
               Top             =   240
               Width           =   2355
            End
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Output Colour Depth:"
      Height          =   1215
      Left            =   240
      TabIndex        =   38
      Top             =   4560
      Width           =   3135
      Begin VB.PictureBox picColourDepthOptions 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   240
         ScaleHeight     =   1095
         ScaleWidth      =   2535
         TabIndex        =   39
         Top             =   120
         Width           =   2535
         Begin VB.OptionButton optColourDepth 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&Black and White"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   2835
         End
         Begin VB.OptionButton optColourDepth 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&16 Colour"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   42
            Top             =   240
            Width           =   2835
         End
         Begin VB.OptionButton optColourDepth 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&256 Colour"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   2
            Left            =   0
            TabIndex        =   41
            Top             =   480
            Width           =   2835
         End
         Begin VB.OptionButton optColourDepth 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            Caption         =   "&True Colour"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   315
            Index           =   3
            Left            =   0
            TabIndex        =   40
            Top             =   720
            Value           =   -1  'True
            Width           =   2835
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   240
      TabIndex        =   30
      Top             =   1680
      Width           =   3135
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Windows MetaFile (*wmf)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Jpeg Images (*.jpeg)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Tiff Images (*.tiff)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Bitmap Images (*.bmp)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "True Vision Images (*.tga)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Cumpuserve  Images (*.gif)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Portable Net. Graphics (*.png)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   13320
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Convert"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8760
      Width           =   3255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   8295
      Left            =   10680
      Max             =   255
      SmallChange     =   100
      TabIndex        =   4
      Top             =   360
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   3600
      SmallChange     =   100
      TabIndex        =   3
      Top             =   8760
      Width           =   6975
   End
   Begin VB.PictureBox Picture3 
      Height          =   8055
      Left            =   3720
      ScaleHeight     =   7980
      ScaleMode       =   0  'User
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   480
      Width           =   6735
      Begin VB.PictureBox picCurrent1 
         AutoSize        =   -1  'True
         Height          =   8475
         Left            =   -120
         ScaleHeight     =   8415
         ScaleWidth      =   6795
         TabIndex        =   20
         Top             =   -120
         Width           =   6855
         Begin VB.PictureBox picCurrent 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   7920
            Left            =   120
            ScaleHeight     =   7920
            ScaleWidth      =   6645
            TabIndex        =   21
            Top             =   120
            Width           =   6645
         End
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   4200
      ScaleHeight     =   315
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   5235
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   240
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   240
      TabIndex        =   52
      Top             =   8160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00E0E0E0&
      Height          =   8295
      Left            =   3600
      Top             =   360
      Width           =   6975
   End
   Begin VB.Label LblPica 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   12600
      TabIndex        =   59
      Top             =   2520
      Width           =   105
   End
   Begin VB.Label lblInches 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   12600
      TabIndex        =   58
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   360
      Left            =   11280
      TabIndex        =   56
      Top             =   6480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00E0E0E0&
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      X1              =   120
      X2              =   3360
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output Colour Reduction Method"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   120
      TabIndex        =   55
      Top             =   6120
      Width           =   3345
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   3480
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output Colour Depth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   120
      TabIndex        =   54
      Top             =   3960
      Width           =   2085
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   120
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output Format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   240
      TabIndex        =   53
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   13560
      TabIndex        =   29
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   12360
      TabIndex        =   28
      Top             =   5040
      Width           =   120
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   735
      Left            =   11280
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X = "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11880
      TabIndex        =   27
      Top             =   5040
      Width           =   390
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   13080
      TabIndex        =   26
      Top             =   5040
      Width           =   345
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scaling and Point Coordinates"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11280
      TabIndex        =   25
      Top             =   4320
      Width           =   3150
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   11280
      X2              =   14400
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00E0E0E0&
      Height          =   2895
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   7
      Left            =   12600
      TabIndex        =   24
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   9
      Left            =   12600
      TabIndex        =   19
      Top             =   2160
      Width           =   105
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   8
      Left            =   12600
      TabIndex        =   18
      Top             =   3840
      Width           =   105
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   0
      Left            =   12600
      TabIndex        =   17
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   1
      Left            =   12600
      TabIndex        =   16
      Top             =   1440
      Width           =   105
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   3
      Left            =   12240
      TabIndex        =   15
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   4
      Left            =   12240
      TabIndex        =   14
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   6
      Left            =   12360
      TabIndex        =   13
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "File Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11520
      TabIndex        =   12
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inches :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11760
      TabIndex        =   11
      Top             =   2880
      Width           =   675
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pica :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11760
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pixel :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11760
      TabIndex        =   9
      Top             =   2160
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Images size :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11400
      TabIndex        =   8
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "File Type :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11400
      TabIndex        =   7
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "File Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11400
      TabIndex        =   6
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   11280
      X2              =   14280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "File Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   240
      Left            =   11280
      TabIndex        =   5
      Top             =   360
      Width           =   1605
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu SelectFile 
         Caption         =   "&Select Picture"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Back to Main Menu"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      Begin VB.Menu About 
         Caption         =   "&About the Program"
      End
      Begin VB.Menu Instruction 
         Caption         =   "&Instruction"
      End
   End
End
Attribute VB_Name = "FileConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program and Design by GERBERT PAGTAMA
'for more information Email : gerbert_p@yahoo.com


Option Explicit
Private m_cDIB As cDIBSection
Dim WithEvents cGif As Class
Attribute cGif.VB_VarHelpID = -1
Dim FileExt As String
Dim Original_Path, Convert_Path

Private mGrad As New Gradients
Public mWhatGradient As Integer
Public mWhatColor As Integer

Dim pressed As Boolean
Dim colpressed As Boolean
Dim point1 As pointapi
Dim point2 As pointapi
Dim New_FileExt, opt1_enb1, opt1_enb

Private Type pointapi
    X As Double
    Y As Double
End Type


Private Property Get SelectedColourDepth() As EDSSColourDepthConstants
   Select Case True
   Case optColourDepth(0).Value
      SelectedColourDepth = edss2Colour
   Case optColourDepth(1).Value
      SelectedColourDepth = edss16Colour
   Case optColourDepth(2).Value
      SelectedColourDepth = edss256Colour
   Case optColourDepth(3).Value
      SelectedColourDepth = edssTrueColour
   End Select
End Property



Private Sub About_Click()
 FrmAbout.Show
End Sub

Private Sub Command2_Click()
 SelectFile_Click
End Sub

Private Sub Exit_Click()
  FrmMain.Show
  Unload Me
End Sub

Private Sub Form_Load()
    optColourDepth_Click 3
    Set m_cDIB = New cDIBSection
    picCurrent.ScaleWidth = 255
    picCurrent.ScaleHeight = 255
    HScroll1.Enabled = False
    VScroll1.Enabled = False
    picCurrent.AutoSize = True
    picCurrent.ScaleMode = vbPixels
    Picture2.BorderStyle = 1
    picCurrent.BorderStyle = 0
    
     mWhatGradient = WhatGradient.Gradient01
    mWhatColor = WhatColor.Blue
    mGrad.Gradient02 Me, mWhatColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
 FrmMain.Show
End Sub

Private Sub Instruction_Click()
 FrmHelpConverter.Show
End Sub

Private Sub picCurrent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Label17.Caption = X
   Label11.Caption = Y
End Sub

Private Sub SelectFile_Click()
On Error Resume Next
   Dim pic_w, pic_h
   Dim oPic As New StdPicture
   Dim s As String, sFile As String
   Dim chk_check
   Dim ValueTrue, ValueTrue1
   CommonDialog2.InitDir = App.Path
   CommonDialog2.Filter = "All Picture Files| *.png; *.gif; *.cel; *.pic; *.cut; *.pal; *.tga; *.pcx; *.bmp; *.tiff; *.jpg; *.jpeg; *.wmf|Portable Network Grphics  (*.png)|*.png|CompuServe Images  (*.gif)|*.gif|Autodesk Images  (*.cel,*.pic)|*.cel,*pic|DrHalo Images  (*.cut,*.pal)|*.cut,*.pal|True Vision Images  (*.tga)|*.tga|ZSoft Paintbrush  (*.pcx)|*.pcx|Bitmap Images  (*.bmp)|*.bmp|Tiff Images  (*.Tiff)|*.Tiff|Jpeg Images  (*.jpeg)|*.jpeg|Windos MetaFile  (*.wmf)|*.wmf|"
   CommonDialog2.DialogTitle = "Open File"
   CommonDialog2.ShowOpen
   cmdSave.Enabled = True
    
   For ValueTrue = 0 To 6
     Option1(ValueTrue).Enabled = True
   Next ValueTrue
   For ValueTrue1 = 0 To 3
      optColourDepth(ValueTrue1).Enabled = True
   Next ValueTrue1
   
   Set oPic = LoadPicture(Trim(CommonDialog2.FileName))
   
   m_cDIB.CreateFromPicture oPic
   picCurrent.Refresh
   
   sFile = Trim(CommonDialog2.FileName)
   s = CInt(FileLen(sFile) / 1024) & "K"
   With picCurrent
       '.AutoRedraw = True
       .FontBold = True
       .Picture = LoadPicture(sFile)
       .CurrentX = picCurrent.Width / 2 - .TextWidth(s) / 2 + 480
       .CurrentY = picCurrent.Height / 2 - .TextHeight(s) / 2
       .ForeColor = vbRed
   End With
   
   Lbl(7).Caption = s 'File Size
   
   HScroll1.Value = 0
   VScroll1.Value = 0
   If picCurrent.Width < Picture3.Width Then
      HScroll1.Enabled = False
      Else: HScroll1.Enabled = True
   End If

   If picCurrent.Height < Picture3.Height Then
       VScroll1.Enabled = False
        Else: VScroll1.Enabled = True
   End If
   
   Lbl(0).Caption = CommonDialog2.FileTitle
   Lbl(9).Caption = picCurrent.ScaleWidth & "  X " & picCurrent.ScaleHeight
   
   Lbl(1).Caption = Split(CommonDialog2.FileName, ".")(1)
   
  ' file  convertion
   pic_w = CDbl((picCurrent.ScaleWidth / 795) * 10)
   
   pic_h = CDbl((picCurrent.ScaleHeight / 795) * 10)
   pic_w = Round(pic_w, 2)
   pic_h = Round(pic_h, 2)
   
   
   
   lblInches.Caption = pic_h & " x " & pic_w
      
      
   
   lblPica.Caption = Round((pic_h * 0.167), 2) & " x " & Round((pic_w * 0.167), 2)
   'end of convertion
   
   For chk_check = 0 To 6
       Option1(chk_check).Enabled = True
   Next chk_check
   For opt1_enb1 = 0 To 3
     optColourDepth.Item(opt1_enb1).Enabled = True
   Next opt1_enb1
   optReduceMethod.Item(0).Enabled = True
   optReduceMethod.Item(1).Enabled = True
   optReduceMethod.Item(2).Enabled = True
   optFloydStucciType.Item(0).Enabled = True
   optFloydStucciType.Item(1).Enabled = True
   
   If Lbl(1) = "png" Then
      Option1(0).Enabled = False
      Option1(5).Value = True
      ElseIf Lbl(1) = "gif" Then
             Option1(1).Enabled = False
             Option1(5).Value = True
      ElseIf Lbl(1) = "tga" Then
             Option1(2).Enabled = False
             Option1(5).Value = True
      ElseIf Lbl(1) = "bmp" Then
             Option1(3).Enabled = False
             Option1(5).Value = True
      ElseIf Lbl(1) = "tiff" Then
             Option1(4).Enabled = False
             Option1(5).Value = True
      ElseIf Lbl(1) = "jpeg" Or Lbl(1) = "jpg" Then
             Option1(5).Enabled = False
             Option1(1).Value = True
      ElseIf Lbl(1) = "wmf" Then
             Option1(6).Enabled = False
             Option1(5).Value = True
   End If
   
   
End Sub

Private Sub VScroll1_Change()
    VScroll1.Max = picCurrent.Height - Picture3.Height
    picCurrent.Top = -VScroll1.Value
End Sub

Private Sub HScroll1_Change()
    HScroll1.Max = picCurrent.Width - Picture3.Width
    picCurrent.Left = -HScroll1.Value
End Sub

Private Sub optColourDepth_Click(Index As Integer)
Dim i As Long
   If optColourDepth(3).Value Then
      optReduceMethod(0).Value = True
      For i = 1 To 2
         optReduceMethod(i).Enabled = False
      Next i
   Else
      optReduceMethod(0).Value = True
      optReduceMethod(1).Enabled = True
      optReduceMethod(2).Enabled = (optColourDepth(2).Value)
   End If
   optFloydStucciType(0).Enabled = optColourDepth(2).Value And optReduceMethod(1).Value
   optFloydStucciType(1).Enabled = optColourDepth(2).Value And optReduceMethod(1).Value
   If Not (optColourDepth(2).Value) Then
      optFloydStucciType(0).Value = False
      optFloydStucciType(1).Value = False
   Else
      If Not (optFloydStucciType(0).Value Or optFloydStucciType(1).Value) Then
         optFloydStucciType(0).Value = True
      End If
   End If
End Sub
Private Sub optReduceMethod_Click(Index As Integer)
   optFloydStucciType(0).Enabled = optColourDepth(2).Value And optReduceMethod(1).Value
   optFloydStucciType(1).Enabled = optColourDepth(2).Value And optReduceMethod(1).Value

End Sub

Private Sub picCurrent_Paint()
   If Not m_cDIB Is Nothing Then
      m_cDIB.PaintPicture picCurrent.hDC
   End If

End Sub

Private Sub cmdSave_Click()

   Label20.Visible = True
   MsgBox "Picture Conversion May Take a Few Minutes!" + (Chr$(13) + Chr$(10)) + "Please Wait..", , "Picture File Converter-Editor"
   Dim s As String, sFile As String
   Dim sI As String
   Dim eD As EDSSColourDepthConstants
   Dim eM As EDSSColourReductionConstants
   Dim New_filename, Old_Filename, New_ExtensionName
   
   
   If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False And Option1(3).Value = False And Option1(4).Value = False And Option1(5).Value = False And Option1(6).Value = False Then
      MsgBox "Please Select Image Format to Convert"
       Else
       Me.MousePointer = 11
      If Option1.Item(0).Value = True Then '*.png
            FileExt = Split(CommonDialog2.FileName, ".")(0) + ".png"
            ElseIf Option1.Item(1).Value = True Then '*gif
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".gif"
            ElseIf Option1.Item(2).Value = True Then '*.tga
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".tga"
            ElseIf Option1.Item(3).Value = True Then '*.bmp
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".bmp"
            ElseIf Option1.Item(4).Value = True Then 'tiff
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".tiff"
            ElseIf Option1.Item(5).Value = True Then '*.jpeg
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".jpeg"
            ElseIf Option1.Item(6).Value = True Then '*.wmf
                   FileExt = Split(CommonDialog2.FileName, ".")(0) + ".wmf"
        End If
        sI = (FileExt)
       If Not (sI = "") Then
          Dim cDIBSave As New cDIBSectionSave ' ' call class module for save
          eD = SelectedColourDepth()
         ' 256 colour has most options:
          If eD = edss256Colour Then
             Select Case True
                   Case optReduceMethod(0).Value
                      ' Let the system do the default when mapping the palette:
                      cDIBSave.Save sI, m_cDIB, , eD, edssSystemDefault
                      
                   Case optReduceMethod(1).Value
                      ' Floyd-Stucci reduce to the selected 256 colour palette:
                      Dim cP As New cPalette
                      Select Case True
                      Case optFloydStucciType(0).Value
                         cP.CreateHalfTone
                         cDIBSave.Save sI, m_cDIB, cP, eD, edssUsePalette
                      Case optFloydStucciType(1).Value
                         cP.CreateWebSafe
                         cDIBSave.Save sI, m_cDIB, cP, eD, edssUsePalette
                      End Select
                      
                   Case optReduceMethod(2).Value
                      ' Octree Colour quantise to generate the optimal palette:
                      cDIBSave.Save sI, m_cDIB, , eD, edssGeneratePalette
                      
                   End Select
                   
            Else
               ' Other colour depths are simpler:
               
                  If optReduceMethod(0).Value Then
                     eM = edssSystemDefault
                  Else
                     eM = edssGeneratePalette
                  End If
                  cDIBSave.Save sI, m_cDIB, , eD, eM
                  
               End If
               
            End If

   Me.MousePointer = 0
   Label20.Visible = False
      
    Old_Filename = CommonDialog2.FileName
    
     If Option1.Item(0).Value = True Then '*.png
            New_filename = Split(CommonDialog2.FileName, ".")(0) + ".png"
            Lbl(1).Caption = "png"
            Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".png"
            ElseIf Option1.Item(1).Value = True Then '*gif
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".gif"
                   Lbl(1).Caption = "gif"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".gif"
            ElseIf Option1.Item(2).Value = True Then '*.tga
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".tga"
                   Lbl(1).Caption = "tga"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".tga"
            ElseIf Option1.Item(3).Value = True Then '*.bmp
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".bmp"
                   Lbl(1).Caption = "bmp"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".bmp"
            ElseIf Option1.Item(4).Value = True Then 'tiff
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".tiff"
                   Lbl(1).Caption = "tiff"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".tiff"
            ElseIf Option1.Item(5).Value = True Then '*.jpeg
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".jpeg"
                   Lbl(1).Caption = "jpeg"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".jpeg"
            ElseIf Option1.Item(6).Value = True Then '*.wmf
                   New_filename = Split(CommonDialog2.FileName, ".")(0) + ".wmf"
                   Lbl(1).Caption = "wmf"
                   Lbl(0).Caption = Split(Lbl(0).Caption, ".")(0) & ".wmf"
        End If
    
    New_ExtensionName = Split(CommonDialog2.FileName, ".")(1)
    sFile = Trim(New_filename)
    s = CInt(FileLen(sFile) / 1024) & "K"
   
    Lbl(7).Caption = s 'file size
     
   For opt1_enb = 0 To 6
      Option1.Item(opt1_enb).Enabled = False
   Next opt1_enb
   
   For opt1_enb = 0 To 3
    optColourDepth.Item(opt1_enb).Enabled = False
   Next opt1_enb
   optReduceMethod.Item(0).Enabled = False
   optReduceMethod.Item(1).Enabled = False
   optReduceMethod.Item(2).Enabled = False
   optFloydStucciType.Item(0).Enabled = False
   optFloydStucciType.Item(1).Enabled = False
   cmdSave.Enabled = True
   picCurrent.Refresh
   MsgBox "Succesfully Converted!", , "Picture File Converter-Editor"
  End If
End Sub

