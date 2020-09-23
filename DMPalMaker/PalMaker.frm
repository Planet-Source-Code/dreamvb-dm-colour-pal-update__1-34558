VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Colour Pal 2002"
   ClientHeight    =   5130
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5490
   Icon            =   "PalMaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5490
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txthex 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   1140
      Width           =   405
   End
   Begin VB.TextBox txthex 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4140
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   1140
      Width           =   405
   End
   Begin VB.TextBox txthex 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4605
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   1140
      Width           =   405
   End
   Begin VB.TextBox txtvbcol 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   1110
      Width           =   1320
   End
   Begin Project1.Tray Tray1 
      Left            =   315
      Top             =   4020
      _ExtentX        =   529
      _ExtentY        =   529
      Icon            =   "PalMaker.frx":0CCA
   End
   Begin Project1.Flat2 Flat23 
      Height          =   4560
      Left            =   75
      TabIndex        =   52
      Top             =   105
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   8043
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   31
      Left            =   4755
      MouseIcon       =   "PalMaker.frx":19A4
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   48
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   3780
      MouseIcon       =   "PalMaker.frx":1CAE
      MousePointer    =   99  'Custom
      Picture         =   "PalMaker.frx":1FB8
      ScaleHeight     =   795
      ScaleWidth      =   1065
      TabIndex        =   47
      Top             =   1890
      Width           =   1095
   End
   Begin VB.HScrollBar hsbblue 
      Height          =   255
      Left            =   840
      Max             =   255
      TabIndex        =   43
      Top             =   2460
      Width           =   2295
   End
   Begin VB.HScrollBar hsbgreen 
      Height          =   255
      Left            =   840
      Max             =   255
      TabIndex        =   42
      Top             =   2190
      Width           =   2295
   End
   Begin VB.HScrollBar hsbred 
      Height          =   255
      Left            =   840
      Max             =   255
      TabIndex        =   41
      Top             =   1920
      Width           =   2295
   End
   Begin VB.PictureBox colview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   2430
      ScaleHeight     =   1125
      ScaleWidth      =   690
      TabIndex        =   37
      Top             =   240
      Width           =   720
   End
   Begin VB.TextBox txthtm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   750
      Width           =   1320
   End
   Begin VB.TextBox txtrgb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   270
      Width           =   405
   End
   Begin VB.TextBox txtrgb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   270
      Width           =   405
   End
   Begin VB.TextBox txtrgb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   270
      Width           =   405
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   3510
      Top             =   2595
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   30
      Left            =   4545
      MouseIcon       =   "PalMaker.frx":4CB2
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   210
      TabIndex        =   30
      Top             =   870
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   29
      Left            =   4335
      MouseIcon       =   "PalMaker.frx":4FBC
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   29
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   28
      Left            =   4125
      MouseIcon       =   "PalMaker.frx":52C6
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   27
      Left            =   3915
      MouseIcon       =   "PalMaker.frx":55D0
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   26
      Left            =   3705
      MouseIcon       =   "PalMaker.frx":58DA
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   25
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":5BE4
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   24
      Left            =   3285
      MouseIcon       =   "PalMaker.frx":5EEE
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   24
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   23
      Left            =   4755
      MouseIcon       =   "PalMaker.frx":61F8
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   22
      Left            =   4545
      MouseIcon       =   "PalMaker.frx":6502
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   21
      Left            =   4335
      MouseIcon       =   "PalMaker.frx":680C
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   20
      Left            =   4125
      MouseIcon       =   "PalMaker.frx":6B16
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   19
      Left            =   3915
      MouseIcon       =   "PalMaker.frx":6E20
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   18
      Left            =   3705
      MouseIcon       =   "PalMaker.frx":712A
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   17
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":7434
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   16
      Left            =   3285
      MouseIcon       =   "PalMaker.frx":773E
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   660
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   15
      Left            =   4755
      MouseIcon       =   "PalMaker.frx":7A48
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   14
      Left            =   4545
      MouseIcon       =   "PalMaker.frx":7D52
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   13
      Left            =   4335
      MouseIcon       =   "PalMaker.frx":805C
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   4125
      MouseIcon       =   "PalMaker.frx":8366
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   11
      Left            =   3915
      MouseIcon       =   "PalMaker.frx":8670
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   10
      Left            =   3705
      MouseIcon       =   "PalMaker.frx":897A
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":8C84
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   3285
      MouseIcon       =   "PalMaker.frx":8F8E
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   450
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   4755
      MouseIcon       =   "PalMaker.frx":9298
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   4545
      MouseIcon       =   "PalMaker.frx":95A2
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   6
      Top             =   240
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   4335
      MouseIcon       =   "PalMaker.frx":98AC
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   4125
      MouseIcon       =   "PalMaker.frx":9BB6
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   3915
      MouseIcon       =   "PalMaker.frx":9EC0
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   3705
      MouseIcon       =   "PalMaker.frx":A1CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":A4D4
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   3285
      MouseIcon       =   "PalMaker.frx":A7DE
      MousePointer    =   99  'Custom
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   450
      Left            =   1680
      Top             =   3255
      Width           =   465
   End
   Begin VB.Label lblhexval 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Hex Values"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2685
      MouseIcon       =   "PalMaker.frx":AAE8
      TabIndex        =   65
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblvbcol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy VB Colour"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2685
      MouseIcon       =   "PalMaker.frx":B7B2
      TabIndex        =   64
      Top             =   3810
      Width           =   1470
   End
   Begin VB.Label lblh 
      AutoSize        =   -1  'True
      Caption         =   "Hex"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3285
      TabIndex        =   60
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "VB Colour"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   59
      Top             =   1140
      Width           =   705
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   150
      X2              =   4875
      Y1              =   1725
      Y2              =   1725
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   150
      X2              =   4875
      Y1              =   1710
      Y2              =   1710
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   165
      X2              =   4890
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   165
      X2              =   4890
      Y1              =   2865
      Y2              =   2865
   End
   Begin VB.Label lblhide 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Move to Tray"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2685
      MouseIcon       =   "PalMaker.frx":C47C
      TabIndex        =   57
      Top             =   4020
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM Colour Pal for Windows 95, 95x,win2k, XP"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1095
      TabIndex        =   56
      Top             =   4800
      Width           =   3285
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   75
      TabIndex        =   55
      Top             =   4725
      Width           =   4425
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2400
      X2              =   2400
      Y1              =   3000
      Y2              =   4395
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   2385
      X2              =   2385
      Y1              =   3000
      Y2              =   4380
   End
   Begin VB.Label lblhex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Html Colour"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2685
      MouseIcon       =   "PalMaker.frx":D146
      TabIndex        =   54
      Top             =   3600
      Width           =   1680
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copy RGB Values"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   2685
      MouseIcon       =   "PalMaker.frx":DE10
      TabIndex        =   53
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      DragIcon        =   "PalMaker.frx":EADA
      Height          =   480
      Left            =   1665
      Picture         =   "PalMaker.frx":EDE4
      Top             =   3240
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the image then choose a colour from the screen"
      ForeColor       =   &H00000000&
      Height          =   810
      Index           =   1
      Left            =   225
      TabIndex        =   51
      Top             =   3195
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "R"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   50
      Top             =   285
      Width           =   120
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "HTML"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   49
      Top             =   810
      Width           =   450
   End
   Begin VB.Label lblblue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   46
      Top             =   2460
      Width           =   435
   End
   Begin VB.Label lblgreen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   45
      Top             =   2190
      Width           =   435
   End
   Begin VB.Label lblred 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   44
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Blue"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   40
      Top             =   2460
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   39
      Top             =   2190
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   38
      Top             =   1920
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "B"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   495
      TabIndex        =   32
      Top             =   285
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "G"
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   360
      TabIndex        =   31
      Top             =   285
      Width           =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   -15
      X2              =   825
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   840
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Pallet"
      End
      Begin VB.Menu mnunewpal 
         Caption         =   "&Create new Pallet"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnublank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
    hsbred.Value = 255
    hsbgreen.Value = 255
    hsbblue.Value = 255
    LoadPallet FixPath(App.Path) & "pallets\coolXp.pal", Form1
    PutFormOnTop Form1.hwnd, True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = False
    lblhex.FontUnderline = False
    lblhide.FontUnderline = False
    lblvbcol.FontUnderline = False
    lblhexval.FontUnderline = False
    
    
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = Form1.ScaleWidth
    Line1(1).X2 = Form1.ScaleWidth
    Label5.Width = Form1.ScaleWidth - Label5.Left - 90
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set frmabout = Nothing
    Set frmnew = Nothing
    
    For I = 0 To Picture1.Count - 1
        Set Picture1(I) = Nothing
    Next
    I = 0
    Tray1.Visible = False
    Unload Form1
    
    
    
    
End Sub

Private Sub hsbblue_Change()
    lblblue.Caption = hsbblue.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
    
    
    
    
End Sub

Private Sub hsbblue_Scroll()
    hsbblue_Change
    
End Sub

Private Sub hsbgreen_Change()
    lblgreen.Caption = hsbgreen.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
    txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
    
    
    
    
End Sub

Private Sub hsbgreen_Scroll()
    hsbgreen_Change
    
End Sub

Private Sub hsbred_Change()
    lblred.Caption = hsbred.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
    
    
    
    
End Sub

Private Sub hsbred_Scroll()
    hsbred_Change
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form1.Hide
    DoEvents
    frmSc.Show
    
End Sub

Private Sub lblhex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = False
    lblhex.FontUnderline = True
    lblhide.FontUnderline = False
    lblvbcol.FontUnderline = False
    lblhexval.FontUnderline = False
    
    
End Sub

Private Sub lblhex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText txthtm.Text
    MsgBox "All HTML colour codes have been copied to the clipboard", vbInformation, Form1.Caption
    
End Sub

Private Sub lblhexval_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = False
    lblhex.FontUnderline = False
    lblhide.FontUnderline = False
    lblvbcol.FontUnderline = False
    lblhexval.FontUnderline = True
    
End Sub

Private Sub lblhexval_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText txthex(0) & "," & txthex(1) & "," & txthex(2)
    MsgBox "All HEX values have been copied to the clipboard", vbInformation, Form1.Caption
    
    
End Sub

Private Sub lblhide_Click()
    Form1.Visible = False
    Tray1.ToolTip = Form1.Caption
    Tray1.Visible = True

End Sub

Private Sub lblhide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = False
    lblhex.FontUnderline = False
    lblhide.FontUnderline = True
    lblvbcol.FontUnderline = False
    lblhexval.FontUnderline = False
    
End Sub

Private Sub lblrgb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = True
    lblhex.FontUnderline = False
    lblhide.FontUnderline = False
    lblvbcol.FontUnderline = False
    lblhexval.FontUnderline = False
    
    
End Sub

Private Sub lblrgb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText txtrgb(0) & "," & txtrgb(1) & "," & txtrgb(2)
    MsgBox "All RGB Vales have been copied to the clipboard", vbInformation, Form1.Caption
    
    
End Sub

Private Sub lblvbcol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblrgb.FontUnderline = False
    lblhex.FontUnderline = False
    lblhide.FontUnderline = False
    lblvbcol.FontUnderline = True
    lblhexval.FontUnderline = False
    
End Sub

Private Sub lblvbcol_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText txtvbcol.Text
    MsgBox "The VB colour value has been copied to the clipboard", vbInformation, Form1.Caption
    
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, Form1
    
End Sub

Private Sub mnuexit_Click()
Dim ans
    ans = MsgBox("Are you sure you would like to quit this program now", vbYesNo Or vbQuestion)
    If ans = vbNo Then Exit Sub
    Unload Form1
    
End Sub

Private Sub mnunewpal_Click()
    PutFormOnTop Form1.hwnd, False
    frmnew.Show vbModal, Form1
        
End Sub

Private Sub mnuopen_Click()
Dim FileExt As String

    Cdialog.DialogTitle = "Open DM pallet"
    Cdialog.Filter = "DM Pallet Files(*.pal)|*.pal"
    Cdialog.InitDir = FixPath(App.Path)
    Cdialog.ShowOpen
    
    FileExt = Right(UCase(Cdialog.FileName), 3)
    If Len(Cdialog.FileName) <= 0 Then Exit Sub
    If Not FileExt = "PAL" Then
        MsgBox "This is not a vaild DM Pallet filename", vbInformation, Form1.Caption
        Exit Sub
    Else
        LoadPallet Cdialog.FileName, Form1
        FileExt = ""
    End If
    
End Sub
Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    LongToRgb Picture1(Index).BackColor
    txtrgb(0).Text = T_RGB.Red
    txtrgb(1).Text = T_RGB.green
    txtrgb(2).Text = T_RGB.blue
    txthtm.Text = "#" & RGBtoHEX(Picture1(Index).BackColor)
    colview.BackColor = Picture1(Index).BackColor
    txtvbcol.Text = Picture1(Index).BackColor
    txtlng.Text = Picture1(Index).BackColor
    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.green
    hsbblue.Value = T_RGB.blue
    
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim T_ColVal As Long
On Error Resume Next
    If Button = vbLeftButton Then
        T_ColVal = Picture2.Point(X, Y)
        txtvbcol.Text = T_ColVal
        LongToRgb T_ColVal
        colview.BackColor = T_ColVal
        hsbred.Value = T_RGB.Red
        hsbgreen.Value = T_RGB.green
        hsbblue.Value = T_RGB.blue
End If

    
End Sub

Private Sub Tray1_DblClick()
    Tray1.Visible = False
    Form1.Visible = True
    
End Sub

Private Sub txtrgb_Change(Index As Integer)
On Error Resume Next
    txtvbcol.Text = txtrgb(2) * 65536 + txtrgb(1) * 256 + txtrgb(0)
    
End Sub

Private Sub txtvbcol_Change()
Dim StrHexCode As String
    StrHexCode = RGBtoHEX(txtvbcol.Text)
    txthex(0).Text = Mid(StrHexCode, 1, 2)
    txthex(1).Text = Mid(StrHexCode, 3, 2)
    txthex(2).Text = Mid(StrHexCode, 5, 2)
    
End Sub
