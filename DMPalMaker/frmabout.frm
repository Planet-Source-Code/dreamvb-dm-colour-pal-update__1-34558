VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Colour Pal 2002"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   1020
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmabout.frx":0000
      Top             =   1785
      Width           =   4440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   1920
      TabIndex        =   3
      Top             =   2955
      Width           =   1020
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   4785
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ver 2.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3555
         TabIndex        =   6
         Top             =   495
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Colour Pal 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   5
         Top             =   150
         Width           =   2565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Colour Pal 2002"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1155
         TabIndex        =   4
         Top             =   165
         Width           =   2565
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   105
         Picture         =   "frmabout.frx":00B4
         Top             =   105
         Width           =   480
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "E-Mail dreamvb@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   1485
      Width           =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Designed by Ben Jones"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   1245
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "#102"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   945
      Width           =   375
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload frmabout
    
End Sub

Private Sub Form_Load()
    Label1.Caption = "DM Colour Picker " & App.Major & "." & App.Minor & "." & App.Revision
    PutFormOnTop Form1.hwnd, False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PutFormOnTop Form1.hwnd, True
    Set frmabout = Nothing
    
End Sub
