VERSION 5.00
Begin VB.Form frmSc 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   Icon            =   "frmSc.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSc.frx":030A
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   33
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmSc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()
Dim DeskTopWnd As Long, THdc As Long
    DeskTopWnd = GetDesktopWindow()
    THdc = GetDC(DeskTopWnd)
    BitBlt frmSc.hDC, 0, 0, Screen.Width, Screen.Height, THdc, 0, 0, vbSrcCopy
    frmSc.Top = 0
    frmSc.Left = 0
    frmSc.Refresh
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Module1.LongToRgb frmSc.Point(X, Y)
    Form1.hsbred = T_RGB.Red
    Form1.hsbgreen = T_RGB.green
    Form1.hsbblue = T_RGB.blue
    Unload frmSc
    Form1.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSc.Picture = Nothing
    Set frmSc = Nothing
    
End Sub

