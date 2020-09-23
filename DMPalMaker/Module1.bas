Attribute VB_Name = "Module1"
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1


Private Type T_RGB
    Red As Integer
    green As Integer
    blue As Integer
End Type

Public T_RGB As T_RGB


Public Function RGBtoHEX(RGBValue)
Dim Ival As Integer, iHex
    iHex = Hex(RGBValue)
    Ival = Len(iHex)
    If Ival = 5 Then iHex = "0" & iHex
    If Ival = 4 Then iHex = "00" & iHex
    If Ival = 3 Then iHex = "000" & iHex
    If Ival = 2 Then iHex = "0000" & iHex
    If Ival = 1 Then iHex = "00000" & iHex
    RGBtoHEX = iHex
    
End Function

Public Sub LongToRgb(lngCol As Long)
    T_RGB.Red = lngCol And (Not &HFFFFFF00)
    T_RGB.green = (lngCol And (Not &HFFFF00FF)) \ &H100&
    T_RGB.blue = (lngCol And Not (&HFF00FFFF)) \ &HFFFF&
End Sub
Public Function rgbToLong(R, G, B As Integer) As Long
Dim VbLng As Long
    VbLng = B * 65536 + G * 256 + R
    
End Function
Function FixPath(lzpath As String)
    If Right(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
    
End Function

Function FindFile(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then FindFile = False Else FindFile = True
    
End Function

Public Function SavePallet(lzFileName As String, PalData As String)
    FileNum = FreeFile
    Open lzFileName For Binary As #FileNum
        Put #FileNum, , PalData
    Close #FileNum
    
End Function

Public Function LoadPallet(lzFile As String, Frm As Form)
Dim PalData As String, FileNum As Long, RgbRow As Variant, _
I, mRed, mGreen, mBlue As Integer
On Error Resume Next

        FileNum = FreeFile
        Open lzFile For Binary As #FileNum
            PalData = Space(LOF(FileNum))
            Get #FileNum, , PalData
        Close #FileNum

        RgbRow = Split(PalData, ",")
        For I = LBound(RgbRow) To UBound(RgbRow)
            mRed = Asc(Mid(RgbRow(I), 1, 1))
            mGreen = Asc(Mid(RgbRow(I), 2, 1))
            mBlue = Asc(Mid(RgbRow(I), 3, 1))
            Frm.Picture1(I).BackColor = RGB(mRed, mGreen, mBlue)
            RgbRow(I) = Nothing
        Next
        PalData = "": FileExt = "": I = 0
        mRed = 0: mGreen = 0: mBlue = 0
        
End Function

Public Function PutFormOnTop(FrmHwnd As Long, OnTop As Boolean)
' just makes a form stay on top of all others
    If OnTop Then
        SetWindowPos FrmHwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
        Exit Function
   Else
        SetWindowPos FrmHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
    End If
    
End Function
