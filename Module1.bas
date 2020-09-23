Attribute VB_Name = "mdlMod"
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Public Type BITMAPFILEHEADER
    bfType As Integer       'must be 19778 = "BM"
    bfSize As Long          'size of file in bytes LOF(%bf)
    bfReserved1 As Integer  'Reserved must be set to zero
    bfReserved2 As Integer  'Reserved must be set to zero
    bfOffBits As Long       'the space between this struct and the begining of the actual bmp data
End Type

Public Type BITMAPINFOHEADER '40 bytes
    biSize As Long              'Len(bmih)
    biWidth As Long             'Width of Bitmap Image
    biHeight As Long            'Height of Bitmap Image
    biPlanes As Integer         'Number of Planes for Target Device,must be set to 1
    biBitCount As Integer       'Number of Bits Per Pixel must be either:1(Monochrome),4(16clrs),8(256color),24(RGBQUADS=16777216 colors)
    biCompression As Long       'Compression Modes can be either:BI_bitfields,BI_JPEG,BI_PNG,BI_RLE4,BI_RLE8
    biSizeImage As Long         'Size in bytes of image,can be set to zero if biCompression = BI_RGB
    biXPelsPerMeter As Long     'Horizonal Resolution in Pixels Per Meter
    biYPelsPerMeter As Long     'Vertical Resolution in Pixels Per Meter
    biClrUsed As Long           'the number of colors used by bitmap if its 0 then all colors are used
    biClrImportant As Long      'the number of colors required to display this bitmap if its 0 then their all required
End Type

Public Type RGBTRIBLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbUnused As Byte
End Type

Public Const BI_bitfields = 3&  'UNKNOWN
Public Const BI_JPEG = 4&       'UNKNOWN
Public Const BI_PNG = 5&        'UNKNOWN
Public Const BI_RGB = 0&        '(uncompressed) THIS IS THE ONLY ONE SUPPORTED IN THIS MODULE
Public Const BI_RLE4 = 2&       'RLE RunLength Compression per 4bits(1/2 byte)
Public Const BI_RLE8 = 1&       'RLE RunLength Compression per 8bits(1bytes)

'1) BITMAPFILEHEADER (bmfh)
'2) BITMAPINFOHEADER (bmih)
'3) RGBQUAD          aColors()
'4) BYTE             aBitmapBits()
'
'bmfh,bmih,acolors,abitmapbits
Dim bmfh As BITMAPFILEHEADER
Dim bmih As BITMAPINFOHEADER
Dim aColors(0 To 255) As RGBQUAD

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
     
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath

End Function

Public Function rgbRed(RGBCol As Long) As Long
    'Return the Red component from an RGB Co
    '     lor
    rgbRed = RGBCol And &HFF
End Function


Public Function rgbGreen(RGBCol As Long) As Long
    'Return the Green component from an RGB
    '     Color
    rgbGreen = ((RGBCol And &H100FF00) / &H100)
End Function


Public Function rgbBlue(RGBCol As Long) As Long
    'Return the Blue component from an RGB C
    '     olor
    rgbBlue = (RGBCol And &HFF0000) / &H10000
End Function


Public Sub GetPal(Filename As String, ByRef Pl() As RGBQUAD)
Dim I As Long
Open Filename For Binary Access Read As #1
Get #1, , bmfh
Get #1, , bmih
If bmih.biBitCount <> 8 Then MsgBox "This picture is not 256 colours!!": Close #1: Exit Sub
Get #1, , Pl
Close #1

End Sub

