Attribute VB_Name = "mdlBMP"
'
'BITMAP FORMAT OVERVIEW
'1) BITMAPFILEHEADER (bmfh)
'2) BITMAPINFOHEADER (bmih)
'3) RGBQUAD          aColors()
'4) BYTE             aBitmapBits() 'this is not avaliable in 24bit bitmaps the aColors replaces this
'
'
'THIS MODULE IS IN BETA STATE
'AND DOESNT SUPPORT THE FOLLOWING:
'1 compressed bitmaps(RLE4,RLE8,JPEG,PNG)
'2 any bitmaps that are not saved in 24-bit DIB
'3 doesnt seem to be working right on some bitmaps
'
'any bugfixes are appreciated,votes also.
'
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
Dim aColors() As RGBTRIBLE
'Dim aColors() as RGBQUAD
Dim aBitmapBits() As Byte

Sub SavePalAsBitmap(Filename As String, Pal() As RGBQUAD)
Dim F%
If Filename = "" Then Exit Sub
With bmfh
.bfOffBits = 1078
.bfReserved1 = 0
.bfReserved2 = 0
.bfSize = 41078
.bfType = 19778
End With

With bmih
.biBitCount = 8
.biClrImportant = 256
.biClrUsed = 256
.biCompression = 0
.biHeight = 200
.biPlanes = 1
.biSize = 40
.biSizeImage = 40000
.biWidth = 200
.biXPelsPerMeter = 2867
.biYPelsPerMeter = 2867
End With

F = FreeFile
Open Filename For Binary Access Write As F
Put F, , bmfh
Put F, , bmih
Put F, , Pal
ReDim aBitmapBits(1 To (bmih.biWidth * bmih.biHeight))
Put F, , aBitmapBits
Close F

End Sub
