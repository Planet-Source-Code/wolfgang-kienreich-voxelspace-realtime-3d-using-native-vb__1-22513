Attribute VB_Name = "modV"
' OPTION SETTINGS ...

    ' Enforce variable declarations
    Option Explicit

' PUBLIC CONSTANTS ...

    ' Pifactor for angle conversion
    Public Const PI180 = 1.74532925199433E-02
    
' PUBLIC TYPES ...

    ' General data for DI Bitmaps
    Public Type BITMAPINFOHEADER
        biSize          As Long                ' Size of structure
        biWidth         As Long                ' Width
        biHeight        As Long                ' Height
        biPlanes        As Integer             ' Planes
        biBitCount      As Integer             ' Bits per pixel
        biCompression   As Long                ' Compression settings
        biSizeImage     As Long                ' Memory size of image
        biXPelsPerMeter As Long                ' -
        biYPelsPerMeter As Long                ' -
        biClrUsed       As Long                ' -
        biClrImportant  As Long                ' -
    End Type
    
    ' RGB data for DI Bitmaps
    Public Type RGBQUAD
        rgbBlue         As Byte                ' Blue
        rgbGreen        As Byte                ' Green
        rgbRed          As Byte                ' Red
        rgbReserved     As Byte                ' -
    End Type
    
    ' Bitmap info for DI Bitmaps
    Public Type BITMAPINFO
        bmiHeader       As BITMAPINFOHEADER    ' Bitmap info
        bmiColors       As RGBQUAD             ' Color table
    End Type
        
    ' Non-DIB bitmap type
    Public Type BITMAP
        bmType As Long                         ' Type of bitmap
        bmWidth As Long                        ' Width
        bmHeight As Long                       ' Height
        bmWidthBytes As Long                   ' Width in bytes
        bmPlanes As Integer                    ' # planes
        bmBitsPixel As Integer                 ' bit depth
        bmBits As Long                         ' memory pointer
    End Type
    
    ' GDI Point type
    Public Type POINTAPI
        X               As Long                ' Position X
        Y               As Long                ' Position Y
    End Type
    
    ' GDI Rectangle type
    Public Type RECT
        Left            As Long                ' Left
        Top             As Long                ' Top
        Right           As Long                ' Right
        Bottom          As Long                ' Bottom
    End Type
        
    ' Array bound type for DMA to VB arrays
    Public Type SAFEARRAYBOUND
        cElements       As Long                ' # elements
        lLbound         As Long                ' Bound
    End Type
    
    ' Array type for DMA to VB arrays
    Public Type SAFEARRAY1D
        cDims           As Integer             ' # dimensions
        fFeatures       As Integer             ' -
        cbElements      As Long                ' -
        cLocks          As Long                ' -
        pvData          As Long                ' Pointer to data
        Bounds(0 To 0)  As SAFEARRAYBOUND      ' Bounds
    End Type
    
    
' PUBLIC API FUNCTION CALLS ...

    ' General utility functions
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
    Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
    Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
    Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

    ' Graphics related functions
    Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
    Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function SaveDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function RestoreDC Lib "gdi32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
    Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Public Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
    Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long


