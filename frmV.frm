VERSION 5.00
Begin VB.Form frmV 
   Appearance      =   0  '2D
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "VoxelSpace Demo"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmV.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   365
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Label lblMove 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   75
      MousePointer    =   5  'Größenänderung
      TabIndex        =   5
      Top             =   75
      Width           =   1290
   End
   Begin VB.Label lblButton 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   4
      Left            =   3585
      TabIndex        =   4
      Top             =   270
      Width           =   2280
   End
   Begin VB.Label lblButton 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   3
      Left            =   3420
      TabIndex        =   3
      Top             =   75
      Width           =   2280
   End
   Begin VB.Label lblButton 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   1
      Left            =   1740
      TabIndex        =   2
      Top             =   255
      Width           =   1650
   End
   Begin VB.Label lblButton 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   0
      Left            =   1575
      TabIndex        =   1
      Top             =   60
      Width           =   1650
   End
   Begin VB.Label lblButton 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   0
      Top             =   465
      Width           =   1650
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   120
      Index           =   4
      Left            =   3615
      Shape           =   2  'Oval
      Top             =   315
      Width           =   135
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   120
      Index           =   3
      Left            =   3450
      Shape           =   2  'Oval
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Ausgefüllt
      Height          =   120
      Index           =   2
      Left            =   1965
      Shape           =   2  'Oval
      Top             =   525
      Width           =   135
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00404040&
      FillStyle       =   0  'Ausgefüllt
      Height          =   120
      Index           =   1
      Left            =   1800
      Shape           =   2  'Oval
      Top             =   315
      Width           =   135
   End
   Begin VB.Shape shpButton 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   120
      Index           =   0
      Left            =   1620
      Shape           =   2  'Oval
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "frmV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' VOXELSPACE
' Realtime 3D using native VB
'
' (C) 2001 NLS - Nonlinear solutions
' http://www.members.aon.at/nls
' mailto:nonlinear@aon.at
'
' Thanks to David Brebner of Unlimited Realities for
' the original Voxel algorithm he provided back in 1997.
'


' OPTION SETTINGS ...

    ' Enforce variable declarations
    Option Explicit
    
' INSTANCE TYPES ...

    ' Animation for waves
    Private Type ANI
        Index       As Long     ' Index of pixel
        Phase       As Byte     ' Phase of pixel
    End Type
    
' INSTANCE VARIABLES ...

    Private I_udtMouseStart     As POINTAPI     ' Offset for window movement on screen
    Private I_udtWindowStart    As POINTAPI     ' Offset for window movement on screen

    Private I_udtBD             As BITMAPINFO   ' Display bitmap descriptor
    Private I_lngDC             As Long         ' Display device context handle
    Private I_lngBM             As Long         ' Display bitmap handle
    Private I_lngMP             As Long         ' Centroider to display memory
    
    Private I_lngDCSkyN         As Long         ' Sky device context (night)
    Private I_lngDCSkyD         As Long         ' Sky device context (day)
        
    Private I_bytTex()          As Byte         ' Texture data
    Private I_bytAlt()          As Byte         ' Altitude data
    
    Private I_udtAni()          As ANI          ' Animation data
    Private I_bytPhase()        As Byte         ' Animation wave sine lookup
    Private I_bytAvg()          As Byte         ' Averaging byte value lookup
    
    Private I_lngMapSize        As Long         ' Size of map
    
    Private I_lngHeading        As Long         ' Heading of motion
    Private I_sngVelocity       As Single       ' Velocity of motion
    Private I_sngPositionX      As Single       ' Current position
    Private I_sngPositionY      As Single       ' Current position
    
    Private I_bytBn1()          As Byte         ' Banner data
    Private I_bytBn2()          As Byte         ' Banner data
    
    Private I_lngFrameCount     As Long         ' Number of frames processed
    Private I_bytFrameRate      As Long         ' Number of frames per second rendered
    Private I_bytRateHistory()  As Byte         ' History of framerate
        
    Private I_blnStateNight     As Boolean      ' Nighttime state
    Private I_blnStateLight     As Boolean      ' Lightcast state
    Private I_blnStateAni       As Boolean      ' Animation state
    Private I_blnStateIlv       As Boolean      ' Interleave
    
' INSTANCE CODE ...
    
    
    '
    ' FORM_LOAD: Initialise application on startup
    '
    Private Sub Form_Load()
    
        ' Declare local variables ...
        
            Dim L_objVBImage    As StdPicture       ' VB picture for loading data
            Dim L_udtBD         As BITMAP           ' GDI bitmap descriptor
            
            Dim L_bytTmp()      As Byte             ' Byte array holding loaded data for translation
            Dim L_lngIndex      As Long             ' Index to run over data
            
        ' Code ...
        
            ' Initialise form ...
            
                ' Set size
                Me.Width = 500 * Screen.TwipsPerPixelX
                Me.Height = 420 * Screen.TwipsPerPixelY
                
                ' Set shape
                SetWindowRgn Me.hWnd, CreateRoundRectRgn(Me.Top / Screen.TwipsPerPixelX, Me.Left / Screen.TwipsPerPixelY, 500, 420, 16, 16), True
                
                ' Set background
                Set Me.Picture = LoadPicture(App.Path + "\interface.jpg")
                
            ' Generate display DI bitmap ...
            
                ' Setup description ...
            
                    With I_udtBD
                        .bmiHeader.biHeight = 360
                        .bmiHeader.biWidth = 480
                        .bmiHeader.biBitCount = 24
                        .bmiHeader.biPlanes = 1
                        .bmiHeader.biSize = Len(.bmiHeader)
                    End With
        
                ' Initialize DI bitmap ...
                                
                    ' Acquire DC
                    I_lngDC = CreateCompatibleDC(Me.hdc)
                    
                    ' Save current DC state on stack
                    SaveDC I_lngDC
                    
                    ' Create DI bitmap
                    I_lngBM = CreateDIBSection(I_lngDC, I_udtBD, 0, I_lngMP, 0, 0)
                    
                    ' Select it into the DC
                    SelectObject I_lngDC, I_lngBM
                    
                    ' Create display memory
                    ReDim L_bytDat(I_udtBD.bmiHeader.biSizeImage) As Byte

            ' Load data from description files ...
            
                ' Altitude ...
                    
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\altitude.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                    
                    ' Reserve space
                    ReDim I_bytAlt(L_udtBD.bmWidth * L_udtBD.bmHeight) As Byte
                                                        
                    ' Remember map size
                    I_lngMapSize = L_udtBD.bmWidth
                                                        
                    ' Translate ...
                    
                        ' Reserve temporary space
                        ReDim L_bytTmp(L_udtBD.bmWidthBytes * L_udtBD.bmHeight) As Byte
                        
                        ' Copy memory to temporary space
                        CopyMemory L_bytTmp(0), ByVal (L_udtBD.bmBits), UBound(L_bytTmp)
                        
                        ' Translate into altitude
                        For L_lngIndex = 0 To UBound(I_bytAlt)
                            I_bytAlt(L_lngIndex) = L_bytTmp(L_lngIndex * 3)
                        Next
                        
                ' Texture ...
                    
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\texture.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                
                    ' Reserve space
                    ReDim I_bytTex(L_udtBD.bmWidthBytes * L_udtBD.bmHeight) As Byte
                                    
                    ' Copy memory
                    CopyMemory I_bytTex(0), ByVal (L_udtBD.bmBits), UBound(I_bytTex)
            
                ' Sky ...
                
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\skyn.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                    
                    ' Create according device context
                    I_lngDCSkyN = CreateCompatibleDC(Me.hdc)
                    
                    ' Save device context state
                    SaveDC I_lngDCSkyN
                    
                    ' Select bitmap into DC
                    SelectObject I_lngDCSkyN, L_objVBImage.handle
                    
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\skyd.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                    
                    ' Create according device context
                    I_lngDCSkyD = CreateCompatibleDC(Me.hdc)
                    
                    ' Save device context state
                    SaveDC I_lngDCSkyD
                    
                    ' Select bitmap into DC
                    SelectObject I_lngDCSkyD, L_objVBImage.handle
                    
                ' Banners ...
                    
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\banner1.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                
                    ' Reserve space
                    ReDim I_bytBn1(L_udtBD.bmWidthBytes * L_udtBD.bmHeight) As Byte
                                    
                    ' Copy memory
                    CopyMemory I_bytBn1(0), ByVal (L_udtBD.bmBits), UBound(I_bytBn1)
                    
                    ' Load image from file
                    Set L_objVBImage = LoadPicture(App.Path + "\banner2.jpg")
                    
                    ' Acquire info
                    GetObject L_objVBImage.handle, Len(L_udtBD), L_udtBD
                
                    ' Reserve space
                    ReDim I_bytBn2(L_udtBD.bmWidthBytes * L_udtBD.bmHeight) As Byte
                                    
                    ' Copy memory
                    CopyMemory I_bytBn2(0), ByVal (L_udtBD.bmBits), UBound(I_bytBn2)
                    
            ' Generate additional data ...
            
                ' Sea animation ...
                    
                    ' Reserve initial memory
                    ReDim I_udtAni(0)
                    
                    ' Run over heightmap
                    For L_lngIndex = 0 To UBound(I_bytAlt) - 1
                    
                        ' If at sea level, add to animation lookup
                        If I_bytAlt(L_lngIndex) = 0 Then
                            I_udtAni(UBound(I_udtAni)).Index = L_lngIndex
                            I_udtAni(UBound(I_udtAni)).Phase = Int((360 / I_lngMapSize) * (L_lngIndex \ 512))
                            ReDim Preserve I_udtAni(UBound(I_udtAni) + 1)
                        End If
                    Next
                    
                    ' Optimize memory
                    ReDim Preserve I_udtAni(UBound(I_udtAni) - 1)
                    
                ' Lookup: Animation phase ...
                
                    ' Reserve memory for wave sine lookup table
                    ReDim I_bytPhase(360)
                    
                    ' Run over lookup table, calculate
                    For L_lngIndex = 0 To 360
                        I_bytPhase(L_lngIndex) = 10 + Int(Sin(L_lngIndex * PI180 * 18) * 10)
                    Next
                
                ' Lookup: Averaging values ...
                
                    ' Reserve memory for byte average lookup table
                    ReDim I_bytAvg(255, 255)
                    
                    ' Run over lookup table, calculate
                    For L_lngIndex = 0 To 65535
                        I_bytAvg(L_lngIndex Mod 256, L_lngIndex \ 256) = (L_lngIndex Mod 256) / 2 + (L_lngIndex \ 256) / 2
                    Next
                    
                ' Interface ...
                
                    ' Reserve space for framerate history
                    ReDim I_bytRateHistory(27)
                    
            ' Setup initial state ...
            
                ' Position
                I_sngPositionX = I_lngMapSize / 1.1
                I_sngPositionY = I_lngMapSize / 2
            
                ' Motion
                I_lngHeading = 172800
                I_sngVelocity = 1
                
                ' State
                I_blnStateNight = False
                I_blnStateLight = False
                I_blnStateAni = True
                I_blnStateIlv = True
                
    End Sub
            
    '
    ' FORM_ACTIVATE: Execute master loop
    '
    Private Sub Form_Activate()
    
        ' Declare local variables ...
        
            Dim L_intNextFrameTime  As Long     ' Time at which to commence next frame
            Dim L_intNextFrameRate  As Long     ' Time at which to take next framerate
            Dim L_intFrameCount     As Long     ' Local frame counter
            
            Dim L_lngIndex          As Long     ' Index over altitude (water animation)
            
        ' Code ...
        
            Do
                
                ' Frame time and statistics calculations ...
                
                    ' Try to stabilise for 50 fps
                    L_intNextFrameTime = timeGetTime + 20
                    
                    ' Count frames ellapsed
                    I_lngFrameCount = I_lngFrameCount + 1
                    
                    ' Increase local frame count
                    L_intFrameCount = L_intFrameCount + 1
                        
                    ' Check for next frame timing imminent
                    If L_intNextFrameRate < L_intNextFrameTime Then
                    
                        ' Remember next frame timing time
                        L_intNextFrameRate = L_intNextFrameTime + 1000
                        
                        ' Store frame rate
                        I_bytFrameRate = L_intFrameCount
                        
                        ' Store old frame rates
                        CopyMemory I_bytRateHistory(0), I_bytRateHistory(1), (UBound(I_bytRateHistory))
                        I_bytRateHistory(UBound(I_bytRateHistory)) = I_bytFrameRate
                        
                        ' Render framerate display ...
                            
                            ' Render
                            For L_lngIndex = 0 To UBound(I_bytRateHistory)
                                Me.Line (432 + L_lngIndex * 2, 40)-(432 + L_lngIndex * 2, 40 - 27), RGB(64, 64, 64)
                                Me.Line (432 + L_lngIndex * 2, 40)-(432 + L_lngIndex * 2, 40 - I_bytRateHistory(L_lngIndex) * 0.5), RGB(192, 192, 192)
                            Next
                        
                        ' Reset local frame count
                        L_intFrameCount = 0
                        
                    End If
                    
                ' Update ...
                            
                    ' Update water animation ...
                        
                        If I_blnStateAni Then
                            
                            ' Run over animated pixel in heightmap, update altitude
                            For L_lngIndex = 0 To UBound(I_udtAni)
                                I_bytAlt(I_udtAni(L_lngIndex).Index) = I_bytPhase((I_udtAni(L_lngIndex).Phase + I_lngFrameCount) Mod 360)
                            Next
                            
                        End If
                        
                    ' Update viewport ...
                    
                        ' Update position according to motion
                        I_sngPositionX = I_sngPositionX + I_sngVelocity * Cos((I_lngHeading + I_udtBD.bmiHeader.biWidth / 2) / 360)
                        I_sngPositionY = I_sngPositionY + I_sngVelocity * Sin((I_lngHeading + I_udtBD.bmiHeader.biWidth / 2) / 360)
                    
                    ' Render next frame
                    Render
                    
                ' Process user input
                
                    ' Escape: Terminate application
                    If (GetAsyncKeyState(&H1B) And 32768) = 32768 Then Unload Me
                    
                    ' Left: Turn
                    If (GetAsyncKeyState(&H25) And 32768) = 32768 Then I_lngHeading = I_lngHeading - 9
                    
                    ' Right: Turn
                    If (GetAsyncKeyState(&H27) And 32768) = 32768 Then I_lngHeading = I_lngHeading + 9
                    
                    ' Up: Accellerate
                    If (GetAsyncKeyState(&H26) And 32768) = 32768 Then I_sngVelocity = I_sngVelocity + IIf(I_sngVelocity < 2, 0.1, 0)
                    
                    ' Down: Decellerate
                    If (GetAsyncKeyState(&H28) And 32768) = 32768 Then I_sngVelocity = I_sngVelocity - IIf(I_sngVelocity > 0.2, 0.1, 0)
                    
                ' Process windows events, await next frame ...
                
                    ' Loop until frame time over
                    Do
                        ' Request window events
                        DoEvents
                    Loop Until L_intNextFrameTime < timeGetTime
                
            Loop
        
    End Sub
        
    '
    ' FORM_UNLOAD: Cleanup application on shutdown
    '
    Private Sub Form_Unload(Cancel As Integer)
    
        ' Declare local variables ...
        
            ' (none)
            
        ' Code ...
        
            ' Cleanup display ...
    
                ' Restore DC states
                RestoreDC I_lngDC, -1
                
                ' Delete DC
                DeleteDC I_lngDC
                
                ' Delete DI bitmap
                DeleteObject I_lngBM
                
            ' Cleanup sky ...
            
                ' Restore DC states
                RestoreDC I_lngDCSkyN, -1
                RestoreDC I_lngDCSkyD, -1
                
                ' Delete DC
                DeleteDC I_lngDCSkyN
                DeleteDC I_lngDCSkyD
                
            ' End application
            End
                
    End Sub
   
    '
    ' RENDER: Renders data onto surface
    '
    Private Sub Render()
    
        ' Declare local variables ...
            
            Dim L_udtArrayDesc  As SAFEARRAY1D         ' Array descriptor
            Dim L_bytDat()      As Byte                ' Byte data array
            
            Dim L_lngPX         As Long                ' Pixel column
            Dim L_lngPY         As Long                ' Pixel row
                        
            Dim L_sngRX         As Single              ' Ray position
            Dim L_sngRY         As Single              ' Ray position
            Dim L_sngRZ         As Single              ' Ray position
            
            Dim L_sngDX         As Single              ' Ray delta
            Dim L_sngDY         As Single              ' Ray delta
            Dim L_sngDZ         As Single              ' Ray delta
            
            Dim L_lngVX         As Long                ' Voxel position
            Dim L_lngVY         As Long                ' Voxel position
            
            Dim L_sngDelta      As Single              ' Slope of ray
            Dim L_sngScale      As Single              ' Scale of voxels
            Dim L_sngAlt        As Single              ' Current altitude
            
            Dim L_bytColor(2)   As Byte                ' Current voxel color
            
            Dim L_sngFogFactor  As Single              ' Fog factor
            Dim L_bytFogValue   As Byte                ' Fog value
                
            Dim L_lngStep       As Integer             ' Step counter for ray
            Dim L_bytStepWidth  As Byte                ' Step width for interleaved mode
            
            Dim L_lngBase       As Long                ' Base adress
            Dim L_lngSource     As Long                ' Source adress
            
            Dim L_intTop()      As Byte                ' Top rendered pixel in column
            
        ' Code ...
                    
            ' Initialise bitmap data DMA ...
            
                ' Fill array description
                With L_udtArrayDesc
                    .cbElements = 1
                    .cDims = 1
                    .Bounds(0).lLbound = 0
                    .Bounds(0).cElements = I_udtBD.bmiHeader.biHeight * I_udtBD.bmiHeader.biWidth * 3
                    .pvData = I_lngMP
                End With
                
                ' Let array point to bitmap data
                CopyMemory ByVal VarPtrArray(L_bytDat), VarPtr(L_udtArrayDesc), 4&
                
            ' Initialise voxel rendering ...
            
                ' Set sky ...
                
                    ' All visible: Just blit
                    If (I_lngHeading Mod 960) < (960 - I_udtBD.bmiHeader.biWidth) Then
                        BitBlt I_lngDC, 0, 0, I_udtBD.bmiHeader.biWidth - 1, 300, IIf(I_blnStateNight, I_lngDCSkyN, I_lngDCSkyD), I_lngHeading Mod 960, 0, vbSrcCopy
                        
                    ' At edge: Blend rollover
                    Else
                        BitBlt I_lngDC, 0, 0, I_udtBD.bmiHeader.biWidth - ((I_lngHeading Mod 960) - I_udtBD.bmiHeader.biWidth), 300, IIf(I_blnStateNight, I_lngDCSkyN, I_lngDCSkyD), I_lngHeading Mod 960, 0, vbSrcCopy
                        BitBlt I_lngDC, I_udtBD.bmiHeader.biWidth - ((I_lngHeading Mod 960) - I_udtBD.bmiHeader.biWidth), 0, I_udtBD.bmiHeader.biWidth - (I_udtBD.bmiHeader.biWidth - ((I_lngHeading Mod 960) - I_udtBD.bmiHeader.biWidth)), 300, IIf(I_blnStateNight, I_lngDCSkyN, I_lngDCSkyD), 0, 0, vbSrcCopy
                    End If
                
                ' Reserve memory for topmost rendered pixel values
                ReDim L_intTop(I_udtBD.bmiHeader.biWidth)
                
                ' Set initial raycast slope
                L_sngDelta = 0.05
                
                ' Set step width
                L_bytStepWidth = IIf(I_blnStateIlv, 2, 1)
                
            ' Render voxels (even rows) ...
            
                ' Run over pixel columns
                For L_lngPX = 0 To I_udtBD.bmiHeader.biWidth - 1 Step L_bytStepWidth
                    
                    ' Initialise ray position
                    L_sngRX = I_sngPositionX
                    L_sngRY = I_sngPositionY
                    L_sngRZ = 384
                
                    ' Initialise ray delta
                    L_sngDX = Cos((I_lngHeading + L_lngPX) / 360)
                    L_sngDY = Sin((I_lngHeading + L_lngPX) / 360)
                    L_sngDZ = L_sngDelta * -200
                    
                    ' Initialise voxel detection scale
                    L_sngScale = 0
                    
                    ' Initialise pixel row
                    L_lngPY = 0
                    
                    ' Ray cast loop
                    For L_lngStep = 0 To 128
                    
                        ' Set voxel position, trim to map size
                        L_lngVX = L_sngRX And (I_lngMapSize - 1)
                        L_lngVY = L_sngRY And (I_lngMapSize - 1)
                        
                        ' Get current altitude
                        L_sngAlt = I_bytAlt(L_lngVX + L_lngVY * I_lngMapSize)
                        
                        ' Voxel hit?
                        If L_sngAlt > L_sngRZ Then
                    
                            ' Acquire color from texture map ...
                            
                                ' Calculate base adress
                                L_lngBase = (Int(L_lngVX) + Int(L_lngVY) * I_lngMapSize) * 3
                                
                                ' Aquire pixel color
                                L_bytColor(0) = I_bytTex(L_lngBase + 0)
                                L_bytColor(1) = I_bytTex(L_lngBase + 1)
                                L_bytColor(2) = I_bytTex(L_lngBase + 2)
                                
                            ' Adjust color ...
                            
                                If I_blnStateNight Then
                                    
                                    ' Calculate fog factors
                                    If I_blnStateLight Then
                                        ' Lighting enabled
                                        L_sngFogFactor = 1 - (L_lngStep * ((Abs(L_lngPX - I_udtBD.bmiHeader.biWidth / 2)) / (I_udtBD.bmiHeader.biWidth * 0.5)) / 128)
                                    Else
                                        ' Lighting disabled
                                        L_sngFogFactor = 1 - (L_lngStep / 128)
                                    End If
                                    
                                    ' Calculate actual fog value
                                    L_bytFogValue = L_lngStep / 8
                                                                            
                                    ' Calculate pixel value
                                    L_bytColor(0) = L_bytColor(0) * L_sngFogFactor + L_bytFogValue
                                    L_bytColor(1) = L_bytColor(1) * L_sngFogFactor + L_bytFogValue
                                    L_bytColor(2) = L_bytColor(2) * L_sngFogFactor + L_bytFogValue + L_bytFogValue
                                        
                                End If
                            
                            ' Acquire pixel base output adress
                            L_lngBase = (L_lngPX + L_lngPY * I_udtBD.bmiHeader.biWidth) * 3
                            
                            ' Run over voxel
                            Do
                                
                                ' Set pixel in viewport
                                L_bytDat(L_lngBase + 0) = L_bytColor(0)
                                L_bytDat(L_lngBase + 1) = L_bytColor(1)
                                L_bytDat(L_lngBase + 2) = L_bytColor(2)
                                
                                ' Advance in ray delta, z component
                                L_sngDZ = L_sngDZ + L_sngDelta
                
                                ' Advance in ray position, z component
                                L_sngRZ = L_sngRZ + L_sngScale
                                
                                ' Advance in pixel adress, row component
                                L_lngBase = L_lngBase + I_udtBD.bmiHeader.biWidth * 3
                                
                                ' Advance in viewport position, row component
                                L_lngPY = L_lngPY + 1
                                
                            ' Exit if altitude expired
                            Loop Until L_sngRZ > L_sngAlt
                            
                            ' Remember maximum altitude
                            L_intTop(L_lngPX) = L_lngPY
                            
                        End If
                        
                        ' Advance in ray position
                        L_sngRX = L_sngRX + L_sngDX
                        L_sngRY = L_sngRY + L_sngDY
                        L_sngRZ = L_sngRZ + L_sngDZ
                
                        ' Advance in voxel scale
                        L_sngScale = L_sngScale + L_sngDelta
                
                    Next
                    
                Next
                            
            ' Interpolate pixels (odd rows)
                
                If I_blnStateIlv Then
                
                    ' Run over pixel columns
                    For L_lngPX = 1 To I_udtBD.bmiHeader.biWidth - 2 Step 2
                        For L_lngPY = 0 To L_intTop(L_lngPX - 1)
                        
                            ' Calculate base adress
                            L_lngBase = (L_lngPX + L_lngPY * I_udtBD.bmiHeader.biWidth) * 3
                            
                            ' Acquire color by interpolation
                            L_bytDat(L_lngBase + 0) = I_bytAvg(L_bytDat(L_lngBase - 3), L_bytDat(L_lngBase + 3))
                            L_bytDat(L_lngBase + 1) = I_bytAvg(L_bytDat(L_lngBase - 2), L_bytDat(L_lngBase + 4))
                            L_bytDat(L_lngBase + 2) = I_bytAvg(L_bytDat(L_lngBase - 1), L_bytDat(L_lngBase + 5))
                            
                        Next
                    Next
                
                End If
                
            ' Overlay banners ...
            
                ' Only at program startup
                If I_lngFrameCount < 360 Then
                    
                    ' Calculate factor
                    L_sngFogFactor = Abs(Sin(I_lngFrameCount * PI180))
                    
                    ' Run over columns
                    For L_lngPY = 0 To 57
                    
                        ' Calculate base adress
                        L_lngBase = (I_udtBD.bmiHeader.biWidth * (L_lngPY + 160) + 60) * 3
                        L_lngSource = (L_lngPY * 360) * 3
                        
                        ' Run over rows
                        For L_lngPX = 0 To 358
                            
                            ' Mix banner into image
                            If I_lngFrameCount < 180 Then
                                L_bytDat(L_lngBase + 0) = L_bytDat(L_lngBase + 0) * (1 - L_sngFogFactor) + I_bytBn1(L_lngSource + 0) * L_sngFogFactor
                                L_bytDat(L_lngBase + 1) = L_bytDat(L_lngBase + 1) * (1 - L_sngFogFactor) + I_bytBn1(L_lngSource + 1) * L_sngFogFactor
                                L_bytDat(L_lngBase + 2) = L_bytDat(L_lngBase + 1) * (1 - L_sngFogFactor) + I_bytBn1(L_lngSource + 2) * L_sngFogFactor
                            Else
                                L_bytDat(L_lngBase + 0) = L_bytDat(L_lngBase + 0) * (1 - L_sngFogFactor) + I_bytBn2(L_lngSource + 0) * L_sngFogFactor
                                L_bytDat(L_lngBase + 1) = L_bytDat(L_lngBase + 1) * (1 - L_sngFogFactor) + I_bytBn2(L_lngSource + 1) * L_sngFogFactor
                                L_bytDat(L_lngBase + 2) = L_bytDat(L_lngBase + 1) * (1 - L_sngFogFactor) + I_bytBn2(L_lngSource + 2) * L_sngFogFactor
                            End If
                            
                            ' Advance in base adress
                            L_lngBase = L_lngBase + 3
                            L_lngSource = L_lngSource + 3
                            
                        Next
                    Next
                
                End If
            
            ' Display resultes ...
            
                ' Blit to screen
                BitBlt Me.hdc, 10, 50, 490, 390, I_lngDC, 0, 0, vbSrcCopy
                
    End Sub
    
'
    ' LBLCOMMAND_CLICK: Toggle options
    '
    Private Sub lblButton_Click(Index As Integer)
    
        ' Declare local variables ...
        
            ' (none)
            
        ' Code ...
                    
            ' Change state flags
            Select Case Index
            
                ' Option 0: Day
                Case 0
                    I_blnStateNight = False
                    I_blnStateLight = False
            
                ' Option 1: Night
                Case 1
                    I_blnStateNight = True
                    I_blnStateLight = False
            
                ' Option 2: Night with light
                Case 2
                    I_blnStateNight = True
                    I_blnStateLight = Not I_blnStateLight
            
                ' Option 3: Interleave mode
                Case 3
                    I_blnStateIlv = Not I_blnStateIlv
            
                ' Option 4: Animations
                Case 4
                    I_blnStateAni = Not I_blnStateAni
            
            End Select
                
            ' Set lamps ...
            
                Me.shpButton(0).FillColor = IIf(Not I_blnStateNight, &HE0E0E0, &H404040)
                Me.shpButton(1).FillColor = IIf(I_blnStateNight, &HE0E0E0, &H404040)
                Me.shpButton(2).FillColor = IIf(I_blnStateLight, &HE0E0E0, &H404040)
                Me.shpButton(3).FillColor = IIf(I_blnStateIlv, &HE0E0E0, &H404040)
                Me.shpButton(4).FillColor = IIf(I_blnStateAni, &HE0E0E0, &H404040)
    
    End Sub
        
    '
    ' LBLMOVE_MOUSEDOWN: Remember window and mouse start coordinates
    '
    Private Sub lblMove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        ' Declare local variables ...
        
            ' (none)
            
        ' Code ...
        
            ' Only move with left button pressed...
            If Button = 1 Then
                
                ' Get mouse and window position
                GetCursorPos I_udtMouseStart
                I_udtWindowStart.X = Me.Left \ Screen.TwipsPerPixelX
                I_udtWindowStart.Y = Me.Top \ Screen.TwipsPerPixelY
                
            End If
        
    End Sub
    
    '
    ' LBLMOVE_MOUSEMOVE: Adjust window to current coordinates of mouse
    '
    Private Sub lblMove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        ' Declare local variables ...
        
            Dim L_udtMouseNow   As POINTAPI         ' Hold current mouse position
            Dim L_udtWindowNow  As POINTAPI         ' Hold current window position
        
        ' Code ...
        
            ' Only move with left button pressed...
            If Button = 1 Then
                
                ' Get mouse and window position
                GetCursorPos L_udtMouseNow
                L_udtWindowNow.X = Me.Left \ Screen.TwipsPerPixelX
                L_udtWindowNow.Y = Me.Top \ Screen.TwipsPerPixelY
                
                ' Set window to new position
                Me.Left = (I_udtWindowStart.X + (L_udtMouseNow.X - I_udtMouseStart.X)) * Screen.TwipsPerPixelX
                Me.Top = (I_udtWindowStart.Y + (L_udtMouseNow.Y - I_udtMouseStart.Y)) * Screen.TwipsPerPixelY
                
            End If
        
    End Sub
