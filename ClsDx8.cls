VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDx8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'DX 8 Header stuff


'Dx Enumeration
Dim nAdapters As Long
Dim AdapterInfo As D3DADAPTER_IDENTIFIER8

Dim nModes As Long


'FPS Stuff
Private Declare Function GetTickCount Lib "kernel32" () As Long '//This is used to get the frame rate.
Dim LastTimeCheckFPS As Long
Dim FramesDrawn As Long
Dim FrameRate As Long

Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR


'ColorKeyVal = &HFF000000 '//Black
'ColorKeyVal = &HFFFF0000 '//Red
'ColorKeyVal = &HFF00FF00 '//Green
'ColorKeyVal = &HFF0000FF '//Blue
'ColorKeyVal = &HFFFF00FF '//Magenta
'ColorKeyVal = &HFFFFFF00 '//Yellow
'ColorKeyVal = &HFF00FFFF '//Cyan
'ColorKeyVal = &HFFFFFFFF '//White

Const DefColorKey = &HFF000000

'#########
'## FONTS ##
'#########
Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim TextRect As RECT
Dim fnt As New StdFont



'Temporary data for storing screen data saved in a listbox

Public Function Initialize(Form As Form, Width As Integer, Height As Integer, BPP As Integer, Fullscreen As Boolean) As Boolean
On Error GoTo ErrHandler:

Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS

Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()
Set D3DX = New D3DX8 '//Create our helper library...

If Fullscreen = False Then

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode '//Retrieve the current display Mode

D3DWindow.Windowed = 1 '//Tell it we're using Windowed Mode
D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC '//We'll refresh when the monitor does
D3DWindow.BackBufferFormat = DispMode.Format '//We'll use the format we just retrieved...

Else

DispMode.Format = D3DFMT_X8R8G8B8
DispMode.Width = ScrnWidth(FrmEnum.cmbRes.ListIndex)
DispMode.Height = ScrnHeight(FrmEnum.cmbRes.ListIndex)
D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferCount = 1 '//1 backbuffer only
D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
D3DWindow.BackBufferHeight = ScrnHeight(FrmEnum.cmbRes.ListIndex)
D3DWindow.BackBufferWidth = ScrnWidth(FrmEnum.cmbRes.ListIndex)
D3DWindow.hDeviceWindow = Form.hWnd

End If

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Form.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                        D3DWindow)

D3DDevice.SetVertexShader FVF

D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

Initialise = True
Exit Function

ErrHandler:

MsgBox "Error Number Returned: " & Err.Number
Initialise = False
Unload
End
End Function

Public Function EnumerateAdapters(ListBox)

Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()

    Dim I As Integer, sTemp As String, J As Integer
        
    nAdapters = D3D.GetAdapterCount
    
    ListBox.Clear
    
    For I = 0 To nAdapters - 1
        
        D3D.GetAdapterIdentifier I, 0, AdapterInfo
        
        sTemp = ""
        For J = 0 To 511
            sTemp = sTemp & Chr$(AdapterInfo.Description(J))
        Next J
        sTemp = Replace(sTemp, Chr$(0), " ")
        ListBox.AddItem sTemp
    Next I
    
    Set Dx = Nothing
    Set D3D = Nothing
    
End Function

Public Function EnumerateDevices(ListBox, Renderer As Long)
On Local Error Resume Next

Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()

Dim Caps As D3DCAPS8

D3D.GetDeviceCaps Renderer, D3DDEVTYPE_HAL, Caps
    
ListBox.Clear
    
    If Err.Number = D3DERR_NOTAVAILABLE Then
        
        ListBox.AddItem "Reference Rasterizer (REF)"
    Else
        ListBox.AddItem "Hardware Acceleration (HAL)"
        ListBox.AddItem "Reference Rasterizer (REF)"
    End If
        
    Set Dx = Nothing
    Set D3D = Nothing
    
End Function

Public Sub EnumerateDisplayModes(ListBox As ComboBox, DisplayDeviceLst As ComboBox, HardwareRendererLst As ComboBox)

ListBox.Clear '//Remove any existing entries...

Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()

Dim I As Integer, ModeTemp As D3DDISPLAYMODE

Dim Renderer As Long

Renderer = HardwareRendererLst.ListIndex + 1

nModes = D3D.GetAdapterModeCount(DisplayDeviceLst.ListIndex)

For I = 0 To nModes - 1

    Call D3D.EnumAdapterModes(DisplayDeviceLst.ListIndex, I, ModeTemp)
    
    ScrnWidth(I) = ModeTemp.Width
    ScrnHeight(I) = ModeTemp.Height
    
    
    If ModeTemp.Format = D3DFMT_R8G8B8 Or ModeTemp.Format = D3DFMT_X8R8G8B8 Or ModeTemp.Format = D3DFMT_A8R8G8B8 Then
        
        If D3D.CheckDeviceType(DisplayDeviceLst.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
                
                ScrnBPP(I) = 32
                    
        End If
    Else
        If D3D.CheckDeviceType(DisplayDeviceLst.ListIndex, Renderer, ModeTemp.Format, ModeTemp.Format, False) >= 0 Then
                
                ScrnBPP(I) = 16
        
        End If
    End If
    
ListBox.AddItem ScrnWidth(I) & "x" & ScrnHeight(I) & " " & ScrnBPP(I) & " Bit"
    
Next I

ListBox.ListIndex = ListBox.ListCount - 7

       
    Set Dx = Nothing
    Set D3D = Nothing
    
End Sub

Public Function Unload()
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
Debug.Print "All Objects Destroyed"
End Function


Public Function GetFPS() As Long
    If GetTickCount - LastTimeCheckFPS >= 1000 Then
        LastTimeCheckFPS = GetTickCount
        FrameRate = FramesDrawn '//Store the frame count
        FramesDrawn = 0 '//Reset the counter
    End If
FramesDrawn = FramesDrawn + 1
GetFPS = FrameRate
End Function

Public Function ClearDevice()
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, &HCCCCFF, 1#, 0
End Function

Public Function Begin3D()
D3DDevice.BeginScene
End Function

Public Function End3D()
D3DDevice.EndScene
End Function

Public Function FlipPrimary()
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Function

Public Function GenerateTexture(Surface As Direct3DBaseTexture8, Filepath As String, W As Integer, H As Integer, ColorKey As Boolean)

Dim ColorKeyVal As Long

If ColorKey = True Then
ColorKeyVal = DefColorKey
Else
ColorKeyVal = 0
End If

If ColorKey = True Then

Set Surface = D3DX.CreateTextureFromFileEx(D3DDevice, Filepath, W, H, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)

Else

Set Surface = D3DX.CreateTextureFromFile(D3DDevice, Filepath)

End If

End Function

Public Function SetFont(StrName As String, IntSize As Integer)
fnt.Name = StrName
fnt.Size = IntSize
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
End Function

Public Function DrawText(X As Integer, Y As Integer, StrText As String, IntColor As Long, Optional TextMode As Long)

'Text modes:
'DT_LEFT          - The text appears along the left of the RECT (Left Justify)
'DT_TOP           - The text appears across the top of the RECT
'DT_CENTER        - The text is centered horizontally in the RECT (Center Justify)
'DT_RIGHT         - The text is along the right edge of the RECT (Right Justify)
'DT_VCENTER       - The text is centered vertically
'DT_BOTTOM        - The text appears along the bottom edge
'DT_SINGLELINE    - The string passed is a single line - line feeds and carriage returns are ignored

TextRect.Top = Y
TextRect.Left = X
TextRect.bottom = Y + (fnt.Size * 2)
TextRect.Right = Len(StrText) * fnt.Size
D3DX.DrawText MainFont, &HFFCCCCFF, StrText, TextRect, TextMode

End Function
