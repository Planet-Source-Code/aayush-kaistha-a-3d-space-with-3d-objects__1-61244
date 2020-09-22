VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################################################
'
'       Walk through 3d space containing 3d objects
'
'       Written By: Aayush Kaistha
'       Place:      UIET, Panjab University, Chandigarh
'       Contact:    aayushk_007@yahoo.com
'
'   Special thanx 2 Jack Hoxley (externalweb.exhedra.com/directx4vb)
'   for teaching me everything related to direct 3d
'
'########################################################

'3d objects were created using 3d studio max and exported
'in .3ds format which were then converted into .x format
'using the program conv3ds.exe available with direct x sdk

Option Explicit

Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim D3DX As D3DX8
Dim D3DDevice As Direct3DDevice8

Dim bRunning As Boolean
Dim UpKey As Boolean, DownKey As Boolean
Dim LeftKey As Boolean, RightKey As Boolean

Const FVF_VERTEX = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private Type VERTEX
    P As D3DVECTOR
    N As D3DVECTOR
    T As D3DVECTOR2
End Type

'all these variables r required to load 3d objects in directX
Private Type Object3D
    nMaterials As Long
    Materials() As D3DMATERIAL8
    Textures() As Direct3DTexture8
    TextureFile As String
    Mesh As D3DXMesh
End Type
    
Private Type Plyr_Data
    Pos As D3DVECTOR
    Rotation As Integer
End Type

Private Declare Function GetTickCount Lib "kernel32" () As Long

'this only holds data req to calculate frames per second
Private Type FPS_data
    Count As Long
    Value As Long
    Last As Long
End Type

Dim fps As FPS_data

Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim fnt As New StdFont

'this holds the no of 3d objects we r loading. we have
'no_obj + 1 objects in our prog
Const no_obj = 4

Dim Proj(360) As D3DVECTOR
Dim Player As Plyr_Data
Dim Obj(no_obj) As Object3D 'array of 3d objects

'we req a matrix 4 each object. we use matrices to modify
'vertices of an object so that it can b rotated, translated
'or scaled
Dim matObj(no_obj) As D3DMATRIX

Dim matProj As D3DMATRIX 'this holds the camera settings
Dim matView As D3DMATRIX 'this tells where the camera is n where it is looking at
Dim matWorld As D3DMATRIX 'this holds the reference coordinates of entire 3d world

Const PI = 3.14159
Const RAD = PI / 180

Private Sub CreateProjTable()

'if we calculate sines and cosines at run time, it slows
'down the prog. we instead store the values of sine n cos
'in an array so that we do not have to cal them at run time
Dim I As Integer, Ang As Double

For I = 0 To 360
    Ang = I * RAD
    Proj(I).X = Cos(Ang)
    Proj(I).Z = Sin(Ang)
Next

End Sub

Private Function Initialize() As Boolean

On Error GoTo Hell:

Dim D3DWindow As D3DPRESENT_PARAMETERS

'initialize and allocate memory 4 directX variables
Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate
Set D3DX = New D3DX8

'this sets our resolution settings. here we r using 16-bit
'screen format so that it runs on older computers as well
With D3DWindow
    .BackBufferCount = 1
    .BackBufferFormat = D3DFMT_R5G6B5
    .BackBufferWidth = 640
    .BackBufferHeight = 480
    .hDeviceWindow = frmMain.hWnd
    .AutoDepthStencilFormat = D3DFMT_D16
    .EnableAutoDepthStencil = 1
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
End With

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

With D3DDevice
    .SetVertexShader FVF_VERTEX
    .SetRenderState D3DRS_LIGHTING, 1
    .SetRenderState D3DRS_AMBIENT, D3DColorXRGB(255, 255, 255)
    .SetRenderState D3DRS_ZENABLE, 1
    .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
End With

'initialize our matrices
D3DXMatrixIdentity matWorld
D3DDevice.SetTransform D3DTS_WORLD, matWorld

D3DXMatrixLookAtLH matView, MakeVector(-2, 2, -2), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
D3DDevice.SetTransform D3DTS_VIEW, matView

D3DXMatrixPerspectiveFovLH matProj, PI / 3, 1, 0.1, 75
D3DDevice.SetTransform D3DTS_PROJECTION, matProj

'font settings
fnt.Name = "Verdana"
fnt.Size = 10
fnt.Bold = True
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)

CreateWorld
Initialize = True

Exit Function
Hell:
MsgBox "ERROR initializing D3D ", vbCritical, "ERROR"
Initialize = False

End Function

Private Function MakeVector(X As Single, Y As Single, Z As Single) As D3DVECTOR
    MakeVector.X = X: MakeVector.Y = Y: MakeVector.Z = Z
End Function

Private Function MakeRect(Left As Single, Right As Single, Top As Single, Bottom As Single) As RECT

MakeRect.Left = Left
MakeRect.Right = Right
MakeRect.Top = Top
MakeRect.Bottom = Bottom
 
End Function

Private Sub CreateWorld()

On Error GoTo Out:

Dim mtrlBuffer As D3DXBuffer
Dim I As Long, j As Integer, fname As String

'this loop loads 3d objects in obj() array
For j = 0 To no_obj
    fname = App.Path + "\" + Trim$(Str(j)) + ".x"
    Set Obj(j).Mesh = D3DX.LoadMeshFromX(fname, D3DXMESH_MANAGED, D3DDevice, Nothing, mtrlBuffer, Obj(j).nMaterials)
    
    ReDim Obj(j).Materials(Obj(j).nMaterials) As D3DMATERIAL8
    ReDim Obj(j).Textures(Obj(j).nMaterials) As Direct3DTexture8

    For I = 0 To Obj(j).nMaterials - 1
        D3DX.BufferGetMaterial mtrlBuffer, I, Obj(j).Materials(I)
        Obj(j).Materials(I).Ambient = Obj(j).Materials(I).diffuse
        Obj(j).TextureFile = D3DX.BufferGetTextureName(mtrlBuffer, I)
        If Obj(j).TextureFile <> "" Then
            Set Obj(j).Textures(I) = D3DX.CreateTextureFromFile(D3DDevice, App.Path + "\" + Obj(j).TextureFile)
        End If
    Next
Next

Exit Sub
Out:
    MsgBox "Error loading models", vbCritical, "ERROR"
End Sub

Private Sub Form_Click()
bRunning = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = True
If KeyCode = vbKeyDown Then DownKey = True
If KeyCode = vbKeyLeft Then LeftKey = True
If KeyCode = vbKeyRight Then RightKey = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then UpKey = False
If KeyCode = vbKeyDown Then DownKey = False
If KeyCode = vbKeyLeft Then LeftKey = False
If KeyCode = vbKeyRight Then RightKey = False

End Sub

Private Sub Form_Load()

Dim matTemp As D3DMATRIX, j As Integer, LastUpdated As Long

Player.Pos.X = 0
Player.Pos.Y = 4
Player.Pos.Z = 0
Player.Rotation = 0

CreateProjTable
bRunning = Initialize

'scale all objects to 1/4th of their size
For j = 0 To no_obj
    D3DXMatrixScaling matObj(j), 1, 1, 1
    D3DXMatrixIdentity matTemp
    D3DXMatrixScaling matTemp, 0.25, 0.25, 0.25
    D3DXMatrixMultiply matObj(j), matObj(j), matTemp
Next

'we use translation matrix to move our objects to desired
'location in space
D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 40, 0, 0
D3DXMatrixMultiply matObj(0), matObj(0), matTemp

D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, -20, 0, 0
D3DXMatrixMultiply matObj(1), matObj(1), matTemp

D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 10, 0, 10
D3DXMatrixMultiply matObj(2), matObj(2), matTemp

D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 5, 0, 40
D3DXMatrixMultiply matObj(3), matObj(3), matTemp

D3DXMatrixIdentity matTemp
D3DXMatrixTranslation matTemp, 20, 0, 5
D3DXMatrixMultiply matObj(4), matObj(4), matTemp

fps.Last = GetTickCount
LastUpdated = GetTickCount

Do While bRunning
    'limit the fps to max of 100 so that our prog does not
    'run too fast
    If ((GetTickCount - LastUpdated) >= 10) Then
        LastUpdated = GetTickCount
        CheckKeys
    
        'as we move using arrow keys, make the camera
        'follow our position
        D3DXMatrixLookAtLH matView, Player.Pos, MakeVector(Player.Pos.X + (Proj(Player.Rotation).X * 10), 2, Player.Pos.Z + (Proj(Player.Rotation).Z * 10)), MakeVector(0, 1, 0)
        D3DDevice.SetTransform D3DTS_VIEW, matView
    
        D3DXMatrixScaling matWorld, 1, 1, 1
        D3DDevice.SetTransform D3DTS_WORLD, matWorld
    
        Render
    
        fps.Count = fps.Count + 1
        If ((GetTickCount - fps.Last) >= 1000) Then
            fps.Value = fps.Count
            fps.Count = 0
            fps.Last = GetTickCount
        End If
    
        DoEvents
    End If
Loop
    
Set D3DX = Nothing
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing

End

End Sub

Private Sub Render()
Dim I As Long, j As Integer

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0, 1#, 0

D3DDevice.BeginScene
    For j = 0 To no_obj
        D3DDevice.SetTransform D3DTS_WORLD, matObj(j)
        For I = 0 To Obj(j).nMaterials - 1
            D3DDevice.SetTexture 0, Obj(j).Textures(I)
            D3DDevice.SetMaterial Obj(j).Materials(I)
            Obj(j).Mesh.DrawSubset I
        Next
    Next
    
    D3DX.DrawText MainFont, &HFFFFCC00, "Position : [ " + Str(Player.Pos.X) + " , " + Str(Player.Pos.Z) + " ]", MakeRect(10, 640, 0, 15), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "Rotation : " + Str(Player.Rotation), MakeRect(10, 640, 20, 35), DT_TOP Or DT_LEFT
    D3DX.DrawText MainFont, &HFFFFCC00, "FPS : " + Str(fps.Value), MakeRect(10, 640, 40, 55), DT_TOP Or DT_LEFT
D3DDevice.EndScene

D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Private Sub CheckKeys()

If LeftKey Then Player.Rotation = Player.Rotation + 2
If RightKey Then Player.Rotation = Player.Rotation - 2
If Player.Rotation < 0 Then Player.Rotation = 360
If Player.Rotation > 360 Then Player.Rotation = 0

If UpKey Then
    Player.Pos.X = Player.Pos.X + Proj(Player.Rotation).X
    Player.Pos.Z = Player.Pos.Z + Proj(Player.Rotation).Z
End If
If DownKey Then
    Player.Pos.X = Player.Pos.X - Proj(Player.Rotation).X
    Player.Pos.Z = Player.Pos.Z - Proj(Player.Rotation).Z
End If

If Player.Pos.X > 50 Then Player.Pos.X = 50
If Player.Pos.X < -50 Then Player.Pos.X = -50
If Player.Pos.Z > 50 Then Player.Pos.Z = 50
If Player.Pos.Z < -50 Then Player.Pos.Z = -50

End Sub
