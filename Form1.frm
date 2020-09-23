VERSION 5.00
Object = "{08216199-47EA-11D3-9479-00AA006C473C}#2.1#0"; "RMCONTROL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin RMControl7.RMCanvas RMCanvas 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'RM Control Sample v1.0
'Author: Dustin Davis
'VB-Live.com
'
'I wrote this to show you how EASY it is to use the RM control
'and that you can make really cool stuff with this.
'Enjoy!
'
'I used 3D studio max R3 to create the scene and then converted
'the .3ds files to .x files using the conv3ds.exe program that
'comes with DirectX SDK
'mssdk\bin\dxutils\xfiles
'
'NOTE: I turned this into a screen saver, so there are
'Screensaver forms and code for this to be a screensaver
'Just ignor it.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Declare our main variables
Dim m_Rm As Direct3DRM3
Dim g_Dx As New DirectX7
'declare our frame animation variables
Dim m_NameAnimation As Direct3DRMAnimationSet2
Dim m_CircleAnimation As Direct3DRMAnimationSet2
Dim m_Circle2Animation As Direct3DRMAnimationSet2
'declare our frame variables
Dim m_Circle As Direct3DRMFrame3
Dim m_rootFrame As Direct3DRMFrame3
Dim m_NameFrame As Direct3DRMFrame3
Dim m_Circle2 As Direct3DRMFrame3
'declare our light variables
Dim m_light As Direct3DRMLight
Dim c_light As Direct3DRMLight
Dim c_light2 As Direct3DRMLight
'declare loop variables
Dim KeepGoing As Boolean

Public Sub Init()
'This is where the magic happens!

'Tell the rm control to start windowed. I dont want to deal with
'Direct Draw to change the screen modes to full right now
b = RMCanvas.StartWindowed
'Check if it can be loaded
If b = False Then
    MsgBox "Cant start 3D window"
    End
End If

'Set the scene background color
RMCanvas.SceneFrame.SetSceneBackgroundRGB 0, 0, 0

'Make the RM control
Set m_Rm = g_Dx.Direct3DRMCreate

'Set the frames
Set m_rootFrame = m_Rm.CreateFrame(Nothing)
Set m_NameFrame = m_Rm.CreateFrame(m_rootFrame)
Set m_Circle = m_Rm.CreateFrame(m_rootFrame)
Set m_Circle2 = m_Rm.CreateFrame(m_rootFrame)

'Create the frames
Set m_NameFrame = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
Set m_Circle = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)
Set m_Circle2 = RMCanvas.D3DRM.CreateFrame(RMCanvas.SceneFrame)

'create animation frames
Set m_NameAnimation = RMCanvas.D3DRM.CreateAnimationSet()
Set m_CircleAnimation = RMCanvas.D3DRM.CreateAnimationSet()
Set m_Circle2Animation = RMCanvas.D3DRM.CreateAnimationSet()

'Set the animation frame properties
m_NameAnimation.LoadFromFile "d.x", 0, 0, Nothing, Nothing, m_NameFrame
m_CircleAnimation.LoadFromFile "n4u.x", 0, 0, Nothing, Nothing, m_Circle
m_Circle2Animation.LoadFromFile "n4u.x", 1, 0, Nothing, Nothing, m_Circle2

'Set the orientation of the frames       -1 = reverse
m_NameFrame.SetOrientation Nothing, 0, 0, 1, 0, 1, 0
m_Circle.SetOrientation Nothing, 0, 0, 1, 0, 1, 0
m_Circle2.SetOrientation Nothing, 0, 0, 1, 0, 1, 0

'Add scale to the scenes
m_NameFrame.AddScale D3DRMCOMBINE_AFTER, 0.3, 0.3, 0.3
m_Circle.AddScale D3DRMCOMBINE_AFTER, 0.3, 0.3, 0.3
m_Circle2.AddScale D3DRMCOMBINE_AFTER, 0.3, 0.3, 0.3

'Set position and viewport of camera
RMCanvas.CameraFrame.SetPosition Nothing, x, Y, z
RMCanvas.Viewport.SetBack 1000

'Create lights
Set m_light = m_Rm.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0, 0, 0.3)
Set c_light = m_Rm.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0, 0.4, 0)
Set c_light2 = m_Rm.CreateLightRGB(D3DRMLIGHT_DIRECTIONAL, 0.4, 0, 0)

'Attatch lights to frames
m_NameFrame.AddLight m_light
m_Circle.AddLight c_light
m_Circle2.AddLight c_light2

'Set the position of the frames
m_Circle.SetPosition m_rootFrame, 0, 8, 60
m_Circle2.SetPosition m_rootFrame, 10, 0, 60
m_NameFrame.SetPosition m_rootFrame, 0, 5, 120

'Set scene speed, this is for animation, but since we
'only rotate the frames, this isnt needed
RMCanvas.SceneSpeed = 30

'Set the rotation of the frames, this controls how they spin
'and how fast                    X  Y  Z  Speed
m_NameFrame.SetRotation Nothing, 0, 1, 0, 0.02
m_Circle.SetRotation Nothing, 0, -1, 0, 0.3
m_Circle2.SetRotation Nothing, 0, 1, 0, 0.2
    
'Start the loop
Do Until KeepGoing = False
        'Update the scene
        RMCanvas.Update
        DoEvents
Loop

Exit Sub

End Sub


Private Sub Form_Activate()
'for screen saver purposes, really not needed
If App.PrevInstance = True Then Unload Me

AlwaysOnTop Me, True

Static b As Boolean
    KeepGoing = True
    If b = True Then End
    b = True
    
    Init
    Me.Show
    
    End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
End
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
    RMCanvas.Width = Me.ScaleWidth
    RMCanvas.Height = Me.ScaleHeight
    RMCanvas.Viewport.SetBack 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
m_running = False
KeepGoing = False
End Sub

Private Sub RMCanvas_KeyPress(KeyAscii As Integer)
End
End Sub

Private Sub RMCanvas_SceneMove(delta As Single)
m_time = m_time + delta
m_NameAnimation.SetTime m_time
End Sub

Sub AlwaysOnTop(FrmID As Form, OnTop As Integer)
'This function came from planet-source-code.com!


    ' This function uses an argument to dete
    '     rmine whether
    ' to make the specified form always on t
    '     op or not
    Const SWP_NOMOVE = 2
    Const SWP_NOSIZE = 1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2


    If OnTop Then
        OnTop = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        OnTop = SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

