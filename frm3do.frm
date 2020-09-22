VERSION 5.00
Begin VB.Form ThreeDObject 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "3D Object Render - Graphics Engine"
   ClientHeight    =   5040
   ClientLeft      =   1080
   ClientTop       =   1470
   ClientWidth     =   8070
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frm3do.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   538
   Begin VB.PictureBox picBrownRed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   3360
      ScaleHeight     =   465
      ScaleWidth      =   4695
      TabIndex        =   3
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdRender 
         Caption         =   "Render!"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdWireFrame 
         Caption         =   "Wire Frame"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdShade 
         Caption         =   "Shaded"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.PictureBox Pict 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   305
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   539
      TabIndex        =   0
      Top             =   480
      Width           =   8085
      Begin VB.Label lblObject 
         Caption         =   " Object Name: None"
         Height          =   255
         Left            =   5460
         TabIndex        =   9
         Top             =   3960
         Width           =   2535
      End
      Begin VB.Label lblRender 
         Caption         =   " Render Mode: Wireframe"
         Height          =   255
         Left            =   5460
         TabIndex        =   8
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label lblCurrentSettings 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Current Settings:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5760
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
   End
End
Attribute VB_Name = "ThreeDObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'3D Object RENDER - A SIMPLE GRAPHICS RENDERING ENGINE
'  Bradley Liang -> http://prgmrsruin.hypermart.net
'  Services -> Patrice Scribe's Win32 API type lib

'  WHAT IS THIS?
'This is my last attempt at a graphic-rendering engine before
'I learn Microsoft's DirectX 7 programming.  I hadn't realized
'how much faster DirectX is at coding than my engine.  For now,
'this engine is scrapped and I revamped the code one last time
'adding comments for submission.  Still not commented as well
'as I'd like, but if you're looking at this, you probably know
'a lot of the mathematics behind it.

'CONCEPT:
'This is a loading program to what was supposed to be a
'3D Chess game with idea of moving objects from an isometric
'view.  This ran into complications and as a result, I have
'removed the coding.  The loading of many tiny figures would
'take minutes, plus would have to account for each piece's
'coordinates and polygons (I tried an array, this is TOO SLOW)
'movement of the pieces took longer than usual and with the
'computer's AI algorithm running all the time (I used this
'approach in order to have the computer move ASAP.  Very
'Bug-ridden creation, and SLOW), it ate too many resources.
'Perhaps using a simple, 2D board and 2D pieces are a better
'approach when one doesn't have the greatest programming skill.

'THIS SHOULD NEVER BE USED AS THE FUNDAMENTALS OF ANY 3D ENGINE
'DirectX is the way to go (though I don't know it yet).  Still,
'knowing the basics helps when going into more advanced and/or
'higher-level programming.  This engine is relatively efficient
'for not using any references or components, but it can still
'be improved (especially the error handling)

'  (*.DAT) CODING FORMAT:
'This is not the .3do file format, though it can be altered
'for it. This coding is based on the .3do (ver 3) format with
'the exception of the Color Fill data input.  Polygons are
'grouped into one section, and vertices in another, adding
'speed as well as organization, though less convenience.
'This engine is not a 8-square quadrant engine as most 3d
'ones are.  the first quadrant, with each side accomadating
'new values.
' -1      0      1
'  |      |      |
'  V      V      V
'   _ _ _ _______    /__ 1
' /      /       /|  \         All the objects fit into this
'       /       / |            one (1) quadrant with the def-
' _ _  /______ /  |  /__ 0     -initions of spacing between
'     |       |  /   \         them defined by the highest and
'     |   I   | /              lowest values scanned by the
' _ _ |_______|/     /__ 1     datafile.
'                    \
'     ^       ^                RESULT: The object fills the
'     0       1                target screen size at one given
'                              size for all objects.
'
'This doesn't mean that a line in the datafile reading
' --> 7 16.003 0.0400 0.303 is incorrect, because the 16.003
'(X scale) could be on a scale of -20 to 20 while Y and Z values
'are out of 1.000  I recommend that you keep one fixed max
'and min because the plotting may get confusing.

'FORMAT OF EACH LINE:
'Using the above example, (7 16.003 0.04 0.303) we see the data
'format for plotting the object.  Each of these lines is a
'coordinate. The [7] is the vertex number, used for the shading
'values later.  [16.003] is the X value on the quadrant.
'[0.04] is the Y value, and [0.303] is the Z value on the quad.
'each of these coords are stored into an array and then plotted
'one by one (speed?).
'When you come to a line with FACES: you should know that the
'rest of the lines will be for polygonal faces.  These use
'references from the vertices. A line of:  3 4 9 1. (notice
'the space before the 3 and the period after the 1) will plot
'a polygon touching the 1st, 3rd, 4th, and 9th coordinates.

'WIREFRAME:
'The wireframe rendering is simply drawing lines at each side
'of each polygonal face.  These lines neither reflect depth
'nor faces while being clearer with the smaller objects.

'SHADING THE OBJECT:
'This can be found within painter.bas (paint3d) and uses a
'simple algorithm to calculate the lighting vector (which I
'know I should have had an option on here for, it is currently
'always set to a certain degree angle) and then shading in
'certain parts with the angle value that it is if it was tilted
'to the degree and the light was from the top.  This idea was
'taken from a graphics tutorial I read several months ago
'(I'm sorry, I forgot the author and the link so I can't give
'out recognition to the person who designed the first stage
'of it).



Dim RenderType As Integer ' 0 Wire, 1 = Shade
Dim FileOpen As String

Sub Render()
 If Len(FileOpen) = 0 Then
    MsgBox "Select a valid coordinate file!", 16
    Exit Sub
 End If
 
 Pict.Cls
 Render3d Pict, RenderType
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdOpen_Click()
Dim strFilter As String
Dim strFileName As String, sFileTitle As String
Dim ofn As OPENFILENAME

'Please wait screen...
Load frmWait
Me.Visible = False

'OPEN DIALOG BOX --> Patrice Scribe
' Filter
strFilter = "Coordinate File (*.DAT)"
' File name buffer
strFileName = String$(255, 0)
With ofn
    .lStructSize = LenB(ofn)
    .hwndOwner = hWnd
    .nFilterIndex = 1
    ' Pointer to ANSI string (Visual Basic strings are Unicode)
    .lpstrFilter = StrPtr(StrConv(strFilter, vbFromUnicode))
    .nMaxFile = Len(strFileName)
    .lpstrFile = StrPtr(strFileName)
End With
        
If GetOpenFileName(ofn) Then
    ' Convert the string buffer from ANSI to Unicode
    strFileName = StrConv(strFileName, vbUnicode)
    ' Keep to first vbNullChar
    strFileName = Left$(strFileName, InStr(strFileName, vbNullChar) - 1)
    
    FileOpen = strFileName
    
    frmWait.Visible = True
    SetWindowPos frmWait.hWnd, -1, 0, 0, 0, 0, &H2 Or &H1
    frmWait.lblStatus = "Reading Data"
    
    'extract the file title
    sFileTitle = strFileName
    While InStr(1, sFileTitle, "\")
        sFileTitle = Mid(sFileTitle, InStr(1, sFileTitle, "\") + 1)
    Wend
    For I = 0 To 300: DoEvents: Next
    
    'show user the New data
    frmWait.lblStatus = "Implanting Data"
    lblObject.Caption = "Object Name: " + Mid(sFileTitle, 1, Len(sFileTitle) - 4)
    
    'Invalid File
    If Not LoadFile(strFileName) Then
        MsgBox "File not valid", 16, "Error opening file"
        lblObject.Caption = "Object Name: None"
        FileOpen = ""
        Exit Sub
    End If
    For I = 0 To 600: DoEvents: Next

'Set COlors of shading + RENDER
 frmWait.lblStatus = "Rendering"
 Me.Visible = True
 SetColors ThreeDObject.Pict
 Render
 For I = 0 To 800: DoEvents: Next
    
frmWait.lblStatus = "Done"
frmWait.Visible = False
End If

Pict.SetFocus
End Sub

Private Sub cmdRender_Click()
Render
Pict.SetFocus
End Sub

Private Sub cmdShade_Click()
lblRender.Caption = "Render Mode: Shading"
RenderType = 1
Pict.SetFocus
End Sub

Private Sub cmdWireFrame_Click()
lblRender.Caption = "Render Mode: Wireframe"
RenderType = 0
Pict.SetFocus
End Sub

Private Sub Form_Load()
  X__max = Pict.ScaleWidth
  Y__max = Pict.ScaleHeight
End Sub

Private Sub Pict_KeyPress(KeyAscii As Integer)
If Asc(Chr(13)) = KeyAscii Then
    Render
End If
End Sub
