Attribute VB_Name = "fStream"
'Fundamental Data Scan.  -Read each line of code and scan
'    for seperate vertices' coordinates.

Type DatCoord               ' Data Coord Array
  I As Integer              ' Vertex
  X As Double
  Y As Double
  Z As Double
End Type

Type DatVertex              ' Vertex Array data
     Count As Integer       ' Number of Vertices
     Vert(100) As Integer   ' Pointers to DatCoord
End Type

'Arrays
Public FileCoord() As DatCoord
Public FileVertex() As DatVertex

'Maximum/Minimum number of vertices in a proj.
Public MaxVertNr As Integer
Public MinVertNr As Integer

Sub GetVertex(strInput As String, FV As DatVertex)
 ' Get Vertices (simple X, Y, Z format with array from *.DAT)
 
 Dim iCurChar As Integer
 Dim strVal As String
 Dim numVal As Integer
 Dim b As Integer
 Dim Ch As String * 1

   'scan the string
   For iCurChar = 1 To Len(strInput)
       Ch = Mid$(strInput, iCurChar, 1)
       If Ch <> " " Then
          strVal = strVal + Ch
       Else
          If Len(strVal) > 0 Then
               numVal = Val(strVal)
               b = b + 1
               FV.Vert(b) = Val(strVal)
               strVal = ""
          End If
      End If
  Next iCurChar
   
' Complete last Vertex
    If Len(strVal) > 0 Then
       b = b + 1
       FV.Vert(b) = Val(strVal)
    End If
    
    FV.Count = b
End Sub

Function LoadFile(File As String) As Integer
  Dim strChar As String
  Dim nn As Integer
  Dim Facce As Integer
  Dim I As Integer
  Dim X As Double, Y As Double, Z As Double
  Dim Pl As Integer
  Dim m As Integer
  Dim Ps As Integer
  Dim Vrt
  
  Erase FileCoord
  Erase FileVertex
  
  LoadFile = True
  
  On Error Resume Next
  nn = FreeFile
  Open File For Input As nn
  
  If Err <> 0 Then
     LoadFile = False
     Exit Function
  End If
  FiletoOpen = nn
  
 On Error GoTo 0
 ReDim FileCoord(1)

 Do Until EOF(nn)
   
   Line Input #nn, strChar
   
   'Commented Lines should be ignored
   If Mid$(strChar, 1, 1) = "`" Then GoTo IgnoreComment
   
   'Is it a Polygon Face or a Vertex?
   If Mid$(strChar, 1, 6) = "Faces:" Then
      'Polygon
      Facce = True
      Line Input #nn, strChar
   End If
      
   If Not Facce Then
      'Vertex
      Vrt = Vrt + 1
      
      Call GetCoord(strChar, I, X, Y, Z)
      
      'Invalid Format
      If St = "FILE NOT VALID" Then
         LoadFile = False
         Exit Function
      End If
      
      'Make sure enough space in array FileCoord
      If Vrt > UBound(FileCoord) Then ReDim Preserve FileCoord(Vrt)
      
      'Set Vertex Coord
      FileCoord(Vrt).I = I
      FileCoord(Vrt).X = X
      FileCoord(Vrt).Y = Y
      FileCoord(Vrt).Z = Z
   Else
        'it's a polygon
         Pl = Pl + 1
         ReDim Preserve FileVertex(Pl)
         GetVertex strChar, FileVertex(Pl)
   End If
   
IgnoreComment:
Loop
   
Close nn%

SetLimits
End Function

Sub GetCoord(strInput As String, I As Integer, X As Double, Y As Double, Z As Double)
On Error Resume Next
'Get the coord of a vertex from a line with template:
' --> [Vertex Number] [X] [Y] [Z] with X,Y,Z values for 3d graph
 
Dim iCurChar As Integer
Dim strVal As String
Dim numVal As Double
Dim b As Integer
Dim Ch As String
 
   'Begin loop for going through each character
   For iCurChar = 1 To Len(strInput)
   
       'Current Character
       Ch = Mid$(strInput, iCurChar, 1)
       'vbNullString?
       If Ch <> " " Then
          strVal = strVal + Ch
       Else
          If Len(strVal) > 0 Then
               'Get integer val of string
               numVal = Val(strVal)
               b = b + 1
               Select Case b
                 'get character respondent
                 Case 1
                    I = numVal
                 Case 2
                    X = numVal
                 Case 3
                    Y = numVal
                 Case 4
                    Z = numVal
              End Select
              'reset string
              strVal = ""
          End If
      End If
  Next iCurChar
   
' Complete the Z
If Len(strVal) > 0 Then Z = Val(strVal)
'Error in file.
If I + X + Y + Z = 0 Then strInput = "FILE NOT VALID"
End Sub


Sub SetLimits()
' Return max vertices
Dim I As Integer
Dim K As Integer
Dim X As Double
Dim Y As Double
Dim Z As Double

' Assign total vetex and object dimension (min,max)

xmin = BIG
xmax = -BIG
ymin = BIG
ymax = -BIG
zmax = -BIG
zmin = BIG

For K = 1 To UBound(FileCoord)
    'For each coord
    I = FileCoord(K).I
    X = FileCoord(K).X
    Y = FileCoord(K).Y
    Z = FileCoord(K).Z
    
    'Reset values if invalid
    If (I > MaxVertNr) Then MaxVertNr = I
    If (I < MinVertNr) Then MinVertNr = I
    If (X < xmin) Then xmin = X
    If (X > xmax) Then xmax = X
    If (Y < ymin) Then ymin = Y
    If (Y > ymax) Then ymax = Y
    If (Z < zmin) Then zmin = Z
    If (Z > zmax) Then zmax = Z
Next
End Sub
