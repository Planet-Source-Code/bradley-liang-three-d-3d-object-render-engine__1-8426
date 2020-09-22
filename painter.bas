Attribute VB_Name = "paint3d"
' Painter's algorithm - calculate light vector and angle
' degrees for brightness.
Type Vec_Int
   X As Integer
   Y As Integer
End Type

Type Vertex
   X As Integer
   Y As Integer
   Z As Integer
   Used As Integer
End Type
Public v() As Vertex

Type TriaData
   Normal As Vec3
   h As Single
   Color As Integer
End Type

Type Tria
    Anr As Integer
    Bnr As Integer
    Cnr As Integer
    Z As Integer
    PTria As TriaData ' Val Triadata
End Type

Public Triangles() As Tria
Public Ptriangles As Integer

Type TriaNode
    trnr As Integer
    NextNodo As Integer
End Type


Public Const LARGE = 32000
Public Const LARGE1 = 32001
Public Const PIdiv180 = 0.0174532
Public Const BIG = 1E+30

Public X__max As Integer ' real coordinate of window
Public Y__max As Integer

Public f As Double
Public zfactor As Double
Public rcolormin As Double
Public rcolormax As Double
Public delta As Double
Public zemin As Double
Public zemax As Double
Public xsC As Double
Public ysC As Double
Public XLCreal As Double
Public YLCreal As Double

Public kEye As Integer
Public hK As Integer
Public XLC As Integer
Public YLC As Integer

Public Vt() As Vec3
Public LightVector As Vec3

Public TriaNodeNext() As TriaNode

Public PStart As Integer
Public PEnd As Integer


Sub Complete_Triangles(n As Integer, offset As Integer, nrs_tr() As Trianrs)

' Complete triangles[offset],..., triangles[offset+n-1].
' Vertex numbers: nrs_tr[0],..., nrs_tr[n-1].
' This triangles are of same polygon. EtDatauation of plane is
' nx . x + ny . y + nz . z = h.
  
  Dim I As Integer
  Dim Anr As Integer, Bnr As Integer, Cnr As Integer
  Dim ZA As Integer, ZB As Integer, ZC As Integer
  Dim zmin As Single, zmax As Single

  Dim nx As Double, ny As Double, nz As Double
  Dim ux As Double, uy As Double, uz As Double
  Dim vx As Double, vy As Double, vz As Double
  
  Dim factor As Double
  Dim h As Double
  
  Dim Ax As Double, Ay As Double, Az As Double
  Dim Bx As Double, By As Double, Bz As Double
  Dim Cx As Double, Cy As Double, Cz As Double
  
  Dim p As Integer, tData As TriaData
  

  For I = n \ 2 To n
     Anr = nrs_tr(I).a
     Bnr = nrs_tr(I).b
     Cnr = nrs_tr(I).C
     If (SetOrient(Anr, Bnr, Cnr) > 0) Then Exit For
   Next

   ZA = v(Anr).Z
   ZB = v(Bnr).Z
   ZC = v(Cnr).Z

   Az = ZEye(ZA): Bz = ZEye(ZB): Cz = ZEye(ZC)
   Ax = XScreen(v(Anr).X) * Az: Ay = YScreen(v(Anr).Y) * Az
   Bx = XScreen(v(Bnr).X) * Bz: By = YScreen(v(Bnr).Y) * Bz
   Cx = XScreen(v(Cnr).X) * Cz: Cy = YScreen(v(Cnr).Y) * Cz
   
   ux = Bx - Ax: uy = By - Ay: uz = Bz - Az
   vx = Cx - Ax: vy = Cy - Ay: vz = Cz - Az
   nx = uy * vz - uz * vy: ny = uz * vx - ux * vz: nz = ux * vy - uy * vx
   
   h = nx * Ax + ny * Ay + nz * Az
   factor = 1 / Sqr(nx * nx + ny * ny + nz * nz)
   
   tData.Normal.X = nx * factor
   tData.Normal.Y = ny * factor
   tData.Normal.Z = nz * factor
   tData.h = h * factor
   
   For I = 0 To n - 1
      p = offset + I
      Triangles(p).Anr = nrs_tr(I).a
      Triangles(p).Bnr = nrs_tr(I).b
      Triangles(p).Cnr = nrs_tr(I).C
      Triangles(p).PTria = tData
      
      zmin = v(Triangles(p).Anr).Z
      zmax = zmin
      ZB = v(Triangles(p).Bnr).Z
      ZC = v(Triangles(p).Cnr).Z
      If (ZB < zmin) Then
         zmin = ZB
      ElseIf (ZB > zmax) Then
         zmax = ZB
      End If
      If (ZC < zmin) Then
         zmin = ZC
      ElseIf (ZC > zmax) Then
         zmax = ZC
      End If
      Triangles(p).Z = (zmin + zmax) / 2
   Next

End Sub

Sub DeleteList(Start() As TriaNode)

   Dim p As TriaNode
'   Do While (Start <> Null)
'     p = start;
'     start = start->next;
End Sub

Function Distance(ITria As Integer, X As Integer, Y As Integer) As Double
' we consider the line passing trought view point
' and the x,y of window. we keep the point wich line intersect
' iTria triangle. Will be returned the ze coordinate of point

  Static Dist0 As Double
  Static X0 As Integer
  Static Y0 As Integer
  Static PTria0 As Integer
   
  Dim a As Double, b As Double, C As Double
  Dim h As Double
  Dim xs As Double, ys As Double
  
'  Dist0 = 0: X0 = 0: Y0 = 0
  Dim TriaPtr As Integer, PTria As Integer
  
  TriaPtr = ITria
  
   If (PTria0 <> TriaPtr Or X <> X0 Or Y <> Y0) Then
      a = Triangles(TriaPtr).PTria.Normal.X
      b = Triangles(TriaPtr).PTria.Normal.Y
      C = Triangles(TriaPtr).PTria.Normal.Z
      h = Triangles(TriaPtr).PTria.h
      xs = XScreen(X)
      ys = YScreen(Y)
      Dist0 = h * Sqr(xs * xs + ys * ys + 1) / (a * xs + b * ys + C)
      X0 = X
      Y0 = Y
      PTria0 = TriaPtr
   End If
   
   Distance = Dist0
End Function

Sub DrawWireFrame(Pic As PictureBox)
Dim I As Integer, K As Integer, j As Integer
  
For K = 1 To UBound(FileVertex)
 I = Abs(FileVertex(K).Vert(1))
 If I > 0 Then
    Xl = TranslateR(v(I).X)
    Yl = TranslateR(v(I).Y)
    x1 = Xl
    y1 = Yl
    For j = 1 To FileVertex(K).Count
        I = FileVertex(K).Vert(j)
        If I > 0 Then
           X = TranslateR(v(I).X)
           Y = TranslateR(v(I).Y)
           Pic.Line (x1, y1)-(X, Y)
           x1 = X: y1 = Y
        End If
    Next j
 End If
          
Next

End Sub

Sub Fill_Triangle(Pic As PictureBox, I As Integer)
'  Fill the triangle i
 
 Dim Triangle(2) As CornerRec
 Dim Anr As Integer
 Dim Bnr As Integer
 Dim Cnr As Integer
 
 Anr = Triangles(I).Anr
 Bnr = Triangles(I).Bnr
 Cnr = Triangles(I).Cnr
 
 Triangle(0).X = TranslateR(v(Anr).X)
 Triangle(0).Y = TranslateR(v(Anr).Y)
 
 Triangle(1).X = TranslateR(v(Bnr).X)
 Triangle(1).Y = TranslateR(v(Bnr).Y)
 
 Triangle(2).X = TranslateR(v(Cnr).X)
 Triangle(2).Y = TranslateR(v(Cnr).Y)
 
 Shade% = Triangles(I).PTria.Color

 Call DrawTriangle(Pic, Triangle(), Shade%)
End Sub

Sub FindRange(I As Integer)
   Dim Normal As Vec3
   Normal.X = Triangles(I).PTria.Normal.X
   Normal.Y = Triangles(I).PTria.Normal.Y
   Normal.Z = Triangles(I).PTria.Normal.Z
   Dim rcolor As Single
   rcolor = DotProduct(Normal, LightVector)
   If (rcolor < rcolormin) Then rcolormin = rcolor
   If (rcolor > rcolormax) Then rcolormax = rcolor
End Sub


Function Inside_Triangle(X As Integer, Y As Integer, XA As Integer, YA As Integer, XB As Integer, YB As Integer, xC As Integer, yC As Integer) As Integer
'   (X, Y) giace out or into ABC triangle?
   
 Inside_Triangle = Orientation(XB - XA, YB - YA, X - XA, Y - YA) >= 0 And _
                  Orientation(xC - XB, yC - YB, X - XB, Y - YB) >= 0 And _
                  Orientation(XA - xC, YA - yC, X - xC, Y - yC) >= 0
End Function

Function Int_TranslateR(X As Double)
  ' Transormation of X
  Int_TranslateR = (X + hK) / K
End Function

Function IntersectOrizontal(a As Vec_Int, b As Vec_Int, Y As Integer, xxMin As Integer, xxmax As Integer) As Integer
' AB segment has some points belong with horizontal segment (Xmin, Y) - (Xmax, Y)?

Dim XA As Integer, YA As Integer
Dim XB As Integer, YB As Integer
Dim dx As Long, dy As Long, yDx As Long

   XA = a.X
   YA = a.Y
   XB = b.X
   YB = b.Y

   If (YA < Y And YB < Y Or YA > Y And YB > Y) Then
      IntersectOrizontal = 0
      Exit Function
   End If
   
   If (YA = Y And XA >= xxMin And XA <= xxmax Or _
       YB = Y And XB >= xxMin And XB <= xxmax) Then
       IntersectOrizontal = 1
      Exit Function
   End If
   
   If (YA = YB) Then
      IntersectOrizontal = YA = Y And (CLng(XA - xxmax) * (XB - xxmax) < 0 Or CLng(XA - xxMin) * (XB - xxMin) < 0)
      Exit Function
   End If
      
   If (YA > YB) Then
      Swap XA, XB
      Swap YA, YB
   End If
   
   dx = XB - XA
   dy = YB - YA
   XdY = XA * dy + (Y - YA) * dx
   
   IntersectOrizontal = XdY >= xmin * dy And XdY <= xmax * dy
End Function

Function IntersectVertical(a As Vec_Int, b As Vec_Int, X As Integer, yyMin As Integer, yymax As Integer) As Integer
' AB segment has some points belong with vertical segment (Xmin, Y) - (Xmax, Y)?

Dim XA As Integer, YA As Integer, XB As Integer, YB As Integer
Dim dx As Long, dy As Long, yDx As Long

   XA = a.X
   YA = a.Y
   XB = b.X
   YB = b.Y
   
   If (XA < X And XB < X Or XA > X And XB > X) Then
      IntersectVertical = 0
      Exit Function
   End If
 
   If (XA = X And YA >= yyMin And YA <= yymax Or _
       XB = X And YB >= yyMin And YB <= yymax) Then
      IntersectVertical = 1
      Exit Function
   End If
   
   If (XA = XB) Then
      IntersectVertical = XA = X And (CLng(YA - yymax) * (YB - yymax) < 0 Or CLng(YA - yyMin) * (YB - yyMin) < 0)
      Exit Function
   End If
       
   If (XA > XB) Then
      Swap XA, XB
      Swap YA, YB
   End If
   
   dx = XB - XA
   dy = YB - YA
   yDx = YA * dx + (X - XA) * dy
   IntersectVertical = yDx >= yyMin * dx And yDx <= yymax * dx
End Function

Sub LoadVec3(VectorPt() As Vec3, I As Integer, X As Double, Y As Double, Z As Double)
    VectorPt(I).X = X
    VectorPt(I).Y = Y
    VectorPt(I).Z = Z
End Sub

Sub Render3d(Pic As PictureBox, Come As Integer)
'The actual plotting of the object - vertices and polygons
'are previously stored in the loading of the datafile
'(See fstream.bas).
Dim vertexnr As Integer, maxnpoly As Integer, totntria As Integer
Dim code As Integer, ntr As Integer
Dim Poly() As Integer, nPoly As Integer
Dim k1 As Integer, k2 As Integer
Dim XLMax As Integer, YLMax As Integer
Dim nvertex As Integer, Pnr As Integer
Dim Qnr As Integer, Orient As Integer, testtria(3) As Integer
Dim nrs_tr() As Trianrs, t1 As Single, t2 As Single

Dim fx As Double, fy As Double
Dim rho As Double, Theta As Double, Phi As Double
Dim xs As Double, ys As Double, xe As Double, ye As Double, ze As Double
Dim xsRange As Double, ysRange As Double
Dim xsmin As Double, xsmax As Double, ysmin As Double, ysmax As Double

Dim X As Double, Y As Double, Z As Double
Dim I As Integer, K As Integer
Dim Ch As String
Dim Method As String
Dim St As String

rcolormin = BIG
rcolormax = -BIG

OrientAlgorithm = 0 ' for Orientation
   
   nvertex = MaxVertNr + 1
   ReDim Vt(nvertex)
   
   SetVista rho, Theta, Phi
   LightVector = AssignVec3(-1, 1, 0)
   SetLimitiVista xsmin, xsmax, ysmin, ysmax, nvertex, Vt()
   
'  calculate constant of screen
  xsRange = xsmax - xsmin
  ysRange = ysmax - ysmin

  xsC = 0.5 * (xsmin + xsmax)
  ysC = 0.5 * (ysmin + ysmax)
  k1 = LARGE / (X__max + 1)
  k2 = LARGE / (Y__max + 1)
  kEye = Min2(k1, k2)
  hK = kEye / 2             '  // k = 50, hk = 25 with VGA
  XLMax = kEye * (X__max + 1)
  YLMax = kEye * (Y__max + 1)
  
  ' Pixel coordinate : Xpix = TranslateR(X) and Ypix = TranslateR(Y)
  XLC = XLMax / 2
  YLC = YLMax / 2
  XLCreal = XLC + 0.5
  YLCreal = YLC + 0.5
  fx = XLMax / xsRange
  fy = YLMax / ysRange
  If fx < fy Then
     f = 0.95 * fx
  Else
     f = 0.95 * fy
  End If
  zfactor = LARGE / (zemax - zemin)
   
  ReDim v(nvertex)
  
 '   Init array of vertex:
   For I = 0 To nvertex - 1
      If (Vt(I).Z < -100000#) Then
         v(I).Used = False        ' V[i] not used
      Else
         v(I).Used = True
            
         xs = Vt(I).X / Vt(I).Z
         ys = Vt(I).Y / Vt(I).Z
         
         v(I).X = XLarge(xs)
         v(I).Y = YLarge(ys)
         v(I).Z = ZLarge(Vt(I).Z)
      End If
   Next
   
   If Come = 0 Then
      DrawWireFrame Pic
      Exit Sub
   End If
   
   Erase Vt
   
' Find max number vertices of a polygon and total number of triangles
' not rear face
maxnpoly = 0
totntria = 0

nPoly = 0
For K = 1 To UBound(FileVertex)
 nPoly = 0
 I = Abs(FileVertex(K).Vert(1))
 If I > 0 Then
    For j = 1 To FileVertex(K).Count
        I = Abs(FileVertex(K).Vert(j))
              
        If I >= nvertex Or Not v(I).Used Then
           MsgBox "Vertex nr." & CStr(I) & " not defined"
           End
        End If
        If nPoly < 3 Then testtria(nPoly) = I
    
        nPoly = nPoly + 1
    Next j
         
         If (nPoly > maxnpoly) Then maxnpoly = nPoly
         If Not (nPoly < 3) Then  '  Ignore segment 'free'
            If (SetOrient(testtria(0), testtria(1), testtria(2)) >= 0) Then totntria = totntria + nPoly - 2
         End If
         
 End If
          
Next

  ReDim Triangles(totntria)
  ReDim Poly(maxnpoly)
  ReDim nrs_tr(maxnpoly - 2)

' Read object faces and store into triangles
For K = 1 To UBound(FileVertex)
         
    nPoly = 0
    For j = 1 To FileVertex(K).Count
        
        I = Abs(FileVertex(K).Vert(j))
        If nPoly = maxnpoly Then
           MsgBox "Error  maxnpoly"
           End
        End If
        Poly(nPoly) = I
        nPoly = nPoly + 1
    
    Next j
    
   
    If (nPoly >= 3) Then
    
    Pnr = Abs(Poly(0))
    Qnr = Abs(Poly(1))
    For I = 2 To nPoly - 1
      Orient = SetOrient(Pnr, Qnr, Abs(Poly(I)))
      If (Orient <> 0) Then Exit For ' Normally, i = 2
    Next
   End If
   
    If (Orient >= 0) Then   '  Not rear face
   
   ' Subdivide a polygon into triangles
      code = Triangul(Poly(), nPoly, nrs_tr(), Orient)
      If (code > 0) Then
        If (ntr + code > totntria) Then
             MsgBox "Error: totntria"
             End
        End If
        Call Complete_Triangles(code, ntr, nrs_tr())
        ntr = ntr + code
      End If
   End If
Next K
   
Erase Poly
Erase nrs_tr

' Calculate shade
   For I = ntr - 1 To 0 Step -1
       FindRange I
   Next

   ncolors = 12
   delta = 0.999 * (ncolors - 1) / (rcolormax - rcolormin + 0.001)
   
   For I = ntr - 1 To 0 Step -1
     Call Set_Tr_Color(I)
   Next
   
     ntr_b% = ntr
     Call Q_Sort(Triangles(), 0, ntr_b%)   '  triangles[0] is the neighboring triangle
   
     For I = ntr - 1 To 0 Step -1
       Fill_Triangle Pic, I
     Next
End Sub

Function Max2(I As Integer, j As Integer) As Integer
     If I > j Then Max2 = I Else Max2 = j
End Function

Function Max3(I As Integer, j As Integer, K As Integer) As Integer
    Max3 = Max2(I, Max2(j, K))
End Function

Function Min2(I As Integer, j As Integer) As Integer
     If I < j Then Min2 = I Else Min2 = j
End Function

Function Min3(I As Integer, j As Integer, K As Integer) As Integer
        Min3 = Min2(I, Min2(j, K))
End Function

Function SetOrient(Pnr As Integer, Qnr As Integer, Rnr As Integer) As Integer
  If OrientAlgorithm = 0 Then ' Render3d
     SetOrient = Orientation(v(Qnr).X - v(Pnr).X, v(Qnr).Y - v(Pnr).Y, v(Rnr).X - v(Pnr).X, v(Rnr).Y - v(Pnr).Y)
  Else
     SetOrient = LOrienta(Pnr, Qnr, Rnr)
  End If
End Function

Function Orientation(u1 As Integer, U2 As Integer, v1 As Integer, v2 As Integer) As Long
   Dim Det As Long

   Det = CLng(u1) * v2 - CLng(U2) * v1
   If Det < -250 Then
      Det = -1
   ElseIf Det > 250 Then
      Det = 1
   End If

   Orientation = Det
End Function

Sub Q_Sort(a() As Tria, Ptr As Integer, n As Integer)
 ' Quick Sort
 ' a = Triangles()
 ' Ptr = Pointer to a()
 ' n = current element corrente to sort

    Dim I As Integer, j As Integer
    Dim X As Tria
    Dim w As Tria

   Do
      I = Ptr
      j = n - 1
      X = a(j / 2)
      Do
         Do While (a(I).Z < X.Z): I = I + 1: Loop
         Do While (a(j).Z > X.Z): j = j - 1: Loop

         If (I < j) Then
              w = a(I)
              a(I) = a(j)
              a(j) = w
          End If
          I = I + 1
          j = j - 1
      Loop While I <= j
          
      If I = j + 3 Then
         I = I - 1
         j = j + 1
      End If

      If j + 1 < n - I Then
         If j > 0 Then Q_Sort a(), 0, j + 1
        ' Ptr = Ptr + i
         n = n - I
       Else
         Pt% = I
         If I < n - 1 Then Q_Sort a(), Pt%, n - I
         n = j + 1
       End If
  
  Loop While n > 1
End Sub

Sub Set_Tr_Color(I As Integer)
   Dim Color As Integer
   Dim rcolor As Double
   Dim Normal As Vec3
   
   Normal.X = Triangles(I).PTria.Normal.X
   Normal.Y = Triangles(I).PTria.Normal.Y
   Normal.Z = Triangles(I).PTria.Normal.Z
   
   rcolor = DotProduct(Normal, LightVector)
   Color = 1 + (rcolor - rcolormin) * delta
   
   If (Color < 0) Then MsgBox ("negative color!"): Debug.Print "negative color?"
   If (Color >= 16) Then MsgBox ("color too big"): Debug.Print "color too big?"
   Triangles(I).PTria.Color = Color
End Sub

Sub SetLimitiVista(xsmin As Double, xsmax As Double, ysmin As Double, ysmax As Double, nvertex As Integer, Vt() As Vec3)
 Dim PNew As Vec3
 Dim Ve As Vec3, Vi As Vec3, Va As Vec3
 Dim I As Integer, K As Integer
   
   For I = 0 To nvertex
      Vt(I).Z = -1000000# ' Not used
   Next
   
   xsmin = BIG: ysmin = BIG: zemin = BIG
   xsmax = -BIG: ysmax = -BIG: zemax = -BIG

For K = 1 To UBound(FileCoord)
      I = FileCoord(K).I
      
      Vi.X = FileCoord(K).X
      Vi.Y = FileCoord(K).Y
      Vi.Z = FileCoord(K).Z
      
  If I > 0 Then
      If (I >= nvertex) Then
         MsgBox "too many vertices"
         End
      End If
      
      PNew.X = Vi.X - ObjPoint.X
      PNew.Y = Vi.Y - ObjPoint.Y
      PNew.Z = Vi.Z - ObjPoint.Z
      
      Call Eyecoord(PNew, Ve)
      Va.X = Ve.X
      Va.Y = Ve.Y
      Va.Z = Ve.Z

      If (Va.Z < 0) Then
         MsgBox "Point 0 of object is a vertex" & Chr(10) & " on different edges of viewpoint E." & Chr(10) & "Try with greater value of rho."
         Exit Sub
      End If

      xs = Va.X / Va.Z
      ys = Va.Y / Va.Z

      If (xs < xsmin) Then xsmin = xs
      If (xs > xsmax) Then xsmax = xs
      If (ys < ysmin) Then ysmin = ys
      If (ys > ysmax) Then ysmax = ys
      If (Va.Z < zemin) Then zemin = Va.Z
      If (Va.Z > zemax) Then zemax = Va.Z
      Vt(I) = Ve
  End If
Next K
      
If (xsmin = BIG) Then
 MsgBox "wrong input file"
 End
End If
End Sub

Sub SetVista(rho As Double, Theta As Double, Phi As Double)
   ObjPoint = AssignVec3(0.5 * (xmin + xmax), 0.5 * (ymin + ymax), 0.5 * (zmin + zmax))
   rho = xmax - xmin
   
   If (ymax - ymin > rho) Then rho = ymax - ymin
   If (zmax - zmin > rho) Then rho = zmax - zmin
   
   rho = rho * 3: Theta = 20: Phi = -65

   Call Coeff(rho, Theta * PIdiv180, Phi * PIdiv180)
End Sub

Sub Swap(X As Integer, Y As Integer)
   Dim t As Integer
   t = X
   X = Y
   Y = t
End Sub

Function TranslateR(X As Integer) As Integer
'  Translate real x to pixel x
 TranslateR = (X + hK) / kEye
End Function

Function XLarge(xs As Double) As Integer
    XLarge = Int(XLCreal + f * (xs - xsC))
End Function

Function XScreen(X As Integer) As Double
   XScreen = xsC + (X - XLC) / f
End Function

Function YLarge(ys As Double) As Integer
     YLarge = Int(YLCreal + f * (ys - ysC))
End Function

Function YScreen(Y As Integer) As Double
     YScreen = ysC + (Y - YLC) / f
End Function

Function ZEye(Z As Integer) As Double
       ZEye = Z / zfactor + zemin
End Function

Function ZLarge(ze As Single) As Integer
   ZLarge = Int((ze - zemin) * zfactor + 0.5)
End Function
