Attribute VB_Name = "LineRoutines"
'ORIGINAL CODING OF MODULE NOT BY ME.  I PIECED TOGETHER
'SEVERAL DIFFERENT ENGINE SOURCES FROM C++ AND USED THEM
'These Functions and Subs are only subroutines.

Public OrientAlgorithm As Integer ' 0 - Render3d, 1 - LinePlot (serve per SetOrient)

Public Const Nscreen = 10
Public density As Double

Public d As Double
Public c1 As Double
Public c2 As Double
Public xfactor As Double
Public yfactor As Double

Public Xrange As Double
Public Yrange As Double
Public Xvp_range As Double
Public Yvp_range As Double

Public xmin As Double
Public xmax As Double
Public ymin As Double
Public ymax As Double
Public zmin As Double
Public zmax As Double

Public deltax As Double
Public deltay As Double
Public denom As Double

Public eps1 As Double
Public trset() As Integer
Public dummy As Integer
Public vertexcount As Integer

Public x_center As Double
Public y_center As Double
Public r_max As Double
Public x_max As Double
Public y_max As Double
Public x_min As Double
Public y_min As Double

Type Vertexes
     Vt As Vec_Int
     Z As Double
     Connect(5) As Integer
End Type

Public VV() As Vertexes
Public pVertex As Integer

Type Nodo
  idx As Integer
  jtr As Integer
  nextn As Integer
End Type

Public VScreen(Nscreen, Nscreen) As Nodo

Type Point
  Pntscr As Vec_Int
  zPnt As Double
  nrPnt As Integer
End Type

Type linked_stack
    p As Point
    q As Point
    k0 As Integer
    nextn As Integer
End Type
    
Public stptr(1) As linked_stack


Sub add_linesegment(Pr As Integer, Qr As Integer)
Dim iaux As Integer
Dim p As Integer
Dim I As Integer
Dim n As Integer
Dim Pt(3) As Integer
Dim p_old(3) As Integer
Dim Pnr As Integer
Dim Qnr As Integer

 Pnr = Pr
 Qnr = Qr
  

   If (Pnr > Qnr) Then
      iaux = Pnr
      Pnr = Qnr
      Qnr = iaux
   End If

   p = VV(Pnr).Connect(0)
   If (p = 0) Then
       VV(Pnr).Connect(0) = 1
       VV(Pnr).Connect(1) = Qnr
       Exit Sub
   End If
   
   n = VV(Pnr).Connect(0)
   For I = 1 To n
      If VV(Pnr).Connect(I) = Qnr Then Exit Sub
   Next I
   
   n = n + 1
   If (n Mod 3 = 0) Then
      p_old(0) = VV(Pnr).Connect(0)
      p_old(1) = VV(Pnr).Connect(1)
      p_old(2) = VV(Pnr).Connect(2)
    
      For I = 1 To n - 1
          VV(Pnr).Connect(I) = p_old(I)
      Next
      VV(Pnr).Connect(0) = n
      VV(Pnr).Connect(n) = Qnr '  // n is a multiple of 3
                               '  // *p=n, p[1],..., p[n]
                               '  // (p[n+1], p[n+2])
   Else
      VV(Pnr).Connect(0) = n
      VV(Pnr).Connect(n) = Qnr ' // n isn't a multiple of 3 (eg n > 1)
   End If


End Sub

Function ColNr(X As Integer) As Integer
         ColNr = (CLng(X) * Nscreen) / LARGE1
End Function

Sub dealwithlinkedstack()

Dim Pt As linked_stack
Dim p As Point
Dim q As Point
Dim k0 As Integer
Dim Ptr As Integer

Ptr = 1
Do While Ptr <> 0
    Pt = stptr(Ptr)
    p = Pt.p
    q = Pt.q
    k0 = Pt.k0
    Ptr = Pt.nextn
    linesegment ThreeDObject.Pict, p, q, k0
Loop

End Sub


Sub LinePlot(Pic As PictureBox)
Dim I As Integer
Dim Pnr As Integer
Dim Qnr As Integer
Dim ii As Integer
Dim vertexnr As Integer
Dim Ptr As Integer
Dim iconnect As Integer
Dim code As Integer
Dim ntr As Integer
Dim i_i As Integer
Dim j_j As Integer
Dim jtop As Integer
Dim jbot As Integer
Dim jI As Integer
Dim trnr As Integer
Dim jtr As Integer
Dim Poly() As Integer
Dim nPoly As Integer
Dim iLeft As Integer
Dim iRight As Integer
Dim nvertex As Integer
Dim ntrset As Integer
Dim maxntrset As Integer
Dim VLOWER(Nscreen) As Integer
Dim VUPPER(Nscreen) As Integer
Dim Orient As Integer
Dim maxnpoly As Integer
Dim totntria As Integer
Dim testtria(3) As Integer

Dim xsmin As Double
Dim xsmax As Double
Dim ysmin As Double
Dim ysmax As Double

Dim nrs_tr() As Trianrs

Dim deltax As Long
Dim deltay As Long

Dim rho As Double
Dim Theta As Double
Dim Phi As Double
Dim X As Double
Dim Y As Double
Dim Z As Double
Dim xe As Double
Dim ye As Double
Dim ze As Double
Dim xx As Double
Dim yy As Double
Dim fx As Double
Dim fy As Double
Dim Xcenter As Double
Dim Ycenter As Double

Dim Ps As Vec_Int
Dim Qs As Vec_Int
Dim vLeft As Vec_Int
Dim vRight As Vec_Int

Dim p As Vec3

Dim pNode As Integer

minvertex = 32000
maxntrset = 400
OrientAlgorithm = 1 ' Per Funct. SetOrient

Erase stptr


   nvertex = MaxVertNr + 1
   ReDim Vt(nvertex)
   
   SetVista rho, Theta, Phi
   SetLimitiVista xsmin, xsmax, ysmin, ysmax, nvertex, Vt()

'INItiaize Grid
   
   x_max = 10
   density = X__max / (x_max - x_min)
   y_max = y_min + Y__max / density
   x_center = 0.5 * (x_min + x_max)
   y_center = 0.5 * (y_min + y_max)

   zfactor = LARGE / (zemax - zemin)
   eps1 = 0.001 * (zemax - zemin)
   
'   // Calculate the Constants within the Grid/Form
   Xrange = xsmax - xsmin
   Yrange = ysmax - ysmin
   
   Xvp_range = x_max - x_min
   Yvp_range = y_max - y_min
   fx = Xvp_range / Xrange
   fy = Yvp_range / Yrange
   If fx < fy Then
      d = 0.95 * fx
   Else
      d = 0.95 * fy
   End If
   
   Xcenter = 0.5 * (xsmin + xsmax)
   Ycenter = 0.5 * (ysmin + ysmax)
   c1 = x_center - d * Xcenter
   c2 = y_center - d * Ycenter
   deltax = Xrange / Nscreen
   deltay = Yrange / Nscreen
   
   xfactor = LARGE / Xrange
   yfactor = LARGE / Yrange
   
   
   
   ReDim VV(nvertex)
   
' INitialize the vertices array
   
   For I = 0 To nvertex
      If Vt(I).Z < -100000# Then
         Erase VV(I).Connect
      Else
         Erase VV(I).Connect
         VV(I).Vt.X = xIntScr(Vt(I).X / Vt(I).Z, xsmin)
         VV(I).Vt.Y = yIntScr(Vt(I).Y / Vt(I).Z, ysmin)
         VV(I).Z = Vt(I).Z
  '     DEBUGGING
  '       MsgBox "x= " & VV(i).Vt.X & "y= " & VV(i).Vt.Y & "z= " & VV(i).Z
       End If
  Next I
  
  Erase Vt

' The number for each polygon can be very big
' to limit it, change the values below
maxnpoly = 0
totntria = 0
         
nPoly = 0
For K = 1 To UBound(FileVertex)
 nPoly = 0
 I = Abs(FileVertex(K).Vert(1))
 If I > 0 Then
    For j = 1 To FileVertex(K).Count
        I = Abs(FileVertex(K).Vert(j))
              
        If I >= nvertex Then
           MsgBox "Vertice nr." & CStr(I) & " indefinito"
           End
        End If
        If nPoly < 3 Then testtria(nPoly) = I
    
        nPoly = nPoly + 1
    Next j
         
         If (nPoly > maxnpoly) Then maxnpoly = nPoly
         If Not (nPoly < 3) Then
            If (SetOrient(testtria(0), testtria(1), testtria(2)) >= 0) Then totntria = totntria + nPoly - 2
         End If
       
 End If
          
Next K
         
  ReDim Triangles(totntria)
  ReDim Poly(maxnpoly)
  ReDim nrs_tr(maxnpoly - 2)


'Triangulation
For K = 1 To UBound(FileVertex)
         
    nPoly = 0
    For j = 1 To FileVertex(K).Count
        
        I = Abs(FileVertex(K).Vert(j))
        If nPoly = maxnpoly Then
           MsgBox "Errore di programmazione maxnpoly"
           End
        End If
        Poly(nPoly) = I
        nPoly = nPoly + 1
    
    Next j
    
   
  '  If (nPoly = 1) Then
      '  MsgBox "Only one vertex for a polygon? Error #22"
      '  End
  '  End If
    
    If nPoly = 2 Then
      Call add_linesegment(Poly(0), Poly(1))
    Else
    
       Pnr = Abs(Poly(0))
       Qnr = Abs(Poly(1))
       For s = 2 To nPoly - 1
          Orient = LOrienta(Pnr, Qnr, Abs(Poly(s)))
          If (Orient <> 0) Then Exit For ' // it should be s = 2
       Next
   
       If (Orient >= 0) Then   ' ; // No problems
   
          For s = 1 To nPoly
             i_i = s Mod nPoly
             code = Poly(i_i)
             vertexnr = Abs(code)
             If code < 0 Then
                Poly(i_i) = vertexnr
              Else
                Call add_linesegment(Poly(s - 1), vertexnr)
             End If
          Next s
   
   
'   //SUbdivide the polygon in triangulation
          code = Triangul(Poly(), nPoly, nrs_tr(), Orient)
          If (code > 0) Then
             If (ntr + code > totntria) Then
                  MsgBox "Errore di programmazione: totntria"
                  End
             End If
             Call LComplete_Triangles(code, ntr, nrs_tr())
             ntr = ntr + code
          End If
      
       End If
    End If
    
Next K
   
Erase Poly
Erase nrs_tr

 Call setupscreenlist(Triangles, ntr)
 ReDim trset(maxntrset)
   
  For Pnr = MinVertNr To MaxVertNr
      Ptr = VV(Pnr).Connect(0)
      
      If Ptr > 0 Then
         Ps = VV(Pnr).Vt
         For iconnect = 1 To Ptr
             Qnr = VV(Pnr).Connect(iconnect)
             Qs = VV(Qnr).Vt
                        
             If (Ps.X < Qs.X Or (Ps.X = Qs.X And Ps.Y < Qs.Y)) Then
                vLeft = Ps: vRight = Qs
             Else
                vLeft = Qs: vRight = Ps
             End If
             iLeft = ColNr(vLeft.X)
             iRight = ColNr(vRight.X)
         
             If (iLeft <> iRight) Then
                 deltay = vRight.Y - vLeft.Y
                 deltax = vRight.X - vLeft.X
             End If
         
             jbot = RowNr(vLeft.Y)
             jtop = jbot
             
             For ii = iLeft To iRight
                If ii = iRight Then
                   jI = RowNr(vRight.Y)
                Else
                   hh& = vLeft.Y + (xCoord(ii + 1) - vLeft.X) * deltay / deltax
                   If hh& > 32000 Then
                      jI = Nscreen
                   Else
                      jI = RowNr(CInt(hh&))
                   End If
                End If
             
                VLOWER(ii) = Min2(jbot, jI)
                jbot = jI
                VUPPER(ii) = Max2(jtop, jI)
                jtop = jI
             
             Next ii
         
         Next iconnect
         
         ntrset = 0
         For I = iLeft To iRight
             For j = VLOWER(I) To VUPPER(I)
                 pNode = VScreen(I, j).idx
              '   Do While pNode <> 0
                    trnr = VScreen(I, j).jtr
                     trset(ntrset) = trnr
                     jtr = 0
                     Do While trset(jtr) <> trnr: jtr = jtr + 1: Loop
                     If (jtr = ntrset) Then
                         ntrset = ntrset + 1
                         If (ntrset = maxntrset) Then
                          '  P = jtr
                            maxntrset = maxntrset + 200
                            ReDim Preserve trset(maxntrset)
                     
                            For s = 0 To ntrset - 1
                                trset(s) = trset(jtr + s)
                            Next s
                         End If
                     End If
                    
                     pNode = VScreen(I, j).nextn
                    
               '  Loop ' Pnode <> 0
             
             Next j
         Call linesegment(Pic, SetPoint(Ps, VV(Pnr).Z, Pnr), SetPoint(Qs, VV(Qnr).Z, Qnr), ntrset)
         
         dealwithlinkedstack
         Next I
      
      End If ' Ptr = 0
      
   Next Pnr

End Sub

Sub LComplete_Triangles(n As Integer, offset As Integer, nrs_tr() As Trianrs)

' Complete triangles[offset],..., triangles[offset+n-1].
' # of Vertices: nrs_tr[0],..., nrs_tr[n-1].
' For triangulation, the equation for polygons (basic one)
' is -  nx . x + ny . y + nz . z = h.
  Dim I As Integer
  Dim Anr As Integer
  Dim Bnr As Integer
  Dim Cnr As Integer
  Dim ZA As Integer
  Dim ZB As Integer
  Dim ZC As Integer
 ' Dim zmin As Single
 ' Dim zmax As Single

  Dim nx As Double
  Dim ny As Double
  Dim nz As Double
  Dim ux As Double
  Dim uy As Double
  Dim uz As Double
  Dim vx As Double
  Dim vy As Double
  Dim vz As Double
  Dim factor As Double
  Dim h As Double
  Dim Ax As Double
  Dim Ay As Double
  Dim Az As Double
  Dim Bx As Double
  Dim By As Double
  Dim Bz As Double
  Dim Cx As Double
  Dim Cy As Double
  Dim Cz As Double
  
  Dim p As Integer
  Dim q As TriaData
  
  For I = n \ 2 To n
     Anr = nrs_tr(I).a
     Bnr = nrs_tr(I).b
     Cnr = nrs_tr(I).C
     If (SetOrient(Anr, Bnr, Cnr) > 0) Then Exit For
   Next

   ZA = VV(Anr).Z
   ZB = VV(Bnr).Z
   ZC = VV(Cnr).Z

   Az = zFloat(ZA)
   Bz = zFloat(ZB)
   Cz = zFloat(ZC)
   Ax = xFloat(VV(Anr).Vt.X) * Az
   Ay = yFloat(VV(Anr).Vt.Y) * Az
   Bx = xFloat(VV(Bnr).Vt.X) * Bz
   By = yFloat(VV(Bnr).Vt.Y) * Bz
   Cx = xFloat(VV(Cnr).Vt.X) * Cz
   Cy = yFloat(VV(Cnr).Vt.Y) * Cz
   ux = Bx - Ax
   uy = By - Ay
   uz = Bz - Az
   vx = Cx - Ax
   vy = Cy - Ay
   vz = Cz - Az
   nx = uy * vz - uz * vy
   ny = uz * vx - ux * vz
   nz = ux * vy - uy * vx
   h = nx * Ax + ny * Ay + nz * Az
   factor = 1 / Sqr(nx * nx + ny * ny + nz * nz)
   q.Normal.X = nx * factor
   q.Normal.Y = ny * factor
   q.Normal.Z = nz * factor
   q.h = h * factor
   For I = 0 To n - 1
      p = offset + I
      Triangles(p).Anr = nrs_tr(I).a
      Triangles(p).Bnr = nrs_tr(I).b
      Triangles(p).Cnr = nrs_tr(I).C
      Triangles(p).PTria = q
   Next

End Sub


Sub linesegment(Pic As PictureBox, p As Point, q As Point, k0 As Integer)
   Dim Ps As Vec_Int
   Dim Qs As Vec_Int
   Dim Ass As Vec_Int
   Dim Bs As Vec_Int
   Dim Cs As Vec_Int
   Dim Temp As Vec_Int
   Dim Iss As Vec_Int
   Dim Js As Vec_Int
   
   
   Dim x1 As Single
   Dim x2 As Single
   Dim y1 As Single
   Dim y2 As Single
   Dim xP As Double
   Dim yP As Double
   Dim xQ As Double
   Dim yQ As Double
   Dim zP As Double
   Dim zQ As Double
   Dim xI As Double
   Dim yI As Double
   Dim hP As Double
   Dim hQ As Double
   Dim xJ As Double
   Dim yJ As Double
   Dim lam_min As Double
   Dim lam_max As Double
   Dim lambda As Double
   Dim mu As Double
   Dim hh As Double
   Dim h1 As Double
   Dim h2 As Double
   Dim zI As Double
   Dim zJ As Double
   Dim ZA As Double
   Dim ZB As Double
   Dim ZC As Double
   Dim zmaxPQ As Double
   
   
   Dim Pnr As Integer
   Dim Qnr As Integer
   Dim kk As Integer
   Dim j As Integer
   Dim Anr As Integer
   Dim Bnr As Integer
   Dim Cnr As Integer
   Dim I As Integer
   Dim Poutside As Integer
   Dim Qoutside As Integer
   Dim Pnear As Integer
   Dim Qnear As Integer
   
   Dim APB As Integer
   Dim AQB As Integer
   Dim BPC As Integer
   Dim BQC As Integer
   Dim CPA As Integer
   Dim CQA As Integer
   Dim xminPQ As Integer
   Dim xmaxPQ As Integer
   Dim yminPQ As Integer
   Dim ymaxPQ As Integer
   Dim X_P As Integer
   Dim Y_P As Integer
   Dim X_Q As Integer
   Dim Y_Q As Integer
   Dim u1 As Integer
   Dim U2 As Integer
   
   Dim denom As Long
   Dim v1 As Long
   Dim v2 As Long
   Dim w1 As Long
   Dim w2 As Long
   
   Dim Normal As Vec3
   
   Ps = p.Pntscr
   Qs = q.Pntscr
   
   zP = p.zPnt
   zQ = q.zPnt
   
   Pnr = p.nrPnt
   Qnr = q.nrPnt
   
   X_P = Ps.X
   Y_P = Ps.Y
   X_Q = Qs.X
   Y_Q = Qs.Y
   u1 = X_Q - X_P
   U2 = Y_Q - Y_P
   kk = k0
   
   
   If (X_P < X_Q) Then
      xminPQ = X_P
      xmaxPQ = X_Q
   Else
      xminPQ = X_Q
      xmaxPQ = X_P
   End If
   If (Y_P < Y_Q) Then
      yminPQ = Y_P
      ymaxPQ = Y_Q
   Else
      yminPQ = Y_Q
      ymaxPQ = Y_P
   End If
   
'8 tests in 2d and 3d rendering for the line plotting

   Do While (kk > 0)
      kk = kk - 1
      j = trset(kk)
      Anr = Triangles(j).Anr
      Bnr = Triangles(j).Bnr
      Cnr = Triangles(j).Cnr

      If ((Pnr = Anr Or Pnr = Bnr Or Pnr = Cnr) And _
          (Qnr = Anr Or Qnr = Bnr Or Qnr = Cnr)) Then GoTo Continua
      
      Ass = VV(Anr).Vt
      Bs = VV(Bnr).Vt
      Cs = VV(Cnr).Vt

      If (xmaxPQ <= Ass.X And xmaxPQ <= Bs.X And xmaxPQ <= Cs.X Or _
          xminPQ >= Ass.X And xminPQ >= Bs.X And xminPQ >= Cs.X Or _
          ymaxPQ <= Ass.Y And ymaxPQ <= Bs.Y And ymaxPQ <= Cs.Y Or _
          yminPQ >= Ass.Y And yminPQ >= Bs.Y And yminPQ >= Cs.Y) Then GoTo Continua
             '  continue;
    
      APB = orientv(Ass, Ps, Bs)
      AQB = orientv(Ass, Qs, Bs)
      If (APB + AQB > 0) Then GoTo Continua
      BPC = orientv(Bs, Ps, Cs)
      BQC = orientv(Bs, Qs, Cs)
      If (BPC + BQC > 0) Then GoTo Continua
      CPA = orientv(Cs, Ps, Ass)
      CQA = orientv(Cs, Qs, Ass)
      If (CPA + CQA > 0) Then GoTo Continua

      If (Abs(orientv(Ps, Qs, Ass) + orientv(Ps, Qs, Bs) + orientv(Ps, Qs, Cs)) > 1) Then GoTo Continua

      ZA = VV(Anr).Z
      ZB = VV(Bnr).Z
      ZC = VV(Cnr).Z
      If zP > zQ Then zmaxPQ = zP Else zmaxPQ = zQ
      
      If (zmaxPQ <= ZA And zmaxPQ <= ZB And zmaxPQ <= ZC) Then GoTo Continua
      Normal = Triangles(j).PTria.Normal
      hh = Triangles(j).PTria.h
      If (hh = 0) Then GoTo Continua
      
      xP = zP * xFloat(X_P)
      yP = zP * yFloat(Y_P)
      xQ = zQ * xFloat(X_Q)
      yQ = zQ * yFloat(Y_Q)
      hP = Normal.X * xP + Normal.Y * yP + Normal.Z * zP
      hQ = Normal.X * xQ + Normal.Y * yQ + Normal.Z * zQ
      h2 = hh + eps1
      If (hP <= h2 And hQ <= h2) Then GoTo Continua

      Poutside = APB = 1 Or BPC = 1 Or CPA = 1
      Qoutside = AQB = 1 Or BQC = 1 Or CQA = 1
      If (Not Poutside And Not Qoutside) Then Exit Sub
      
      h1 = hh - eps1
      Pnear = hP < h1
      Qnear = hQ < h1
      If (Pnear And Not Poutside Or Qnear And Not Qoutside) Then GoTo Continua

      lam_min = 1#
      lam_max = 0#
      For I = 0 To 2
      
         v1 = Bs.X - Ass.X
         v2 = Bs.Y - Ass.Y
         w1 = Ass.X - xP
         w2 = Ass.Y - yP
         denom = u1 * v2 - U2 * v1
         If (denom <> 0) Then
            mu = (U2 * w1 - u1 * w2) / CDbl(denom)
            If (mu > -0.0001 And mu < 1.0001) Then
               lambda = (v2 * w1 - v1 * w2) / CDbl(denom)
               If (lambda > -0.0001 And lambda < 1.0001) Then
                  If (Poutside <> Qoutside And _
                  lambda > 0.0001 And lambda < 0.9999) Then
                     lam_min = lam_max = lambda
                     Exit For
                  End If
                  If (lambda < lam_min) Then lam_min = lambda
                  If (lambda > lam_max) Then lam_max = lambda
               End If
            End If
         End If ' Denom <> 0
         Temp = Ass
         Ass = Bs
         Bs = Cs
         Cs = Temp
      Next I
      
      If (Poutside And lam_min > 0.01) Then
         Iss.X = Int(xP + lam_min * u1 + 0.5)
         Iss.Y = Int(yP + lam_min * U2 + 0.5)
         zI = 1 / (lam_min / zQ + (1 - lam_min) / zP)
         xI = zI * xFloat(Iss.X)
         yI = zI * yFloat(Iss.Y)
         If (Normal.X * xI + Normal.Y * yI + Normal.Z * zI) < h1 Then GoTo Continua
         Call stack_linesegment(SetPoint(Ps, zP, Pnr), SetPoint(Iss, zI, -1), kk)
      End If
      If (Qoutside And lam_max < 0.99) Then
         Js.X = Int(xP + lam_max * u1 + 0.5)
         Js.Y = Int(yP + lam_max * U2 + 0.5)
         zJ = 1 / (lam_max / zQ + (1 - lam_max) / zP)
         xJ = zJ * xFloat(Js.X)
         yJ = zJ * yFloat(Js.Y)
         If (Normal.X * xJ + Normal.Y * yJ + Normal.Z * zJ) < h1 Then GoTo Continua
         
         Call stack_linesegment(SetPoint(Qs, zQ, Qnr), SetPoint(Js, zJ, -1), kk)
      End If
      
      Exit Sub
      
Continua:

   Loop ' While (kk > 0)
   
   x1 = SetX(d * xFloat(X_P) + c1)
   y1 = SetY(d * yFloat(Y_P) + c2)
   x2 = SetX(d * xFloat(X_Q) + c1)
   y2 = SetY(d * yFloat(Y_Q) + c2)
   
   Pic.Line (x1, y1)-(x2, y2)
   
'   Ms = " x1= " & x1 & Chr(10)
'   Ms = Ms & " y1= " & y1 & Chr(10)
'   Ms = Ms & " x2= " & x2 & Chr(10)
'   Ms = Ms & " y2= " & y2
'   MsgBox Ms
   
  ' move(d * xfloat(XP) + c1, d * yfloat(YP) + c2);
  ' draw(d * xfloat(XQ) + c1, d * yfloat(YQ) + c2);
End Sub


Function LOrienta(Pnr As Integer, Qnr As Integer, Rnr As Integer) As Integer
 Dim Ps As Vec_Int
 Dim Qs As Vec_Int
 Dim Rs As Vec_Int
 
 Dim u1 As Integer
 Dim U2 As Integer
 Dim v1 As Integer
 Dim v2 As Integer
 
 Dim Det As Long
 
 Ps = VV(Pnr).Vt
 Qs = VV(Qnr).Vt
 Rs = VV(Rnr).Vt
 
 u1 = Qs.X - Ps.X
 U2 = Qs.Y - Ps.Y
 v1 = Rs.X - Ps.X
 v2 = Rs.Y - Ps.Y
 
 Det = CLng(u1) * v2 - CLng(U2) * v1
 
 If Det < -300 Then
    LOrienta = -1
 Else
    LOrienta = Abs(Det > 300)
 End If
 
End Function

Function orientv(Ps As Vec_Int, Qs As Vec_Int, Rs As Vec_Int) As Integer
 Dim u1 As Integer
 Dim U2 As Integer
 Dim v1 As Integer
 Dim v2 As Integer

 Dim Det As Long
  
 u1 = Qs.X - Ps.X
 U2 = Qs.Y - Ps.Y
 v1 = Rs.X - Ps.X
 v2 = Rs.Y - Ps.Y

 Det = CLng(u1) * v2 - CLng(U2) * v1
   
 If Det < -10 Then
    orientv = -1
 Else
    orientv = Det > 10
 End If

End Function

Function RowNr(Y As Integer) As Integer
         RowNr = (CLng(Y) * Nscreen) / LARGE1
End Function


Function SetPoint(p As Vec_Int, Z As Double, nr As Integer) As Point
   SetPoint.Pntscr = p
   SetPoint.zPnt = Z
   SetPoint.nrPnt = nr
End Function

Sub setupscreenlist(Tr() As Tria, n As Integer)
 Dim I As Integer
 Dim l As Integer
 Dim j As Integer
 Dim iMin As Integer
 Dim iMax As Integer
 Dim j_old As Integer
 Dim jI As Integer
 Dim topcode(2) As Integer
 Dim iLeft As Integer
 Dim iRight As Integer
 Dim LLOWER(Nscreen) As Integer
 Dim LUPPER(Nscreen) As Integer
 
 Dim deltax As Long
 Dim deltay As Long

 Dim Ass As Vec_Int
 Dim Bs As Vec_Int
 Dim Cs As Vec_Int
 Dim vLeft(2) As Vec_Int
 Dim vRight(2) As Vec_Int
 Dim Aux As Vec_Int

 Dim p As Integer
 Dim p_New As Integer
 Dim p_old As Integer
   
 For I = 0 To n - 1
     p = I
     Ass = VV(Tr(p).Anr).Vt
     Bs = VV(Tr(p).Bnr).Vt
     Cs = VV(Tr(p).Cnr).Vt
               
     topcode(0) = Ass.X > Bs.X
     topcode(1) = Cs.X > Ass.X
     topcode(2) = Bs.X > Cs.X
     vLeft(0) = Ass
     vRight(0) = Bs
     vLeft(1) = Ass
     vRight(1) = Cs
     vLeft(2) = Bs
     vRight(2) = Cs
     For l = 0 To 2
        If (vLeft(l).X > vRight(l).X Or _
           (vLeft(l).X = vRight(l).X And vLeft(l).Y > vRight(l).Y)) Then
             Aux = vLeft(l)
             vLeft(l) = vRight(l)
             vRight(l) = Aux
        End If
     Next l
        
     iMin = ColNr(Min3(Ass.X, Bs.X, Cs.X))
     iMax = ColNr(Max3(Ass.X, Bs.X, Cs.X))
     'iMin = Max2(iMin, 0)
      
     For ii = iMin To iMax
         LLOWER(ii) = 32000
         LUPPER(ii) = -32000
     Next ii
       
     For l = 0 To 2
         iLeft = ColNr(vLeft(l).X)
         iRight = ColNr(vRight(l).X)
         If (iLeft <> iRight) Then
           deltay = vRight(l).Y - vLeft(l).Y
           deltax = vRight(l).X - vLeft(l).X
         End If
         j_old = RowNr(vLeft(l).Y)
         For ii = iLeft To iRight
             If ii = iRight Then
                jI = RowNr(vRight(l).Y)
             Else
                g& = vLeft(l).Y + CLng(xCoord(ii + 1) - vLeft(l).X * deltay / deltax)
                If g& > 32000 Then g& = 32000
                If g& < 0 Then g& = 0
                jI = RowNr(CInt(g&))
             End If
            If topcode(l) Then
               LUPPER(ii) = Max3(j_old, jI, LUPPER(ii))
            Else
               LLOWER(ii) = Min3(j_old, jI, LLOWER(ii))
            End If
            j_old = jI
         Next ii
      Next l
      
       For ii = iMin To iMax
          For j = LLOWER(ii) To LUPPER(ii)
             p_New = p_New + 1
             p_old = VScreen(ii, j).idx
  
             VScreen(ii, j).idx = p_New
             VScreen(ii, j).jtr = I
             VScreen(ii, j).nextn = p_old
          Next j
       Next ii
 Next I

'{  int i, l, I, J, Imin, Imax, j_old, jI, topcode[3],
'      ileft, iright, LOWER[Nscreen], UPPER[Nscreen];
'   long deltax, deltay;
'   vec_int As, Bs, Cs, Left[3], Right[3], Aux;
'   tria huge*p;
'   node huge*p_new, huge*p_old;
'   for (i=0; i<n; i++)
'   {  p = TR + i;
'      As = V[p->Anr].VT; Bs = V[p->Bnr].VT; Cs = V[p->Cnr].VT;
'      topcode[0] = As.X > Bs.X; // Per l'orientamento positivo
'      topcode[1] = Cs.X > As.X;
'      topcode[2] = Bs.X > Cs.X;

'      Left[0] = As; Right[0] = Bs;
'      Left[1] = As; Right[1] = Cs;
'      Left[2] = Bs; Right[2] = Cs;
'      for (l=0; l<3; l++)  // l = numero di lati del triangolo
'         if (Left[l].X > Right[l].X ||
'         (Left[l].X == Right[l].X && Left[l].Y > Right[l].Y))
'         {  Aux = Left[l]; Left[l] = Right[l]; Right[l] = Aux;
'         }
'      Imin = colnr(min3(As.X, Bs.X, Cs.X));
'      Imax = colnr(max3(As.X, Bs.X, Cs.X));
'      for (I = Imin; I<=Imax; I++)
'      {  LOWER[I] = INT_MAX; UPPER[I] = INT_MIN;
'      }
'      for (l=0; l<3; l++)
'      {  ileft = colnr(Left[l].X); iright = colnr(Right[l].X);
'         if (ileft != iright)
'         { deltay = Right[l].Y - Left[l].Y;
'           deltax = Right[l].X - Left[l].X;
'         }
'         j_old = rownr(Left[l].Y);
'         for (I=ileft; I<=iright; I++)
'         {  jI = (I == iright ? rownr(Right[l].Y) : rownr(Left[l].Y
'                  + (Xcoord(I+1) - Left[l].X) * deltay / deltax));
'            if (topcode[l])
'                UPPER[I] = max3(j_old, jI, UPPER[I]);
'            else LOWER[I] = min3(j_old, jI, LOWER[I]);
'            j_old = jI;
'         }
'      }
'      // Per la colonna I del video, il triangolo Š associato solo
'      // con i rettangoli delle righe LOWER[I],...,UPPER[I].
'      for (I=Imin; I<=Imax; I++)
'      for (J=LOWER[I]; J<=UPPER[I]; J++)
'      {  p_old = SCREEN[I][J];
'         SCREEN[I][J] = p_new = AllocMem1(node);
'         if (p_new == NULL) memproblem('G');
'         p_new->jtr = i; p_new->next = p_old;
'      }
'   }
'}

End Sub


Function SetX(X As Double) As Integer
    X = Int(density * (X - x_min))
    If (X < 0) Then
       X = 0
       outside = 1
    End If
    If (X > X__max) Then
      X = X__max
      outside = 1
    End If
   SetX = X
End Function

Function SetY(Y As Double) As Integer

   Y = Y__max - Int(density * (Y - y_min))
   If (Y < 0) Then
      Y = 0
      outside = 1
   End If
   
   If (Y > Y__max) Then
      Y = Y__max
      outside = 1
   End If
   
   SetY = Y

End Function


Sub stack_linesegment(p As Point, q As Point, k0 As Integer)

  Dim Pt As Integer
  Dim xP As Integer
  Dim yP As Integer
  Dim xQ As Integer
  Dim yQ As Integer
  xP = p.Pntscr.X
  yP = p.Pntscr.Y
  xQ = q.Pntscr.X
  yQ = q.Pntscr.Y
  If (Abs(xP - xQ) + Abs(yP - yQ) < 50) Then Exit Sub
  
 ' Pt = UBound(stptr) + 1
 ' ReDim Preserve stptr(Pt)
  
  stptr(Pt).p = p
  stptr(Pt).q = q
  stptr(Pt).k0 = k0
  
End Sub


Function xCoord(nr As Integer) As Integer
  xCoord = (CLng(nr) * LARGE) / Nscreen
End Function

Function xFloat(X As Integer) As Double
      xFloat = (X / xfactor) + xmin
End Function

Function zFloat(Z As Integer) As Double
      zFloat = Z / zfactor + zmin
End Function

Function xIntScr(X As Double, xxMin As Double) As Integer
   xIntScr = (X - xxMin) * xfactor + 1 '0.5
End Function

Function yFloat(Y As Integer) As Double
   yFloat = (Y / yfactor) + ymin
End Function

Function yIntScr(Y As Double, yyMin As Double) As Integer
   yIntScr = (Y - yyMin) * yfactor + 1 '  0.5
End Function
