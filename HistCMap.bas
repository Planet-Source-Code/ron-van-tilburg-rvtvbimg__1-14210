Attribute VB_Name = "HistCMap"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'HistCMap.bas   inverse colormapping by color histogram
'* cmatch.c  colour matching to a given color map
'* C code copyright ©1995 Ron van Tilburg, VB code copyright ©2000 Ron van Tilburg
'*
'* Heavily modified by Ron van Tilburg
'*
'* Original Copyright (C) 1991-1995, Thomas G. Lane.
'* This file is derived from quant2.c part of the Independent JPEG Group's
'* software, with additions/alterations by Ron van Tilburg.
'* For conditions of distribution and use, see the accompanying README file.
'*
'* This file contains color mapping routines.
'*
'* The matching pass over the image maps each input pixel to the closest output
'* color (optionally after applying a dithering correction).
'* This mapping is logically trivial, but making it go fast enough requires
'* considerable care.
'*
'* To improve the visual quality of the results, we actually work in scaled
'* RGB space, giving G distances more weight than R, and R in turn more than
'* B.  To do everything in integer math, we must use integer scale factors.
'* The 2/3/1 scale factors used here correspond loosely to the relative
'* weights of the colors in the NTSC grayscale equation.
'* If you want to use this code to quantize a non-RGB color space, you'll
'* probably need to change these scale factors.

Const R_SCALE As Integer = 1             '*2 scale R distances by this much */
Const G_SCALE As Integer = 1             '*3 scale G distances by this much */
Const B_SCALE As Integer = 1             '*1 scale B distances by this much */

'*=========================================================================
'* First we have the histogram data structure and routines for creating it.
'*
'* The number of bits of precision can be adjusted by changing these symbols.
'* We recommend keeping 6 bits for G and 5 each for R and B.
'* If you have plenty of memory and cycles, 6 bits all around gives marginally
'* better results; if you are short of memory, 5 bits all around will save
'* some space but degrade the results.
'* The histogram space is used for pixel mapping data;
'* in that capacity, each cell must be able to store zero to the number of
'* desired colors.  16 bits/cell is plenty for that too.
'* the histogram is allocated as a single array representation of a 3d array
'* arranged as strips of blue, inside a 2d array of red and green of 2^5*2^6*2^5=65536 entries

Const MAXNUMCOLORS As Integer = 256       '* maximum size of colormap */

Const HIST_R_BITS As Integer = 5          '* bits of precision in R histogram */
Const HIST_G_BITS As Integer = 6          '* bits of precision in G histogram */
Const HIST_B_BITS As Integer = 5          '* bits of precision in B histogram */

'* Number of elements along histogram axes. */

Const HIST_R_ELEMS As Integer = (2 ^ HIST_R_BITS) ' (1<<HIST_R_BITS)
Const HIST_G_ELEMS As Integer = (2 ^ HIST_G_BITS) ' (1<<HIST_G_BITS)
Const HIST_B_ELEMS As Integer = (2 ^ HIST_B_BITS) ' (1<<HIST_B_BITS)

'* These are the amounts to shift an input value to get a histogram index. */

Const R_SHIFT As Integer = (8 - HIST_R_BITS)
Const G_SHIFT As Integer = (8 - HIST_G_BITS)
Const B_SHIFT As Integer = (8 - HIST_B_BITS)

Const R_SHIFT_VAL As Integer = (2 ^ R_SHIFT)
Const G_SHIFT_VAL As Integer = (2 ^ G_SHIFT)
Const B_SHIFT_VAL As Integer = (2 ^ B_SHIFT)

'*===== COLOR MATCHING ===================================================*/
'*
' * These routines are concerned with the time-critical task of mapping input
' * colors to the nearest color in the selected colormap.
' *
' * We use the histogram space as an "inverse color map", essentially a
' * cache for the results of nearest-color searches.  All colors within a
' * histogram cell will be mapped to the same colormap entry, namely the one
' * closest to the cell's center.  This may not be quite the closest entry to
' * the actual input color, but it's almost as good.  A zero in the cache
' * indicates we haven't found the nearest color for that cell yet; the array
' * is cleared to zeroes before starting the mapping pass.  When we find the
' * nearest color for a cell, its colormap index plus one is recorded in the
' * cache for future use.  The matching routines call fill_inverse_cmap
' * when they need to use an unfilled entry in the cache.
' *
' * Our method of efficiently finding nearest colors is based on the "locally
' * sorted search" idea described by Heckbert and on the incremental distance
' * calculation described by Spencer W. Thomas in chapter III.1 of Graphics
' * Gems II (James Arvo, ed.  Academic Press, 1991).  Thomas points out that
' * the distances from a given colormap entry to each cell of the histogram can
' * be computed quickly using an incremental method: the differences between
' * distances to adjacent cells themselves differ by a constant.  This allows a
' * fairly fast implementation of the "brute force" approach of computing the
' * distance from every colormap entry to every histogram cell.  Unfortunately,
' * it needs a work array to hold the best-distance-so-far for each histogram
' * cell (because the inner loop has to be over cells, not colormap entries).
' * The work array elements have to be longs, so the work array would need
' * 256Kb at our recommended precision.  This is not feasible in DOS machines.
' *
' * To get around these problems, we apply Thomas' method to compute the
' * nearest colors for only the cells within a small subbox of the histogram.
' * The work array need be only as big as the subbox, so the memory usage
' * problem is solved.  Furthermore, we need not fill subboxes that are never
' * referenced while matching; many images use only part of the color gamut,
' * so a fair amount of work is saved.  An additional advantage of this
' * approach is that we can apply Heckbert's locality criterion to quickly
' * eliminate colormap entries that are far away from the subbox; typically
' * three-fourths of the colormap entries are rejected by Heckbert's criterion,
' * and we need not compute their distances to individual cells in the subbox.
' * The speed of this approach is heavily influenced by the subbox size: too
' * small means too much overhead, too big loses because Heckbert's criterion
' * can't eliminate as many colormap entries.  Empirically the best subbox
' * size seems to be about 1/512th of the histogram (1/8th in each direction).
' */

'* log2(histogram cells in update box) for each axis; this can be adjusted */

Const BOX_R_LOG As Integer = (HIST_R_BITS - 3)
Const BOX_G_LOG As Integer = (HIST_G_BITS - 3)
Const BOX_B_LOG As Integer = (HIST_B_BITS - 3)

Const BOX_R_LOG_VAL As Integer = (2 ^ BOX_R_LOG)    '(1<<BOX_R_LOG)     '* - of hist cells in update box */
Const BOX_G_LOG_VAL As Integer = (2 ^ BOX_G_LOG)    '(1<<BOX_G_LOG)
Const BOX_B_LOG_VAL As Integer = (2 ^ BOX_B_LOG)    '(1<<BOX_B_LOG)

Const BOX_R_ELEMS As Integer = (2 ^ BOX_R_LOG)    '(1<<BOX_R_LOG)     '* - of hist cells in update box */
Const BOX_G_ELEMS As Integer = (2 ^ BOX_G_LOG)    '(1<<BOX_G_LOG)
Const BOX_B_ELEMS As Integer = (2 ^ BOX_B_LOG)    '(1<<BOX_B_LOG)

Const BOX_R_SHIFT As Integer = (R_SHIFT + BOX_R_LOG)
Const BOX_G_SHIFT As Integer = (G_SHIFT + BOX_G_LOG)
Const BOX_B_SHIFT As Integer = (B_SHIFT + BOX_B_LOG)

Const BOX_R_SHIFT_VAL As Integer = (2 ^ BOX_R_SHIFT)
Const BOX_G_SHIFT_VAL As Integer = (2 ^ BOX_G_SHIFT)
Const BOX_B_SHIFT_VAL As Integer = (2 ^ BOX_B_SHIFT)


'* the globals in this routine
Dim Histogram() As Integer
Dim MatchMap() As Byte
Dim NMatchColors As Integer

Public Sub InitColorMappingHistogram(ByRef CMap() As Byte, ByVal NColors As Long)
  Dim i As Long, p As Long
  On Error GoTo 0
  ReDim Histogram(0 To CLng(HIST_R_ELEMS) * CLng(HIST_G_ELEMS) * CLng(HIST_B_ELEMS) - 1) As Integer
  ReDim MatchMap(LBound(CMap) To UBound(CMap))
  MatchMap() = CMap()
  NMatchColors = NColors

  p = 0                                       'the colours of the colormap will be inserted to start with
  For i = 0 To NMatchColors - 1               'GIF files suit line drawings so the colors will match well
    Call MatchColorbyHistogram(MatchMap(p), MatchMap(p + 1), MatchMap(p + 2))
    p = p + 3
  Next
End Sub

Public Sub FreeColorMappingHistogram()
  Erase Histogram
  Erase MatchMap
End Sub

' * The next three routines implement inverse colormap filling.  They could
' * all be folded into one big routine, but splitting them up this way saves
' * some stack space (the mindist[] and bestdist[] arrays need not coexist)
' * and may allow some compilers to produce better code by registerizing more
' * inner-loop variables.
' */

Private Function find_nearby_colors(ByVal minr As Integer, _
                                    ByVal ming As Integer, _
                                    ByVal minb As Integer, _
                                    ByRef colorlist() As Byte) As Integer

  Dim i As Integer, p As Integer, x As Integer, NColors As Integer
  Dim maxr As Integer, maxg As Integer, maxb As Integer
  Dim midr As Integer, midg As Integer, midb As Integer
  Dim minmaxdist As Long, min_dist As Long, max_dist As Long, tdist As Long
  Dim mindist() As Long         '* min distance to colormap entry i */

'   * Locate the colormap entries close enough to an update box to be candidates
'   * for the nearest entry to some cell(s) in the update box.  The update box
'   * is specified by the center coordinates of its first cell.  The number of
'   * candidate colormap entries is returned, and their colormap indexes are
'   * placed in colorlist[].
'   * This routine uses Heckbert's "locally sorted search" criterion to select
'   * the colors that need further consideration.
'   * Compute true coordinates of update box's upper corner and center.
'   * Actually we compute the coordinates of the center of the upper-corner
'   * histogram cell, which are the upper bounds of the volume we care about.
'   * Note that since ">>" rounds down, the "center" values may be closer to
'   * min than to max; hence comparisons to them must be "<=", not "<".
'   */
  
  ReDim mindist(0 To MAXNUMCOLORS - 1) As Long
  
  maxr = minr + (BOX_R_SHIFT_VAL - R_SHIFT_VAL)   '((1 << BOX_R_SHIFT) - (1 << R_SHIFT))
  maxg = ming + (BOX_G_SHIFT_VAL - G_SHIFT_VAL)   '((1 << BOX_G_SHIFT) - (1 << G_SHIFT))
  maxb = minb + (BOX_B_SHIFT_VAL - B_SHIFT_VAL)   '((1 << BOX_B_SHIFT) - (1 << B_SHIFT))
  
  midr = (minr + maxr) \ 2
  midg = (ming + maxg) \ 2
  midb = (minb + maxb) \ 2

  '* For each color in colormap, find:
  '*  1. its minimum squared-distance to any point in the update box
  '*     (zero if color is within update box);
  '*  2. its maximum squared-distance to any point in the update box.
  '* Both of these can be found by considering only the corners of the box.
  '* We save the minimum distance for each color in mindist[];
  '* only the smallest maximum distance is of interest.
  
  minmaxdist = &H7FFFFFFF     'a very large number indeed
  p = 0
  For i = 0 To NMatchColors - 1
  
    '* We compute the squared-r-distance term, then add in the other two. */

    x = MatchMap(p): p = p + 1        'cp->zrgb.r;

    If (x < minr) Then
      tdist = (x - minr) * R_SCALE:   min_dist = tdist * tdist
      tdist = (x - maxr) * R_SCALE:   max_dist = tdist * tdist
    ElseIf (x > maxr) Then
      tdist = (x - maxr) * R_SCALE:   min_dist = tdist * tdist
      tdist = (x - minr) * R_SCALE:   max_dist = tdist * tdist
    Else  '* within cell range so no contribution to min_dist */
      min_dist = 0
      If (x <= midr) Then
        tdist = (x - maxr) * R_SCALE: max_dist = tdist * tdist
      Else
        tdist = (x - minr) * R_SCALE: max_dist = tdist * tdist
      End If
    End If

    x = MatchMap(p): p = p + 1        'cp->zrgb.g;

    If (x < ming) Then
      tdist = (x - ming) * G_SCALE:   min_dist = min_dist + tdist * tdist
      tdist = (x - maxg) * G_SCALE:   max_dist = max_dist + tdist * tdist
    ElseIf (x > maxg) Then
      tdist = (x - maxg) * G_SCALE:   min_dist = min_dist + tdist * tdist
      tdist = (x - ming) * G_SCALE:   max_dist = max_dist + tdist * tdist
    Else '* within cell range so no contribution to min_dist */
      If (x <= midg) Then
        tdist = (x - maxg) * G_SCALE: max_dist = max_dist + tdist * tdist
      Else
        tdist = (x - ming) * G_SCALE: max_dist = max_dist + tdist * tdist
      End If
    End If

    x = MatchMap(p): p = p + 1        'cp->zrgb.b;

    If (x < minb) Then
      tdist = (x - minb) * B_SCALE:   min_dist = min_dist + tdist * tdist
      tdist = (x - maxb) * B_SCALE:   max_dist = max_dist + tdist * tdist
    ElseIf (x > maxb) Then
      tdist = (x - maxb) * B_SCALE:   min_dist = min_dist + tdist * tdist
      tdist = (x - minb) * B_SCALE:   max_dist = max_dist + tdist * tdist
    Else  '* within cell range so no contribution to min_dist */
      If (x <= midb) Then
        tdist = (x - maxb) * B_SCALE: max_dist = max_dist + tdist * tdist
      Else
        tdist = (x - minb) * B_SCALE: max_dist = max_dist + tdist * tdist
      End If
    End If

    mindist(i) = min_dist    '* save away the results */
    If (max_dist < minmaxdist) Then minmaxdist = max_dist
  Next i

  '* Now we know that no cell in the update box is more than minmaxdist
  '* away from some colormap entry.  Therefore, only colors that are
  '* within minmaxdist of some part of the box need be considered.

  NColors = 0
  For i = 0 To NMatchColors - 1
    If mindist(i) <= minmaxdist Then colorlist(NColors) = i: NColors = NColors + 1
  Next i
  
  find_nearby_colors = NColors
End Function

Private Sub find_best_colors(ByVal minr As Integer, ByVal ming As Integer, ByVal minb As Integer, _
                             ByVal numcolors As Integer, _
                             ByRef colorlist() As Byte, _
                             ByRef bestcolor() As Byte)

  Dim bptr As Long         '* pointer into bestdist[] array */
  Dim cptr As Long         '* pointer into bestcolor[] array */

  Dim ir As Integer, ig As Integer, ib As Integer
  Dim i As Integer, p As Integer, icolor As Integer

  Dim dist0 As Long, dist1 As Long                           '* initial distance values */
  Dim dist2 As Long                                          '* current distance in inner loop */
  Dim xx0 As Long, xx1 As Long, xx2 As Long                  '* distance increments */
  Dim incr_r As Long, incr_g As Long, incr_b As Long         '* initial values for increments */
  
  '* This array holds the distance to the nearest-so-far color for each cell */
  Dim bestdist(0 To BOX_R_ELEMS * BOX_G_ELEMS * BOX_B_ELEMS - 1) As Long

  '* Find the closest colormap entry for each cell in the update box,
  '* given the list of candidate colors prepared by find_nearby_colors.
  '* Return the indexes of the closest entries in the bestcolor[] array.
  '* This routine uses Thomas' incremental distance calculation method to
  '* find the distance from a colormap entry to successive cells in the box.

  '* Initialize best-distance for each cell of the update box */

  For i = BOX_R_ELEMS * BOX_G_ELEMS * BOX_B_ELEMS - 1 To 0 Step -1
    bestdist(i) = &H7FFFFFFF
  Next i

  '* For each color selected by find_nearby_colors,
  '* compute its distance to the center of each cell in the box.
  '* If that's less than best-so-far, update best distance and color number.
  
  '* Nominal steps between cell centers ("x" in Thomas article) */

  Const STEP_R As Integer = (R_SHIFT_VAL * R_SCALE)   ' ((1 << R_SHIFT) * R_SCALE)
  Const STEP_G As Integer = (G_SHIFT_VAL * G_SCALE)   ' ((1 << G_SHIFT) * G_SCALE)
  Const STEP_B As Integer = (B_SHIFT_VAL * B_SCALE)   ' ((1 << B_SHIFT) * B_SCALE)
  
  For i = 0 To numcolors - 1

    icolor = colorlist(i)
    p = 3 * icolor
    
    '* Compute (square of) distance from minr/g/b to this color */

    incr_r = (minr - MatchMap(p)) * R_SCALE: p = p + 1    'r
    dist0 = incr_r * incr_r

    incr_g = (ming - MatchMap(p)) * G_SCALE: p = p + 1    'g
    dist0 = dist0 + incr_g * incr_g

    incr_b = (minb - MatchMap(p)) * B_SCALE: p = p + 1    'b
    dist0 = dist0 + incr_b * incr_b

    '* Form the initial difference increments */

    incr_r = incr_r * (2 * STEP_R) + (STEP_R * STEP_R)
    incr_g = incr_g * (2 * STEP_G) + (STEP_G * STEP_G)
    incr_b = incr_b * (2 * STEP_B) + (STEP_B * STEP_B)

    '* Now loop over all cells in box, updating distance per Thomas method */

    bptr = 0
    cptr = 0
    xx0 = incr_r
    For ir = BOX_R_ELEMS - 1 To 0 Step -1
      dist1 = dist0
      xx1 = incr_g
      For ig = BOX_G_ELEMS - 1 To 0 Step -1
        dist2 = dist1
        xx2 = incr_b
        For ib = BOX_B_ELEMS - 1 To 0 Step -1
          If dist2 < bestdist(bptr) Then
            bestdist(bptr) = dist2
            bestcolor(cptr) = icolor
          End If
          dist2 = dist2 + xx2
          xx2 = xx2 + 2 * (STEP_B * STEP_B)
          bptr = bptr + 1
          cptr = cptr + 1
        Next ib
        dist1 = dist1 + xx1
        xx1 = xx1 + 2 * (STEP_G * STEP_G)
      Next ig
      dist0 = dist0 + xx0
      xx0 = xx0 + 2 * (STEP_R * STEP_R)
    Next ir
  Next i
End Sub

Private Sub fill_inverse_cmap(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer)

  Dim cache As Long, cptr As Long
  Dim minr As Integer, ming As Integer, minb As Integer        '* lower left corner of update box */
  Dim ir As Integer, ig As Integer, ib As Integer
  Dim numcolors As Integer                                     '* number of candidate colors */
  Dim bestcolor() As Byte, colorlist() As Byte
    
  '* Array of the actually closest colormap index for each cell. */
  ReDim bestcolor(0 To BOX_R_ELEMS * BOX_G_ELEMS * BOX_B_ELEMS - 1) As Byte
  
  '* array of candidate colormap indices. */
  ReDim colorlist(0 To MAXNUMCOLORS - 1) As Byte
  
  '* Fill the inverse-colormap entries in the update box that contains */
  '* histogram cell r/g/b.  (Only that one cell MUST be filled, but */
  '* we can fill as many others as we wish.) */

  '* Convert cell coordinates to update box ID */

  r = r \ BOX_R_LOG_VAL
  g = g \ BOX_G_LOG_VAL
  b = b \ BOX_B_LOG_VAL

  '* Compute true coordinates of update box's origin corner.
  '* Actually we compute the coordinates of the center of the corner
  '* histogram cell, which are the lower bounds of the volume we care about.
  
  minr = (r * BOX_R_SHIFT_VAL) + (R_SHIFT_VAL \ 2)
  ming = (g * BOX_G_SHIFT_VAL) + (G_SHIFT_VAL \ 2)
  minb = (b * BOX_B_SHIFT_VAL) + (B_SHIFT_VAL \ 2)

  '* Determine which colormap entries are close enough to be candidates
  '* for the nearest entry to some cell in the update box.

  numcolors = find_nearby_colors(minr, ming, minb, colorlist)

  If numcolors = 0 Then MsgBox "PANIC in fill_inverse map: find nearby colors broken" 'SHOULD NOT OCCUR

  '* Determine the actually nearest colors. */

  Call find_best_colors(minr, ming, minb, numcolors, colorlist, bestcolor)

  '* Save the best color numbers (plus 1) in the main cache array */

  r = r * BOX_R_LOG_VAL       '* convert ID back to base cell indexes */
  g = g * BOX_G_LOG_VAL
  b = b * BOX_B_LOG_VAL
  
  cptr = 0
  For ir = 0 To BOX_R_ELEMS - 1
    For ig = 0 To BOX_G_ELEMS - 1
      cache = b + CLng(HIST_B_ELEMS) * (r + ir + (g + ig) * CLng(HIST_R_ELEMS))
      For ib = 0 To BOX_B_ELEMS - 1
        Histogram(cache) = 1 + bestcolor(cptr)
        cache = cache + 1
        cptr = cptr + 1
      Next ib
    Next ig
  Next ir
End Sub

'returns the CMap Index of the nearest colour to r,g,b
Public Function MatchColorbyHistogram(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Integer
  
  Dim c As Integer, p As Long
  
  r = r \ R_SHIFT_VAL
  g = g \ G_SHIFT_VAL
  b = b \ B_SHIFT_VAL
  
  p = b + CLng(HIST_B_ELEMS) * (r + g * CLng(HIST_R_ELEMS))
  
  c = Histogram(p)
  If c = 0 Then       'We have not seen this color before, find nearest colormap entry and update the cache
    Call fill_inverse_cmap(r, g, b)
    c = Histogram(p)
  End If
  If c = 0 Then MsgBox "PANIC in match_color: inverse_cmap faulty"    'SHOULD NEVER HAPPEN
  MatchColorbyHistogram = c - 1
End Function


