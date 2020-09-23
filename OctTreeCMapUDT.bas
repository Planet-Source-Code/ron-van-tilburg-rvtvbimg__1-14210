Attribute VB_Name = "OctTreeCMap"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'OctTreeCMap.bas
'From Wicked Code 1997  Jeff Prosise
'Octree Color Quantization
'This text is abstracted in part from Microsoft Systems Journal. Copyright © 1995 by Miller Freeman, Inc.
'All rights are reserved (hohum).

'In 1988, M. Gervautz and W. Purgathofer of Austria's Technische UniversitŠt Wien published an article entitled "A Simple Method for Color Quantization: Octree Quantization." They proposed an elegant new method for quantizing color bitmap images by employing octrees—tree-like data structures whose nodes contain pointers to up to eight subnodes. Properly implemented, octree color quantization is at least as fast as the median-cut method and more memory-efficient.
'The basic idea in octree color quantization is to graph an image's RGB color values in a hierarchical octree.
'The octree can go up to nine levels deep—a root level plus one level for each bit in an 8-bit red, green, or
'blue value—but it's typically restricted to fewer levels to conserve memory. Lower levels correspond to less
'significant bits in RGB color values, so allowing the octree to grow deeper than five or six levels has
'little or no effect on the output. Leaf nodes (nodes with no children) store pixel counts and running totals
'of the red, green, and blue color components of the pixels encoded there, while intermediate nodes form paths
'from the topmost level in the octree to the leaves. This is an efficient way to count colors and the number
'of occurrences of each color because no memory is allocated for colors that don't appear in the image. If
'the number of leaf nodes happens to be equal to or less than the number of palette colors you want, you can
'fill a palette simply by traversing the octree and copying RGB values from its leaves.

'The beauty of the octree method is what happens when the number of leaf nodes n exceeds the desired number
'of palette colors k. Each time adding a color to the octree creates a new leaf, n is compared to k. If n is
'greater than k, the tree is reduced by merging one or more leaf nodes into the parent. After the operation
'is complete, the parent, which was an intermediate node, is a leaf node that stores the combined color
'information of all its former children.

'Because the octree is trimmed continually to keep the leaf count under k, you end up with an octree
'containing k or fewer leaves whose RGB values make ideal palette colors. No matter how many colors the image
'contains, you can walk the octree and pick leaves off it to formulate a palette. Better yet, the octree never
'requires memory for more than k+1 leaf nodes plus some number of intermediate nodes.

'There are two parts of an octree that I want to study: the parent-child relationship between nodes and the
'significance of the RGB data in each leaf. Figure 1 shows the parent-child relationship for each node. At
'a given level in the tree, a value from zero to 7, derived from the RGB color value, identifies a child node.
'At the uppermost (root) level, bit 7 of the red value is combined with bit 7 of the green value and bit 7
'of the blue value to form a 3-bit index. Bit 7 from the red value becomes bit 2 in the index, bit 7 from the
'green value becomes bit 1 in the index, and bit 7 from the blue value becomes bit zero in the index. At the
'next level, bit 6 is used instead of bit 7, and the bit number keeps decreasing as the level number
'increases. For red, green, and blue color values equal to 109 (binary 01101101), 204 (11001100),
'and 170 (10101010), the index of the first child node is 3 (011), the index of the second child node
'is 6 (110), and so on. This mechanism places the more significant bits of the RGB values at the top of the
'tree. In this example, the octree's depth is restricted to six levels, which allows you to factor in up
'to 5 bits of each 8-bit color component. The remaining bits are effectively averaged together.
  
  'TEXT BY RVT
  'the OctTree structure is implemented as an array of UDTs of Type uOctTreeNode. This is a little less memory
  'efficient but faster than using Classes to implement with. The code is fairly terse and reasonably quick.
  'I deliberately let the tree grow to 512nodes then Prune it down to the required number of colors
  'This is to counter a tendency in the Prune algorithm to merely trim 1 node parents thus effectively
  'cutting the overall bit resolution down by 1
  
  'NOTE: I do get errors when trying to get to 2 or 4 colors, 8 and up is OK. This is because pruning to
  'this small may actually remove all of the colors (7 may be removed per turn) RVT
  
Private Type uOctTreeNode
  isLeaf         As Boolean            ' TRUE if node has no children
  RedSum         As Long               ' Sum of red components
  GreenSum       As Long               ' Sum of green components
  BlueSum        As Long               ' Sum of blue components
  NPixels        As Long               ' Number of pixels represented by this leaf
  pChild(0 To 7) As Integer            ' Pointers to child nodes in variant array
  pNext          As Integer            ' Pointer to next reducible node in reducibles array
  ID             As Integer
End Type

Private Type IRGB
  Index As Byte                        'colorindex
  r     As Byte                        'red
  b     As Byte                        'green
  g     As Byte                        'blue
End Type

  '********************************************************************************************
  '  CreateOctTreeCMap presents an implementation of the Gervautz-Purgathofer octree
  '  color quanitization algorithm that creates optimized color palettes for
  '  for 16, 24, and 32-bit DIB sections. The code is barely recognisable from the C
  '  quoted in the article. In order to speed things up Ive Globalised most of the
  '  variables that were passed recursively (ad nauseum) onto the stack.
  '  I then took the step of removing all the recursion itself other than the tree traversal to get
  '  the colormap at the end. (Every little bit helps)
  '  I also added the JustSeen steps (which really helps on drawings with large blocks
  '  of the same colour.) Also implemented as an Array of UDT Nodes . Some housekeeping is done to
  '  minimise memory use and resuse previously deleted nodes                       RVT
  '********************************************************************************************

  'The routine returns the CMap containing the best nMaxColors (to the nearest power of 2) fitting,
  '(CMAP is redimensioned) the 16BPP, 24BPP or 32BPP data passed in as PixBits()
  'OctTreeDepthBits (5 or 6 usually) will be used in the determination of the palette

Dim OctTree                 As Integer          'The ColorOctTree
Dim NLeaves                 As Long             'the number of Leaves in the Tree
Dim OTNodes()               As uOctTreeNode     'we keep the OTNodes in an array of structures
                                                'OTNodes(0) is the root, all others form there
Dim FirstFreeNode           As Integer          'look for a free node from here
Dim ReducibleNodes(0 To 8)  As Integer          'the list of nodes that could be removed
Dim LastLeafSeen            As Integer          'the Last Leaf accessed
Dim LastRGBUsed             As IRGB             'the last color reference at LastLeafSeen

Dim NColorBits              As Long             'the BitDepth of the OctTree
Dim ID                      As Long

Public Function CreateOctTreeCMap(ByVal Width As Long, ByVal Height As Long, _
                                 ByRef PixBits() As Byte, ByVal BitsPerPixel As Long, _
                                 ByRef CMap() As Byte, ByRef NMapColors As Long, _
                                 ByVal OctTreeDepthBits As Long) As Long
                                 
  Dim ColorIndex As Long
  Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, i As Long
  Dim rgb As IRGB, wColor As Long

  On Error GoTo q
  
   ' Initialize octree variables
  Call DeleteOctTreeCMap            'wipe out any previous tree
  ReDim OTNodes(0 To 511) As uOctTreeNode
  If OctTreeDepthBits < 1 Or OctTreeDepthBits > 8 Then
    NColorBits = 5
  Else
    NColorBits = OctTreeDepthBits
  End If
  
   ' Scan the PixBits() and build the octree
  
   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Function      'will only work for unmapped DIBs
  
  Skip = (BitsPerPixel \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  For y = 0 To Height - 1                                               'assume bmap is right way up
    z = y * RowMod
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row
    
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          rgb.b = (wColor And &H1F&) * 8&
          rgb.g = (wColor And &H3E0&) \ 4&
          rgb.r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          rgb.b = PixBits(x)
          rgb.g = PixBits(x + 1)
          rgb.r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          rgb.r = PixBits(x + 1)
          rgb.g = PixBits(x + 2)
          rgb.b = PixBits(x + 3)
      End Select
      
      'add the color to the OctTree
      If JustSeen(rgb) = True Then  'Speed kludge:  first Check to see if we have just added this color
        Call IncrementLastLeafSeen
      Else
        Call AddColor(rgb)
      End If
      
      'and if its too big prune it back
      If NLeaves > 512 Then Call PruneTree(512)
    Next x
  Next y
  
  Call PruneTree(NMapColors)
  ColorIndex = 0
  Call GetCMapColors(OctTree, CMap, ColorIndex)
  ColorIndex = ColorIndex \ 3      'the nr of colors really used
  If ColorIndex < 2 Then ColorIndex = 2              'no point in a map of 1
  x = 1
  y = 0
  Do While x < ColorIndex
    x = x + x
    y = y + 1
  Loop
  NMapColors = x: BitsPerPixel = y
  ReDim Preserve CMap(0 To 3 * NMapColors - 1) As Byte  'the real size required (only powers of 2)
  
'  MsgBox "OK OctTree " & ColorIndex & " colors made, CMapSize= " & NMapColors & ", " & BitsPerPixel & "BPP"

  CreateOctTreeCMap = BitsPerPixel
  Exit Function
q:
  MsgBox "OctTreeFailed"
  On Error GoTo 0
  CreateOctTreeCMap = 0
End Function

Public Sub DeleteOctTreeCMap()     'cleanup all of the Node references
  Erase ReducibleNodes()
  Erase OTNodes()
  LastLeafSeen = 0
  OctTree = -1
  NLeaves = 0                    'no tree, no leaves
  FirstFreeNode = 0
End Sub

Private Function JustSeen(ByRef rgb As IRGB) As Boolean
  JustSeen = False
  If LastLeafSeen >= 0 Then
    If LastRGBUsed.r = rgb.r Then
      If LastRGBUsed.g = rgb.g Then
        If LastRGBUsed.b = rgb.b Then
          JustSeen = True
        End If
      End If
    End If
  End If
End Function

Private Sub AddColor(ByRef rgb As IRGB)

  Dim pNode As Integer, qNode As Integer
  Dim Idx As Integer, Mask As Integer, Level As Integer

  If OctTree <> 0 Then ID = -1: OctTree = CreateNode(0) 'first time round
  
  pNode = OctTree
  Mask = &H80           '2^(7-Level)
  Level = 0
  Do
    Idx = 0
    If (rgb.r And Mask) <> 0 Then Idx = Idx Or 4
    If (rgb.g And Mask) <> 0 Then Idx = Idx Or 2
    If (rgb.b And Mask) <> 0 Then Idx = Idx Or 1
    qNode = OTNodes(pNode).pChild(Idx)
    Level = Level + 1
    If qNode = 0 Then
      qNode = CreateNode(Level)                  'if the node doesn't exist, create it
      OTNodes(pNode).pChild(Idx) = qNode         'and relink it
    End If
    Mask = Mask \ 2
    pNode = qNode
  Loop Until OTNodes(qNode).isLeaf = True        'must now be a leaf node
  
  LastRGBUsed = rgb
  LastLeafSeen = qNode
  Call IncrementLastLeafSeen                     ' Update color information
End Sub

Private Sub IncrementLastLeafSeen()
  With OTNodes(LastLeafSeen)
    .NPixels = .NPixels + 1
    .RedSum = .RedSum + LastRGBUsed.r
    .GreenSum = .GreenSum + LastRGBUsed.g
    .BlueSum = .BlueSum + LastRGBUsed.b
  End With
End Sub

Private Function CreateNode(ByVal Level As Long) As Integer
  Dim i As Long
  
  ID = FirstFreeNode
  With OTNodes(ID)
    .ID = ID
    .isLeaf = False
    .pNext = 0
    .RedSum = 0
    .GreenSum = 0
    .BlueSum = 0
    .NPixels = 0
    Erase .pChild()
  End With
  
  If Level = NColorBits Then
    OTNodes(ID).isLeaf = True
    NLeaves = NLeaves + 1
  Else                                        ' Add the node to the reducible list for this level
    OTNodes(ID).pNext = ReducibleNodes(Level)
    ReducibleNodes(Level) = ID
  End If
  
  For i = FirstFreeNode + 1 To UBound(OTNodes)              'now find the first free slot
    If i = UBound(OTNodes) Then                             'we are about to run out of space
      ReDim Preserve OTNodes(0 To i + 128) As uOctTreeNode  'upsize it a little
    End If
    If OTNodes(i).ID = 0 Then
      FirstFreeNode = i
      Exit For
    End If
  Next
  
  CreateNode = ID
End Function

Private Sub PruneTree(ByVal MaxLeaves As Long)
  Dim i As Integer
  Dim pNode As Integer
  Dim RedSum As Long, GreenSum As Long, BlueSum As Long, NChildren As Long, NPixels As Long

  Do While NLeaves > MaxLeaves
    ' Find the deepest level containing at least one reducible node
    For i = NColorBits - 1 To 0 Step -1
      If ReducibleNodes(i) <> 0 Then Exit For
    Next i
    
     ' Reduce the first node in the most recently added list at level i
    pNode = ReducibleNodes(i)
    ReducibleNodes(i) = OTNodes(pNode).pNext
  
    NChildren = 0
    RedSum = 0
    GreenSum = 0
    BlueSum = 0
    NPixels = 0
    For i = 0 To 7
      If OTNodes(pNode).pChild(i) <> 0 Then
        With OTNodes(OTNodes(pNode).pChild(i))
          RedSum = RedSum + .RedSum
          GreenSum = GreenSum + .GreenSum
          BlueSum = BlueSum + .BlueSum
          NPixels = NPixels + .NPixels
          If .ID < FirstFreeNode Then FirstFreeNode = .ID
          .ID = 0     'now freed
        End With
        OTNodes(pNode).pChild(i) = 0
        NChildren = NChildren + 1
      End If
    Next i
    With OTNodes(pNode)
      .isLeaf = True
      .RedSum = RedSum
      .GreenSum = GreenSum
      .BlueSum = BlueSum
      .NPixels = NPixels
    End With
    NLeaves = NLeaves - (NChildren - 1)
  Loop
  LastLeafSeen = -1
End Sub

Private Sub GetCMapColors(ByRef pNode As Integer, ByRef CMap() As Byte, ByRef ColorIndex As Long)

  Dim i As Integer, n As Long

  If OTNodes(pNode).isLeaf Then
    With OTNodes(pNode)
      n = .NPixels
      .RedSum = .RedSum / n     'these are the rgb values for the colour at this node
      .GreenSum = .GreenSum / n
      .BlueSum = .BlueSum / n
      .NPixels = ColorIndex \ 3        'we reuse this as the CMAP index
    
      CMap(ColorIndex) = .RedSum:   ColorIndex = ColorIndex + 1 'r mean
      CMap(ColorIndex) = .GreenSum: ColorIndex = ColorIndex + 1 'g mean
      CMap(ColorIndex) = .BlueSum:  ColorIndex = ColorIndex + 1 'b mean
    End With
  Else
    For i = 0 To 7
      If OTNodes(pNode).pChild(i) <> 0 Then Call GetCMapColors(OTNodes(pNode).pChild(i), CMap, ColorIndex)
    Next i
  End If
End Sub

'Back to text by Jeff Prosise
'One implementation detail you should be aware of, especially if you want to modify the code, is how
'PruneTree picks a node to reduce. It's important to do your reductions as deep in the octree as possible
'because deeper levels correspond to subtler variations in tone. Since it's time-consuming to traverse the
'entire tree from top to bottom, I borrowed an idea from an excellent article by Dean Clark that appeared in
'the January 1996 issue of Dr. Dobb's Journal ("Color Quantization Using Octrees"); I implemented an array
'of linked lists named ReducibleNodes containing pointers to all the reducible (non-leaf) nodes in each
'level of the octree. Thus, finding the deepest level with a reducible node is as simple as scanning the
'array from bottom to top until a non-NULL pointer is located. Like Dean's code, mine picks the node most
'recently added to a given level as the one to reduce. You could refine my implementation somewhat by adding
'more intelligence to the node-selection process—for example, by walking the linked list and picking the node
'that represents the fewest or the most colors.
'Jeff Prosise: 72241.44@compuserve.com

Private Function MatchColorbyOctTree(ByRef rgb As IRGB) As Integer

  Dim Idx As Integer, Mask As Integer, Level As Integer, Node As Integer

  MatchColorbyOctTree = -1
  If OctTree = 0 Then
    Node = OctTree
    Mask = &H80           '2^(7-Level)
    Level = 0
    Do
      Idx = 0
      If (rgb.r And Mask) <> 0 Then Idx = Idx Or 4
      If (rgb.g And Mask) <> 0 Then Idx = Idx Or 2
      If (rgb.b And Mask) <> 0 Then Idx = Idx Or 1
      Node = OTNodes(Node).pChild(Idx)
      If Node = 0 Then Exit Function           'this shouldnt happen
      Level = Level + 1
      Mask = Mask \ 2
    Loop Until OTNodes(Node).isLeaf = True            'must now be a leaf node
    
    LastLeafSeen = Node
    rgb.Index = OTNodes(LastLeafSeen).NPixels
    LastRGBUsed = rgb
    MatchColorbyOctTree = rgb.Index
  End If
End Function
  
'=======================================================================================================
'RVT NOTE: one cannot reliably colormap against an Octree with a fixed CMap. When colors in the picture are
'not in the tree map there will be no match found. So we use another method for mapping see CHist
'I left this one here so that you can experiment
'======================================================================================================

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by direct substitution. This should really be dithered to minimise the
'error propagation of approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly sized without padding
'every entry will be the index into the map

Public Sub OctTreeMapColors(ByVal Width As Long, ByVal Height As Long, _
                            ByRef PixBits() As Byte, _
                            ByVal BitsPerPixel As Long, _
                            ByRef CMap() As Byte, _
                            ByVal NMapColors As Long)
                                 
  Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, i As Long
  Dim rgb As IRGB, wColor As Long, p As Long

   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Sub           'will only work for unmapped DIBs
   
   ' Scan the PixBits() and build the octree
  Skip = (BitsPerPixel \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  p = 0                                                                 'this is where we put the new byte
  For y = 0 To Height - 1                                               'assume bmap is right way up
    z = y * RowMod
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row
    
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          rgb.b = (wColor And &H1F&) * 8&
          rgb.g = (wColor And &H3E0&) \ 4&
          rgb.r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          rgb.b = PixBits(x)
          rgb.g = PixBits(x + 1)
          rgb.r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          rgb.r = PixBits(x + 1)
          rgb.g = PixBits(x + 2)
          rgb.b = PixBits(x + 3)
      End Select
      
      'find the nearest the color to the given color
      If JustSeen(rgb) = True Then
      '  i = LastRGBUsed.Index        'i hasnt changed since last loop
      Else
        i = MatchColorbyOctTree(rgb)
      End If
      
      PixBits(p) = i: p = p + 1
    Next x
  Next y
    
  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To Width * Height - 1) As Byte
'  MsgBox "OK Remap"
End Sub
