VERSION 5.00
Begin VB.Form frmMemCmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compile for True Tests"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1260
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   3510
      Visible         =   0   'False
      Width           =   4605
      Begin VB.TextBox Text2 
         Height          =   600
         Index           =   1
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   585
         Width           =   4530
      End
      Begin VB.Label Label2 
         Caption         =   "Simple Loop with byte to byte comparison"
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   4485
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1260
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   3510
      Width           =   4605
      Begin VB.TextBox Text2 
         Height          =   600
         Index           =   0
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   585
         Width           =   4530
      End
      Begin VB.Label Label2 
         Caption         =   "If NTDLL.dll is available, this is the time that RtlCompareMemory takes to compare the same array"
         Height          =   465
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   4485
      End
   End
   Begin VB.CheckBox chkMakeUnequal 
      Caption         =   "Random Unequal"
      Height          =   495
      Index           =   1
      Left            =   3465
      TabIndex        =   4
      Top             =   2123
      Width           =   1125
   End
   Begin VB.CheckBox chkDblSize 
      Caption         =   "Double the Array Size?"
      Height          =   450
      Left            =   3180
      TabIndex        =   3
      Top             =   360
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Compare UDT array (LOGFONT) -- Random Amount"
      Height          =   345
      Index           =   2
      Left            =   315
      TabIndex        =   2
      Top             =   825
      Width           =   4140
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Simulate a full screen 8 bit DIB"
      Height          =   345
      Index           =   1
      Left            =   315
      TabIndex        =   1
      Top             =   465
      Width           =   2715
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Simulate a full screen 24 bit DIB"
      Height          =   345
      Index           =   0
      Left            =   315
      TabIndex        =   0
      Top             =   105
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      Height          =   600
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2655
      Width           =   4530
   End
   Begin VB.CheckBox chkMakeUnequal 
      Caption         =   "Make Last Elements in Arrays Unequal"
      Height          =   540
      Index           =   0
      Left            =   1485
      TabIndex        =   5
      Top             =   2100
      Width           =   1830
   End
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare Now"
      Height          =   495
      Left            =   105
      TabIndex        =   6
      Top             =   2115
      Width           =   1350
   End
   Begin VB.CheckBox chkOthers 
      Caption         =   "Simple Loop Comparison"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   7
      Top             =   3270
      Width           =   2175
   End
   Begin VB.CheckBox chkOthers 
      Caption         =   "NTDLL API call - fastest"
      Height          =   270
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   3270
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   825
      Left            =   300
      TabIndex        =   10
      Top             =   1215
      Width           =   4155
   End
End
Attribute VB_Name = "frmMemCmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' VB answer to the unavailable C memcmp function & RtlCompareMemory API.
'   FYI: RtlCompareMemory is only available on NT-based machines (NT4 and above)

' Tested on Win95, 98, 2K, XP Pro
' Not usable for comparing string arrays because the memory in a string array is
' not contiguous. The string array data is actually pointers to the strings, not
' the strings themselves.  However, if you wish to compare the string arrays to
' see if they actually contain the same pointers, then this would work for you too.

' The amazing thing, quite surprising to me, is that this routine has nearly
' identical speed to RtlCompareMemory (+ 1>3 ms) when the following conditions met:
'   1. In project property (Compile Tab, Optimizations button),
'           "Remove Array Bounds Check" is checked/selected
'   2. Compile the project for maximum speed
'   FYI> Do not "Remove Array Bounds Check" if you are relying on Error trapping to
'        check for uninitialized arrays or if you are using it to test for
'        array out of bounds errors. Without that optimization, the routines are
'        still very fast, but are faster with the optimization.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY1D        ' used to overlay array on a memory address
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound As SAFEARRAYBOUND
End Type


' This API may not be on your system. Error checks in routine will validate.
' Additionally, this API & following boolean are not needed for the vbMemCompare routine
Private Declare Function CompMemory Lib "ntdll.dll" Alias "RtlCompareMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private bCanUseNTDLL As Boolean ' flag indicating we can use RtlCompareMemory

' following UDT used only for testing
Private Type LOGFONT
    lfHeight As Long                        ' 4 bytes
    lfWidth As Long                         '+4
    lfEscapement As Long                    '+4
    lfOrientation As Long                   '+4
    lfWeight As Long                        '+4 = 20 bytes
    lfItalic As Byte                        '+1
    lfUnderline As Byte                     '+1
    lfStrikeOut As Byte                     '+1
    lfCharSet As Byte                       '+1
    lfOutPrecision As Byte                  '+1
    lfClipPrecision As Byte                 '+1
    lfQuality As Byte                       '+1
    lfPitchAndFamily As Byte                '+1 = 28 bytes
    lfFaceName(1 To 32) As Byte             '+32= 60 bytes total
    
    ' the above UDT can be compared reliably using vbMemCompare
    ' but another version of this UDT member follows & cannot be compared because
    ' the lfFaceName member will be a string pointer, it won't actually contain the
    ' string therefore 2 UDTs having Tahoma as the font name won't equal each other
    ' because the string pointers will likely be different between the 2 UDTs...
    
    'lfFaceName As Long
    
    ' now for another variation... with the following valid member definition as a
    ' fixed string, you must account for the 2-byte per character string format that
    ' VB uses internally when passing fixed strings to our vbMemCompare function.
    
    'lfFaceName As String * 32
    
    ' If the above UDT was passed to an API, VB will convert lfFaceName to 32 bytes
    ' for the API, but internally, it is 64 bytes. Therefore, to use vbMemCompare
    ' with fixed length strings within UDTs, use LenB(UDT) vs Len(UDT). The above
    ' variation run against Len() is 60 bytes, but against LenB() is 92 bytes
    
End Type




Private Function vbMemCompare(ByVal MemAddr1 As Long, ByVal MemAddr2 As Long, _
                              ByVal nrBytesToCompare As Long, _
                              Optional ByRef unEqualLocation As Long) As Boolean
    
    ' a flexible memcmp / RtlCompareMemory substitute for VB6
    
    ' WARNING: JUST LIKE CopyMemory, RtlCompareMemory, & MemCmp
    '   this routine does NO safety checks, it can't. The pointers you pass must be
    '   valid & number of bytes to be compared must be contiguous & not overestimated
    
    ' Requires following Declarations:
    '       Types: SAFEARRAYBOUND, SAFEARRAY1D
    '       APIs:  VarPtrArray, CopyMemory
    
    ' [in] MemAddr1 :: memory address (i.e., VarPtr) to compare against MemAddr2
    ' [in] MemAddr2 :: memory address (i.e., VarPtr) to compare against MemAddr1
    ' [in] nrBytesToCompare :: contiguous bytes to be compared
    '      The bytes used starting at MemAddr1,2 must be => nrBytesToCompare
    ' [out] unEqualLocation (Optional) :: byte where inequality occurred (if any)
    '       See end of this routine for tips on using this value if desired
    ' [out] Return Value :: True if comparision proves identical else False
    
    ' tips on use:
    '   The 1st element to be compared in the byte arrays can be anything you like, not necessarily zero
    ' Compare 2 byte arrays: vbMemCompare(VarPtr(aByte1(0)), VarPtr(aByte1(0)), UBound(aByte1) + 1)
    ' Compare 2 Long arrays: vbMemCompare(VarPtr(aLong1(0)), VarPtr(aLong2(0)), (UBound(aLong1) + 1) * 4&)
    ' Compare mix arrays: let's say compare 1000 Longs (4 bytes each) against 4000 bytes:
    '       vbMemCompare(VarPtr(aLong(0)), VarPtr(aByte(0)), 4000&)
    ' Compare 2 DIBs same format. vbMemCompare(dibPtr1, dibPtr2, (Dib ScanWidth * Dib Height))
    ' Compare DIB against DDB bytes returned by GetDIBits:
    '       vbMemCompare(dibPtr, VarPtr(dibBits(0)), (UBound(dibBits) + 1))
    ' Compare within same array: example of comparing 1st 5000 bytes with last 5000 bytes
    '       vbMemCompare(VarPtr(aByte1(0)), VarPtr(aByte1(5000)), 5000&)
    
    ' basic sanity checks
    If nrBytesToCompare < 1 Then
        Exit Function
    ElseIf MemAddr1 = 0 Then
        Exit Function
    ElseIf MemAddr2 = 0 Then
        Exit Function
    End If
    
    Dim Index As Long ' loop variable
    Const ScanSize As Long = 8& ' the tScan1/tScan2 array types must match this size
    ' FYI: using Date or Double for the tScan array types can result in Overflows
    
    ' array overlays
    Dim tSA1 As SAFEARRAY1D, tSA2 As SAFEARRAY1D
    Dim tScan1() As Currency, tScan2() As Currency      ' 8 byte scans
    Dim tBytes1() As Byte, tBytes2() As Byte            ' 1 byte scans
    
    If nrBytesToCompare >= ScanSize Then
    
        ' set up ScanSize byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .pvData = MemAddr1
            .cbElements = ScanSize
            .rgSABound.cElements = (nrBytesToCompare \ ScanSize) ' truncate for now
        End With
        ' set up ScanSize byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2
        
        ' overlay now
        CopyMemory ByVal VarPtrArray(tScan1), VarPtr(tSA1), 4&
        CopyMemory ByVal VarPtrArray(tScan2), VarPtr(tSA2), 4&
        
        ' compare, ScanSize bytes at a time. Wish VB had a 16, 32 or 64 variable type
        For Index = 0 To UBound(tScan1)
            If Not tScan1(Index) = tScan2(Index) Then Exit For
            ' bug out when inequality is found
        Next
        ' remove the overlays
        CopyMemory ByVal VarPtrArray(tScan1), 0&, 4&
        CopyMemory ByVal VarPtrArray(tScan2), 0&, 4&
            
        Index = Index * ScanSize ' set Index = actual byte to be checked next
        
    End If
    
    If Not Index = nrBytesToCompare Then
        ' unequal if all bytes were checked....
        ' locate exact byte position where inequality was located
        ' This also will check any bytes not compared due to non-ScanSize alignment
        
        ' set up 1 byte scan over the 1st passed memory pointer
        With tSA1
            .cDims = 1
            .cbElements = 1
            .pvData = MemAddr1 + Index  ' move memory pointer to where Index left off
            .rgSABound.lLbound = Index  ' adjust LBound to where Index left off
            .rgSABound.cElements = (nrBytesToCompare - Index)  ' nr elements remaining
        End With
        ' set up 1 byte scan over the 2nd passed memory pointer
        tSA2 = tSA1
        tSA2.pvData = MemAddr2 + Index  ' move memory ponter to where Index left off
        ' overlay now
        CopyMemory ByVal VarPtrArray(tBytes1), VarPtr(tSA1), 4&
        CopyMemory ByVal VarPtrArray(tBytes2), VarPtr(tSA2), 4&
    
        ' do the comparison and/or check last n bytes
        For Index = Index To nrBytesToCompare - 1 '(max of ScanSize loops)
            If Not tBytes1(Index) = tBytes2(Index) Then Exit For
            ' bug out when inequality is found
        Next
        ' remove overlays
        CopyMemory ByVal VarPtrArray(tBytes1), 0&, 4&
        CopyMemory ByVal VarPtrArray(tBytes2), 0&, 4&

    End If
    
    ' return result(s)
    unEqualLocation = Index
    vbMemCompare = (unEqualLocation = nrBytesToCompare)
    
    ' If you wish to identify where in your passed memory the inequality occured in
    ' relation to arrays, pixels, or memory addresses...
    
    ' This routine has no way of knowing whether you passed it a long, byte, integer
    ' array or whether you passed it memory addresses like DIB pointers. Suggest
    ' using following algos with the returned unEqualLocation parameter.
    
    ' Arrays: Note aStartA & aStartB are array elements passed to this
    '   function (i.e., byteArrayA(0), byteArrayB(1) where aStartA=0, aStartB=1)
    
    ' Long arrays.
    '       Loc = unEqualLocation\4 ' Calculate Loc
    '       LongA(Loc + aStartA) <> LongB(Loc + aStartB))
    
    ' Integer Arrays.
    '       Loc = unEqualLocation\2 ' Calculate Loc
    '       IntegerA(Loc + aStartA) <> IntegerB(Loc + aStartB))
    
    ' Byte Arrays
    '       ByteA(unEqualLocation + aStartA) <> ByteB(unEqualLocation + aStartB))
    
    ' Date/Double/Currency arrays.
    '       Loc = unEqualLocation\8 ' Calculate Loc
    '       DateA(Loc + aStartA) <> DateB(Loc + aStartB))
    
    ' UDTs Arrays.
    '   Yes it is possible, but not if UDT contains pointers to other memory addresses
    '   For example, VarLen string members are pointers in UDT memory & using
    '   vbMemCompare should return inequality every time since pointers won't be the same
    '       Loc = unEqualLocation\Len(UDT) ' Calculate Loc
    '       aUDT1(Loc + aStartA) <> aUDT2(Loc + aStartB))
    '   but which member of the UDT proved inequality? Using the following you
    '   can partially determine, but maybe comparing the 2 UDT members might be best:
    '       byte within aUDT1/aUDT2 at location: unEqualLocation-(Len(UDT)*Loc))
    
    ' DIB pointers: depending on bit depth, tweak result
    ' (remember to adjust result for bottom-up DIBs if needed)
    '   -- Pixel Colors:
    '       8 bit: unEqualLocation is the palette index
    '       24 bit: unEqualLocation\3 is 1st byte of the pixel
    '       32 bit: unEqualLocation\4  is 1st byte of the pixel
    '   -- Pixel Location (DIB row & column)
    '       8 bit:  Row = unEqualLocation\Bitmap.ScanWidth
    '               Column = unEqualLocation-(Row * Bitmap.ScanWidth)
    '      24 bit:  Row = unEqualLocation\Bitmap.ScanWidth
    '               Column = (unEqualLocation-(Row * Bitmap.ScanWidth))\3
    '      32 bit:  Row = unEqualLocation\Bitmap.ScanWidth
    '               Column = (unEqualLocation-(Row * Bitmap.ScanWidth))\4
    
    ' Memory Pointers: simply, memPointer + unEqualLocation
    
End Function



Private Sub cmdCompare_Click()

    Dim a() As Byte, b() As Byte
    Dim X As Long
    Dim nrBytes As Long, bUDTexample As Boolean
    
    Dim lgFont1() As LOGFONT, lgFont2() As LOGFONT

    ' determine number of bytes across width of "simualted screen" for bit depth selected
    Select Case True
    Case Option1(0): nrBytes = ByteAlignOnWord(24, Screen.Width \ Screen.TwipsPerPixelX)
    Case Option1(1): nrBytes = ByteAlignOnWord(8, Screen.Width \ Screen.TwipsPerPixelX)
    Case Option1(2): bUDTexample = True
    End Select
    
    If bUDTexample = True Then
        ' Late entry example. Simply shows that many UDTs can be compared,
        ' including possibly your custom UDTs. Things to keep in mind when
        ' comparing UDTs.  Any member of the UDT cannot contain pointers.
        ' See comments in the LOGFONT declaration near top of this module,
        ' especially regarding comparing fixed length strings
        X = CLng(Rnd * 10000 + 1000)
        ' double the array size if wanted
        If chkDblSize = 1 Then X = X * 2
        
        ReDim lgFont1(1 To X)
        ReDim lgFont2(1 To X)
        nrBytes = Len(lgFont1(1)) * X
        
        If chkMakeUnequal(0) = 1 Then
            lgFont1(X).lfFaceName(32) = 62
        ElseIf chkMakeUnequal(1) = 1 Then
            X = CLng(Rnd * (X - 10) + 1)
            lgFont1(X).lfFaceName(29) = 62
        End If
        ' if UDT used fixed length strings, then we would use
        '   LenB() for vbMemCompare
        ' but would use Len() if passing UDT to an API
        
    Else ' simulated DIB example
        ' multiply by simulated screen height
        nrBytes = nrBytes * (Screen.Height \ Screen.TwipsPerPixelY)
        ' double the array size if wanted
        If chkDblSize = 1 Then nrBytes = nrBytes * 2
        ' create 2 identical arrays (all zeros)
        ReDim a(0 To nrBytes - 1)
        ReDim b(0 To nrBytes - 1)

        ' make one of the arrays unequal (doesn't matter which one) if desired
        If chkMakeUnequal(0) = 1 Then
            a(nrBytes - 1) = 62
        ElseIf chkMakeUnequal(1) = 1 Then
            ' random byte selection
            b(CLng(Rnd * nrBytes)) = 62
        End If
    End If
    
    Dim cTimer As New cTiming, lBytesChecked As Long, bResult As Boolean
    
    cTimer.Reset    ' start a timer
    If bUDTexample = True Then
        ' comparing UDTs
        bResult = vbMemCompare(VarPtr(lgFont1(1)), VarPtr(lgFont2(1)), nrBytes, lBytesChecked)
    Else
        ' comparing arrays
        bResult = vbMemCompare(VarPtr(a(0)), VarPtr(b(0)), nrBytes, lBytesChecked)
    End If
    ' show the results
    Text1.Text = "Arrays equal? {" & bResult & "}, compared in " & Format(cTimer.Elapsed, "Standard") & " ms."
    Text1.Text = Text1.Text & vbNewLine & "Number bytes compared: " & Format(lBytesChecked, "Standard")



    ' simple loop test. Averages at least 8x slower than vbMemCompare
    If bUDTexample = True Then
        ReDim a(0 To nrBytes - 1) ' convert UDT to bytes for this comparison; easier
        ReDim b(0 To UBound(a))
        CopyMemory a(0), lgFont1(1), nrBytes
    End If
    cTimer.Reset ' start a timer
    For X = 0 To nrBytes - 1
        If Not a(X) = b(X) Then Exit For
    Next
    Text2(1).Text = "Arrays equal? {" & CBool(X = nrBytes) & "}, compared in " & Format(cTimer.Elapsed, "Standard") & " ms."
    Text2(1).Text = Text2(1).Text & vbNewLine & "Number bytes compared: " & Format(X, "Standard")
    
    ' use RtlCompareMemory, if available, for another comparison/benchmark
    If bCanUseNTDLL Then
        
        On Error Resume Next
        cTimer.Reset ' start a timer
        If bUDTexample = True Then
            lBytesChecked = CompMemory(lgFont1(1), lgFont2(1), nrBytes)
        Else
            lBytesChecked = CompMemory(a(0), b(0), nrBytes)
        End If
        If Err Then
            ' prevent trying this again
            Text2(0).Text = "NTDLL not available on your system"
            bCanUseNTDLL = False
            Err.Clear
        Else
            Text2(0).Text = "Arrays equal? {" & CBool(lBytesChecked = nrBytes) & "}, compared in " & Format(cTimer.Elapsed, "Standard") & " ms."
            Text2(0).Text = Text2(0).Text & vbNewLine & "Number bytes compared: " & Format(lBytesChecked, "Standard")
        End If
    
    End If
    
    
End Sub

Private Sub chkMakeUnequal_Click(Index As Integer)
    If chkMakeUnequal(Index) = 1 Then
        chkMakeUnequal(Abs(Index - 1)) = 0 ' when one is checked, uncheck the other
    End If
End Sub

Private Sub chkOthers_Click(Index As Integer)
    If chkOthers(Index) = 1 Then
        ' toggle frame visibility
        chkOthers(Abs(Index - 1)) = 0 ' when one is checked, uncheck the other
        Frame1(Abs(Index - 1)).Visible = False
        Frame1(Index).ZOrder
        Frame1(Index).Visible = True
    End If
End Sub

Private Sub Form_Load()
    Option1(2) = True
    bCanUseNTDLL = True
    Randomize Timer
    Label1.Caption = "Simulation will create 2 arrays of equal size.  Depending on the check box below, the arrays will either be equal or not equal. When compiled && optimized, this routine is almost as fast as NTDLL.DLL's  RTLCompareMemory API."
End Sub

Private Function ByteAlignOnWord(ByVal bitDepth As Byte, ByVal Width As Long) As Long
    ' function to align byte range on dWord boundaries
    ByteAlignOnWord = (((Width * bitDepth) + &H1F) And Not &H1F&) \ &H8
End Function

