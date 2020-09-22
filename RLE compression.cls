VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================
'= Run Length Encoding Compression (RLE)                            =
'====================================================================
'=  This class implements basic 8bit RLE compression and 8bit       =
'= decompression. 8bit compression yields the best compression in   =
'= most cases, especially when compressing strings, as opposed to   =
'= 16 or 32 bit compressions.                                       =
'=                                                                  =
'=  To get more information on the RLE algorithm search on the web  =
'= or you can E-mail me if you have any queries regarding this      =
'= implementation.                                                  =
'=                                                                  =
'=  You may redistribute this code by itself as long as it is left  =
'= unmodified. If you do use this code in a commercial or private   =
'= application please give me credit for it. Do not call it your    =
'= own code! I am releasing this piece of code for the public for   =
'= people who wish to learn from it and use it. You are free to do  =
'= so. All I ask in return is that you not take credit for writing  =
'= it.                                                              =
'=                                                                  =
'= Created By   : Eyal Cinamon A.K.A AAA- (aaa_001@hotmail.com)     =
'= Created On   : 12/05/1999                                        =
'= Modified On  : 05/26/2000                                        =
'====================================================================
'= Constant declarations =
'=========================
Private Const MAX_LEN As Byte = &H7F            ' Max length of data segment
Private Const MAX_RUN_HEADER As Byte = &HFF     ' Header description of max runs
Private Const MAX_SEQ_HEADER As Byte = &H7F     ' Header description of max unencoded data
Private Const RUN As Byte = &H80                ' A run follows
Private Const SEQ As Byte = &H0                 ' A Sequence follows

'================
'= API Declares =
'================
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlFillMemory Lib "kernel32.dll" (dest As Any, ByVal numBytes As Long, ByVal c As Byte)

'====================================================================
'= CompressRLE - compresses an array using 8bit RLE compression.
'====================================================================
'=
'= Inputs:
'=  aIn()   -> Array of uncompressed data
'=  aOut()  -> Array that holds the compressed data. The function
'=             assumes it is large enough.
'=  lInSize -> Size of array, (number of bytes to compress)
'=
'= Output:
'=  The size of the output buffer
'=
'= Assumptions:
'=  1) Both arrays must be 0-indexed.
'=  2) The user is responsible for making sure that the aOut() array
'=     is large enough to contain the compressed data. With RLE
'=     compression it possible to have a negative compression, where
'=     the compressed data is actually larger than the original. This
'=     can occur when there are no runs or when there is a checker-
'=     board effect where short runs are followed by short sequences.
'=     The worse case scenerio produces a 4:3 ratio between the
'=     compressed and uncomressed data.
'====================================================================
Public Function CompressRLE(aIn() As Byte, aOut() As Byte, _
    ByVal lInSize As Long) As Long

    Dim bChar       As Integer  ' current character
    Dim bLastChar   As Integer  ' The last character
    Dim lCurrChar   As Long     ' The character count
    Dim lBufferSize As Long     ' The output buffer size

    Dim lSize       As Long
    Dim iState      As Integer  ' Are processing a run or a sequence?

    Dim aSeq(MAX_LEN) As Byte   ' Temp array for sequence

    lBufferSize = 0
    lSize = 0
    bChar = -1
    iState = SEQ    ' Assume sequence to begin

    ' Go through the entire array
    For lCurrChar = 0 To lInSize - 1
        ' Save the last character
        bLastChar = bChar

        ' Get the new character
        bChar = CInt(aIn(lCurrChar))

        If iState = SEQ Then
            ' Normal mode in a sequence
            If bChar = bLastChar Then
                ' seems like the last 2 characters produce a run...

                ' End the sequence and start a run
                lSize = lSize - 1
                
                If lSize Then
                    aOut(lBufferSize) = lSize
                    lBufferSize = lBufferSize + 1
                    
                    ' Copy the buffers
                    RtlMoveMemory aOut(lBufferSize), aSeq(0), lSize
                    lBufferSize = lBufferSize + lSize
                End If
                
                lSize = 2    ' 2 characters in the run
                iState = RUN ' We are in run mode now
            Else
                ' Continue the sequence
                aSeq(lSize) = CByte(bChar)
                lSize = lSize + 1

                If lSize = MAX_LEN Then
                    ' We've reached the max number of
                    ' allowed characters.
                    aOut(lBufferSize) = MAX_SEQ_HEADER
                    lBufferSize = lBufferSize + 1
                    
                    ' Copy the sequence to the buffer
                    RtlMoveMemory aOut(lBufferSize), _
                        aSeq(0), MAX_LEN

                    lBufferSize = lBufferSize + MAX_LEN
                    ' Start fresh
                    lSize = 0
                    bChar = -1
                End If
            End If

        Else
            '===================================
            '= We must be in the middle of a run
            '===================================
                        
            If bChar = bLastChar Then
                lSize = lSize + 1
                
                ' Continue the run
                If lSize = MAX_LEN Then
                    ' We've reached the max number of
                    ' characters allowed in the run
                    ' Write the run & the run character
                    aOut(lBufferSize) = MAX_RUN_HEADER
                    aOut(lBufferSize + 1) = CByte(bLastChar)

                    lBufferSize = lBufferSize + 2
                    ' Start fresh
                    bChar = -1
                    lSize = 0
                    iState = SEQ ' Assume sequence again
                End If
            Else
                ' The run is finished

                ' Write the run length and character
                aOut(lBufferSize) = lSize Or RUN
                aOut(lBufferSize + 1) = CByte(bLastChar)

                lBufferSize = lBufferSize + 2
                aSeq(0) = CByte(bChar) ' Write the character
                lSize = 1
                iState = SEQ ' Assume sequence again
            End If
        End If

    Next

    If lSize Then
        ' Finish a run/sequence
        If iState = SEQ Then
            aOut(lBufferSize) = lSize
            lBufferSize = lBufferSize + 1
            
            ' Copy the buffers
            RtlMoveMemory aOut(lBufferSize), aSeq(0), lSize
            lBufferSize = lBufferSize + lSize   ' Set the new buffer size
            
        Else
            aOut(lBufferSize) = lSize Or RUN
            aOut(lBufferSize + 1) = CByte(bChar)

            lBufferSize = lBufferSize + 2
        End If
    End If
    
    ' Return the size fo the buffer
    CompressRLE = lBufferSize
End Function

'====================================================================
'= DecompressRLE - Decompresses an array using 8bit RLE decompression
'====================================================================
'=
'= Inputs:
'=  aIn()   -> Array of compressed data
'=  aOut()  -> Array that holds the uncompressed data. The function
'=             assumes it is large enough.
'=  lInSize -> Size of array, (number of compressed bytes)
'=
'= Output:
'=  The size of the output buffer
'=
'= Assumptions:
'=  1) Both arrays must be 0-indexed.
'=  2) The user is responsible for making sure that the aOut() array
'=     is large enough to contain the compressed data. With RLE
'=     compression it possible to have a negative compression, where
'=     the compressed data is actually larger than the original. This
'=     can occur when there are no runs or when there is a checker-
'=     board effect where short runs are followed by short sequences.
'=     The worse case scenerio produces a 4:3 ratio between the
'=     compressed and uncomressed data.
'====================================================================
Public Function DecompressRLE(ByRef aIn() As Byte, ByRef aOut() As Byte, _
    ByVal lInSize As Long) As Long

    Dim lCurrIndex   As Long
    Dim bTemp        As Byte
    Dim j            As Long
    Dim lOutputPos   As Long

    lCurrIndex = 0
    lOutputPos = 0
    
    Do While lCurrIndex < lInSize
        bTemp = aIn(lCurrIndex)
        
        If (bTemp And RUN) Then
            ' There is a run... repeat the character
            j = bTemp And &H7F          ' Get the repeat count
            lCurrIndex = lCurrIndex + 1
            bTemp = aIn(lCurrIndex)   ' get repeat char

            ' expand the repeated character in the array
            RtlFillMemory aOut(lOutputPos), j, bTemp
            
        Else
            ' Process the sequence
            
            ' How big is it?
            j = bTemp
            
            ' Copy the sequence over
            RtlMoveMemory aOut(lOutputPos), aIn(lCurrIndex + 1), j
            lCurrIndex = lCurrIndex + j
        End If
        
        lOutputPos = lOutputPos + j
        lCurrIndex = lCurrIndex + 1
    Loop
    
    DecompressRLE = lOutputPos + 1
End Function
