VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRLE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run Length Encoding Demo By: AAA- (aaa_001@hotmail.com)"
   ClientHeight    =   1215
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Decompression Info "
      Height          =   1035
      Index           =   1
      Left            =   4065
      TabIndex        =   3
      Top             =   120
      Width           =   3930
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   3
         Left            =   2340
         TabIndex        =   9
         Top             =   585
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   2
         Left            =   2445
         TabIndex        =   8
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "File size before decompression:"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   315
         Width           =   2310
      End
      Begin VB.Label Label1 
         Caption         =   "File size after decompression:"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   585
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   75
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compression Info "
      Height          =   1035
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   3930
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   2325
         TabIndex        =   7
         Top             =   645
         Width           =   1515
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   0
         Left            =   2325
         TabIndex        =   6
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "File size after compression:"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   615
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "File size before compression:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   315
         Width           =   2175
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   2133.333
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   2133.333
      Y1              =   2
      Y2              =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileCompress 
         Caption         =   "&Compress file"
      End
      Begin VB.Menu mnuFileDecompress 
         Caption         =   "&Decompress file"
      End
      Begin VB.Menu mnuBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmRLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
'=  This form demostrates the use of the clsRLE class to perform 8bit
'= RLE compression.
'=
'=  Created By: Eyal Cinamon A.K.A (aaa_001@hotmail.com)
'=  Email me for questions or comments.
'====================================================================
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const CREATE_ALWAYS = 2
Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private RLE As clsRLE   ' The compressor object

Private Sub Form_Load()
    Set RLE = New clsRLE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RLE = Nothing
End Sub

Private Sub mnuFileCompress_Click()
    ' Choose file to compress, I put in minimal error checking to focus
    ' on the task at hand.
    Dim sNewFileName As String
    Dim InFile As Long, OutFile As Long ' File handles
    Dim lFileSize As Long, lRLEFileSize As Long
    
    Dim aFile() As Byte     ' uncompressed file
    Dim aFileRLE() As Byte  ' compressed file data
    
    Dim lTemp As Long
    
    On Error GoTo Err_Compress
    
    With dlg
        .ShowOpen
        sNewFileName = Left$(.FileName, Len(.FileName) - 3) & "RLE"
        
        ' Open the file to read its contents
        InFile = CreateFile(.FileName, GENERIC_READ, 0&, ByVal 0&, _
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
        
        OutFile = CreateFile(sNewFileName, GENERIC_WRITE, 0&, ByVal 0&, _
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0&)

        If InFile And OutFile Then  ' Did it work?
            ' Output size of file
            lFileSize = GetFileSize(InFile, 0&)
            lblInfo(0).Caption = CStr(lFileSize)
            
            ' Resize the arrays
            ReDim aFile(lFileSize - 1)
            ReDim aFileRLE(CLng((lFileSize - 1) * 1.33))   ' assume worse case
            
            ' Read in the file
            ReadFile InFile, aFile(0), lFileSize, lTemp, ByVal 0&
            
            ' Compress the file
            lRLEFileSize = RLE.CompressRLE(aFile, aFileRLE, lFileSize)
            
            ' Write out the result to the file
            WriteFile OutFile, lFileSize, 4, lTemp, ByVal 0&    ' size of original file
            WriteFile OutFile, aFileRLE(0), lRLEFileSize, lTemp, ByVal 0&
            
            ' Show output
            lblInfo(1).Caption = CStr(lRLEFileSize + 4)
            
            ' Done using the files
            CloseHandle InFile
            CloseHandle OutFile
            
            MsgBox "Saved compressed file as " & sNewFileName
            
        Else
            GoTo Err_Compress
        End If
    End With
    
    Exit Sub
    
Err_Compress:
    ' Close both files just in case
    CloseHandle InFile
    CloseHandle OutFile
    MsgBox "There was an error!" & vbCrLf & _
        Err.Number & ": " & Err.Description, vbExclamation, "error"
End Sub

Private Sub mnuFileDecompress_Click()
    ' choose a file to decompress
    Dim sNewFileName As String
    Dim InFile As Long, OutFile As Long ' File handles
    Dim lFileSize As Long
    
    Dim aFile() As Byte     ' uncompressed file
    Dim aFileRLE() As Byte  ' compressed file data
    
    Dim lTemp As Long
    
    On Error GoTo Err_Decompress
    
    With dlg
        .ShowOpen
        sNewFileName = App.Path & "\temp.dat"
        
        ' Open the file to read its contents
        InFile = CreateFile(.FileName, GENERIC_READ, 0&, ByVal 0&, _
            OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
        
        OutFile = CreateFile(sNewFileName, GENERIC_WRITE, 0&, ByVal 0&, _
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, ByVal 0&)

        If InFile And OutFile Then  ' Did it work?
            ' Get size of decompressed file & output result
            ReadFile InFile, lFileSize, 4&, lTemp, ByVal 0&
            lblInfo(3).Caption = CStr(lFileSize)
            
            ' Resize the arrays
            ReDim aFile(lFileSize - 1)
            lFileSize = GetFileSize(InFile, 0&)
            lblInfo(2).Caption = CStr(lFileSize)
            
            lFileSize = lFileSize - 4 ' -4 for 4byte header
            ReDim aFileRLE(lFileSize - 1)
            
            ' Read in the file
            ReadFile InFile, aFileRLE(0), lFileSize, lTemp, ByVal 0&
            
            ' Decompress the file
            lFileSize = RLE.DecompressRLE(aFileRLE, aFile, lFileSize)
            
            ' Write out the result to the file
            WriteFile OutFile, aFile(0), lFileSize, lTemp, ByVal 0&
            
            ' Done using the files
            CloseHandle InFile
            CloseHandle OutFile
            
            MsgBox "File was restored as " & App.Path & "\Temp.dat"
            
        Else
            GoTo Err_Decompress
        End If
    End With
    
    Exit Sub
    
Err_Decompress:
    ' Close both files just in case
    CloseHandle InFile
    CloseHandle OutFile
    MsgBox "There was an error!" & vbCrLf & _
        Err.Number & ": " & Err.Description, vbExclamation, "error"

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "This demo illustrates the use of the RLE compression" & _
        vbCrLf & "algorithm through an object." & vbCrLf & _
        vbCrLf & "This demo was created by AAA. Email me with" & _
        vbCrLf & "questions or comments at aaa_001@hotmail.com", _
        vbInformation, "About..."
End Sub
