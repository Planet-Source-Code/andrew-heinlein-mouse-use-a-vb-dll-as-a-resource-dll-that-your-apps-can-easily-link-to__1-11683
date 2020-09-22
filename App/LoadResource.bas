Attribute VB_Name = "LoadResource"
'Put together by Mouse
'mouse@theblackhand.net
'www.theblackhand.net

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As Long, ByVal lpType As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadBitmap Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, ByVal lpBitmapName As Long) As Long
Private Declare Function LoadString Lib "user32" Alias "LoadStringA" (ByVal hInstance As Long, ByVal wID As Long, ByVal lpBuffer As String, ByVal nBufferMax As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function LoadAccelerators Lib "user32" Alias "LoadAcceleratorsA" (ByVal hInstance As Long, ByVal lpTableName As Long) As Long
Private Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function SetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Private Const SRCCOPY = &HCC0020
Private Const ONE_K = &H400
Private Const GCW_HCURSOR = (-12)
   
Private Type BITMAP '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Public Function LoadResourceToPicBox(ResourceFile As String, ResourceName As Long, ToPicBox As PictureBox) As Boolean
    Dim RET As Long, RET2 As Long, RET3 As Long, RET4 As Long
    Dim TEMP_BITMAP As BITMAP, TEMP_DC As Long
    
    Set ToPicBox.Picture = Nothing
    
    'Load the file with the .rsrc (DLL, EXE, OCX, etc.) section and return the Instance handle
    RET = LoadLibrary(ResourceFile)
    If RET = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    
    'Load the bitmap by resource name and return a pointer to it in memory
    RET2 = LoadBitmap(RET, ResourceName)
    If RET2 = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    
    'Get the bitmap that is now in memory and store the details to a neat struc
    RET3 = GetObject(RET2, Len(TEMP_BITMAP), TEMP_BITMAP)
    'Create a workspace in memory
    TEMP_DC = CreateCompatibleDC(ToPicBox.hdc)
    'move the bitmap into the workspace
    RET4 = SelectObject(TEMP_DC, RET2)
    'set up the recieving picture box (self explainitory i think)
    With ToPicBox
        .AutoRedraw = True
        .BorderStyle = 0
        .Width = TEMP_BITMAP.bmWidth * 15
        .Height = TEMP_BITMAP.bmHeight * 15
    End With
    'Use the all-mighty BitBlt to "paint" the bitmap stored in memory to the picbox
    BitBlt ToPicBox.hdc, 0, 0, TEMP_BITMAP.bmWidth, TEMP_BITMAP.bmHeight, TEMP_DC, 0, 0, SRCCOPY
    
    'return TRUE if successfull
    LoadResourceToPicBox = True
    
ERROR_H:
    'free up the memory we borrowed.
    'remember: Be nice to your memory... clean up after yourself before you leave. (OK that was cheezy)
    DeleteObject RET4
    DeleteDC TEMP_DC
    FreeLibrary RET
End Function

Public Function LoadResourceString(ResourceFile As String, StringResource As Long) As String
    Dim RET As Long, TEMP_BUFFER As String, TEMP_LEN As Long
    'Load the file with the .rsrc (DLL, EXE, OCX, etc.) section and return the Instance handle
    RET = LoadLibrary(ResourceFile)
    If RET = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    'Buffer a string (I set it to 1024 bytes since that is more than enough.)
    TEMP_BUFFER = String(ONE_K, Chr(0))
    'Load the string from the Instance handle and it retuns the actual real length
    'of the resource.
    TEMP_LEN = LoadString(RET, StringResource, TEMP_BUFFER, Len(TEMP_BUFFER))
    'if it returns ZERO for length then the resouce never existed:
    If TEMP_LEN = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    'return the real string:
    LoadResourceString = Left(TEMP_BUFFER, TEMP_LEN)
ERROR_H:
    'clean up =)
    FreeLibrary RET
End Function

'These features below DO work, but are not used in this example:
Public Function LoadResourceCursor(ResourceFile As String, CurserName As Long, Mee As Form) As Boolean
    Dim RET As Long, RET2 As Long
    
    RET = LoadLibrary(ResourceFile)
    If RET = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    RET2 = LoadCursor(RET, CurserName)
    If RET2 = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    SetClassWord Mee.hwnd, GCW_HCURSOR, RET2
    LoadResourceCursor = True
ERROR_H:
    FreeLibrary RET
End Function

Public Function LoadAccelTable(ResourceFile As String, TableName As Long) As Long
    Dim RET As Long
    RET = LoadLibrary(ResourceFile)
    If RET = 0 Then 'no such animal
        GoTo ERROR_H
    End If
    LoadAccelTable = LoadAccelerators(RET, TableName)
    
ERROR_H:
    FreeLibrary RET
End Function
