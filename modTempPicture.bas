Attribute VB_Name = "modTempPicture"
Option Explicit

Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Const UNIQUE_NAME = &H0
Private Const MAX_FILENAME_LEN = 256

Public Function GetStringFromPic(p As IPictureDisp) As String
   Dim fn As String
   Dim fh As String
   Dim s As String
   
   fn = GetTempFileName
   Kill fn
   Call SavePicture(p, fn)
   
   s = String(FileLen(fn), 0)
   fh = FreeFile(0)
   Open fn For Binary As fh
   Get #fh, , s
   Close #fh
   
   Kill fn
   
   GetStringFromPic = s
End Function

Public Function GetPicFromString(s As String) As IPictureDisp
   Dim fn As String
   Dim fh As String
   Dim p As IPictureDisp
   
   fn = GetTempFileName
   Kill fn
   fh = FreeFile(0)
   Open fn For Binary As fh
   Put #fh, , s
   Close fh
   
   Set p = LoadPicture(fn)
   Kill fn
   
   Set GetPicFromString = p
End Function

Private Function GetTempFileName() As String
   Dim s As String
   Dim s2 As String
   
   s2 = GetTempPath & vbNullChar
   s = Space(Len(s2) + MAX_FILENAME_LEN)
   Call GetTempFileNameA(s2, App.EXEName, UNIQUE_NAME, s)
   GetTempFileName = Left$(s, InStr(s, vbNullChar) - 1)
End Function

'
'  Returns the path to the temp directory.
'
Private Function GetTempPath() As String
   Dim s As String
   Dim i As Integer
   i = GetTempPathA(0, "")
   s = Space(i)
   Call GetTempPathA(i, s)
   GetTempPath = Backslash(Left$(s, i - 1))
End Function

'
'  Adds a backslash if the string doesn't have one already.
'
Private Function Backslash(ByVal s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         Backslash = s + "\"
      Else
         Backslash = s
      End If
   Else
      Backslash = "\"
   End If
End Function


