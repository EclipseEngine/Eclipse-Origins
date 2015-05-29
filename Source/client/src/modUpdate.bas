Attribute VB_Name = "modUpdate"


'Public Function DownloadFile(strURL As String, strDestination As String) As Boolean
Public Function DownloadFile(strURL As String, strDestination As String, Optional updatetext As String, Optional showpercent As Boolean = False) As Boolean
Const CHUNK_SIZE As Long = 1024
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim strHeader As String
Dim b() As Byte
Dim i As Integer
Dim DownloadProgress As Integer

DoEvents

DownloadFile = True

If Trim$(strURL) = "" Then DownloadFile = False: Exit Function

On Error GoTo ErrorHandler:
    
With frmMain.Inet
    
.URL = strURL
.Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
        
While .StillExecuting
DoEvents
Wend

strHeader = .GetHeader
End With
    
    
strHeader = frmMain.Inet.GetHeader("Content-Length")
lngFileLength = Val(strHeader)

DoEvents
    
lngBytesReceived = 0

intFile = FreeFile()

Open strDestination For Binary Access Write As #intFile

Do
b = frmMain.Inet.GetChunk(CHUNK_SIZE, icByteArray)
Put #intFile, , b
lngBytesReceived = lngBytesReceived + UBound(b, 1) + 1

DownloadProgress = (Round((lngBytesReceived / lngFileLength) * 100))
If showpercent = True Then
    frmLoad.lblStatus.Caption = updatetext & " - " & DownloadProgress & "% complete."
End If
DoEvents
Loop While UBound(b, 1) > 0

Close #intFile

Exit Function
ErrorHandler:
 Call HandleError("DownloadFile", "modUpdate", Err.Number, Err.Description, Erl)
 DownloadFile = False
End Function
