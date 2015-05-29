VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Buffer() As Byte
Private BufferSize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()

   On Error GoTo errorhandler

    Flush


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub PreAllocate(ByVal nLength As Long)

   On Error GoTo errorhandler

    WriteHead = 0
    ReadHead = 0
    BufferSize = nLength - 1
    ReDim Buffer(0 To BufferSize)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Allocate(ByVal nLength As Long)

   On Error GoTo errorhandler

    If BufferSize = 0 And nLength > 1 Then nLength = nLength - 1
    BufferSize = BufferSize + nLength
    ReDim Preserve Buffer(0 To BufferSize)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Allocate", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Flush()

   On Error GoTo errorhandler

    WriteHead = 0
    ReadHead = 0
    BufferSize = 0
    ReDim Buffer(0)


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Flush", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub Trim()
    ' If the readhead is past the buffersize, this means everything has been read in the array, flush it

   On Error GoTo errorhandler

    If ReadHead >= Count Then Flush


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "Trim", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WriteByte(ByVal nByte As Byte)


   On Error GoTo errorhandler

    If WriteHead > BufferSize Then Allocate 1
 
    Buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WriteByte", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
         

   On Error GoTo errorhandler

    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > BufferSize Then Allocate nLength
    
    CopyMemory Buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)


   On Error GoTo errorhandler

    If WriteHead + 1 > BufferSize Then Allocate 2
        
    CopyMemory Buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WriteInteger", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WriteLong(ByVal nLong As Long)


   On Error GoTo errorhandler

    If WriteHead + 3 > BufferSize Then Allocate 4
    
    CopyMemory Buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WriteLong", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    

   On Error GoTo errorhandler

    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > BufferSize Then Allocate sLength
      
    CopyMemory Buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength


   On Error GoTo 0
   Exit Sub
errorhandler:
    HandleError "WriteString", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte


   On Error GoTo errorhandler

    If ReadHead > BufferSize Then Exit Function
    
    ReadByte = Buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim data() As Byte


   On Error GoTo errorhandler

    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > BufferSize Then Exit Function
        
    ReDim data(nLength - 1)
        
    CopyMemory data(0), Buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = data


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer


   On Error GoTo errorhandler

    If ReadHead + 1 > BufferSize Then Exit Function
    
    CopyMemory ReadInteger, Buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ReadInteger", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long


   On Error GoTo errorhandler

    If ReadHead + 3 > BufferSize Then Exit Function
    
    CopyMemory ReadLong, Buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ReadLong", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
  

   On Error GoTo errorhandler

    sLength = ReadLong(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), Buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ReadString", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function Count() As Long

   On Error GoTo errorhandler

    Count = (UBound(Buffer) - LBound(Buffer)) + 1


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Count", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function Length() As Long

   On Error GoTo errorhandler

    Length = Count - ReadHead


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "Length", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ToArray() As Byte()

   On Error GoTo errorhandler

    ToArray = Buffer()


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ToArray", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function

Public Function ToString() As String

   On Error GoTo errorhandler

    ToString = StrConv(Buffer, vbUnicode)


   On Error GoTo 0
   Exit Function
errorhandler:
    HandleError "ToString", "clsBuffer", Err.Number, Err.Description, Erl
    Err.Clear
End Function
