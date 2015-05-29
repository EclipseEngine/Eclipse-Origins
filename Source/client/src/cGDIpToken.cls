VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGDIpToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Credit for these GDI+ classes go to LaVolpe
'http://www.vbforums.com/showthread.php?t=598771
' HOW TO USE THIS CLASS
' This class must be created first before any other GDI+ classes call their functions
'   Many of the other GDI+ class' functions/methods expect this class as a parameter.
' Only ONE instance of this class needs to be created and should exist until your close your project
' Recommend making this public in your main form or in a module
'   :: Public cToken As cGDIpToken
'   :: If declared in a form, to access in other forms: MainFormName.cToken
' This is "create and forget" type class. You do not call any functions from this class at all

Option Explicit

Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)

Private m_Token As Long
Private m_Users As Collection

Public Property Get Token() As Long
    Token = m_Token
End Property

Public Sub AddUser(User As Object)
    ' DO NOT CALL THIS FROM YOUR PROJECT
    If m_Users Is Nothing Then Set m_Users = New Collection
    If Not User Is Nothing Then m_Users.Add ObjPtr(User), CStr(ObjPtr(User))
End Sub
Public Sub RemoveUser(User As Object)
    ' DO NOT CALL THIS FROM YOUR PROJECT
    If Not m_Users Is Nothing Then
        If Not User Is Nothing Then m_Users.Remove CStr(ObjPtr(User))
    End If
End Sub

Private Sub Class_Initialize()
    Dim GSI As GdiplusStartupInput
    
    On Error Resume Next
    If m_Token = 0& Then
        GSI.GdiplusVersion = 1&
        Call GdiplusStartup(m_Token, GSI)
    End If

End Sub

Private Sub Class_Terminate()
    If m_Token Then
        Dim u As Long, o As Object, oTmp As Object, oPtr As Long
        If Not m_Users Is Nothing Then
            For u = m_Users.count To 1 Step -1
                oPtr = m_Users(u)
                CopyMemory oTmp, oPtr, 4&
                Set o = oTmp
                CopyMemory oTmp, 0&, 4&
                o.Clear ' classes added to this collection have a .Clear method
            Next
            Set m_Users = Nothing
        End If
        GdiplusShutdown m_Token
        m_Token = 0&
    End If
End Sub
