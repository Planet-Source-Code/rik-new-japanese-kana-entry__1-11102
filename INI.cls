VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Declarations
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' Module wide variables
Private INIFileName As String
Private Ret As String
'

Public Function ReadINI(Section As String, Key As String) As Variant
Ret = Space$(255)
RetLen = GetPrivateProfileString(Section, Key, "", Ret, Len(Ret), INIFileName)
Ret = Left$(Ret, RetLen)
ReadINI = Ret

End Function

Public Sub WriteINI(Section As String, Key As String, Text As String)
WritePrivateProfileString Section, Key, Text, INIFileName
End Sub

Public Property Get inifile() As String
    inifile = INIFileName
End Property

Public Property Let inifile(ByVal vNewValue As String)
    INIFileName = vNewValue
End Property


