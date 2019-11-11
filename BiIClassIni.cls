Option Explicit

Private Declare PtrSafe Function ApiReadKey Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpkeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function ApiWriteKey Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpkeyName As String, ByVal lpDefault As String, ByVal lpFileName As String) As Long

Public Function ReadKey(file$, section$, key$) As String
    Dim value$: value = VBA.Space$(255)
    Dim ret&: ret = ApiReadKey(section$, key$, "", value$, VBA.Len(value$), file$)
    ReadKey = VBA.Left$(value$, ret&)
End Function

Public Function WriteKey(ByVal section$, ByVal key$, ByVal value$, file$)
    Dim ret&: ret = ApiWriteKey(section$, key$, value$, file$)
End Func
