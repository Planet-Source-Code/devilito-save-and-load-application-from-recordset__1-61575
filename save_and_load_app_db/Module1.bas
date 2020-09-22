Attribute VB_Name = "Module1"
Option Explicit
Private Type BROWSEINFO
    hWndOwner As Long
    pidlRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Declare Sub CoTaskMemFree _
                Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat _
                Lib "kernel32" _
                Alias "lstrcatA" (ByVal lpString1 As String, _
                                  ByVal lpString2 As String) As Long

Private Declare Function SHBrowseForFolder _
                Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList _
                Lib "shell32" (ByVal pidList As Long, _
                               ByVal lpBuffer As String) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
                               
Public Function GetFolder(hWnd As Long, _
                          Optional Stitle As String = "Select a Directory:") As String
    Dim iNull As Integer, lpIDList As Long ', Lresult As Long
    Dim sPath As String, udtBI As BROWSEINFO

    With udtBI
        .hWndOwner = hWnd
        .lpszTitle = lstrcat(Stitle, "")
        .ulFlags = 1 Or &H40
    End With

    lpIDList = SHBrowseForFolder(udtBI)

    If lpIDList Then
        sPath = String$(260, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)

        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    GetFolder = sPath
    
End Function

