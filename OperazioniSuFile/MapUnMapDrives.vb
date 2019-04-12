Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class MapUnMapDrives
    Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Integer) As Integer

    Private Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As Integer
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        Public lpLocalName As String
        Public lpRemoteName As String
        Public lpComment As String
        Public lpProvider As String
    End Structure

    Public Function RitornaLetteraLibera() As String
        Dim LetteraDisco As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "LetteraDisco", "")
        If LetteraDisco = "" Then
            Dim LetteraLibera As String = FindNextAvailableDriveLetter()
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\BackupNet", "LetteraDisco", LetteraLibera)
        End If

        Return LetteraDisco
    End Function

    Private Const RESOURCETYPE_DISK As Long = &H1
    Private Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long 'do we force the disconnect or not

    Private Sub Attendi(Secondi As Integer)
        Dim a As Single = Now.Second

        Do While Now.Second - a < Secondi Or a > Now.Second
            Application.DoEvents()
        Loop
    End Sub

    Public Function MappaDiscoDiRete(Percorso As String, sUtente As String, sPassword As String) As Boolean
        Dim Lettera As String = RitornaLetteraLibera()
        Dim sPercorso As String = Percorso

        If Mid(sPercorso, sPercorso.Length, 1) = "\" Then
            sPercorso = Mid(sPercorso, 1, sPercorso.Length - 1)
        End If

        'SMappaDiscoDiRete(Percorso, True)
        'SMappaDiscoDiRete(Lettera, False)

        Dim nr As NETRESOURCE
        Dim strUsername As String
        Dim strPassword As String

        nr = New NETRESOURCE
        nr.lpRemoteName = Percorso
        nr.lpLocalName = Lettera & ":"
        strUsername = sUtente
        strPassword = sPassword
        nr.dwType = RESOURCETYPE_DISK

        Dim result As Integer = 0

        Try
            result = WNetAddConnection2(nr, strPassword, strUsername, 1)
        Catch ex As Exception

        End Try

        If result = 0 Then
            Return True
        Else
            Return False
        End If

        Attendi(2)
    End Function

    <DllImport("mpr.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.Cdecl)> _
    Public Shared Sub SMappaDiscoDiRete()

    End Sub

    Public Function SMappaDiscoDiRete(Cosa As String, Percorso As Boolean) As Boolean
        Dim lngCancel As Long = 0
        Dim Altro As String = ":"

        If Percorso Then
            Altro = ""
        End If

        Try
            lngCancel = WNetCancelConnection(Cosa & altro, True)
        Catch ex As Exception

        End Try

        If lngCancel = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
