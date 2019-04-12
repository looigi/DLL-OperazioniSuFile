Imports System.Windows.Forms
Imports System.IO
Imports Log

Public Class GestioneACCESS
    Private Connessione As String

    Public Function ProvaConnessione() As String
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.Close()

            Conn = Nothing
            Return ""
        Catch ex As Exception

            Return ex.Message
        End Try
    End Function

    Public Function LeggeImpostazioniDiBase(ModalitaAutomatica As Boolean, PercorsoDBTemp As String, Chiave As String, Optional TipoDbPassato As Integer = -1) As Boolean
        Dim PercorsoDB As String
        Dim connectionString As String

        If ModalitaAutomatica Then
            PercorsoDB = PercorsoDBTemp
            connectionString = "Data Source=" & PercorsoDB & ";"
            Connessione = "Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;" & connectionString
        Else
            PercorsoDB = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", "")

            If PercorsoDB = "" Then
                Dim gf As New GestioneFilesDirectory
                gf.CreaAggiornaFile("D:\Errore.txt", "ERRORE: Path DB Non trovato")
                gf = Nothing
                Stop
            End If

            Select Case TipoDbPassato
                Case 1
                    DBSql = TipoDB.Access
                Case 2
                    DBSql = TipoDB.SQLCE
                Case -1
                    ImpostaDBSqlAccess()
            End Select

            Select Case DBSql
                Case TipoDB.Access
                    connectionString = "Data Source=" & PercorsoDB & "\DB\dbBackup.mdb;"
                    Connessione = "Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;" & connectionString
                Case TipoDB.SQLCE
                    connectionString = "Data Source=" & PercorsoDB & "\DB\dbBackup.sdf;"
                    Connessione = "Provider=Microsoft.SQLSERVER.CE.OLEDB.4.0;" & connectionString
                Case TipoDB.Niente
                    Stop
            End Select
        End If

        Return True
    End Function

    Public Function ApreDB(idProc As Integer, clLog As LogCasareccio.LogCasareccio.Logger) As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.CommandTimeout = 0
        Catch ex As Exception
            If Not clLog Is Nothing Then
                ScriveErrore(idProc, ex.Message, clLog)
            End If
        End Try

        Return Conn
    End Function

    Private Sub ScriveErrore(idProc As Integer, Errore As String, cllog As LogCasareccio.LogCasareccio.Logger)
        If Not cllog Is Nothing Then
            ScriveLog(idProc, "ERRORE " & Errore, cllog)
        End If
    End Sub

    Private Function ControllaAperturaConnessione(idProc As Integer, ByRef Conn As Object, clLog As LogCasareccio.LogCasareccio.Logger) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB(idProc, clLog)
        End If

        Return Ritorno
    End Function

    Public Function EsegueSql(idProc As Integer, ByVal Conn As Object, ByVal Sql As String, clLog As LogCasareccio.LogCasareccio.Logger) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(idProc, Conn, clLog)
        Dim Ritorno As Boolean = True
        Dim Errore As String = ""

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
            If Not clLog Is Nothing Then
                ScriveErrore(idProc, "ERRORE in EsegueSQL: " & Sql & " -> " & ex.Message, clLog)
            End If
            Ritorno = False
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function EsegueSqlSenzaTRY(idProc As Integer, ByVal Conn As Object, ByVal Sql As String, clLog As LogCasareccio.LogCasareccio.Logger) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(idProc, Conn, clLog)
        Dim Ritorno As Boolean = True
        Dim Errore As String = ""

        ' Routine che esegue una query sul db senza controllo degli errori
        Conn.Execute(Sql)

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            If Not Conn Is Nothing Then
                Conn.Close()
            End If
        End If
    End Sub

    Public Function LeggeQuery(idProc As Integer, ByVal Conn As Object, ByVal Sql As String, clLog As LogCasareccio.LogCasareccio.Logger) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(idProc, Conn, clLog)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = Nothing

            If Not clLog Is Nothing Then
                ScriveErrore(idProc, "ERRORE in LeggeQuery:" & Sql & " -> " & ex.Message, clLog)
            End If
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function

    Public Sub CompattazioneDb()
        ' RICORDARSI DI INCLUDERE LE JRO... 
        ' Microsoft Jet and Replication Objects 2.6 Library
        ' per la compattazione e...
        ' RICORDARSI CHE LA LIBRERIA FUNZIONA SOLO X86

        Dim PercorsoDB As String

        PercorsoDB = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", "")

        If PercorsoDB = "" Then
            Dim gf As New GestioneFilesDirectory
            gf.CreaAggiornaFile("D:\Errore.txt", "ERRORE: Path DB Non trovato")
            gf = Nothing
            Stop
        End If

        Dim JRO As JRO.JetEngine
        JRO = New JRO.JetEngine

        Dim NomeDB As String = PercorsoDB & "\DB\dbBackup.mdb"
        Dim NomeCompactDB As String = PercorsoDB & "\DB\dbBackupCOMP.mdb"

        Dim source As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NomeDB & ";"
        Dim compact As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NomeCompactDB & ";"

        Try
            JRO.CompactDatabase(source, compact)

            System.IO.File.Delete(NomeDB)

            System.IO.File.Move(NomeCompactDB, NomeDB)
        Catch ex As Exception
            ' MsgBox("Problema nella compattazione del db:" & vbCrLf & vbCrLf & ex.Message, vbExclamation)
        End Try
    End Sub
End Class
