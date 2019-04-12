Imports System.Web
Imports System.Configuration
Imports System.Data.OleDb

Public Class SQLSERVER
    Private ConnessioneSQL As String
    Private ModalitaLocale As Boolean = True

    Public Function ProvaConnessione(Connessione As String) As String
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(Connessione)
            Conn.Close()

            Conn = Nothing
            Return ""
        Catch ex As Exception
            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            'Dim StringaPassaggio As String

            'StringaPassaggio = "?Errore=Apertura DB"
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'StringaPassaggio = StringaPassaggio & "&Errore=" & ex.Message
            'H.Response.Redirect("Errore.aspx" & StringaPassaggio)

            Return ex.Message
        End Try
    End Function

    Public Function ImpostaConnessioneDirettamente(Connessione As String) As Boolean
        Dim Conn As String = Connessione
        Dim Ok As Boolean = True
        Dim Ritorno As String

        If Conn = "" Then
            ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=Impostazioni di connessione al DB non valide&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
            Ok = False
        Else
            Ritorno = ProvaConnessione(Conn)
            If Ritorno <> "" Then
                ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=" & Ritorno & "&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
                Ok = False
            Else
                ConnessioneSQL = Conn
            End If
            ' Impostazioni di base
        End If

        Return Ok
    End Function

    Public Function LeggeImpostazioniDiBase() As Boolean
        'Dim Ritorno As String
        'Dim Ok As Boolean = True
        'Dim CosaCercare As String
        'Dim Conn As String = ""

        'If ModalitaLocale = True Then
        '    CosaCercare = "SQLConnectionStringLOCALE"
        'Else
        '    CosaCercare = "SQLConnectionStringWEB"
        'End If

        '' Impostazioni di base
        'Dim ListaConnessioni As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

        'If ListaConnessioni.Count <> 0 Then
        '    ' Get the collection elements. 
        '    For Each Connessioni As ConnectionStringSettings In ListaConnessioni
        '        Dim Nome As String = Connessioni.Name
        '        Dim Provider As String = Connessioni.ProviderName
        '        Dim connectionString As String = Connessioni.ConnectionString

        '        If Nome = CosaCercare Then
        '            Conn = "Provider=" & Provider & ";" & connectionString

        '            Exit For
        '        End If
        '    Next
        'End If

        'If Conn = "" Then
        '    ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=Impostazioni di connessione al DB non valide&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
        '    Ok = False
        'Else
        '    Ritorno = ProvaConnessione(Conn)
        '    If Ritorno <> "" Then
        '        ' Response.Redirect("errore_ErroreImprevisto.aspx?Errore=" & Ritorno & "&Chiamante=" & Request.CurrentExecutionFilePath.ToUpper.Trim & "&Sql=")
        '        Ok = False
        '    Else
        '        ConnessioneSQL = Conn
        '    End If
        '    ' Impostazioni di base
        'End If

        'Return Ok
    End Function

    Public Function ApreDB() As Object
        ' Routine che apre il DB e vede se ci sono errori
        Dim Conn As Object = CreateObject("ADODB.Connection")

        Try
            Conn.Open(ConnessioneSQL)
            Conn.CommandTimeout = 0
        Catch ex As Exception
            Conn = Nothing
            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            'Dim StringaPassaggio As String

            'StringaPassaggio = "?Errore=Apertura DB"
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'StringaPassaggio = StringaPassaggio & "&Sql="
            'H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try

        Return Conn
    End Function

    Private Function ControllaAperturaConnessione(ByRef Conn As Object) As Boolean
        Dim Ritorno As Boolean = False

        If Conn Is Nothing Then
            Ritorno = True
            Conn = ApreDB(ConnessioneSQL)
        End If

        Return Ritorno
    End Function

    Public Function EsegueSql(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        ' Routine che esegue una query sul db
        Try
            Conn.Execute(Sql)
        Catch ex As Exception
            Ritorno = ex.Message
            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            'Dim StringaPassaggio As String

            'StringaPassaggio = "?Errore=Errore esecuzione query: " & Err.Description
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'StringaPassaggio = StringaPassaggio & "&Sql=" & Sql
            'H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function EsegueSqlSenzaTRY(ByVal Conn As Object, ByVal Sql As String) As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Ritorno As String = ""

        Conn.Execute(Sql)

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Private Sub ChiudeDB(ByVal TipoApertura As Boolean, ByRef Conn As Object)
        If TipoApertura = True Then
            Conn.Close()
        End If
    End Sub

    Public Function LeggeQuery(ByVal Conn As Object, ByVal Sql As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Try
            Rec.Open(Sql, Conn)
        Catch ex As Exception
            Rec = Nothing

            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance
            'Dim StringaPassaggio As String

            'StringaPassaggio = "?Errore=Errore query: " & Err.Description
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("idUtente")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'StringaPassaggio = StringaPassaggio & "&Sql=" & Sql
            'H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function

    Public Function LeggeQuerySenzaTRY(ByVal Conn As Object, ByVal Sql As String) As Object
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        Dim Rec As Object = CreateObject("ADODB.Recordset")

        Rec.Open(Sql, Conn)

        ChiudeDB(AperturaManuale, Conn)

        Return Rec
    End Function

    Public Function ControlloEsistenzaTabella(Conn As Object, NomeTabella As String) As Boolean
        Dim Ritorno As Boolean
        Dim Rec As Object = CreateObject("ADODB.Recordset")
        Dim Sql As String
        Dim AperturaManuale As Boolean = ControllaAperturaConnessione(Conn)
        
        Sql = "SELECT * " & _
            "FROM INFORMATION_SCHEMA.TABLES " & _
            "WHERE TABLE_SCHEMA = N'dbo' " & _
            "AND TABLE_NAME = N'" & NomeTabella & "'"
        Rec.Open(Sql, Conn)
        If Rec.Eof = True Then
            Ritorno = False
        Else
            Ritorno = True
        End If
        Rec.Close()

        ChiudeDB(AperturaManuale, Conn)

        Return Ritorno
    End Function

    Public Function RitornaNumeroColonne(Rec As Object) As Integer
        Dim Ritorno As Integer = Rec.Fields.Count

        Return Ritorno
    End Function

    Public Function RitornaInformazioniCampiRecordset(Rec As Object) As String()
        Dim Campi() As String = {}

        ReDim Campi(Rec.Fields.Count - 1)

        For i As Integer = 0 To Rec.Fields.Count - 1
            Campi(i) = Rec.Fields(i).Name & ";" & Rec.Fields(i).Type & ";"
        Next

        Return Campi
    End Function

    Public Function PrendeChiaviTabella(ConnSql As Object, NomeTabella As String) As String
        Dim Rec As Object = CreateObject("ADODB.Recordset")
        Dim Sql As String = "select COLUMN_NAME from information_schema.KEY_COLUMN_USAGE   where table_name = '" & NomeTabella & "' Order By ORDINAL_POSITION "
        Dim Ritorno As String = ""

        Rec = LeggeQuery(ConnSql, Sql)
        Do Until Rec.Eof
            Ritorno += Rec(0).Value & ";"

            Rec.MoveNext()
        Loop
        Rec.Close()
        If Ritorno <> "" Then
            Ritorno = Mid(Ritorno, 1, Ritorno.Length - 1)
        End If

        Return Ritorno
    End Function
End Class
