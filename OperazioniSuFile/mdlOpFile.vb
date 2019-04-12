Imports System.Text
Imports System.Windows.Forms
Imports System.ServiceProcess
Imports System.Collections.Specialized
Imports System.IO
Imports frmLog

Module mdlOpFile
    Public MetteInPausa As Boolean = False
    Public BloccaTutto As Boolean = False
    Public Skippa As Boolean = False

    Private Contatore1 As Integer = 0
    Private Contatore2 As Integer = 0

    Private ContatoreLog As Long = -1
    Private dbGlobale As GestioneACCESS = Nothing
    Private ConnSqlGlobale As Object = Nothing

    Public Enum TipoDB
        Niente = -1
        Access = 1
        SQLCE = 2
    End Enum

    Public DBSql As Integer = TipoDB.Niente

    Public Sub MetteInPausaLaRoutine()
        While MetteInPausa
            Threading.Thread.Sleep(1000)

            Application.DoEvents()
        End While
    End Sub

    Public Sub ImpostaDBSqlAccess()
        Dim PercorsoDB As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", "")
        If Directory.Exists(PercorsoDB) = False Then
            PercorsoDB = ""
        End If
        If PercorsoDB = "" Then
            My.Computer.Registry.CurrentUser.CreateSubKey("Software\BackupNet")
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\BackupNet", "PathDB", Application.StartupPath)
            PercorsoDB = Application.StartupPath
        End If

        If File.Exists(PercorsoDB & "\DB\dbBackup.sdf") Then
            DBSql = TipoDB.SQLCE
        Else
            DBSql = TipoDB.Access
        End If
    End Sub

    Public Sub Attendi(ByVal gapToWait As Integer)
        Dim o As Integer = Now.Second
        Dim Diff As Integer
        Dim Ancora As Boolean = True

        Do While Ancora = True
            Diff = (Now.Second - o)
            If Diff < 0 Then
                Diff = 60 - Math.Abs(Diff)
            End If

            If Diff > gapToWait Then
                Ancora = False
            Else
            End If

            If BloccaTutto Then
                Exit Do
            End If

            Application.DoEvents()
        Loop
    End Sub

    Public Function FermaServizio(idProc As Integer, Nome As String, clLog As LogCasareccio.LogCasareccio.Logger) As Boolean
        ScriveLog(idProc, "FERMA SERVIZIO:", clLog)

        Dim myController As ServiceController
        Dim Ok As Boolean = True

        myController = New ServiceController(Nome)
        If myController.CanStop Then
            Try
                myController.Stop()
                ScriveLog(idProc, "SERVIZIO FERMATO", clLog)
            Catch ex As Exception
                Ok = False
                ScriveLog(idProc, "ERRORE FERMA SERVIZIO: " & ex.Message, clLog)
            End Try
        Else
            Ok = False
            ScriveLog(idProc, "ERRORE: SERVIZIO NON STOPPABILE", clLog)
        End If

        Return Ok
    End Function

    Public Function FaiPartireServizio(idProc As Integer, Nome As String, Parametro As String, clLog As LogCasareccio.LogCasareccio.Logger) As Boolean
        ScriveLog(idProc, "AVVIA SERVIZIO: " & Nome, clLog)

        Dim myController As ServiceController
        Dim Ok As Boolean = True

        myController = New ServiceController(Nome)
        Try
            myController.Start()
        Catch exp As Exception
            Ok = False
            ScriveLog(idProc, "ERRORE AVVIA SERVIZIO: " & exp.Message, clLog)
        Finally
            ScriveLog(idProc, "SERVIZIO AVVIATO", clLog)
        End Try

        Return Ok
    End Function

    Private Function PrendeDataOra() As String
        Dim Ritorno As String = Format(Now.Day, "00") & "/" & Format(Now.Month, "00") & "/" & Now.Year & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

        Return Ritorno
    End Function

    Private lblOperazione As Label
    Private lblContatore As Label

    Private Sub AddTextOperazione(ByVal str As String)
        lblOperazione.Text = str
    End Sub

    Private Sub AddTextContatore(ByVal str As String)
        lblContatore.Text = str
    End Sub

    Private Delegate Sub DelegateAddTextOperazione(ByVal str As String)
    Private MethodDelegateAddTextOperazione As New DelegateAddTextOperazione(AddressOf AddTextOperazione)

    Private Delegate Sub DelegateAddTextContatore(ByVal str As String)
    Private MethodDelegateAddTextContatore As New DelegateAddTextContatore(AddressOf AddTextContatore)

    Public Sub ScriveOperazione(instance As Form, Subito As Boolean, idProc As Integer, ByRef log As StringBuilder, lblOper As Label,
                                lblCont As Label, Primo As String, Secondo As String, ModalitaServizio As Boolean, clLog As LogCasareccio.LogCasareccio.Logger,
                                Optional ScriveSuFile As Boolean = True)
        If lblOperazione Is Nothing Then
            If Not lblOper Is Nothing Then
                lblOperazione = lblOper
            End If
        End If
        If lblContatore Is Nothing Then
            If Not lblCont Is Nothing Then
                lblContatore = lblCont
            End If
        End If

        Dim Ok As Boolean = True

        If Primo <> "" Then
            If lblOperazione Is Nothing = False Then
                If Not Subito Then
                    Ok = False
                    Contatore1 += 1
                    If Contatore1 >= 50 Then
                        Contatore1 = 0
                        Ok = True
                    End If
                End If
                If Ok Then
                    'lblOperazione.Text = Primo
                    If instance.InvokeRequired Then
                        instance.Invoke(MethodDelegateAddTextOperazione, Primo)
                    Else
                        lblOperazione.Text = Primo
                    End If
                End If
            End If
        End If

        If Secondo <> "" Then
            If lblContatore Is Nothing = False Then
                If Not Subito Then
                    Ok = False
                    Contatore2 += 1
                    If Contatore2 >= 10 Then
                        Contatore2 = 0
                        Ok = True
                    End If
                End If
                If Ok Then
                    'lblContatore.Text = Secondo
                    If instance.InvokeRequired Then
                        instance.Invoke(MethodDelegateAddTextContatore, Secondo)
                    Else
                        lblContatore.Text = Secondo
                    End If
                End If
            End If
        End If

        If ScriveSuFile And (Primo <> "" Or Secondo <> "") Then
            ScriveLog(idProc, Primo.Replace(vbCrLf, " ") & ";" & Secondo.Replace(vbCrLf, " ") & ";", clLog)
        End If

        If ModalitaServizio Then
            'Dim opFiles As New OperazioniSuFile
            clLog.ScriveLogServizio(Primo.Replace(vbCrLf, " ") & ";" & Secondo.Replace(vbCrLf, " ") & ";")
            'opFiles = Nothing
        End If

        If ScriveSuFile And (Primo <> "" Or Secondo <> "") Then
            log.Append(PrendeDataOra() & ";" & Primo.Replace(vbCrLf, " ") & ";" & Secondo.Replace(vbCrLf, " ") & ";" & vbCrLf)
            Application.DoEvents()
        End If

        'If Subito = True Then
        '    Contatore1 = 50
        '    Contatore2 = 10
        'End If
    End Sub

    Private Function PrendeDataOraPerScritturaSuDB() As String
        Dim Ritorno As String = Format(Now.Year, "00") & "-" & Format(Now.Month, "00") & "-" & Now.Day & " " & Format(Now.Hour, "00") & ":" & Format(Now.Minute, "00") & ":" & Format(Now.Second, "00")

        Return Ritorno
    End Function

    Public Sub ChiudeDBGlobale()
        dbGlobale = Nothing
    End Sub

    Private Sub PrendeMaxLog(idproc As Integer)
        'Dim Rec As Object = CreateObject("ADODB.Recordset")
        'Dim Sql As String
        'Dim Cont As Long

        'Sql = "Select Max(Progressivo) From LogOperazioni Where idProc=" & idproc
        'Rec = dbGlobale.LeggeQuery(idproc, ConnSqlGlobale, Sql)
        'If Rec(0).Value Is DBNull.Value = True Then
        '    Cont = 1
        'Else
        '    Cont = Rec(0).Value
        'End If
        'Rec.Close()

        'ContatoreLog = Cont
    End Sub

    Public Sub ScriveLog(idProc As Integer, Log As String, clLog As LogCasareccio.LogCasareccio.Logger)
        'Dim o As OperazioniSuFile = New OperazioniSuFile
        clLog.ScriveLogServizio(Log)

        'Try
        '    Dim GiaAperto As Boolean = True

        '    If dbGlobale Is Nothing Then
        '        GiaAperto = False
        '        dbGlobale = New GestioneACCESS

        '        If dbGlobale.LeggeImpostazioniDiBase("ConnDB") = True Then
        '            ConnSqlGlobale = dbGlobale.ApreDB(idProc)
        '        End If
        '    End If

        '    Dim Rec As Object = CreateObject("ADODB.Recordset")
        '    Dim Sql As String

        '    If ContatoreLog = -1 Then
        '        PrendeMaxLog(idProc)
        '    End If
        '    ContatoreLog += 1

        '    Sql = "Insert Into LogOperazioni Values (" &
        '        " " & idProc & ", " &
        '        " " & ContatoreLog & ", " &
        '        "'" & Log.Replace("'", "''") & "', " &
        '        "'" & PrendeDataOraPerScritturaSuDB() & "' " &
        '        ")"
        '    Try
        '        dbGlobale.EsegueSqlSenzaTRY(idProc, ConnSqlGlobale, Sql)
        '    Catch ex As Exception

        '    End Try

        '    If Not GiaAperto Then
        '        ConnSqlGlobale.Close()
        '        ChiudeDBGlobale()
        '    End If
        'Catch ex As Exception

        'End Try
    End Sub

    Public Function FindNextAvailableDriveLetter() As String
        ' build a string collection representing the alphabet
        Dim alphabet As New StringCollection()

        Dim lowerBound As Integer = Convert.ToInt16("a"c)
        Dim upperBound As Integer = Convert.ToInt16("z"c)
        For i As Integer = lowerBound To upperBound - 1
            Dim driveLetter As Char = ChrW(i)
            alphabet.Add(driveLetter.ToString())
        Next

        ' get all current drives
        Dim drives As DriveInfo() = DriveInfo.GetDrives()
        alphabet.Remove("a")
        alphabet.Remove("b")
        For Each drive As DriveInfo In drives
            alphabet.Remove(drive.Name.Substring(0, 1).ToLower())
        Next

        If alphabet.Count > 0 Then
            Return alphabet(0).ToUpper
        Else
            Return ""
        End If
    End Function

    Public Function ExistsDriveLetter(Lettera As String) As Boolean
        Dim Ok As Boolean = False
        Dim drives As DriveInfo() = DriveInfo.GetDrives()

        For Each drive As DriveInfo In drives
            If drive.Name.Substring(0, 1).ToUpper = Lettera.ToUpper.Trim Then
                Ok = True
                Exit For
            End If
        Next

        Return Ok
    End Function

End Module
