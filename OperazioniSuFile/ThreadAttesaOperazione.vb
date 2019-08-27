Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms
Imports Ionic.Zip

Public Class ThreadAttesaOperazione
    Private trd As Thread
    Private instanceThread As Form
    Private idProcThread As Integer
    Private ModalitaServizioThread As Boolean
    Private clLogThread As LogCasareccio.LogCasareccio.Logger
    Private Ancora As Boolean
    Private FileDaControllare As String
    Private Log As StringBuilder
    Private lblOperazione As Label
    Private lblContatore As Label
    Private Operazione As String
    Private zipFile As ZipFile

    Public Sub EsegueControllo(i As Form, id As Integer, modServ As Boolean, clLog As LogCasareccio.LogCasareccio.Logger, FileDaC As String, L As StringBuilder,
                               lblOp As Label, lblCont As Label, Oper As String, Optional z As ZipFile = Nothing)
        instanceThread = i
        idProcThread = id
        ModalitaServizioThread = modServ
        clLogThread = clLog
        FileDaControllare = FileDaC
        Log = L
        lblOperazione = lblOp
        lblContatore = lblCont
        zipFile = z

        Operazione = Oper

        Ancora = True

		' ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione, ModalitaServizioThread, clLogThread, False)

		If zipFile Is Nothing Then
            trd = New Thread(AddressOf ControlloDimensioni)
            trd.IsBackground = True
            trd.Start()
        Else
            ContatoreUscita = 0

            trd = New Thread(AddressOf ControlloErrore)
            trd.IsBackground = True
            trd.Start()

            AddHandler(zipFile.SaveProgress), New EventHandler(Of SaveProgressEventArgs)(AddressOf zip_SaveProgress)
        End If
    End Sub

    Private ContatoreUscita As Integer

    Private Sub ControlloErrore()
        While Ancora
            Thread.Sleep(1000)

            ContatoreUscita += 1
            If ContatoreUscita = 10 Or BloccaTutto Or Skippa Then
                zipFile.Dispose()
                Ancora = False

                Thread.Sleep(1000)

                Try
                    Kill(FileDaControllare)
                Catch ex As Exception

                End Try
            End If
        End While

        trd.Abort()
    End Sub

    Private Attuale As Integer
    Private Totale As Integer

    Private Sub zip_SaveProgress(ByVal sender As Object, ByVal e As SaveProgressEventArgs)
        Dim gf As New GestioneFilesDirectory

        If e.EventType = ZipProgressEventType.Saving_Started Then
            ContatoreUscita = 0
            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": " & gf.TornaNomeFileDaPath(e.ArchiveName), ModalitaServizioThread, clLogThread, False)
        ElseIf e.EventType = ZipProgressEventType.Saving_BeforeWriteEntry Then
            ContatoreUscita = 0

            Dim ilunghezza As Integer = e.BytesTransferred ' * 100 / e.TotalBytesToTransfer
            Dim Lunghezza As String = ilunghezza.ToString.Trim

            If (Val(Lunghezza) > -1) Then
                If Val(Lunghezza) < 1024 Then
                    Lunghezza = gf.FormattaNumero(Val(Lunghezza), False) & " B."
                Else
                    If Val(Lunghezza) < 1024000 Then
                        Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024, False) & " Kb."
                    Else
                        If Val(Lunghezza) < 1024000000 Then
                            Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000, False) & " Mb."
                        Else
                            Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000000, False) & " Gb."
                        End If
                    End If
                End If
            Else
                Lunghezza = "0 B."
            End If

            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, "File: " & gf.TornaNomeFileDaPath(e.CurrentEntry.FileName) & " " & Lunghezza & vbCrLf & gf.FormattaNumero(e.EntriesSaved, False) & " / " & gf.FormattaNumero(e.EntriesTotal, False), ModalitaServizioThread, clLogThread, False)
            Attuale = e.EntriesSaved
            Totale = e.EntriesTotal
        ElseIf e.EventType = ZipProgressEventType.Saving_AfterCompileSelfExtractor Then
            ContatoreUscita = 0
            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": Attendere prego ", ModalitaServizioThread, clLogThread, False)
        ElseIf e.EventType = ZipProgressEventType.Saving_AfterSaveTempArchive Then
            ContatoreUscita = 0
            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": Creazione archivio temporaneo ", ModalitaServizioThread, clLogThread, False)
        ElseIf e.EventType = ZipProgressEventType.Saving_AfterRenameTempArchive Then
            ContatoreUscita = 0
            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": Rinomina file", ModalitaServizioThread, clLogThread, False)
        ElseIf e.EventType = ZipProgressEventType.Saving_EntryBytesRead Then
            ContatoreUscita = 0

            Dim ilunghezza As Integer = e.BytesTransferred ' * 100 / e.TotalBytesToTransfer
            Dim Lunghezza As String = ilunghezza.ToString.Trim

            If (Val(Lunghezza) > -1) Then
                If Val(Lunghezza) < 1024 Then
                    Lunghezza = gf.FormattaNumero(Val(Lunghezza), False) & " B."
                Else
                    If Val(Lunghezza) < 1024000 Then
                        Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024, False) & " Kb."
                    Else
                        If Val(Lunghezza) < 1024000000 Then
                            Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000, False) & " Mb."
                        Else
                            Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000000, False) & " Gb."
                        End If
                    End If
                End If
            Else
                Lunghezza = "0 B."
            End If

            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, "File: " & gf.TornaNomeFileDaPath(e.CurrentEntry.FileName) & " " & Lunghezza & vbCrLf & Attuale & "/" & Totale, ModalitaServizioThread, clLogThread, False)
        ElseIf e.EventType = ZipProgressEventType.Saving_Completed Then
            ContatoreUscita = 0

            ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": Opreazione effettuata", ModalitaServizioThread, clLogThread, False)

            Ancora = False
        End If
        gf = Nothing
    End Sub

    Public Sub Blocca(lblOper As Label)
        Ancora = False

        ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, "", ModalitaServizioThread, clLogThread, False)
    End Sub

    Private Sub ControlloDimensioni()
        Dim Contatore As Integer = 0

        While Ancora
            If File.Exists(FileDaControllare) Then
                Dim gf As New GestioneFilesDirectory
                Dim Lunghezza As String = gf.TornaDimensioneFile(FileDaControllare)

                If (Val(Lunghezza) > -1) Then
                    If Val(Lunghezza) < 1024 Then
                        Lunghezza = gf.FormattaNumero(Val(Lunghezza), False) & " B."
                    Else
                        If Val(Lunghezza) < 1024000 Then
                            Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024, False) & " Kb."
                        Else
                            If Val(Lunghezza) < 1024000000 Then
                                Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000, False) & " Mb."
                            Else
                                Lunghezza = gf.FormattaNumero(Val(Lunghezza) / 1024000000, False) & " Gb."
                            End If
                        End If
                    End If
                Else
                    Lunghezza = "0 B."
                End If
                gf = Nothing

                Try
                    ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, Operazione & ": " & Lunghezza & " (Secondi impiegati: " & Contatore & ")", ModalitaServizioThread, clLogThread, False)
                Catch ex As Exception

                End Try
            End If
            Contatore += 1

            Thread.Sleep(1000)
        End While

        ScriveOperazione(instanceThread, True, idProcThread, Log, lblOperazione, lblContatore, lblOperazione.Text, "", ModalitaServizioThread, clLogThread, False)

        instanceThread = Nothing
        idProcThread = -1
        clLogThread = Nothing

        trd.Abort()
    End Sub

End Class
