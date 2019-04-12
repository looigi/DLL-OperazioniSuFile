Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports Ionic.Zip
Imports System.Threading

Public Class OperazioniSuFileDettagli
    Private FileOrigine As String
    Private FileDestinazione As String
    Private gf As GestioneFilesDirectory
    Private log As StringBuilder = New StringBuilder
    Private LunghezzaMassimaScritte As Integer = 70

    Sub New()
        gf = New GestioneFilesDirectory
    End Sub

    Protected Overrides Sub Finalize()
        gf = Nothing

        MyBase.Finalize()
    End Sub

    Public Function fAttendi(idProc As Integer, Origine As String, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "ATTENDI: " & Origine & " secondi", clLog)

        log = New StringBuilder

        Attendi(Origine)

        ScriveLog(idProc, "USCITA ATTENDI", clLog)

        Return log
    End Function

    Public Function AvviaEseguibile(idProc As Integer, Origine As String, Parametro As String, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "AVVIA ESEGUIBILE: " & Origine, clLog)

        log = New StringBuilder

        'Dim sparametro As String = Origine
        'If sparametro.Contains(" ") Then
        '    sparametro = Mid(sparametro, sparametro.IndexOf(" ") + 1, sparametro.Length).Trim
        '    Origine = Origine.Replace(sparametro, "").Trim
        'Else
        '    sparametro = ""
        'End If

        Try
            If Parametro = "" Then
                Process.Start(Origine)
            Else
                Process.Start(Origine, Parametro)
            End If
        Catch ex As Exception
            ScriveLog(idProc, "AVVIA ESEGUIBILE: Errore " & ex.Message, clLog)
        End Try

        gf.CreaAggiornaFile("PassaggioService.Dat", "AVVIO APPLICAZIONE;" & Origine & ";")

        ScriveLog(idProc, "USCITA AVVIA ESEGUIBILE", clLog)

        Return log
    End Function

    Public Function FermaEseguibile(idProc As Integer, Origine As String, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "STOPPA ESEGUIBILE: " & Origine, clLog)

        log = New StringBuilder

        Try
            FermaServizio(idProc, Origine, clLog)

            gf.CreaAggiornaFile("PassaggioService.Dat", "CHIUSURA APPLICAZIONE;" & Origine & ";")

            ScriveLog(idProc, "USCITA STOPPA ESEGUIBILE", clLog)
        Catch ex As Exception
            log.Append("ERRORE: " & ex.Message)
        End Try

        Return log
    End Function

    Public Function RiavvioPC(idProc As Integer, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "RIAVVIO PC:", clLog)

        log = New StringBuilder

        Try
            FaiPartireServizio(idProc, "BackupLauncher.NET", "", clLog)

            gf.CreaAggiornaFile("PassaggioService.Dat", "RIAVVIO;")

            ScriveLog(idProc, "USCITA RIAVVIO PC", clLog)
        Catch ex As Exception
            log.Append("ERRORE: " & ex.Message)
        End Try

        Return log
    End Function

    Public Function EliminaFile(instance As Form, idProc As Integer, Origine As String, Filtro As String, SottoDirectory As String, lblOperazione As Label,
                                lblContatore As Label, ModalitaServizio As Boolean, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "ELIMINAZIONE FILE:", clLog)

        log = New StringBuilder

        Try
            If SottoDirectory = "S" Then
                gf.ScansionaDirectorySingola(Origine, instance, Filtro, lblOperazione, False)
            Else
                gf.ScansionaDirectorySingola(Origine, instance, Filtro, lblOperazione, True)
            End If

            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    FileOrigine = Filetti(i)

                    gf.EliminaFileFisico(FileOrigine)
                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Elimina file " & gf.TornaNomeFileDaPath(Filetti(i)), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False), ModalitaServizio, clLog)
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function EliminazioneDirectory(instance As Form, idProc As Integer, Origine As String,
                                          lblOperazione As Label, lblContatore As Label,
                                          ModalitaServizio As Boolean, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "ELIMINAZIONE DIRECTORY:", clLog)

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, instance, "", lblOperazione, False)

            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati
            Dim Cartelle() As String = gf.RitornaDirectoryRilevate
            Dim qCartelle As Long = gf.RitornaQuanteDirectoryRilevate

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    'FileDestinazione = Origine & "\" & gf.TornaNomeFileDaPath(Filetti(i))

                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione file in directory " & gf.TornaNomeFileDaPath(Filetti(i)), i & "/" & qFiletti, ModalitaServizio, clLog)

                    gf.EliminaFileFisico(Filetti(i))
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            If Not BloccaTutto Then
                For k As Integer = 1 To 3
                    For i As Long = qCartelle To 0 Step -1
                        If Cartelle(i) <> "" Then
                            Try
                                RmDir(Cartelle(i))
                                ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione directory " & Cartelle(i), i & "/" & qCartelle, ModalitaServizio, clLog)
                            Catch ex As Exception
                            End Try
                        End If

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If BloccaTutto Then
                            Exit For
                        End If
                    Next

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    If BloccaTutto Then
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function CreaDirectory(instance As Form, idProc As Integer, Origine As String, Filtro As String, lblOperazione As Label, lblContatore As Label,
                                  ModalitaServizio As Boolean, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "CREAZIONE DIRECTORY:", clLog)

        log = New StringBuilder

        Try
            FileOrigine = Origine & "\" & Filtro

            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione directory " & FileOrigine, " ", ModalitaServizio, clLog)

            gf.CreaDirectoryDaPercorso(FileOrigine & "\")
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function Copia(idProc As Integer, Operazione As Integer, Origine As String, Destinazione As String, qFiletti As Long,
                          Filetti() As String, Sovrascrivi As String, lblOperazione As Label, lblContatore As Label, ModalitaServizio As Boolean,
                          instance As Form, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "SPOSTAMENTO FILES: " & Origine & "->" & Destinazione, clLog)

        log = New StringBuilder

        Try
            Dim PathUlteriore As String = ""

            gf.ScansionaDirectorySingola(Origine, instance, "", lblOperazione, False)

            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

            Filetti = gf.RitornaFilesRilevati
            qFiletti = gf.RitornaQuantiFilesRilevati

            For i As Long = 0 To qFiletti
                If Filetti(i) <> "" Then
                    FileOrigine = Filetti(i)
                    PathUlteriore = FileOrigine.Replace(Origine & "\", "")
                    FileDestinazione = Destinazione & "\" & PathUlteriore ' & "\" & gf.TornaNomeFileDaPath(Filetti(i))

                    gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(FileDestinazione) & "\")

                    Dim t As New ThreadAttesaOperazione
                    t.EsegueControllo(instance, idProc, ModalitaServizio, clLog, FileDestinazione, log, lblOperazione, lblContatore, "Copia file " & gf.TornaNomeFileDaPath(FileOrigine))

                    gf.CopiaFileFisico(FileOrigine, FileDestinazione, IIf(Sovrascrivi = "S", True, False))

                    t.Blocca(lblOperazione)

                    If Operazione = OperazioniSuFile.TipoOperazione.Spostamento Then
                        gf.EliminaFileFisico(FileOrigine)
                        ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Sposta file " & gf.TornaNomeFileDaPath(Filetti(i)), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False), ModalitaServizio, clLog)
                    Else
                        ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Copia file " & gf.TornaNomeFileDaPath(Filetti(i)), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False), ModalitaServizio, clLog)
                    End If
                End If

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If
            Next
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function Sincronizza(idProc As Integer, Progressivo As Integer, Origine As String, Destinazione As String, Filtro As String,
                                lblOper As Label, lblCont As Label, Intelligente As Boolean, ModalitaServizio As Boolean, instance As Form,
                                clLog As LogCasareccio.LogCasareccio.Logger, ModalitaEsecuzioneAutomatica As Boolean, PercorsoDBTemp As String) As StringBuilder
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

        If Intelligente Then
            ScriveLog(idProc, "SINCRONIZZAZIONE INTELLIGENTE DIRECTORY: " & Origine & "->" & Destinazione, clLog)
        Else
            ScriveLog(idProc, "SINCRONIZZAZIONE DIRECTORY: " & Origine & "->" & Destinazione, clLog)
        End If

        log = New StringBuilder

        'Try
        Dim DB As New GestioneACCESS
        Dim ConnSQL As Object = Nothing

        If DB.LeggeImpostazioniDiBase(ModalitaEsecuzioneAutomatica, PercorsoDBTemp, "ConnDB") = True Then
            ConnSQL = DB.ApreDB(idProc, clLog)

            Dim Rec2 As Object = CreateObject("ADODB.Recordset")
            Dim Sql As String
            Dim CartelleDest() As String = {}
            Dim qCartelleDest As Integer
            Dim LeggiFiles As Boolean

            Dim Filetti() As String = {}
            Dim DimensioneFiletti() As Long = {}
            Dim DataFiletti() As Date = {}
            Dim qFiletti As Long
            Dim CartelleOrig() As String = {}
            Dim qCartelleOrig As Long
            Dim Altro As String
            Dim Datella As Date
            Dim DatellaFile As String
            Dim Massimo As Long = 0
            Dim AggiornataTabellaIntelligente As Boolean

            Dim FilesDaElaborare As Collection
            Dim Dimens As Long
            Dim q As Long
            Dim Tipo As String = ""
            Dim opFiles As New OperazioniSuFile

            ' Lettura origine
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Lettura directory di origine " & vbCrLf & Origine, " ", ModalitaServizio, clLog)

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            If Not BloccaTutto Then
                gf.ScansionaDirectorySingola(Origine, instance, Filtro, lblOperazione, False)

                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If Not BloccaTutto Then
                    Filetti = gf.RitornaFilesRilevati
                    DimensioneFiletti = gf.RitornaDimensioneFilesRilevati
                    DataFiletti = gf.RitornaDataFilesRilevati
                    qFiletti = gf.RitornaQuantiFilesRilevati

                    CartelleOrig = gf.RitornaDirectoryRilevate
                    qCartelleOrig = gf.RitornaQuanteDirectoryRilevate
                End If
            End If

            If MetteInPausa Then
                MetteInPausaLaRoutine()
            End If

            If Not BloccaTutto Then
                Sql = "Delete * From FilesOrigine"
                DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Scrittura dati directory" & vbCrLf & Origine, " ", ModalitaServizio, clLog)
                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Numero Cartelle: " & gf.FormattaNumero(qCartelleOrig, False) & vbCrLf & "Numero files: " & gf.FormattaNumero(qFiletti, False), " ", ModalitaServizio, clLog)

                If Mid(Origine, Origine.Length, 1) = "\" Then
                    Altro = ""
                Else
                    Altro = "\"
                End If

                For i As Long = 0 To qFiletti
                    If Filetti(i) <> "" Then
                        Try
                            Datella = DataFiletti(i)
                            DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                        Catch ex As Exception
                            DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                        End Try

                        Sql = "Insert Into FilesOrigine Values (" &
                                "'" & Filetti(i).Replace("'", "''").Replace(Origine & Altro, "") & "', " &
                                " " & DimensioneFiletti(i) & ", " &
                                "'" & DatellaFile & "' " &
                                ")"
                        DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                        If IsNothing(lblContatore) = False Then
                            If i / 100 = Int(i / 100) Then
                                'lblContatore.Text = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False)
                                If instance.InvokeRequired Then
                                    instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False))
                                Else
                                    lblContatore.Text = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False)
                                End If
                            End If
                        End If

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If BloccaTutto Then
                            Exit For
                        End If

                        Application.DoEvents()
                    End If
                Next

                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                ReDim Filetti(0)
                ReDim DimensioneFiletti(0)
                ReDim DataFiletti(0)

                If Not BloccaTutto Then
                    ' Lettura Destinazione
                    LeggiFiles = True

                    If Intelligente Then
                        ' Nel caso di sincronia intelligente vado a riprendere gli eventuali dati salvati nel db l'ultima volta. Se non ci sono li ricreo
                        LeggiFiles = False

                        Sql = "Select Max(Progressivo)+1 From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                        Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                        If Rec2(0).Value Is DBNull.Value = True Then
                            Massimo = 1
                        Else
                            Massimo = Rec2(0).Value
                        End If
                        Rec2.Close()

                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Progressivo Destinazione: " & gf.FormattaNumero(Massimo, False), " ", ModalitaServizio, clLog)

                        Sql = "Select Count(*) From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                        Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                        If Rec2(0).Value Is DBNull.Value = True Then
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Rilevato nessun file", " ", ModalitaServizio, clLog)

                            LeggiFiles = True
                        Else
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Rilevati files: " & gf.FormattaNumero(Rec2(0).Value, False), " ", ModalitaServizio, clLog)

                            If Rec2(0).Value = 0 Then
                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Rileggo directory destinazione", " ", ModalitaServizio, clLog)

                                LeggiFiles = True
                            End If
                        End If
                        Rec2.Close()
                    End If

                    If LeggiFiles Then
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Lettura directory di destinazione " & vbCrLf & Destinazione, " ", ModalitaServizio, clLog)

                        gf.ScansionaDirectorySingola(Destinazione, instance, Filtro, lblOperazione, False)

                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If Not BloccaTutto Then
                            Filetti = gf.RitornaFilesRilevati
                            DimensioneFiletti = gf.RitornaDimensioneFilesRilevati
                            DataFiletti = gf.RitornaDataFilesRilevati
                            qFiletti = gf.RitornaQuantiFilesRilevati

                            CartelleDest = gf.RitornaDirectoryRilevate
                            qCartelleDest = gf.RitornaQuanteDirectoryRilevate

                            'If Intelligente Then
                            '    Sql = "Delete From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                            'Else
                            Sql = "Delete * From FilesDestinazione"
                            'End If
                            DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Scrittura dati directory" & vbCrLf & gf.TagliaLunghezzaScritta(Destinazione, LunghezzaMassimaScritte), " ", ModalitaServizio, clLog)
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Numero Cartelle: " & gf.FormattaNumero(qCartelleDest, False) & vbCrLf & "Numero files: " & gf.FormattaNumero(qFiletti, False), " ", ModalitaServizio, clLog)

                            If Mid(Destinazione, Destinazione.Length, 1) = "\" Then
                                Altro = ""
                            Else
                                Altro = "\"
                            End If

                            For i As Long = 0 To qFiletti
                                If Filetti(i) <> "" Then
                                    Try
                                        Datella = DataFiletti(i)
                                        DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                    Catch ex As Exception
                                        DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                    End Try

                                    If Intelligente Then
                                        Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                            " " & idProc & ", " &
                                            " " & Progressivo & ", " &
                                            " " & (i + 1) & ", " &
                                            "'" & Filetti(i).Replace("'", "''").Replace(Destinazione & "\", "") & "', " &
                                            " " & DimensioneFiletti(i) & ", " &
                                            "'" & DatellaFile & "' " &
                                            ")"
                                    Else
                                        Sql = "Insert Into FilesDestinazione Values (" &
                                            "'" & Filetti(i).Replace("'", "''").Replace(Destinazione & "\", "") & "', " &
                                            " " & DimensioneFiletti(i) & ", " &
                                            "'" & DatellaFile & "' " &
                                            ")"
                                    End If
                                    DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                    Massimo = i + 1

                                    If IsNothing(lblContatore) = False Then
                                        If i / 100 = Int(i / 100) Then
                                            'lblContatore.Text = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False)
                                            If instance.InvokeRequired Then
                                                instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False))
                                            Else
                                                lblContatore.Text = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(qFiletti, False)
                                            End If
                                        End If
                                    End If

                                    If MetteInPausa Then
                                        MetteInPausaLaRoutine()
                                    End If

                                    If BloccaTutto Then
                                        Exit For
                                    End If

                                    Application.DoEvents()
                                End If
                            Next
                        End If
                    Else
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Lettura directory " & gf.TagliaLunghezzaScritta(Destinazione, LunghezzaMassimaScritte) & " skippata: Sincronizzazione intelligente", " ", ModalitaServizio, clLog)
                    End If

                    If MetteInPausa Then
                        MetteInPausaLaRoutine()
                    End If

                    ReDim Filetti(0)
                    ReDim DimensioneFiletti(0)
                    ReDim DataFiletti(0)

                    ' Copia i dati della sincronia intelligente nella tabella di appoggio per rendere più veloci le ricerche
                    If Not BloccaTutto Then
                        If Intelligente Then
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Pulizia valori per sincronizzazione intelligente", " ", ModalitaServizio, clLog)
                            Sql = "Delete * From DatiSincroniaIntelligente"
                            DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Copia valori per sincronizzazione intelligente", " ", ModalitaServizio, clLog)
                            Sql = "Insert Into DatiSincroniaIntelligente Select * From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
                            DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                        End If
                    End If

                    If Not BloccaTutto Then
                        ' Copia i file che esistono nell'origine ma non nella destinazione
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Conteggio files da copiare verso destinazione", " ", ModalitaServizio, clLog)

                        If Intelligente Then
                            Sql = "Select A.[File] From FilesOrigine A LEFT OUTER JOIN DatiSincroniaIntelligente B " &
                                "On ((A.[File]=B.[File]) And B.idProc=" & idProc & " And B.Operazione=" & Progressivo & ") " &
                                "Where B.[File] Is Null"
                        Else
                            Sql = "SELECT A.[File] " &
                                    "FROM FilesOrigine AS A LEFT OUTER JOIN " &
                                    "FilesDestinazione AS B ON (B.[File] = A.[File]) " &
                                    "WHERE B.[File] Is NULL"
                        End If

                        If ModalitaServizio Then
                            clLog.ScriveLogServizio("SQL: " & Sql)
                        End If

                        Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                        FilesDaElaborare = New Collection
                        q = 0
                        Do Until Rec2.Eof
                            If Not lblContatore Is Nothing Then
                                If q / 10 = Int(q / 10) Then
                                    'lblContatore.Text = gf.FormattaNumero(q, False)
                                    If instance.InvokeRequired Then
                                        instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(q, False))
                                    Else
                                        lblContatore.Text = gf.FormattaNumero(q, False)
                                    End If
                                    Application.DoEvents()
                                End If
                            End If

                            If ModalitaServizio Then
                                If q / 10 = Int(q / 10) Then
                                    clLog.ScriveLogServizio("Riga " & gf.FormattaNumero(q, False) & " - Files rilevati: " & gf.FormattaNumero(FilesDaElaborare.Count, False))
                                End If
                            End If

                            Try
                                FileDestinazione = Destinazione & "\" & Rec2("File").Value
                                If Not File.Exists(FileDestinazione) Then
                                    FilesDaElaborare.Add(Rec2("File").Value)
                                    q += 1
                                End If
                            Catch ex As Exception
                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Problema su lettura file: " & ex.Message, " ", ModalitaServizio, clLog)
                            End Try

                            Rec2.MoveNext()
                        Loop
                        Rec2.Close()

                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Files rilevati: " & gf.FormattaNumero(FilesDaElaborare.Count, False), " ", ModalitaServizio, clLog)

                        If FilesDaElaborare.Count > 0 Then
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Elaborazione files da copiare verso destinazione", " ", ModalitaServizio, clLog)
                            For i As Long = 1 To FilesDaElaborare.Count
                                FileOrigine = Origine & "\" & FilesDaElaborare.Item(i)
                                FileDestinazione = Destinazione & "\" & FilesDaElaborare.Item(i)

                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Controllo file:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte) & vbCrLf & gf.FormattaNumero(Dimens, False) & " " & Tipo, gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)

                                If File.Exists(FileOrigine) Then
                                    gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(FileDestinazione) & "\")
                                    Dim s As String = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False)
                                    Dim RitornoCopia As String = gf.CopiaFileFisico(FileOrigine, FileDestinazione, True, instance, lblContatore, s)
                                    If RitornoCopia <> "SKIPPED" And Not RitornoCopia.Contains("ERRORE") Then
                                        If Intelligente Then
                                            Try
                                                Datella = FileDateTime(FileOrigine)
                                                DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                            Catch ex As Exception
                                                DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                            End Try

                                            Massimo += 1

                                            Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                                " " & idProc & ", " &
                                                " " & Progressivo & ", " &
                                                " " & Massimo & ", " &
                                                "'" & FileOrigine.Replace(Origine & "\", "").Replace("'", "''") & "', " &
                                                " " & FileLen(FileOrigine) & ", " &
                                                "'" & DatellaFile & "' " &
                                                ")"
                                            DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                        End If

                                        Dimens = Int(FileLen(FileOrigine) / 1024)
                                        Tipo = "Kb."
                                        If Dimens = 0 Then
                                            Dimens = FileLen(FileOrigine)
                                            Tipo = "b."
                                        End If
                                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Copiato file:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte) & vbCrLf & gf.FormattaNumero(Dimens, False) & " " & Tipo, gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                    Else
                                        If RitornoCopia.Contains("ERRORE") Then
                                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, RitornoCopia & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                        Else
                                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Skipped file:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                        End If
                                    End If
                                Else
                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "File di origine non presente:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                End If

                                If MetteInPausa Then
                                    MetteInPausaLaRoutine()
                                End If

                                If BloccaTutto Then
                                    Exit For
                                End If

                                'Rec2.MoveNext()
                            Next
                        End If

                        If MetteInPausa Then
                            MetteInPausaLaRoutine()
                        End If

                        If Not BloccaTutto Then
                            ' Elimina i files nella destinazione che non esistono nell'origine
                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Conteggio files da eliminare nella destinazione", " ", ModalitaServizio, clLog)

                            If Intelligente Then
                                Sql = "Select A.[File] From FileDestinazioneIntelligente A " &
                                    "LEFT OUTER JOIN Filesorigine B On ((A.[File]=B.[File]) And A.idProc = " & idProc & " And A.Operazione = " & Progressivo & ") " &
                                    "Where B.[File] Is Null"
                            Else
                                Sql = "SELECT  A.[File] " &
                                        "FROM FilesDestinazione AS A LEFT OUTER JOIN " &
                                        "FilesOrigine AS B ON (B.[File] = A.[File]) " &
                                        "WHERE B.[File] Is NULL"
                            End If

                            If ModalitaServizio Then
                                clLog.ScriveLogServizio("SQL: " & Sql)
                            End If

                            Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                            FilesDaElaborare = New Collection
                            q = 0
                            Do Until Rec2.Eof
                                If Not lblContatore Is Nothing Then
                                    If q / 10 = Int(q / 10) Then
                                        'lblContatore.Text = gf.FormattaNumero(q, False)
                                        If instance.InvokeRequired Then
                                            instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(q, False))
                                        Else
                                            lblContatore.Text = gf.FormattaNumero(q, False)
                                        End If
                                        Application.DoEvents()
                                    End If
                                End If

                                If ModalitaServizio Then
                                    If q / 10 = Int(q / 10) Then
                                        clLog.ScriveLogServizio("Riga " & gf.FormattaNumero(q, False) & " - Files rilevati: " & gf.FormattaNumero(FilesDaElaborare.Count, False))
                                    End If
                                End If

                                Try
                                    FileOrigine = Origine & "\" & Rec2("File").Value
                                    FileDestinazione = Destinazione & "\" & Rec2("File").Value
                                    If Not File.Exists(FileOrigine) And File.Exists(FileDestinazione) Then
                                        FilesDaElaborare.Add(Rec2("File").Value)
                                        q += 1
                                    End If
                                Catch ex As Exception
                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Problema su lettura file: " & ex.Message, " ", ModalitaServizio, clLog)
                                End Try

                                Rec2.MoveNext()
                            Loop
                            Try
                                Rec2.Close()
                            Catch ex As Exception

                            End Try

                            If FilesDaElaborare.Count > 0 Then
                                Dim Ok As Boolean = True
                                'If FilesDaElaborare.Count > 500 Then
                                '    If MsgBox("Troppi files da eliminare. Proseguo con l'operazione ?", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
                                '        Ok = False
                                '    End If
                                'End If
                                If Ok Then
                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Elaborazione files da eliminare nella destinazione", " ", ModalitaServizio, clLog)
                                    For i As Long = 1 To FilesDaElaborare.Count
                                        FileOrigine = Destinazione & "\" & FilesDaElaborare.Item(i)

                                        If Not gf.EliminaFileFisico(FileOrigine).Contains("ERRORE:") Then
                                            If Intelligente Then
                                                Sql = "Delete * From FileDestinazioneIntelligente Where " &
                                                    "[File]='" & FileOrigine.Replace(Origine & "\", "").Replace("'", "''") & "' And " &
                                                    "idProc=" & idProc & " And " &
                                                    "Operazione=" & Progressivo
                                                DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                            End If

                                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Eliminazione file:" & vbCrLf & gf.TagliaLunghezzaScritta(FilesDaElaborare.Item(i), LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                        End If

                                        If MetteInPausa Then
                                            MetteInPausaLaRoutine()
                                        End If

                                        If BloccaTutto Then
                                            Exit For
                                        End If

                                        'Rec2.MoveNext()
                                    Next
                                End If
                            End If
                            'Rec2.Close()

                            If MetteInPausa Then
                                MetteInPausaLaRoutine()
                            End If

                            If Not BloccaTutto Then
                                ' Copia nella destinazione i files che hanno date superiori nell'origine o dimensioni diverse
                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Conteggio files diversi", " ", ModalitaServizio, clLog)

                                If Intelligente Then
                                    Sql = "Select A.[File] From DatiSincroniaIntelligente A Left Outer Join " &
                                        "FilesOrigine B " &
                                        "On ((A.[File]=B.[File]) And A.idProc=" & idProc & " And A.Operazione=" & Progressivo & ") " &
                                        "Where A.Dimensione <> B.Dimensioni Or A.DataFile < B.DataOra"
                                Else
                                    Sql = "Select A.[File] From Filesorigine A Left Outer Join FilesDestinazione B On (A.[File] = B.[File]) " &
                                            "Where A.Dimensioni <> B.Dimensioni " &
                                            "Or A.DataOra > B.DataOra"
                                End If

                                If ModalitaServizio Then
                                    clLog.ScriveLogServizio("SQL: " & Sql)
                                End If

                                Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                                FilesDaElaborare = New Collection
                                q = 0
                                Do Until Rec2.Eof
                                    If Not lblContatore Is Nothing Then
                                        If q / 10 = Int(q / 10) Then
                                            'lblContatore.Text = gf.FormattaNumero(q, False)
                                            If instance.InvokeRequired Then
                                                instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(q, False))
                                            Else
                                                lblContatore.Text = gf.FormattaNumero(q, False)
                                            End If
                                            Application.DoEvents()
                                        End If
                                    End If

                                    If ModalitaServizio Then
                                        If q / 10 = Int(q / 10) Then
                                            clLog.ScriveLogServizio("Riga " & gf.FormattaNumero(q, False) & " - Files rilevati: " & gf.FormattaNumero(FilesDaElaborare.Count, False))
                                        End If
                                    End If

                                    FileOrigine = Origine & "\" & Rec2("File").Value
                                    FileDestinazione = Destinazione & "\" & Rec2("File").Value
                                    If File.Exists(FileOrigine) And File.Exists(FileDestinazione) Then
                                        FilesDaElaborare.Add(Rec2("File").Value)
                                        q += 1
                                    End If

                                    Rec2.MoveNext()
                                Loop
                                Rec2.Close()

                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Files rilevati: " & gf.FormattaNumero(FilesDaElaborare.Count, False), " ", ModalitaServizio, clLog)

                                If FilesDaElaborare.Count > 0 Then
                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Elaborazione files diversi", " ", ModalitaServizio, clLog)

                                    For i As Long = 1 To FilesDaElaborare.Count
                                        FileOrigine = Origine & "\" & FilesDaElaborare.Item(i)
                                        FileDestinazione = Destinazione & "\" & FilesDaElaborare.Item(i)

                                        If File.Exists(FileOrigine) Then
                                            Dim t As New ThreadAttesaOperazione
                                            gf.CreaDirectoryDaPercorso(gf.TornaNomeDirectoryDaPath(FileDestinazione) & "\")

                                            t.EsegueControllo(instance, idProc, ModalitaServizio, clLog, FileDestinazione, log, lblOperazione, lblContatore, "Controllo file " & gf.TornaNomeFileDaPath(FileOrigine))

                                            Dim s As String = gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False)
                                            Dim RitornoCopia As String = gf.CopiaFileFisico(FileOrigine, FileDestinazione, True, instance, lblContatore, s)
                                            If RitornoCopia <> "SKIPPED" And Not RitornoCopia.Contains("ERRORE") Then
                                                If Intelligente Then
                                                    Try
                                                        Datella = FileDateTime(FileOrigine)
                                                        DatellaFile = Datella.Year & "-" & Datella.Month & "-" & Datella.Day & " " & Datella.Hour & ":" & Datella.Minute & ":" & Datella.Second
                                                    Catch ex As Exception
                                                        DatellaFile = Now.Year & "-" & Now.Month & "-" & Now.Day & " " & Now.Hour & ":" & Now.Minute & ":" & Now.Second
                                                    End Try

                                                    Massimo += 1

                                                    Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                                                        " " & idProc & ", " &
                                                        " " & Progressivo & ", " &
                                                        " " & Massimo & ", " &
                                                        "'" & FileOrigine.Replace(Origine & "\", "").Replace("'", "''") & "', " &
                                                        " " & FileLen(FileOrigine) & ", " &
                                                        "'" & DatellaFile & "' " &
                                                        ")"
                                                    DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                                End If

                                                Dimens = Int(FileLen(FileOrigine) / 1024)
                                                Tipo = "Kb."
                                                If Dimens = 0 Then
                                                    Dimens = FileLen(FileOrigine)
                                                    Tipo = "b."
                                                End If
                                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Aggiorna file:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte) & vbCrLf & gf.FormattaNumero(Dimens, False) & " " & Tipo, gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                            Else
                                                If RitornoCopia.Contains("ERRORE") Then
                                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, RitornoCopia & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                                Else
                                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Skip file:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                                End If
                                            End If

                                            t.Blocca(lblOperazione)
                                        Else
                                            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "File di origine non presente:" & vbCrLf & gf.TagliaLunghezzaScritta(FileOrigine, LunghezzaMassimaScritte), gf.FormattaNumero(i, False) & "/" & gf.FormattaNumero(FilesDaElaborare.Count, False), ModalitaServizio, clLog)
                                        End If

                                        If MetteInPausa Then
                                            MetteInPausaLaRoutine()
                                        End If

                                        If BloccaTutto Then
                                            Exit For
                                        End If

                                        'Rec2.MoveNext()
                                    Next
                                End If
                                'Rec2.Close()

                                If MetteInPausa Then
                                    MetteInPausaLaRoutine()
                                End If

                                If Not BloccaTutto Then
                                    ' Elimina/Crea cartelle vuote nella destinazione
                                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Pulizia tabella appoggio origine", "", ModalitaServizio, clLog)
                                    Sql = "Delete From DirectOrig"
                                    DB.EsegueSqlSenzaTRY(idProc, ConnSQL, Sql, clLog)

                                    Dim i As Long = 0

                                    For Each cd As String In CartelleOrig
                                        If cd <> "" Then
                                            If i / 100 = Int(i / 100) Then
                                                ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Scrittura tabella origine", i & "/" & qCartelleOrig, ModalitaServizio, clLog)
                                            End If
                                            cd = cd.Replace(Origine, "")
                                            If cd <> "" Then
                                                Sql = "Insert Into DirectOrig Values ('" & cd.Replace("'", "''") & "')"
                                                DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                            End If
                                        End If
                                        i += 1
                                    Next

                                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Pulizia tabella appoggio destinazione", "", ModalitaServizio, clLog)
                                    Sql = "Delete From DirectDest"
                                    DB.EsegueSqlSenzaTRY(idProc, ConnSQL, Sql, clLog)

                                    i = 0
                                    For Each cd As String In CartelleDest
                                        If cd <> "" Then
                                            If i / 100 = Int(i / 100) Then
                                                ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Scrittura tabella destin.", i & "/" & qCartelleDest, ModalitaServizio, clLog)
                                            End If
                                            cd = cd.Replace(Destinazione, "")
                                            If cd <> "" Then
                                                Sql = "Insert Into DirectDest Values ('" & cd.Replace("'", "''") & "')"
                                                DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                            End If
                                        End If
                                        i += 1
                                    Next

                                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Lettura directory da eliminare", "", ModalitaServizio, clLog)
                                    Sql = "Select * From DirectDest Where Nome Not In (Select Nome From DirectOrig)"
                                    Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                                    Dim CartelleDaEliminare As New ArrayList
                                    i = 0
                                    Do Until Rec2.Eof
                                        If i / 100 = Int(i / 100) Then
                                            ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione tabella" & gf.TagliaLunghezzaScritta(Destinazione & Rec2("Nome").Value, LunghezzaMassimaScritte), i & "/" & CartelleDaEliminare.Count, ModalitaServizio, clLog)
                                        End If
                                        i += 1
                                        CartelleDaEliminare.Add(Destinazione & Rec2("Nome").Value)
                                        Rec2.MoveNext
                                    Loop
                                    Rec2.Close

                                    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Lettura directory da creare", "", ModalitaServizio, clLog)
                                    Sql = "Select * From DirectOrig Where Nome Not In (Select Nome From DirectDest)"
                                    Rec2 = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                                    Dim CartelleDaAggiungere As New ArrayList
                                    i = 0
                                    Do Until Rec2.Eof
                                        If i / 100 = Int(i / 100) Then
                                            ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Creazione tabella" & gf.TagliaLunghezzaScritta(Destinazione & Rec2("Nome").Value, LunghezzaMassimaScritte), i & "/" & CartelleDaEliminare.Count, ModalitaServizio, clLog)
                                        End If
                                        i += 1
                                        CartelleDaAggiungere.Add(Destinazione & Rec2("Nome").Value)
                                        Rec2.MoveNext
                                    Loop
                                    Rec2.Close

                                    i = 0
                                    Dim qc As Long = CartelleDaEliminare.Count
                                    For Each c As String In CartelleDaEliminare
                                        ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione cartella: " & vbCrLf & gf.TagliaLunghezzaScritta(c, LunghezzaMassimaScritte), i & "/" & qc, ModalitaServizio, clLog)
                                        Directory.Delete(c, True)
                                    Next

                                    i = 0
                                    qc = CartelleDaAggiungere.Count
                                    For Each c As String In CartelleDaAggiungere
                                        ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione cartella: " & vbCrLf & gf.TagliaLunghezzaScritta(c, LunghezzaMassimaScritte), i & "/" & qc, ModalitaServizio, clLog)
                                        Directory.CreateDirectory(c)
                                    Next

                                    ' For k As Integer = 1 To 3
                                    'If Not CartelleDest Is Nothing Then
                                    '        For i As Long = qCartelleDest - 1 To 0 Step -1
                                    '            If i / 10 = Int(i / 10) Then
                                    '                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Controllo cartelle vuote " & i, " ", ModalitaServizio, clLog)
                                    '            End If
                                    '            If CartelleDest(i) <> "" Then

                                    '                'If ModalitaServizio Then
                                    '                '    ScriveOperazione(True, idProc, log, lblOperazione, lblContatore, "Cartella: " & CartelleDest(i), " ", ModalitaServizio)
                                    '                'End If

                                    '                If Directory.Exists(CartelleDest(i)) = True Then
                                    '                    Dim Ok As Boolean = True

                                    '                    If Not CartelleOrig Is Nothing Then
                                    '                        For kk As Long = 1 To qCartelleOrig - 1
                                    '                            If CartelleOrig(kk).Replace(Origine & "\", "").Trim.ToUpper = CartelleDest(i).Replace(Destinazione & "\", "").Trim.ToUpper Then
                                    '                                Ok = False
                                    '                                Exit For
                                    '                            End If
                                    '                        Next
                                    '                    End If

                                    '                    If Ok Then
                                    '                        Dim FilesInFolder As Integer = Directory.GetFiles(CartelleDest(i), "*.*").Count
                                    '                        Dim FoldersInFolder As Integer = Directory.GetDirectories(CartelleDest(i), "*.*").Count

                                    '                        If FilesInFolder = 0 And FoldersInFolder = 0 Then
                                    '                            Try
                                    '                                RmDir(CartelleDest(i))
                                    '                                ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione cartella:" & vbCrLf & gf.TagliaLunghezzaScritta(CartelleDest(i), LunghezzaMassimaScritte), " ", ModalitaServizio, clLog)
                                    '                            Catch ex As Exception
                                    '                                ' ScriveOperazione(idProc, log, lblOperazione, lblContatore, "      " & CartelleDest(i), " ERRORE: " & ex.Message)
                                    '                            End Try
                                    '                        End If
                                    '                    End If
                                    '                End If
                                    '            End If

                                    '            If MetteInPausa Then
                                    '                MetteInPausaLaRoutine()
                                    '            End If

                                    '            If BloccaTutto Then
                                    '                Exit For
                                    '            End If
                                    '        Next

                                    '        If MetteInPausa Then
                                    '            MetteInPausaLaRoutine()
                                    '        End If

                                    '    'If BloccaTutto Then
                                    '    '    Exit For
                                    '    'End If
                                    'End If
                                    '' Next

                                    If MetteInPausa Then
                                        MetteInPausaLaRoutine()
                                    End If

                                    If Not BloccaTutto Then
                                        ' Crea Cartelle nella destinazione in caso ce ne siano di vuote nell'origine
                                        'Dim sCartella As String

                                        'For i As Long = 0 To qCartelleOrig - 1
                                        '    If i / 100 = Int(i / 100) Then
                                        '        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione cartelle vuote di origine " & i & "/" & qCartelleOrig, " ", ModalitaServizio, clLog)
                                        '    End If
                                        '    If CartelleOrig(i) <> "" Then
                                        '        sCartella = CartelleOrig(i).Replace(Origine, "")
                                        '        If sCartella <> "" Then
                                        '            If Mid(sCartella, 1, 1) = "\" Then
                                        '                sCartella = Mid(sCartella, 2, sCartella.Length)
                                        '            End If

                                        '            'If ModalitaServizio Then
                                        '            '    ScriveOperazione(True, idProc, log, lblOperazione, lblContatore, "Cartella: " & sCartella, " ", ModalitaServizio)
                                        '            'End If

                                        '            If Directory.Exists(Destinazione & "\" & sCartella) = False Then
                                        '                ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Creazione cartella:" & vbCrLf & gf.TagliaLunghezzaScritta(Destinazione & "\" & sCartella & "\", LunghezzaMassimaScritte), " ", ModalitaServizio, clLog)

                                        '                gf.CreaDirectoryDaPercorso(Destinazione & "\" & sCartella & "\")
                                        '            End If
                                        '        End If
                                        '    End If

                                        '    If MetteInPausa Then
                                        '        MetteInPausaLaRoutine()
                                        '    End If

                                        '    If BloccaTutto Then
                                        '        Exit For
                                        '    End If
                                        'Next
                                        'ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione cartelle vuote", " ", ModalitaServizio, clLog)

                                        'Dim CartelleDaAggiungere As New ArrayList
                                        'i = 0
                                        'qc = CartelleOrig.Count

                                        'For Each cD As String In CartelleOrig
                                        '    If cD <> "" Then
                                        '        i += 1
                                        '        If i / 50 = Int(i / 50) Then
                                        '            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione cartelle vuote", "Controllo " & i & "/" & qc, ModalitaServizio, clLog)
                                        '        End If
                                        '        Ok = True
                                        '        For Each cO As String In CartelleDest
                                        '            If cO <> "" Then
                                        '                If cD.Replace(Origine, "") = cO.Replace(Destinazione, "") Then
                                        '                    Ok = False
                                        '                    Exit For
                                        '                End If
                                        '            End If
                                        '        Next
                                        '        If Ok Then
                                        '            CartelleDaAggiungere.Add(cD.Replace(Origine & "\", ""))
                                        '        End If
                                        '    End If
                                        'Next

                                        'i = 0
                                        'qc = CartelleDaAggiungere.Count
                                        'For Each c As String In CartelleDaAggiungere
                                        '    i += 1
                                        '    ScriveOperazione(instance, False, idProc, log, lblOperazione, lblContatore, "Eliminazione cartella: " & vbCrLf & gf.TagliaLunghezzaScritta(c, LunghezzaMassimaScritte), i & "/" & qc, ModalitaServizio, clLog)
                                        '    gf.CreaDirectoryDaPercorso(Destinazione & "\" & c)
                                        'Next

                                        If MetteInPausa Then
                                            MetteInPausaLaRoutine()
                                        End If

                                        If Not BloccaTutto Then
                                            ' Pulizia tabelle di appoggio e compattazione DB
                                            If Intelligente Then
                                                AggiornataTabellaIntelligente = True
                                                AggiornaTabellaIntelligente(instance, idProc, Progressivo, DB, ConnSQL, lblOperazione, lblContatore, ModalitaServizio, clLog)
                                            Else
                                                If MetteInPausa Then
                                                    MetteInPausaLaRoutine()
                                                End If

                                                If Not BloccaTutto Then
                                                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Pulizia tabelle", " ", ModalitaServizio, clLog)

                                                    Sql = "Delete * From FilesOrigine"
                                                    DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                                                    Sql = "Delete * From FilesDestinazione"
                                                    DB.EsegueSql(idProc, ConnSQL, Sql, clLog)

                                                    Sql = "Delete * From DatiSincroniaIntelligente"
                                                    DB.EsegueSql(idProc, ConnSQL, Sql, clLog)
                                                End If
                                            End If

                                            If MetteInPausa Then
                                                MetteInPausaLaRoutine()
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If BloccaTutto Then
                ScriveLog(idProc, "SINCRONIZZAZIONE BLOCCATA", clLog)
            Else
                If Intelligente Then
                    If AggiornataTabellaIntelligente = False Then
                        AggiornaTabellaIntelligente(instance, idProc, Progressivo, DB, ConnSQL, lblOperazione, lblContatore, ModalitaServizio, clLog)
                    End If
                End If
            End If

            DB.ChiudeDB(True, ConnSQL)

            ConnSQL = Nothing
            DB = Nothing

            Dim GA As New GestioneACCESS
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Compattazione DB", " ", ModalitaServizio, clLog)
            GA.CompattazioneDb()
            GA = Nothing
        Else
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Impossibile aprire il DB", " ", ModalitaServizio, clLog)
        End If
        'Catch ex As Exception
        '    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        'End Try

        Return log
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

    Private Sub AggiornaTabellaIntelligente(instance As Form, idProc As Integer, Progressivo As Integer, DB As GestioneACCESS, ConnSql As Object, lblOperazione As Label,
                                            lblContatore As Label, ModalitaServizio As Boolean, clLog As LogCasareccio.LogCasareccio.Logger)
        Dim Sql As String = ""

        ' Copio la tabella di origine in quella di destinazione in caso di sincronia intelligente
        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Aggiornamento tabella di sincronia", " ", ModalitaServizio, clLog)

        Sql = "Delete * From FileDestinazioneIntelligente Where idProc=" & idProc & " And Operazione=" & Progressivo
        DB.EsegueSql(idProc, ConnSql, Sql, clLog)

        Dim c As Long = 0
        Dim d As Date
        Dim mese As String
        Dim giorno As String
        Dim ora As String
        Dim minuti As String
        Dim secondi As String
        Dim Rec2 As Object = CreateObject("ADODB.Recordset")
        Dim DatellaFile As String = ""
        Dim gf As New GestioneFilesDirectory

        Sql = "Select * From FilesOrigine"
        Rec2 = DB.LeggeQuery(idProc, ConnSql, Sql, clLog)
        Do Until Rec2.Eof
            c += 1
            d = Rec2(2).Value

            mese = d.Month.ToString.Trim : If mese.Length = 1 Then mese = "0" & mese
            giorno = d.Day.ToString.Trim : If giorno.Length = 1 Then giorno = "0" & giorno
            ora = d.Hour.ToString.Trim : If ora.Length = 1 Then ora = "0" & ora
            minuti = d.Minute.ToString.Trim : If minuti.Length = 1 Then minuti = "0" & minuti
            secondi = d.Second.ToString.Trim : If secondi.Length = 1 Then secondi = "0" & secondi

            DatellaFile = d.Year & "-"
            DatellaFile &= mese & "-"
            DatellaFile &= giorno & " "
            DatellaFile &= ora & ":"
            DatellaFile &= minuti & ":"
            DatellaFile &= secondi

            Sql = "Insert Into FileDestinazioneIntelligente Values (" &
                 " " & idProc & ", " &
                 " " & Progressivo & ", " &
                 " " & c & ", " &
                 "'" & Replace(Rec2(0).Value, "'", "''") & "', " &
                 " " & Rec2(1).Value & ", " &
                 "'" & DatellaFile & "' " &
                 ")"
            DB.EsegueSql(idProc, ConnSql, Sql, clLog)

            Rec2.MoveNext()

            If c / 50 = Int(c / 50) Then
                If Not lblContatore Is Nothing Then
                    ' lblContatore.Text = gf.FormattaNumero(c, False)
                    If instance.InvokeRequired Then
                        instance.Invoke(MethodDelegateAddTextContatore, gf.FormattaNumero(c, False))
                    Else
                        lblContatore.Text = gf.FormattaNumero(c, False)
                    End If
                End If

                Application.DoEvents()
            End If
        Loop
        Rec2.Close()

        Sql = "Delete * From FilesOrigine"
        DB.EsegueSql(idProc, ConnSql, Sql, clLog)

        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Aggiornate righe sincronia: " & gf.FormattaNumero(c, False) & " - idProc: " & idProc & " - Operazione: " & Progressivo, " ", ModalitaServizio, clLog)
    End Sub

    Public Function EsegueZip(idProc As Integer, Origine As String, Destinazione As String, lblOperazione As Label, lblContatore As Label, ModalitaServizio As Boolean,
                              instance As Form, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "ZIP FILES: " & Origine & "->" & Destinazione, clLog)

        Dim PathUlteriore As String = ""

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, instance, "", lblOperazione, False)

            'lblOperazione.Text = ""
            'Application.DoEvents()

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati

            Try
                File.Delete(Destinazione)
            Catch ex As Exception

            End Try

            Try
                ' Using zip As New ZipFile()
                Dim zip As New ZipFile()
                Dim FileZippante As ZipEntry

                'If qFiletti > 60000 Then
                '    zip.UseZip64WhenSaving = Zip64Option.Always
                'End If
                zip.ParallelDeflateThreshold = -1

                Dim t As New ThreadAttesaOperazione
                t.EsegueControllo(instance, idProc, ModalitaServizio, clLog, Destinazione, log, lblOperazione, lblContatore, "Creazione file zip", zip)

                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Aggiunta file a zip", " ", ModalitaServizio, clLog)

                Dim Dest2 As String = Destinazione
                Dim Progressivo As Integer = 0
                Dim i2 As Long = 0

                For i As Long = 1 To qFiletti
                    i2 += 1
                    If (i2 / 100 = Int(i2 / 100)) Then
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Aggiunta file a zip: " & gf.FormattaNumero ( i,False) & " / " & gf.FormattaNumero (qfiletti,False), " ", ModalitaServizio, clLog)
                    End If
                    FileZippante = zip.AddFile(Filetti(i))

                    If i2 > 32000 Then
                        i2 = 0
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione zip " & Origine, " ", ModalitaServizio, clLog)

                        Try
                            File.Delete(dest2)
                        Catch ex As Exception

                        End Try

                        zip.Save(Dest2)

                        Progressivo += 1
                        Dim este As String = gf.TornaEstensioneFileDaPath(Destinazione)

                        Dest2 = Destinazione.Replace(este, "") & "." & Format(Progressivo, "000") & este

                        t.Blocca(lblOperazione)

                        zip = New ZipFile()
                        FileZippante = New ZipEntry

                        t = New ThreadAttesaOperazione
                        t.EsegueControllo(instance, idProc, ModalitaServizio, clLog, Dest2, log, lblOperazione, lblContatore, "Creazione file zip", zip)
                    End If

                    If BloccaTutto Then
                        Exit For
                    End If
                Next

                If Not BloccaTutto Then
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione zip " & Origine, " ", ModalitaServizio, clLog)

                    zip.Save(Dest2)
                End If

                t.Blocca(lblOperazione)

                FileZippante = Nothing
                zip = Nothing
                ' End Using
            Catch ex1 As System.Exception
                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su zip: " + ex1.Message, " ", ModalitaServizio, clLog)
            End Try
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function ListaFiles(idProc As Integer, Origine As String, Destinazione As String, lblOperazione As Label, lblContatore As Label, ModalitaServizio As Boolean,
                               instance As Form, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "LISTA FILES: " & Origine & "->" & Destinazione, clLog)

        Dim PathUlteriore As String = ""

        log = New StringBuilder

        Try
            gf.ScansionaDirectorySingola(Origine, instance, "", lblOperazione, False)

            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "", "", ModalitaServizio, clLog)

            Dim Filetti() As String = gf.RitornaFilesRilevati
            Dim qFiletti As Long = gf.RitornaQuantiFilesRilevati
            Dim qDirectory As Long = gf.RitornaQuanteDirectoryRilevate
            Dim Percorso As String
            Dim oldPercorso As String = ""
            Dim Filetto As String
            Dim qDett As Integer = 0
            Dim qBytes As Long = 0
            Dim qBytesTotali As Long = 0

            Try
                File.Delete(Destinazione)
            Catch ex As Exception

            End Try

            Try
                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Creazione lista files", " ", ModalitaServizio, clLog)

                gf.ApreFileDiTestoPerScrittura(Destinazione)
                For i As Long = 1 To qFiletti
                    Percorso = gf.TornaNomeDirectoryDaPath(Filetti(i))
                    Filetto = gf.TornaNomeFileDaPath(Filetti(i))

                    If oldPercorso <> Percorso Then
                        If oldPercorso <> "" Then
                            gf.ScriveTestoSuFileAperto("")
                            gf.ScriveTestoSuFileAperto("     Files: " & gf.FormattaNumero(qDett, False) & " - MBytes: " & gf.FormattaNumero(qBytes / 1024 / 1024, False))
                            gf.ScriveTestoSuFileAperto("")
                        End If
                        gf.ScriveTestoSuFileAperto(Percorso)
                        gf.ScriveTestoSuFileAperto("")
                        oldPercorso = Percorso
                        qDett = 0
                        qBytes = 0
                    End If

                    gf.ScriveTestoSuFileAperto("     " & Filetto)
                    qDett += 1
                    qBytes += gf.TornaDimensioneFile(Filetti(i))
                    qBytesTotali += gf.TornaDimensioneFile(Filetti(i))

                    If BloccaTutto Then
                        Exit For
                    End If
                Next
                gf.ScriveTestoSuFileAperto("")
                gf.ScriveTestoSuFileAperto("     Directories: " & gf.FormattaNumero(qDirectory, False) & " - Files: " & gf.FormattaNumero(qFiletti, False) & " - MBytes: " & gf.FormattaNumero(qBytesTotali / 1024 / 1024, False))
                gf.ChiudeFileDiTestoDopoScrittura()
            Catch ex1 As System.Exception
                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su Lista Files: " + ex1.Message, " ", ModalitaServizio, clLog)
            End Try
        Catch ex As Exception
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "ERRORE: " & ex.Message, " ", ModalitaServizio, clLog)
        End Try

        Return log
    End Function

    Public Function EsegueSQL(idProc As Integer, Origine As String, Destinazione As String, lblOperazione As Label, lblContatore As Label, ModalitaServizio As Boolean,
                               instance As Form, clLog As LogCasareccio.LogCasareccio.Logger) As StringBuilder
        ScriveLog(idProc, "ESEGUE SQL: " & Origine & "->" & Destinazione, clLog)

        log = New StringBuilder

        If Not File.Exists(Destinazione) Then
            ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su esegue sql: file '" & Destinazione & "' non esistente ", "", ModalitaServizio, clLog)
        Else
            Dim sql As String = gf.LeggeFileIntero(Destinazione)
            Dim s As New SQLSERVER
            If s.ImpostaConnessioneDirettamente(Origine) Then
                Dim conn As Object = s.ApreDB()
                If conn Is Nothing Then
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su apertura connessione: " & Origine, " ", ModalitaServizio, clLog)
                Else
                    Dim ritorno As String
                    Dim righe() As String = sql.Split("|")
                    Dim riga As Integer = 0
                    Dim erro As Boolean = False
                    For Each ss As String In righe
                        riga += 1
                        If ss <> "" Then
                            ritorno = s.EsegueSql(conn, ss)
                            If ritorno <> "" Then
                                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su esecuzione sql alla riga " & riga & ": " & ss, " ", ModalitaServizio, clLog)
                                erro = True
                            End If
                        End If
                    Next
                    If Not erro Then
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Esecuzione sql OK", " ", ModalitaServizio, clLog)
                    Else
                        ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Esecuzione sql conculsa con errori", " ", ModalitaServizio, clLog)
                    End If
                End If
            Else
                ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Errore su connessione: " & Origine, " ", ModalitaServizio, clLog)
            End If
        End If

        Return log
    End Function

End Class
