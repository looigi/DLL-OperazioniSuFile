Imports System.Windows.Forms
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.ServiceProcess
Imports System.IO
Imports frmLog

Public Class OperazioniSuFile
    Private DB As New GestioneACCESS
    Private ConnSQL As Object

    Private NumeroCampiTabella As Integer = 13
    Public RigheProcedura(200, NumeroCampiTabella) As String

    'Private NomeFileLog As String = ""

    Public Enum TipoOperazione
        Nulla = 0
        Copia = 1
        Spostamento = 2
        Sincronizzazione = 3
        Eliminazione = 4
        CreaDirectory = 5
        EliminaDirectory = 6
        RiavvioPC = 7
        AvvioServizio = 8
        FermaServizio = 9
        AvviaEseguibile = 10
        FermaEseguibile = 11
        Attendi = 12
        SincroniaIntelligente = 13
        Zip = 14
        ListaFiles = 15
        Messaggio = 16
        EsegueSQL = 17
    End Enum

    Public Function TornaListaOperazioni() As String()
        Dim Righe(17) As String

        Righe(0) = "Nulla"
        Righe(1) = "Copia"
        Righe(2) = "Sposta"
        Righe(3) = "Sincronizza"
        Righe(4) = "Elimina Files"
        Righe(5) = "Crea dir"
        Righe(6) = "Elimina dir"
        Righe(7) = "Riavvio PC"
        Righe(8) = "Avvia Servizio"
        Righe(9) = "Ferma Servizio"
        Righe(10) = "Avvia EXE"
        Righe(11) = "Ferma EXE"
        Righe(12) = "Attendi"
        Righe(13) = "Sincronia Intelligente"
        Righe(14) = "Zip"
        Righe(15) = "Lista Files"
        Righe(16) = "Messaggio"
        Righe(17) = "Esegue SQL Server"

        Return Righe
    End Function

    Sub New()
        MetteInPausa = False
        BloccaTutto = False

        ImpostaDBSqlAccess()
    End Sub

    'Public Function RitornaNomeFileDiLog() As String
    '    Return NomeFileLog
    'End Function

    Public Sub ImpostaPausa(Attiva As Boolean)
        MetteInPausa = Attiva
    End Sub

    Public Function PrendeLetteraDiscoLibero() As String
        Return FindNextAvailableDriveLetter()
    End Function

    Public Function EsisteLetteraDisco(Lettera As String) As Boolean
        Return ExistsDriveLetter(Lettera)
    End Function

    Public Sub ChiudeDBLog()
        ChiudeDBGlobale()
    End Sub

    Public Sub ScriveLogDaRemoto(idProc As Integer, Log As String, clLog As LogCasareccio.LogCasareccio.Logger)
        ScriveLog(idProc, Log, clLog)
    End Sub

    Public Sub ImpostaSkippa(Cosa As Boolean)
        Skippa = Cosa
    End Sub

    Public Sub ImpostaBlocco()
        BloccaTutto = True
    End Sub

    Private Sub PulisceArray()
        For i As Integer = 0 To 200
            For k As Integer = 0 To NumeroCampiTabella
                RigheProcedura(i, k) = ""
            Next
        Next
    End Sub

    Public Function TornaRecordsetRigheProcedura(ModalitaEsecuzioneAutomatica As Boolean, PercorsoDBTemp As String, idProc As Integer, clLog As LogCasareccio.LogCasareccio.Logger) As String(,)
        PulisceArray()

        Try
            Dim Db As New GestioneACCESS
            Dim ConnSQL As Object = CreateObject("ADODB.Connection")

            If Db.LeggeImpostazioniDiBase(ModalitaEsecuzioneAutomatica, PercorsoDBTemp, "ConnDB") = True Then
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                ConnSQL = Db.ApreDB(idProc, clLog)
                Dim Riga As Integer = 0
                Dim Sql As String

                Sql = "Select * From DettaglioProcedure Where idProc=" & idProc & " Order By Progressivo"
                Rec = Db.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                Do Until Rec.Eof
                    For i As Integer = 0 To NumeroCampiTabella
                        RigheProcedura(Riga, i) = Rec(i).Value.ToString
                    Next
                    Riga += 1

                    Rec.MoveNext()
                Loop
                Rec.Close()
                RigheProcedura(Riga, 0) = "***"

                Db.ChiudeDB(True, ConnSQL)
                Db = Nothing
            End If
        Catch ex As Exception

        End Try

        Return RigheProcedura
    End Function

    Public Sub RilasciaOggetti()
        ConnSQL.close()
        ConnSQL = Nothing

        DB = Nothing
    End Sub

    Private Function SistemaLunghezzaCampo(Campo As String, Lunghezza As String) As String
        Dim Ritorno As String = Campo.Trim

        If Ritorno.Length < Lunghezza Then
            For i As Integer = Ritorno.Length To Lunghezza
                Ritorno = Ritorno & " "
            Next
        Else
            Ritorno = Mid(Ritorno, 1, (Lunghezza / 2)) & "..." & Mid(Ritorno, Ritorno.Length - (Lunghezza / 2) + 3, Lunghezza)
        End If

        Return Ritorno
    End Function

    Public Function CaricaDatiProcedura(ModalitaEsecuzioneAutomatica As Boolean, PercorsoDBTemp As String, idProc As Integer, clLog As LogCasareccio.LogCasareccio.Logger) As String
        Dim Ritorno As String = ""

        Try
            Dim InvioMail As String = "N"
            Dim DB As New GestioneACCESS

            If DB.LeggeImpostazioniDiBase(ModalitaEsecuzioneAutomatica, PercorsoDBTemp, "ConnDB") = True Then
                Dim ConnSQL As Object = DB.ApreDB(idProc, clLog)
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String = "Select * From NomiProcedure Where idProc=" & idProc

                Rec = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                If Rec.Eof = False Then
                    InvioMail = Rec("InvioMail").Value
                End If
                Rec.Close()

                Ritorno = InvioMail & ";"

                DB.ChiudeDB(True, ConnSQL)
            End If

            DB = Nothing
        Catch ex As Exception

        End Try

        Return Ritorno
    End Function

    Public Function CaricaRigheProcedura(ModalitaEsecuzioneAutomatica As Boolean, PercorsoDBTemp As String, Procedura As String,
                                         Optional lstOperazioni As ListBox = Nothing, Optional clLog As LogCasareccio.LogCasareccio.Logger = Nothing) As Integer
        Dim idProc As Integer

        Try
            Dim DB As New GestioneACCESS

            If DB.LeggeImpostazioniDiBase(ModalitaEsecuzioneAutomatica, PercorsoDBTemp, "ConnDB") = True Then
                Dim ConnSQL As Object = DB.ApreDB(idProc, clLog)
                Dim Rec As Object = CreateObject("ADODB.Recordset")
                Dim Sql As String

                idProc = -1

                If lstOperazioni Is Nothing = False Then
                    lstOperazioni.Items.Clear()
                End If

                Sql = "Select idProc From NomiProcedure Where NomeProcedura='" & Procedura & "'"
                Rec = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                If Rec.Eof = False Then
                    idProc = Rec(0).Value
                End If
                Rec.Close()

                If idProc <> -1 Then
                    Dim sTipoOperazione As String = ""
                    Dim sOrigine As String = ""
                    Dim sDestinazione As String = ""
                    Dim sSovrascrivi As String = ""
                    Dim sSottoDirectory As String = ""
                    Dim sFiltro As String = ""
                    Dim sProgressivo As String = ""

                    Sql = "Select * From DettaglioProcedure Where idProc=" & idProc & " Order By Progressivo"
                    Rec = DB.LeggeQuery(idProc, ConnSQL, Sql, clLog)
                    Do Until Rec.Eof
                        sProgressivo = Rec("Progressivo").Value
                        If sProgressivo.Length = 1 Then
                            sProgressivo = " " & sProgressivo
                        End If

                        Select Case Rec("idOperazione").Value
                            Case TipoOperazione.Nulla
                                sTipoOperazione = ""
                            Case TipoOperazione.Copia
                                sTipoOperazione = "Copia"
                            Case TipoOperazione.CreaDirectory
                                sTipoOperazione = "Crea dir"
                            Case TipoOperazione.EliminaDirectory
                                sTipoOperazione = "Elimina dir"
                            Case TipoOperazione.Eliminazione
                                sTipoOperazione = "Elimina Files"
                            Case TipoOperazione.Sincronizzazione
                                sTipoOperazione = "Sincronizza"
                            Case TipoOperazione.SincroniaIntelligente
                                sTipoOperazione = "Sincronia Intelligente"
                            Case TipoOperazione.Spostamento
                                sTipoOperazione = "Sposta"
                            Case TipoOperazione.RiavvioPC
                                sTipoOperazione = "Riavvio"
                            Case TipoOperazione.AvvioServizio
                                sTipoOperazione = "Avvia Servizio"
                            Case TipoOperazione.FermaServizio
                                sTipoOperazione = "Ferma Servizio"
                            Case TipoOperazione.AvviaEseguibile
                                sTipoOperazione = "Avvia EXE"
                            Case TipoOperazione.FermaEseguibile
                                sTipoOperazione = "Ferma EXE"
                            Case TipoOperazione.Attendi
                                sTipoOperazione = "Attendi"
                            Case TipoOperazione.Zip
                                sTipoOperazione = "Zip"
                            Case TipoOperazione.ListaFiles
                                sTipoOperazione = "Lista Files"
                            Case TipoOperazione.Messaggio
                                sTipoOperazione = "Messaggio"
                            Case TipoOperazione.EsegueSQL
                                sTipoOperazione = "Esegue SQL Server"
                        End Select

                        For i As Integer = sTipoOperazione.Length To 23
                            sTipoOperazione = sTipoOperazione & " "
                        Next

                        sOrigine = SistemaLunghezzaCampo(Rec("Origine").Value, 40)
                        sDestinazione = SistemaLunghezzaCampo(Rec("Destinazione").Value, 40)

                        sSottoDirectory = Rec("Sottodirectory").Value
                        sSovrascrivi = Rec("Sovrascrivi").Value

                        sFiltro = Rec("Filtro").Value

                        If lstOperazioni Is Nothing = False Then
                            lstOperazioni.Items.Add(sProgressivo & " " & sTipoOperazione & " " & sOrigine & " " & sDestinazione & " " & sSottoDirectory & " " & sSovrascrivi & " " & Rec("Attivo").Value & " " & sFiltro)
                        End If

                        Rec.MoveNext()
                    Loop
                    Rec.Close()
                End If

                ConnSQL.close()
                ConnSQL = Nothing

                DB.ChiudeDB(True, ConnSQL)
            End If

            DB = Nothing
        Catch ex As Exception
            idProc = -1
        End Try

        Return idProc
    End Function

    Public Function EsegueOperazione(ModalitaEsecuzioneAutomatica As Boolean, PercorsoDBTemp As String, instance As Form, idProc As Integer, Progressivo As Integer, Operazione As Integer, Origine As String, Destinazione As String, Filtro As String,
                            Sovrascrivi As String, SottoDirectory As String, Optional lblOper As Label = Nothing,
                            Optional lblCont As Label = Nothing, Optional UtenzaOrigine As String = "", Optional PasswordOrigine As String = "",
                            Optional UtenzaDest As String = "", Optional PasswordDest As String = "", Optional ModalitaServizio As Boolean = False,
                            Optional clLog As LogCasareccio.LogCasareccio.Logger = Nothing) As StringBuilder
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

        Dim Gf As New GestioneFilesDirectory
        Dim Filetti() As String = {}
        Dim Cartelle() As String = {}
        Dim qFiletti As Long = -1
        Dim qCartelle As Long = -1
        Dim LeggiCartelle As Boolean = False
        Dim log As StringBuilder = New StringBuilder
		Dim NomeFileUltimaOperazione As String = Application.StartupPath & "\UltimaOperazione.txt"

		Select Case Operazione
            Case TipoOperazione.Copia
                LeggiCartelle = True
            Case TipoOperazione.Spostamento
                LeggiCartelle = True
            Case TipoOperazione.EliminaDirectory
                LeggiCartelle = True
        End Select

        If Not BloccaTutto Then
            Dim sOrigine As String = ""
            Dim sDestinazione As String = ""
            Dim mud As New MapUnMapDrives
            Dim Oper As New OperazioniSuFileDettagli

            If ModalitaServizio Then
                Dim discoOriginale As String = Mid(Origine, 1, 2)
                Dim PathOriginale As String = Gf.PrendePercorsoDiReteDelDisco(discoOriginale)
                If PathOriginale <> "" Then
                    Origine = Origine.Replace(discoOriginale, PathOriginale)

                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Rilevato disco di rete di origine:", " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Disco originale: " & discoOriginale, " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Percorso disco:" & PathOriginale, " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Origine: " & Origine, " ", ModalitaServizio, clLog)
                End If

                discoOriginale = Mid(Destinazione, 1, 2)
                PathOriginale = Gf.PrendePercorsoDiReteDelDisco(discoOriginale)
                If PathOriginale <> "" Then
                    Destinazione = Destinazione.Replace(discoOriginale, PathOriginale)

                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Rilevato disco di rete di destinazione:", " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Disco originale: " & discoOriginale, " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Percorso disco:" & PathOriginale, " ", ModalitaServizio, clLog)
                    ScriveOperazione(instance, True, idProc, log, lblOperazione, lblContatore, "Destinazione: " & Destinazione, " ", ModalitaServizio, clLog)
                End If
            End If

            If Not BloccaTutto Then
                Select Case Operazione
                    Case TipoOperazione.Copia, TipoOperazione.Spostamento
                        log = Oper.Copia(idProc, Operazione, Origine, Destinazione, qFiletti, Filetti, Sovrascrivi, lblOperazione, lblContatore, ModalitaServizio, instance, clLog)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.CreaDirectory
                        log = Oper.CreaDirectory(instance, idProc, Origine, Filtro, lblOperazione, lblContatore, ModalitaServizio, clLog)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.EliminaDirectory
                        log = Oper.EliminazioneDirectory(instance, idProc, Origine, lblOperazione, lblContatore, ModalitaServizio, clLog)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.Eliminazione
                        log = Oper.EliminaFile(instance, idProc, Origine, Filtro, SottoDirectory, lblOperazione, lblContatore, ModalitaServizio, clLog)
                        ' ----------------------------------------------------------------------------------------------

                    Case TipoOperazione.SincroniaIntelligente
						log = Oper.Sincronizza(idProc, Progressivo, Origine, Destinazione, Filtro, lblOperazione, lblContatore, True, ModalitaServizio,
											   instance, clLog, ModalitaEsecuzioneAutomatica, PercorsoDBTemp, Gf, NomeFileUltimaOperazione)

						' ----------------------------------------------------------------------------------------------

					Case TipoOperazione.Sincronizzazione
						log = Oper.Sincronizza(idProc, Progressivo, Origine, Destinazione, Filtro, lblOperazione, lblContatore, False, ModalitaServizio,
											   instance, clLog, ModalitaEsecuzioneAutomatica, PercorsoDBTemp, Gf, NomeFileUltimaOperazione)

						' ----------------------------------------------------------------------------------------------
					Case TipoOperazione.RiavvioPC
                        log = Oper.RiavvioPC(idProc, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.FermaServizio
                        FermaServizio(idProc, Origine, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.AvvioServizio
                        FaiPartireServizio(idProc, Origine, Filtro, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.AvviaEseguibile
                        log = Oper.AvviaEseguibile(idProc, Origine, Filtro, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.FermaEseguibile
                        log = Oper.FermaEseguibile(idProc, Origine, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Attendi
                        log = Oper.fAttendi(idProc, Origine, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Zip
                        log = Oper.EsegueZip(idProc, Origine, Destinazione, lblOperazione, lblContatore, ModalitaServizio, instance, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.ListaFiles
                        log = Oper.ListaFiles(idProc, Origine, Destinazione, lblOperazione, lblContatore, ModalitaServizio, instance, clLog)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.Messaggio
                        MsgBox(Filtro, vbInformation)

                        ' ----------------------------------------------------------------------------------------------
                    Case TipoOperazione.EsegueSQL
                        log = Oper.EsegueSQL(idProc, Origine, Destinazione, lblOperazione, lblContatore, ModalitaServizio, instance, clLog)

                        ' ----------------------------------------------------------------------------------------------
                End Select
            End If

            If lblOperazione Is Nothing = False Then
                'lblOperazione.Text = ""
                If instance.InvokeRequired Then
                    instance.Invoke(MethodDelegateAddTextOperazione, "")
                Else
                    lblOperazione.Text = ""
                End If
            End If

            If lblContatore Is Nothing = False Then
                'lblContatore.Text = ""
                If instance.InvokeRequired Then
                    instance.Invoke(MethodDelegateAddTextContatore, "")
                Else
                    lblContatore.Text = ""
                End If
            End If
            Application.DoEvents()

			ScriviUltimaOperazione(Gf, Operazione, Origine, Destinazione, NomeFileUltimaOperazione)

			mud = Nothing
            Oper = Nothing
            Gf = Nothing
        End If

        Return log
    End Function

	Private Sub ScriviUltimaOperazione(Gf As GestioneFilesDirectory, Operazione As Integer, Origine As String, Destinazione As String, NomeFileUltimaOperazione As String)
		Dim Cosa As String = Operazione & ";" & Origine.Replace(";", "***PV***") & ";" & Destinazione.Replace(";", "***PV***") & ";"
		Gf.EliminaFileFisico(NomeFileUltimaOperazione)
		Gf.CreaAggiornaFile(NomeFileUltimaOperazione, cosa)
	End Sub

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
End Class
