Imports System.IO
Imports System.Text
Imports System.Management
Imports System.Windows.Forms
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Timers

Public Structure ModalitaDiScan
    Dim TipologiaScan As Integer
    Const SoloStruttura = 0
    Const Elimina = 1
End Structure

Public Class GestioneFilesDirectory
    Private barra As String = "\"

    Private DirectoryRilevate() As String
    Private FilesRilevati() As String
    Private DimensioneFilesRilevati() As Long
    Private DataFilesRilevati() As Date
    Private QuantiFilesRilevati As Long
    Private QuanteDirRilevate As Long
    Private RootDir As String
    Private Eliminati As Boolean
    Private Percorso As String

    Public Const NonEliminareRoot As Boolean = False
    Public Const EliminaRoot As Boolean = True
    Public Const NonEliminareFiles As Boolean = False
    Public Const EliminaFiles As Boolean = True

    Private DimensioniArrayAttualeDir As Long
    Private DimensioniArrayAttualeFiles As Long

    Private RefreshLabel As Integer = 100
    Private Conta As Integer

    Private Declare Function WNetGetConnection Lib "mpr.dll" Alias _
             "WNetGetConnectionA" (ByVal lpszLocalName As String,
             ByVal lpszRemoteName As String, ByRef cbRemoteName As Integer) As Integer
    Private outputFile As StreamWriter

    Public Sub PrendeRoot(R As String)
        RootDir = R
    End Sub

    Public Function RitornaFilesRilevati() As String()
        Return FilesRilevati
    End Function

    Public Function RitornaDataFilesRilevati() As Date()
        Return DataFilesRilevati
    End Function

    Public Function RitornaDimensioneFilesRilevati() As Long()
        Return DimensioneFilesRilevati
    End Function

    Public Function RitornaDirectoryRilevate() As String()
        Return DirectoryRilevate
    End Function

    Public Function RitornaQuantiFilesRilevati() As Long
        Return QuantiFilesRilevati
    End Function

    Public Function RitornaQuanteDirectoryRilevate() As Long
        Return QuanteDirRilevate
    End Function

    Public Sub ImpostaPercorsoAttuale(sPercorso As String)
        Percorso = sPercorso
    End Sub

    Public Function TornaDimensioneFile(NomeFile As String) As Long
		If File.Exists(NomeFile) Then
			Try
				Dim infoReader As System.IO.FileInfo
				infoReader = My.Computer.FileSystem.GetFileInfo(NomeFile)
				Dim Dime As Long

				Dime = infoReader.Length

				infoReader = Nothing

				Return Dime
			Catch ex As Exception
				Return -1
			End Try
		Else
			Return -1
        End If
    End Function

    Public Function NomeFileEsistente(NomeFile As String) As String
        Dim NomeFileDestinazione As String = NomeFile
        Dim gf As New GestioneFilesDirectory
        Dim Estensione As String = gf.TornaEstensioneFileDaPath(NomeFileDestinazione)
        If Estensione <> "" Then
            NomeFileDestinazione = NomeFileDestinazione.Replace(Estensione, "")
        End If

        Dim Contatore As Integer = 1

        Do While File.Exists(NomeFileDestinazione & "_" & Format(Contatore, "0000") & Estensione) = True
            Contatore += 1
        Loop

        NomeFileDestinazione = NomeFileDestinazione & "_" & Format(Contatore, "0000") & Estensione
        gf = Nothing

        Return NomeFileDestinazione
    End Function

    Public Function EliminaFileFisico(NomeFileOrigine As String) As String
        If File.Exists(NomeFileOrigine) Then
            Dim Ritorno As String = ""

            If NomeFileOrigine.Trim <> "" Then
                Try
                    Dim f As New FileAttribute

                    f = PrendeAttributiFile(NomeFileOrigine)
                    If (f And System.IO.FileAttributes.System) = System.IO.FileAttributes.System Or
                        (f And System.IO.FileAttributes.Hidden) = System.IO.FileAttributes.Hidden Then
                        ImpostaAttributiFile(NomeFileOrigine, FileAttribute.Normal)
                    End If

                    File.Delete(NomeFileOrigine)
                Catch ex As Exception
                    Ritorno = "ERRORE: " & ex.Message
                End Try
            End If

            Return Ritorno
        Else
            Return ""
        End If
    End Function

    Public Function PrendeAttributiFile(Filetto As String) As FileAttribute
        If File.Exists(Filetto) Then
            Dim attributes As FileAttributes
            attributes = File.GetAttributes(Filetto)

            Return attributes
        Else
            Return Nothing
        End If
    End Function

    Public Sub ImpostaAttributiFile(Filetto As String, Attributi As FileAttribute)
        If File.Exists(Filetto) Then
            Try
                File.SetAttributes(Filetto, Attributi)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private tmr As System.Timers.Timer
    Private instance As Form
    Private NomDest As String
    Private NomDestCompleto As String
    Private NomOrig As String
    Private Secondi As Integer
    Private QuantiFiles As String

    Private Sub hdnlCopia(ByVal sender As Object, ByVal e As ElapsedEventArgs)
        If IsNothing(lblAggiornamento) = False Then
			Secondi += 1
			If Secondi <= 500 Then
				Dim dime As Long = TornaDimensioneFile(NomDestCompleto)
				Dim mess As String = "Copia file " & QuantiFiles & ": " & TagliaLunghezzaScritta(NomOrig, 45) & vbCrLf &
				"Dimensioni: " & FormattaNumero(dime, False) & " - Secondi: " & Secondi
				If instance.InvokeRequired Then
					instance.Invoke(MethodDelegateAddText, mess)
				Else
					Me.lblAggiornamento.Text = mess
				End If
			Else
				If Not tmr Is Nothing Then
					tmr.Stop()
					tmr.Enabled = False
				End If
			End If
		End If
	End Sub

	Public Function CopiaFileFisico(NomeFileOrigine As String, NomeFileDestinazione As String, SovraScrittura As Boolean, Optional I As Form = Nothing, Optional lblC As Label = Nothing, Optional S As String = "") As String
		Dim Ritorno As String = ""

		NomDest = TornaNomeFileDaPath(NomeFileDestinazione)
		NomDestCompleto = NomeFileDestinazione
		NomOrig = TornaNomeFileDaPath(NomeFileOrigine)

		If File.Exists(NomeFileOrigine) Then
			If NomeFileOrigine.Trim <> "" And NomeFileDestinazione.Trim <> "" And NomeFileOrigine.Trim.ToUpper <> NomeFileDestinazione.Trim.ToUpper Then
				Dim Ok As Boolean = True

				Dim fOrigine As New FileAttribute

				fOrigine = PrendeAttributiFile(NomeFileOrigine)
				If (fOrigine And System.IO.FileAttributes.System) = System.IO.FileAttributes.System Or
						(fOrigine And System.IO.FileAttributes.Hidden) = System.IO.FileAttributes.Hidden Then
					ImpostaAttributiFile(NomeFileOrigine, FileAttribute.Normal)
				End If

				If File.Exists(NomeFileDestinazione) Then
					If SovraScrittura = False Then
						NomeFileDestinazione = NomeFileEsistente(NomeFileDestinazione)
					Else
						Dim flO As Integer = FileLen(NomeFileOrigine)
						Dim flD As Integer = FileLen(NomeFileDestinazione)
						Dim duaO As Date = FileDateTime(NomeFileOrigine)
						Dim duaD As Date = FileDateTime(NomeFileDestinazione)

						Dim diff As Integer = Math.Abs(DateDiff(DateInterval.Second, duaO, duaD))
						If flO = flD And diff < 60 Then
							Ritorno = "SKIPPED"
							Ok = False
						Else
							EliminaFileFisico(NomeFileDestinazione)
						End If
					End If
				End If

				If Ok Then
					If Not lblC Is Nothing Then
						lblAggiornamento = lblC
						instance = I

						Secondi = 0
						tmr = New System.Timers.Timer(1000)
						AddHandler tmr.Elapsed, New ElapsedEventHandler(AddressOf hdnlCopia)
						tmr.Enabled = True
						QuantiFiles = S
					End If

					Dim dataUltimoAccesso As Date = TornaDataUltimoAccesso(NomeFileOrigine)
					' Dim attr As FileAttribute = PrendeAttributiFile(NomeFileOrigine)
					' ImpostaAttributiFile(NomeFileOrigine, FileAttribute.Normal)

					Try
						File.Copy(NomeFileOrigine, NomeFileDestinazione, True)

						ImpostaAttributiFile(NomeFileDestinazione, fOrigine)
						Ritorno = TornaNomeFileDaPath(NomeFileDestinazione)
					Catch ex As Exception
						Ritorno = "ERRORE: " & ex.Message
					End Try

					ImpostaAttributiFile(NomeFileOrigine, fOrigine)
					Try
						File.SetLastAccessTime(NomeFileOrigine, dataUltimoAccesso)
					Catch ex As Exception

					End Try
				End If
			End If

			If Not lblC Is Nothing Then
				If Not tmr Is Nothing Then
					tmr.Stop()
					tmr.Enabled = False
				End If
			End If

			Return Ritorno
		Else
			If Not tmr Is Nothing Then
				tmr.Stop()
				tmr.Enabled = False
			End If

			Return "ERRORE: File di origine non presente"
		End If
    End Function

    Public Function TornaNomeFileDaPath(Percorso As String) As String
        Dim Ritorno As String = ""

        For i As Integer = Percorso.Length To 1 Step -1
            If Mid(Percorso, i, 1) = "/" Or Mid(Percorso, i, 1) = barra Then
                Ritorno = Mid(Percorso, i + 1, Percorso.Length)
                Exit For
            End If
        Next

        Return Ritorno
    End Function

    Public Function TornaEstensioneFileDaPath(Percorso As String) As String
        Dim Ritorno As String = ""

        For i As Integer = Percorso.Length To 1 Step -1
            If Mid(Percorso, i, 1) = "." Then
                Ritorno = Mid(Percorso, i, Percorso.Length)
                Exit For
            End If
        Next
        If Ritorno.Length > 5 Then
            Ritorno = ""
        End If

        Return Ritorno
    End Function

    Public Function TornaNomeDirectoryDaPath(Percorso As String) As String
        Dim Ritorno As String = ""

        For i As Integer = Percorso.Length To 1 Step -1
            If Mid(Percorso, i, 1) = "/" Or Mid(Percorso, i, 1) = barra Then
                Ritorno = Mid(Percorso, 1, i - 1)
                Exit For
            End If
        Next

        Return Ritorno
    End Function

    Public Sub CreaAggiornaFile(NomeFile As String, Cosa As String)
        Try
            Dim path As String

            If Percorso <> "" Then
                path = Percorso & barra & NomeFile
            Else
                path = NomeFile
            End If

            path = path.Replace(barra & barra, barra)

            ' Create or overwrite the file.
            Dim fs As FileStream = File.Create(path)

            ' Add text to the file.
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(Cosa)
            fs.Write(info, 0, info.Length)
            fs.Close()
        Catch ex As Exception
            'Dim StringaPassaggio As String
            'Dim H As HttpApplication = HttpContext.Current.ApplicationInstance

            'StringaPassaggio = "?Errore=Errore CreaAggiornaFileVisMese: " & Err.Description.Replace(" ", "%20").Replace(vbCrLf, "")
            'StringaPassaggio = StringaPassaggio & "&Utente=" & H.Session("Nick")
            'StringaPassaggio = StringaPassaggio & "&Chiamante=" & H.Request.CurrentExecutionFilePath.ToUpper.Trim
            'H.Response.Redirect("Errore.aspx" & StringaPassaggio)
        End Try
    End Sub

    Private objReader As StreamReader

    Public Sub ApreFilePerLettura(NomeFile As String)
        objReader = New StreamReader(NomeFile)
    End Sub

    Public Function RitornaRiga() As String
        Return objReader.ReadLine()
    End Function

    Public Sub ChiudeFile()
        objReader.Close()
    End Sub

    Public Function LeggeFileIntero(NomeFile As String) As String
        If File.Exists(NomeFile) Then
            Dim objReader As StreamReader = New StreamReader(NomeFile)
            Dim sLine As String = ""
            Dim Ritorno As String = ""

            Do
                sLine = objReader.ReadLine()
                Ritorno += sLine
            Loop Until sLine Is Nothing
            objReader.Close()

            Return Ritorno
        Else
            Return ""
        End If
    End Function

    Private lblAggiornamento As Label
    Private Delegate Sub DelegateAddText(ByVal str As String)
    Private MethodDelegateAddText As New DelegateAddText(AddressOf AddText)

    Public Sub ScansionaDirectorySingola(Percorso As String, Instance As Form, Optional Filtro As String = "", Optional lblAgg As Label = Nothing, Optional SoloRoot As Boolean = False)
        If lblAggiornamento Is Nothing Then
            If Not lblAgg Is Nothing Then
                lblAggiornamento = lblAgg
            End If
        End If

        Eliminati = False

        PulisceInfo()

        If Not Directory.Exists(Percorso) Then
            If Mid(Percorso, Percorso.Length, Percorso.Length) <> "\" Then
                CreaDirectoryDaPercorso(Percorso & "\")
            Else
                CreaDirectoryDaPercorso(Percorso)
            End If
        End If

        QuanteDirRilevate += 1
        DirectoryRilevate(QuanteDirRilevate) = Percorso

        If QuantiFilesRilevati / RefreshLabel = Int(QuantiFilesRilevati / RefreshLabel) Then
            If IsNothing(lblAggiornamento) = False Then
                Dim mess As String = "Lettura " & TagliaLunghezzaScritta(Percorso, 55) & vbCrLf & "Cartelle: " & FormattaNumero(QuanteDirRilevate, False) & " - Files: " & FormattaNumero(QuantiFilesRilevati, False)
                If Instance.InvokeRequired Then
                    Instance.Invoke(MethodDelegateAddText, mess)
                Else
                    Me.lblAggiornamento.Text = mess
                End If
            End If
            Application.DoEvents()
        End If

        LeggeFilesDaDirectory(Instance, Percorso, Filtro, lblAggiornamento)

        If SoloRoot = False Then
            LeggeTutto(Instance, Percorso, Filtro, lblAggiornamento)
        End If

		If IsNothing(lblAggiornamento) = False Then
			Dim mess As String = "Lettura " & TagliaLunghezzaScritta(Percorso, 55) & vbCrLf & "Cartelle: " & FormattaNumero(QuanteDirRilevate, False) & " - Files: " & FormattaNumero(QuantiFilesRilevati, False)
			If Instance.InvokeRequired Then
				Instance.Invoke(MethodDelegateAddText, mess)
			Else
				Me.lblAggiornamento.Text = mess
			End If
		End If
		Application.DoEvents()
	End Sub

	Public Sub ScansionaDirectorySingolaVeloce(Percorso As String, Instance As Form, Optional Filtro As String = "", Optional lblAgg As Label = Nothing, Optional SoloRoot As Boolean = False)
        If lblAggiornamento Is Nothing Then
            If Not lblAgg Is Nothing Then
                lblAggiornamento = lblAgg
            End If
        End If

        Eliminati = False

        PulisceInfo()

        If Not Directory.Exists(Percorso) Then
            If Mid(Percorso, Percorso.Length, Percorso.Length) <> "\" Then
                CreaDirectoryDaPercorso(Percorso & "\")
            Else
                CreaDirectoryDaPercorso(Percorso)
            End If
        End If

        'QuanteDirRilevate += 1
        'DirectoryRilevate(QuanteDirRilevate) = Percorso

        'If QuantiFilesRilevati / RefreshLabel = Int(QuantiFilesRilevati / RefreshLabel) Then
        '    If IsNothing(lblAggiornamento) = False Then
        '        Dim mess As String = "Lettura " & TagliaLunghezzaScritta(Percorso, 70) & vbCrLf & "Cartelle: " & FormattaNumero(QuanteDirRilevate, False) & vbCrLf & "Files: " & FormattaNumero(QuantiFilesRilevati, False)
        '        If Instance.InvokeRequired Then
        '            Instance.Invoke(MethodDelegateAddText, mess)
        '        Else
        '            Me.lblAggiornamento.Text = mess
        '        End If
        '    End If
        '    Application.DoEvents()
        'End If

        'LeggeFilesDaDirectory(Instance, Percorso, Filtro, lblAggiornamento)

        'If SoloRoot = False Then
        '    LeggeTutto(Instance, Percorso, Filtro, lblAggiornamento)
        'End If

        ' Dim f As FileData() = FastDirectoryEnumeratorModule.EnumerateFiles(Dir)
    End Sub

    Private Sub AddText(ByVal str As String)
        lblAggiornamento.Text = str
    End Sub

    Public Function TagliaLunghezzaScritta(Scritta As String, Lunghezza As Integer) As String
        Dim Ritorno As String = ""

        If Scritta.Length > Lunghezza Then
            Ritorno = Mid(Scritta, 1, (Lunghezza / 2) - 2) & "..." & Mid(Scritta, Scritta.Length - ((Lunghezza / 2) - 2), Scritta.Length)
        Else
            Ritorno = Scritta
        End If

        Return Ritorno
    End Function

    Private Sub LeggeTutto(instance As Form, Percorso As String, Filtro As String, Optional lblAggiornamento As Label = Nothing)
        If BloccaTutto Then
            Exit Sub
        End If

        If Directory.Exists(Percorso) Then
            Dim di As New IO.DirectoryInfo(Percorso)
            Dim diar1 As IO.DirectoryInfo() = di.GetDirectories
            Dim dra As IO.DirectoryInfo

            For Each dra In diar1
                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                If BloccaTutto Then
                    Exit For
                End If

                Conta += 1
                If Conta >= RefreshLabel Then
                    Conta = 0
                    If lblAggiornamento Is Nothing = False Then
                        Dim mess As String = "Lettura " & TagliaLunghezzaScritta(Percorso, 55) & vbCrLf & "Cartelle: " & FormattaNumero(QuanteDirRilevate, False) & " - Files: " & FormattaNumero(QuantiFilesRilevati, False)
                        If instance.InvokeRequired Then
                            instance.Invoke(MethodDelegateAddText, mess)
                        Else
                            Me.lblAggiornamento.Text = mess
                        End If
                        Application.DoEvents()
                    End If
                End If

                QuanteDirRilevate += 1
                If QuanteDirRilevate > DimensioniArrayAttualeDir Then
                    DimensioniArrayAttualeDir += 10000
                    ReDim Preserve DirectoryRilevate(DimensioniArrayAttualeDir)
                End If
                DirectoryRilevate(QuanteDirRilevate) = dra.FullName

                LeggeFilesDaDirectory(instance, dra.FullName, Filtro, lblAggiornamento)

                LeggeTutto(instance, dra.FullName, Filtro, lblAggiornamento)
            Next
        End If
    End Sub

    Public Sub PulisceInfo()
        Erase FilesRilevati
        Erase DimensioneFilesRilevati
        Erase DataFilesRilevati
        QuantiFilesRilevati = 0

        Erase DirectoryRilevate
        QuanteDirRilevate = 0

        DimensioniArrayAttualeDir = 10000
        DimensioniArrayAttualeFiles = 10000

        ReDim DirectoryRilevate(DimensioniArrayAttualeDir)
        ReDim FilesRilevati(DimensioniArrayAttualeFiles)
        ReDim DimensioneFilesRilevati(DimensioniArrayAttualeFiles)
        ReDim DataFilesRilevati(DimensioniArrayAttualeFiles)
    End Sub

    Public Function RitornaEliminati() As Boolean
        Return Eliminati
    End Function

    Public Sub LeggeFilesDaDirectory(instance As Form, Percorso As String, Optional Filtro As String = "", Optional lblAggiornamento As Label = Nothing)
        If BloccaTutto Then
            Exit Sub
        End If

        If Directory.Exists(Percorso) Then
            Dim di As New IO.DirectoryInfo(Percorso)
            Dim fiar1 As IO.FileInfo() = di.GetFiles
            Dim fra As IO.FileInfo
            Dim Ok As Boolean = True
            Dim Filtri() As String = Filtro.Split(";")

            For Each fra In fiar1
                If MetteInPausa Then
                    MetteInPausaLaRoutine()
                End If

                Ok = True
                If Filtri.Length > 0 Then
                    Ok = False
                    For Each f In Filtri
                        If fra.FullName.ToUpper.Trim.Contains(f.ToUpper.Trim.Replace("*", "")) Then
                            Ok = True
                            Exit For
                        End If
                    Next
                End If

                If Ok Then
                    QuantiFilesRilevati += 1
                    If QuantiFilesRilevati > DimensioniArrayAttualeFiles Then
                        DimensioniArrayAttualeFiles += 10000
                        ReDim Preserve FilesRilevati(DimensioniArrayAttualeFiles)
                        ReDim Preserve DimensioneFilesRilevati(DimensioniArrayAttualeFiles)
                        ReDim Preserve DataFilesRilevati(DimensioniArrayAttualeFiles)
                    End If
                    FilesRilevati(QuantiFilesRilevati) = fra.FullName
                    DimensioneFilesRilevati(QuantiFilesRilevati) = fra.Length
                    DataFilesRilevati(QuantiFilesRilevati) = fra.LastWriteTime
                End If

                If BloccaTutto Then
                    Exit For
                End If

                If QuantiFilesRilevati / RefreshLabel = Int(QuantiFilesRilevati / RefreshLabel) Then
                    If IsNothing(lblAggiornamento) = False Then
                        Dim mess As String = "Lettura " & TagliaLunghezzaScritta(Percorso, 55) & vbCrLf & "Cartelle: " & FormattaNumero(QuanteDirRilevate, False) & " - Files: " & FormattaNumero(QuantiFilesRilevati, False)
                        If instance.InvokeRequired Then
                            instance.Invoke(MethodDelegateAddText, mess)
                        Else
                            Me.lblAggiornamento.Text = mess
                        End If
                    End If
                    Application.DoEvents()
                End If
            Next
        End If
    End Sub

    Public Sub CreaDirectoryDaPercorso(Percorso As String)
        Dim Ritorno As String = Percorso

        For i As Integer = 1 To Ritorno.Length
            If Mid(Ritorno, i, 1) = barra Then
                Try
                    MkDir(Mid(Ritorno, 1, i))
                Catch ex As Exception

                End Try
            End If
        Next
    End Sub

    Public Function Ordina(Filetti() As String) As String()
        If Filetti Is Nothing Then
            Return {}
            Exit Function
        End If

        Dim Appoggio() As String = Filetti
        Dim Appo As String

        For i As Integer = 1 To Appoggio.Count - 1
            For k As Integer = i + 1 To Appoggio.Count - 1
                If Appoggio(i).ToUpper.Trim > Appoggio(k).ToUpper.Trim Then
                    Appo = Appoggio(i)
                    Appoggio(i) = Appoggio(k)
                    Appoggio(k) = Appo
                End If
            Next
        Next

        Return Appoggio
    End Function

    Public Sub EliminaAlberoDirectory(Percorso As String, EliminaRoot As Boolean, EliminaFiles As Boolean, instance As Form)
        ScansionaDirectorySingola(Percorso, instance, 0)

        DirectoryRilevate = Ordina(DirectoryRilevate)

        If EliminaFiles = True Then
            FilesRilevati = Ordina(FilesRilevati)

            For i As Integer = FilesRilevati.Length - 1 To 1 Step -1
                Try
                    EliminaFileFisico(FilesRilevati(i))
                Catch ex As Exception

                End Try
            Next
        End If

        For i As Integer = DirectoryRilevate.Length - 1 To 1 Step -1
            Try
                RmDir(RitornaDirectoryRilevate(i))
            Catch ex As Exception

            End Try
        Next

        If EliminaRoot = True Then
            Try
                RmDir(Percorso)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Public Function TornaDataDiCreazione(NomeFile As String) As Date
        Dim info As New FileInfo(NomeFile)
        Return info.CreationTime
    End Function

    Public Function TornaDataDiUltimaModifica(NomeFile As String) As Date
        Dim info As New FileInfo(NomeFile)
        Return info.LastWriteTime
    End Function

    Public Function TornaDataUltimoAccesso(NomeFile As String) As Date
        Dim info As New FileInfo(NomeFile)
        Return info.LastAccessTime
    End Function

    Public Function FormattaNumero(Numero As Single, ConVirgola As Boolean, Optional Lunghezza As Integer = -1) As String
        Dim Ritorno As String
        Dim Formattazione As String

        Select Case ConVirgola
            Case True
                Formattazione = "0,000.00"
            Case False
                Formattazione = "0,000"
            Case Else
                Formattazione = "0"
        End Select

        Ritorno = Numero.ToString(Formattazione, CultureInfo.InvariantCulture)

        Do While Left(Ritorno, 1) = "0"
            Ritorno = Mid(Ritorno, 2, Ritorno.Length)
        Loop
        If ConVirgola = True Then
            If Left(Ritorno.Trim, 1) = "." Then
                Ritorno = "0" & Ritorno
            End If
        Else
            If Left(Ritorno.Trim, 1) = "," Then
                Ritorno = Mid(Ritorno, 2, Ritorno.Length)
                For i As Integer = 1 To Ritorno.Length
                    If Mid(Ritorno, i, 1) = "0" Then
                        Ritorno = Mid(Ritorno, 1, i - 1) & "*" & Mid(Ritorno, i + 1, Ritorno.Length)
                    Else
                        Exit For
                    End If
                Next
                Ritorno = Ritorno.Replace("*", "")
                If Ritorno = "" Then Ritorno = "0"
            End If
        End If

        Ritorno = Ritorno.Replace(",", "+")
        Ritorno = Ritorno.Replace(".", ",")
        Ritorno = Ritorno.Replace("+", ".")

        If Ritorno = ".000" Then
            Ritorno = "0"
        End If

        If Lunghezza <> -1 Then
            Dim Spazi As String = ""

            If Ritorno.Length < Lunghezza Then
                For i As Integer = 1 To Lunghezza - Ritorno.Length
                    Spazi += " "
                Next
                Ritorno = Spazi & Ritorno
            End If
        End If

        Return Ritorno
    End Function

    Public Function PrendePercorsoDiReteDelDisco(Lettera As String) As String
        Dim ret As Integer
        Dim out As String = New String(" ", 260)
        Dim len As Integer = 260

        Try
            ret = WNetGetConnection(Lettera, out, len)
        Catch ex As Exception

        End Try

        Return out.Replace(Chr(0), "").Trim
    End Function

    Public Sub ApreFileDiTestoPerScrittura(Percorso As String)
        outputFile = New StreamWriter(Percorso, True)
    End Sub

    Public Sub ScriveTestoSuFileAperto(Cosa As String)
        outputFile.WriteLine(Cosa)
    End Sub

    Public Sub ChiudeFileDiTestoDopoScrittura()
        outputFile.Flush()
        outputFile.Close()
    End Sub
End Class
