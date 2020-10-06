Module Module1

    Sub Main()

        Dim connShare As New ADODB.Connection
        Dim connESolver As New ADODB.Connection

        Dim cmdShare As New ADODB.Command
        Dim cmdESolver As New ADODB.Command

        Dim rstShare As New ADODB.Recordset
        Dim rstESolver As New ADODB.Recordset

        Dim iniFilePath As String = My.Application.Info.DirectoryPath
        Dim iniFile As String = iniFilePath & "\SharePDM.ini"

        Dim ServerESolver As String             ' ==> Server eSolver
        Dim UserESolver As String               ' ==> User eSolver
        Dim PwdESolver As String                ' ==> Password eSolver
        Dim CatESolver As String                ' ==> Catalogo eSolver

        Dim ServerShare As String               ' ==> Server Share PDM
        Dim UserShare As String                 ' ==> User Share PDM
        Dim PwdShare As String                  ' ==> Password Share PDM
        Dim CatShare As String                  ' ==> Catalogo Share PDM

        Dim GruppoArchivi As String

        ' MATERIA_PRIMA_RADICE
        Dim CodValuta As String
        Dim Cambio As Double
        Dim Costo As Double
        Dim PesoSpecifico As Double
        Dim AltezzaTessuto As Double
        Dim GruppoMat As String
        Dim TipoMateriale As String
        Dim FlagTgMis As String
        Dim DescrMateriale As String
        Dim DescrEstesaMateriale
        Dim DescrCompletaMateriale
        Dim FatConvConsDiba As Double
        Dim NumDecCosto As Long
        Dim NumDecPesoSpecifico As Long
        Dim NumDecAltezzaTessuto As Long
        Dim DataVar As Date
        Dim Annullato As Boolean

        ' MATERIA_PRIMA_COLORE
        Dim Descrizione As String

        ' ESOLVER_FORNITORI
        Dim RagSociale As String
        Dim RagSocialeEstesa As String
        Dim Indirizzo As String
        Dim Localita As String
        Dim Tipo As String

        ' TAB_COMPOSIZIONI
        Dim DescrizioneEstesa As String

        ' SEMAFORO_MATERIALI
        Dim Semaforo As Byte
        Dim RecSemaforo As Boolean
        Dim Data_Esolver
        Dim Data_Lectra

        Dim DataEsolver As Date
        Dim DataOggi As Date
        Dim DataNull As Date

        Dim SQL As String

        ' -- LETTURA FILE INI
        Dim sb As System.Text.StringBuilder
        sb = New System.Text.StringBuilder(500)
        Dim Sezione As String = "Sezione1"
        Dim VL_FileName As String = iniFile

        GetPrivateProfileString(Sezione, "dseSolver", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            ServerESolver = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "usereSolver", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            UserESolver = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "pwdeSolver", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            PwdESolver = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "cateSolver", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CatESolver = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "gruppo", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            GruppoArchivi = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "dsSharePDM", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            ServerShare = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "userSharePDM", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            UserShare = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "pwdSharePDM", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            PwdShare = sb.ToString
        End If

        GetPrivateProfileString(Sezione, "catSharePDM", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CatShare = sb.ToString
        End If

        ' ==> Collegamento al Database eSolver
        Dim ConnessioneEsolver As String = "Provider=sqloledb;user id = '" & UserESolver & "'; password = '" & PwdESolver & "' ; initial catalog= '" & CatESolver & "' ; data source = '" & ServerESolver & "'"
        connESolver.Open(ConnessioneEsolver)

        cmdESolver.ActiveConnection = connESolver

        ' ==> Collegamento al Database Share PDM
        Dim ConnessionePDM As String = "Provider=sqloledb;user id = '" & UserShare & "'; password = '" & PwdShare & "' ; initial catalog= '" & CatShare & "' ; data source = '" & ServerShare & "'"
        connShare.Open(ConnessionePDM)

        cmdShare.ActiveConnection = connShare

        'MsgBox("Connessione")

        ' ==> Verifico il semaforo SEMAFORO_MATERIALI
        SQL = "SELECT * FROM SEMAFORO_MATERIALI"
        rstShare.Open(SQL, connShare, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        If rstShare.RecordCount = 0 Then
            Semaforo = 1
            RecSemaforo = False
        Else
            Do Until rstShare.EOF
                RecSemaforo = True
                Semaforo = rstShare.Fields("VALORE").Value
                If Not IsDBNull(rstShare.Fields("Data_Esolver").Value) Then
                    Data_Esolver = rstShare.Fields("Data_Esolver").Value
                    Data_Lectra = rstShare.Fields("Data_Lectra").Value

                    ' sottraggo 5 minuti per essere sicura di non perdere dei dati avendo problemi nell'impostazione
                    ' della data esolver (imposto la data/ora corrente al termine della procedura mentre sarebbe piu'
                    ' corretto riportare la data/ora all'inizio della procedura
                    Data_Esolver = DateAdd("n", -5, Data_Esolver)

                    DataEsolver = CDate(Day(Data_Esolver) & "/" & Month(Data_Esolver) & "/" & Year(Data_Esolver))
                Else
                    DataEsolver = #1/1/1800#
                End If
                Exit Do
            Loop
        End If
        rstShare.Close()

        ' Se il semaforo è rosso (=0) significa che Lectra non ha ancora importato i dati dell'ultimo aggiornamento
        ' quindi non eseguo nulla
        If Semaforo = 0 Then
            GoTo Chiusura_Connessione
        End If

        ' altrimenti procedo con l'aggiornamento dei dati modificati dalla data Data_Esolver
        DataOggi = Date.Now
        DataNull = #1/1/1800#

        ' Imposto il semaforo a rosso per Lectra (=1)
        If RecSemaforo = False Then
            cmdShare.CommandText = "INSERT INTO SEMAFORO_MATERIALI (VALORE) Values(1);"
            cmdShare.Execute()
        End If

        ' ==> Svuoto e ripopolo la tabella ESOLVER_MATERIA_PRIMA_RADICE
        cmdShare.CommandText = "DELETE FROM ESOLVER_MATERIA_PRIMA_RADICE;"
        cmdShare.Execute()

        'MsgBox("ESOLVER_MATERIA_PRIMA_RADICE")

        ' Select con filtro tipologia articoli = Materia prima
        ' Venivano esclusi i semi lavorati che ora vengono passati
        SQL = "SELECT ArtAnagrafica.CodArt, ArtAnagrafica.DesArt, ArtAnagrafica.CodFamiglia, ArtAnagrafica.CodFamiglia, ArtAnagrafica.FirmaUltVarData, ModaArticoli.CostoSchedaProd, ModaArticoli.AltezzaTessuto, ModaArticoli.PesoSpecifico, ArtAnagrafica.DesArt, ArtAnagrafica.DesArt, ModaArticoli.CodComposizione1, ModaTabComposizioni.Descrizione_1, ArtAnagrafica.TecniciUm, ArtAnagrafica.AcqCodForAbituale, ArtAnagrafica.AcqUm, ModaArticoli.CodClassifComponente, ModaArticoli.CodiceTabellaTaglie, TipoArt, ArtAnagrafica.DesEstesa, StatoArt " &
        "FROM UnitaDiMisura RIGHT JOIN (((ArtAnagrafica LEFT JOIN ModaArticoli ON (ArtAnagrafica.CodArt = ModaArticoli.CodiceArticolo) AND (ArtAnagrafica.DBGruppo = ModaArticoli.DBGruppo)) LEFT JOIN ModaTabComposizioni ON (ModaArticoli.DBGruppo = ModaTabComposizioni.DBGruppo) AND (ModaArticoli.CodComposizione1 = ModaTabComposizioni.CodiceComposizione)) LEFT JOIN Famiglia ON (ArtAnagrafica.DBGruppo = Famiglia.DBGruppo) AND (ArtAnagrafica.CodFamiglia = Famiglia.CodFamiglia)) ON UnitaDiMisura.CodUnitaMisura = ArtAnagrafica.TecniciUm " &
        "WHERE ArtAnagrafica.DBGruppo='" & GruppoArchivi & "' AND ArtAnagrafica.DesArt Is Not Null AND ArtAnagrafica.TipoAnagr=1 AND ArtAnagrafica.DesArt>' '  and (TipoArt=0 or TipoArt=1) AND " &
        "ArtAnagrafica.FirmaUltVarData>='" & DataEsolver & "'"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        CodValuta = "EUR"
        Cambio = 1

        Do Until rstESolver.EOF
            GruppoMat = Left(rstESolver.Fields("CodFamiglia").Value, 3)

            TipoMateriale = ""
            FlagTgMis = ""
            CodValuta = ""
            Cambio = 0
            DescrMateriale = ""
            DescrEstesaMateriale = ""
            DescrCompletaMateriale = ""
            FatConvConsDiba = 1
            Annullato = False

            If rstESolver.Fields("StatoArt").Value = 1 Then
                Annullato = True
            Else
                Annullato = False
            End If

            If Left(rstESolver.Fields("CodArt").Value, 3) < "011" Then
                TipoMateriale = "F"
            Else
                TipoMateriale = "A"
            End If

            If Len(rstESolver.Fields("CodiceTabellaTaglie").Value) > 0 Then
                FlagTgMis = "1"
            Else
                FlagTgMis = ""
            End If

            If Not IsDBNull(rstESolver.Fields("DesArt").Value) Then
                DescrMateriale = ""
                DescrMateriale = Replace(rstESolver.Fields("DesArt").Value, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("DesEstesa").Value) Then
                DescrEstesaMateriale = ""
                DescrEstesaMateriale = Replace(rstESolver.Fields("DesEstesa").Value, "'", " ")
            End If

            DescrCompletaMateriale = DescrMateriale & " " & DescrEstesaMateriale

            If Not IsDBNull(rstESolver.Fields("CostoSchedaProd").Value) Then
                If InStr(1, CStr(rstESolver.Fields("CostoSchedaProd").Value), ",") > 0 Then
                    NumDecCosto = Len(CStr(rstESolver.Fields("CostoSchedaProd").Value)) - InStr(1, CStr(rstESolver.Fields("CostoSchedaProd").Value), ",")
                    Costo = TogliDec(rstESolver.Fields("CostoSchedaProd").Value, NumDecCosto)
                Else
                    NumDecCosto = 0
                    Costo = rstESolver.Fields("CostoSchedaProd").Value
                End If
            Else
                NumDecCosto = 0
                Costo = 0
            End If

            If Not IsDBNull(rstESolver.Fields("PesoSpecifico").Value) Then
                If InStr(1, CStr(rstESolver.Fields("PesoSpecifico").Value), ",") > 0 Then
                    NumDecPesoSpecifico = Len(CStr(rstESolver.Fields("PesoSpecifico").Value)) - InStr(1, CStr(rstESolver.Fields("PesoSpecifico").Value), ",")
                    PesoSpecifico = TogliDec(rstESolver.Fields("PesoSpecifico").Value, NumDecPesoSpecifico)
                Else
                    NumDecPesoSpecifico = 0
                    PesoSpecifico = rstESolver.Fields("PesoSpecifico").Value
                End If
            Else
                NumDecPesoSpecifico = 0
                PesoSpecifico = 0
            End If

            If Not IsDBNull(rstESolver.Fields("AltezzaTessuto").Value) Then
                If InStr(1, CStr(rstESolver.Fields("AltezzaTessuto").Value), ",") > 0 Then
                    NumDecAltezzaTessuto = Len(CStr(rstESolver.Fields("AltezzaTessuto").Value)) - InStr(1, CStr(rstESolver.Fields("AltezzaTessuto").Value), ",")
                    AltezzaTessuto = TogliDec(rstESolver.Fields("AltezzaTessuto").Value, NumDecAltezzaTessuto)
                Else
                    NumDecAltezzaTessuto = 0
                    AltezzaTessuto = rstESolver.Fields("AltezzaTessuto").Value
                End If
            Else
                NumDecAltezzaTessuto = 0
                AltezzaTessuto = 0
            End If
            ' !!!! AGGIUNGERE QUI LA GESTIONE DELL'ANNULLAMENTO MATERIALE
            ' !!!! QUANDO LECTRA MI DICE IL NOME DEL CAMP DEL FLAG DI ANNULLAMENTO SETTARLO CON LA VARIABILE ANNULLATO
            If IsDBNull(rstESolver.Fields("FirmaUltVarData").Value) = True Then
                ' Non passo la data variazione
                cmdShare.CommandText = "INSERT INTO ESOLVER_MATERIA_PRIMA_RADICE ( MATR_CODICE, MATR_DESCR_MATERIALE, MATR_GRUPPO_MAT, MATR_SUBGRUPPO_MAT, MATR_COSTO, MATR_COD_VALUTA, MATR_CAMBIO, MATR_ALTEZZA_STD, MATR_PESO_AREA, MATR_DESCR1, MATR_DESCR2, MATR_COD_COMPOSIZIONE, MATR_COMPOSIZIONE, MATR_UMED, MATR_FAT_CONV_CONS_DIBA, MATR_CODICE_FORNITORE, MATR_UMACQ, MATR_TIPO_MATERIALE, MATR_FASE_DEFAULT, MATR_FLAG_TG_MIS) " &
                                       "Values('" & rstESolver.Fields("CodArt").Value & "', '" & DescrMateriale & "', '" & GruppoMat & "', '" & rstESolver.Fields("CodFamiglia").Value & "', " & Costo & ", '" & CodValuta & "', " & Cambio & ", " & AltezzaTessuto & ", " & PesoSpecifico & ", '" & Left(DescrEstesaMateriale, 500) & "', '" & Left(DescrCompletaMateriale, 1000) & "', '" & rstESolver.Fields("CodComposizione1").Value & "', '" & rstESolver.Fields("Descrizione_1").Value & "', '" & rstESolver.Fields("TecniciUm").Value & "', " & FatConvConsDiba & ", '" & rstESolver.Fields("AcqCodForAbituale").Value & "', '" & rstESolver.Fields("AcqUm").Value & "', '" & TipoMateriale & "', '" & rstESolver.Fields("CodClassifComponente").Value & "', '" & FlagTgMis & "');"

            Else
                DataVar = rstESolver.Fields("FirmaUltVarData").Value

                cmdShare.CommandText = "INSERT INTO ESOLVER_MATERIA_PRIMA_RADICE ( MATR_CODICE, MATR_DESCR_MATERIALE, MATR_GRUPPO_MAT, MATR_SUBGRUPPO_MAT, MATR_COSTO, MATR_COD_VALUTA, MATR_CAMBIO, MATR_ALTEZZA_STD, MATR_PESO_AREA, MATR_DESCR1, MATR_DESCR2, MATR_COD_COMPOSIZIONE, MATR_COMPOSIZIONE, MATR_UMED, MATR_FAT_CONV_CONS_DIBA, MATR_CODICE_FORNITORE, MATR_UMACQ, MATR_TIPO_MATERIALE, MATR_FASE_DEFAULT, MATR_FLAG_TG_MIS, MATR_DATAORA_ERP) " &
                                       "Values('" & rstESolver.Fields("CodArt").Value & "', '" & DescrMateriale & "', '" & GruppoMat & "', '" & rstESolver.Fields("CodFamiglia").Value & "', " & Costo & ", '" & CodValuta & "', " & Cambio & ", " & AltezzaTessuto & ", " & PesoSpecifico & ", '" & Left(DescrEstesaMateriale, 500) & "', '" & Left(DescrCompletaMateriale, 1000) & "', '" & rstESolver.Fields("CodComposizione1").Value & "', '" & rstESolver.Fields("Descrizione_1").Value & "', '" & rstESolver.Fields("TecniciUm").Value & "', " & FatConvConsDiba & ", '" & rstESolver.Fields("AcqCodForAbituale").Value & "', '" & rstESolver.Fields("AcqUm").Value & "', '" & TipoMateriale & "', '" & rstESolver.Fields("CodClassifComponente").Value & "', '" & FlagTgMis & "', '" & DataVar & "');"
            End If

            cmdShare.Execute()

            Call MettiDec("MATR_COSTO", NumDecCosto, rstESolver.Fields("CodArt").Value, connShare, cmdShare)
            Call MettiDec("MATR_PESO_AREA", NumDecPesoSpecifico, rstESolver.Fields("CodArt").Value, connShare, cmdShare)
            Call MettiDec("MATR_ALTEZZA_STD", NumDecAltezzaTessuto, rstESolver.Fields("CodArt").Value, connShare, cmdShare)

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("Fine radice")

        ' ==> Svuoto e ripopolo la tabella ESOLVER_MATERIA_PRIMA_COLORE
        ' Come costo inizialmente passo quello di anagrafica materiale
        cmdShare.CommandText = "DELETE FROM ESOLVER_MATERIA_PRIMA_COLORE;"
        cmdShare.Execute()

        'MsgBox("ESOLVER_MATERIA_PRIMA_COLORE")

        ' Select con filtro tipologia articoli = Materia prima
        ' Venivano esclusi i semi lavorati che ora vengono passati
        SQL = "SELECT ArtConfigVariante.CodArt, ArtConfigVariante.VarianteArt, ArtConfigVariante.Descrizione, 1 AS Espr1, StatoArt, ArtConfigVariante.DataFineValidita, ModaArticoli.CostoSchedaProd " &
          "FROM (ArtConfigVariante INNER JOIN ArtAnagrafica ON (ArtConfigVariante.CodArt = ArtAnagrafica.CodArt) AND (ArtConfigVariante.DBGruppo = ArtAnagrafica.DBGruppo)) LEFT JOIN ModaArticoli ON (ArtAnagrafica.CodArt = ModaArticoli.CodiceArticolo) AND (ArtAnagrafica.DBGruppo = ModaArticoli.DBGruppo) " &
          "WHERE ArtConfigVariante.DBGruppo ='" & GruppoArchivi & "' AND ArtAnagrafica.DesArt Is Not Null AND ArtAnagrafica.TipoAnagr=1 AND ArtAnagrafica.DesArt>' 'and (TipoArt=0 or TipoArt=1) AND ArtConfigVariante.VarianteArt Is Not Null AND ArtConfigVariante.VarianteArt>'' AND " &
          "ArtAnagrafica.FirmaUltVarData>='" & DataEsolver & "'"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstESolver.EOF
            Descrizione = ""
            Annullato = False

            If rstESolver.Fields("StatoArt").Value = 1 Then
                Annullato = True
            Else
                If rstESolver.Fields("DataFineValidita").Value < Date.Now And rstESolver.Fields("DataFineValidita").Value <> "01/01/1800" Then
                    Annullato = True
                Else
                    Annullato = False
                End If
            End If

            If Not IsDBNull(rstESolver.Fields("Descrizione").Value) Then
                Descrizione = Replace(rstESolver.Fields("Descrizione").Value, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("CostoSchedaProd").Value) Then
                If InStr(1, CStr(rstESolver.Fields("CostoSchedaProd").Value), ",") > 0 Then
                    NumDecCosto = Len(CStr(rstESolver.Fields("CostoSchedaProd").Value)) - InStr(1, CStr(rstESolver.Fields("CostoSchedaProd").Value), ",")
                    Costo = TogliDec(rstESolver.Fields("CostoSchedaProd").Value, NumDecCosto)
                Else
                    NumDecCosto = 0
                    Costo = rstESolver.Fields("CostoSchedaProd").Value
                End If
            Else
                NumDecCosto = 0
                Costo = 0
            End If

            cmdShare.CommandText = "INSERT INTO ESOLVER_MATERIA_PRIMA_COLORE ( MATC_CODICE_MATERIALE, MATC_CODICE_COLORE, MATC_DESCR_COLOR, MATC_ABILITATO, MATC_COSTO_COLORE  ) " &
                                   "Values('" & rstESolver.Fields("CodArt").Value & "', '" & rstESolver.Fields("VarianteArt").Value & "', '" & Descrizione & "', 1, " & Costo & ");"
            cmdShare.Execute()

            Call Colori_MettiDec("MATC_COSTO_COLORE", NumDecCosto, rstESolver.Fields("CodArt").Value, rstESolver.Fields("VarianteArt").Value, connShare, cmdShare)

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("TEST 1")

        ' Aggiorno il costo con quello per variante (se presente)
        SQL = "SELECT ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE, ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_COLORE " &
              "FROM ESOLVER_MATERIA_PRIMA_COLORE;"

        rstShare.Open(SQL, connShare, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstShare.EOF
            ' Sostituita select su LISTINI_RIGHE con vista JSV_LIS_PRZVALDATA per migrazione 4.2
            SQL = "SELECT JSV_LIS_PRZVALDATA.DBGruppo, JSV_LIS_PRZVALDATA.ClasseListino, JSV_LIS_PRZVALDATA.CodListino, JSV_LIS_PRZVALDATA.CodArticolo, JSV_LIS_PRZVALDATA.VarArt, JSV_LIS_PRZVALDATA.PrezzoLis " &
                  "From JSV_LIS_PRZVALDATA " &
                  "WHERE JSV_LIS_PRZVALDATA.DBGruppo='" & GruppoArchivi & "' AND JSV_LIS_PRZVALDATA.ClasseListino=2 AND JSV_LIS_PRZVALDATA.CodListino='CSKP' AND JSV_LIS_PRZVALDATA.CodArticolo='" & rstShare.Fields("MATC_CODICE_MATERIALE").Value & "' AND JSV_LIS_PRZVALDATA.VarArt='" & rstShare.Fields("MATC_CODICE_COLORE").Value & "';"

            rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

            If rstESolver.RecordCount > 0 Then
                If Not IsDBNull(rstESolver.Fields("PrezzoLis").Value) And rstESolver.Fields("PrezzoLis").Value > 0 Then

                    Costo = rstESolver.Fields("PrezzoLis").Value

                    Dim CostoTXT = CStr(rstESolver.Fields("PrezzoLis").Value)
                    CostoTXT = Replace(CostoTXT, ",", ".")

                    SQL = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET MATC_COSTO_COLORE=" & CostoTXT & " where MATC_CODICE_MATERIALE='" & rstESolver.Fields("CodArticolo").Value & "' AND MATC_CODICE_COLORE='" & rstESolver.Fields("VarArt").Value & "'"
                    cmdShare.CommandText = SQL
                    cmdShare.Execute()
                End If
            End If
            rstESolver.Close()
            rstShare.MoveNext()
        Loop
        rstShare.Close()

        'MsgBox("fine colore")

        ' ==> Elimino ESOLVER_MATERIA_PRIMA_RADICE i materiali che non hanno record in tabella colore
        SQL = "SELECT ESOLVER_MATERIA_PRIMA_RADICE.*, ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE " &
              "FROM ESOLVER_MATERIA_PRIMA_RADICE LEFT JOIN ESOLVER_MATERIA_PRIMA_COLORE ON ESOLVER_MATERIA_PRIMA_RADICE.MATR_CODICE = ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE " &
              "WHERE ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE Is Null;"

        rstShare.Open(SQL, connShare, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstShare.EOF
            cmdShare.CommandText = "DELETE FROM ESOLVER_MATERIA_PRIMA_COLORE " &
                                   "WHERE ESOLVER_MATERIA_PRIMA_RADICE.MATR_CODICE='" & rstShare.Fields("MATR_CODICE").Value & "';"
            cmdShare.Execute()

            'MsgBox("ESOLVER_MATERIA_PRIMA_COLORE")

            rstShare.MoveNext()
        Loop
        rstShare.Close()

        ' ==> Svuoto e ripopolo la tabella ESOLVER_MATERIA_PRIMA_MISURE
        SQL = "DELETE FROM ESOLVER_MATERIA_PRIMA_MISURE;"
        cmdShare.CommandText = SQL
        cmdShare.Execute()

        'MsgBox("ESOLVER_MATERIA_PRIMA_MISURE")

        ' Accodo riga in tabella ESOLVER_MATERIA_PRIMA_MISURE
        SQL = "SELECT ModaArticoli.CodiceArticolo, ModaArticoli.CodiceTabellaTaglie, FlagTagliaValida_1, CodiciTaglie_1, DescrizioneTaglie_1, FlagTagliaValida_2, CodiciTaglie_2, DescrizioneTaglie_2, FlagTagliaValida_3, CodiciTaglie_3, DescrizioneTaglie_3, FlagTagliaValida_4, CodiciTaglie_4, ModaTabellaTaglie.DescrizioneTaglie_4, FlagTagliaValida_5, CodiciTaglie_5, DescrizioneTaglie_5, " &
              "FlagTagliaValida_6, CodiciTaglie_6, DescrizioneTaglie_6, FlagTagliaValida_7, CodiciTaglie_7, DescrizioneTaglie_7, FlagTagliaValida_8, CodiciTaglie_8, DescrizioneTaglie_8, FlagTagliaValida_9, CodiciTaglie_9, DescrizioneTaglie_9, FlagTagliaValida_10, CodiciTaglie_10 , DescrizioneTaglie_10, " &
              "FlagTagliaValida_11, CodiciTaglie_11, DescrizioneTaglie_11, FlagTagliaValida_12, CodiciTaglie_12, DescrizioneTaglie_12, FlagTagliaValida_13, CodiciTaglie_13, DescrizioneTaglie_13, FlagTagliaValida_14, CodiciTaglie_14, DescrizioneTaglie_14, FlagTagliaValida_15, CodiciTaglie_15, DescrizioneTaglie_15 , " &
              "FlagTagliaValida_16 , CodiciTaglie_16, DescrizioneTaglie_16, FlagTagliaValida_17, CodiciTaglie_17, DescrizioneTaglie_17, FlagTagliaValida_18, CodiciTaglie_18, DescrizioneTaglie_18, FlagTagliaValida_19, CodiciTaglie_19, DescrizioneTaglie_19 , FlagTagliaValida_20, CodiciTaglie_20, DescrizioneTaglie_20, " &
              "FlagTagliaValida_21, CodiciTaglie_21, DescrizioneTaglie_21, FlagTagliaValida_22, CodiciTaglie_22, DescrizioneTaglie_22, FlagTagliaValida_23, CodiciTaglie_23, DescrizioneTaglie_23, FlagTagliaValida_24, CodiciTaglie_24, DescrizioneTaglie_24, FlagTagliaValida_25 , CodiciTaglie_25, DescrizioneTaglie_25, " &
              "FlagTagliaValida_26, CodiciTaglie_26, DescrizioneTaglie_26, FlagTagliaValida_27, CodiciTaglie_27, DescrizioneTaglie_27, FlagTagliaValida_28, CodiciTaglie_28, DescrizioneTaglie_28, FlagTagliaValida_29, CodiciTaglie_29, DescrizioneTaglie_29, FlagTagliaValida_30, CodiciTaglie_30, DescrizioneTaglie_30, " &
              "ArtAnagrafica.TipoArt, ArtAnagrafica.DesArt " &
              "FROM ArtAnagrafica INNER JOIN (ModaArticoli LEFT JOIN ModaTabellaTaglie ON ModaArticoli.CodiceTabellaTaglie = ModaTabellaTaglie.CodiceTabellaTaglie) ON (ArtAnagrafica.CodArt = ModaArticoli.CodiceArticolo) AND (ArtAnagrafica.DBGruppo = ModaArticoli.DBGruppo) " &
              "WHERE (((ArtAnagrafica.DBGruppo)='" & GruppoArchivi & "') AND ((ArtAnagrafica.TipoAnagr)=1) AND ((ArtAnagrafica.TipoArt)=0 OR (ArtAnagrafica.TipoArt)=1) AND ((ArtAnagrafica.DesArt) Is Not Null And (ArtAnagrafica.DesArt)>' ')) AND " &
              "ArtAnagrafica.FirmaUltVarData>='" & DataEsolver & "'"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstESolver.EOF
            ' Se è un articolo a taglie
            If Len(rstESolver.Fields("CodiceTabellaTaglie").Value) > 0 Then
                Dim articolo = rstESolver.Fields("CodiceArticolo").Value
                If rstESolver.Fields("FlagTagliaValida_1").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_1").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 1, rstESolver.Fields("CodiciTaglie_1").Value, rstESolver.Fields("DescrizioneTaglie_1").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_2").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_2").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 2, rstESolver.Fields("CodiciTaglie_2").Value, rstESolver.Fields("DescrizioneTaglie_2").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_3").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_3").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 3, rstESolver.Fields("CodiciTaglie_3").Value, rstESolver.Fields("DescrizioneTaglie_3").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_4").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_4").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 4, rstESolver.Fields("CodiciTaglie_4").Value, rstESolver.Fields("DescrizioneTaglie_4").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_5").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_5").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 5, rstESolver.Fields("CodiciTaglie_5").Value, rstESolver.Fields("DescrizioneTaglie_5").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_6").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_6").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 6, rstESolver.Fields("CodiciTaglie_6").Value, rstESolver.Fields("DescrizioneTaglie_6").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_7").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_7").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 7, rstESolver.Fields("CodiciTaglie_7").Value, rstESolver.Fields("DescrizioneTaglie_7").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_8").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_8").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 8, rstESolver.Fields("CodiciTaglie_8").Value, rstESolver.Fields("DescrizioneTaglie_8").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_9").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_9").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 9, rstESolver.Fields("CodiciTaglie_9").Value, rstESolver.Fields("DescrizioneTaglie_9").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_10").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_10").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 10, rstESolver.Fields("CodiciTaglie_10").Value, rstESolver.Fields("DescrizioneTaglie_10").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_11").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_11").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 11, rstESolver.Fields("CodiciTaglie_11").Value, rstESolver.Fields("DescrizioneTaglie_11").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_12").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_12").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 12, rstESolver.Fields("CodiciTaglie_12").Value, rstESolver.Fields("DescrizioneTaglie_12").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_13").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_13").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 13, rstESolver.Fields("CodiciTaglie_13").Value, rstESolver.Fields("DescrizioneTaglie_13").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_14").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_14").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 14, rstESolver.Fields("CodiciTaglie_14").Value, rstESolver.Fields("DescrizioneTaglie_14").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_15").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_15").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 15, rstESolver.Fields("CodiciTaglie_15").Value, rstESolver.Fields("DescrizioneTaglie_15").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_16").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_16").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 16, rstESolver.Fields("CodiciTaglie_16").Value, rstESolver.Fields("DescrizioneTaglie_16").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_17").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_17").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 17, rstESolver.Fields("CodiciTaglie_17").Value, rstESolver.Fields("DescrizioneTaglie_17").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_18").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_18").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 18, rstESolver.Fields("CodiciTaglie_18").Value, rstESolver.Fields("DescrizioneTaglie_18").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_19").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_19").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 19, rstESolver.Fields("CodiciTaglie_19").Value, rstESolver.Fields("DescrizioneTaglie_19").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_20").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_20").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 20, rstESolver.Fields("CodiciTaglie_20").Value, rstESolver.Fields("DescrizioneTaglie_20").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_21").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_21").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 21, rstESolver.Fields("CodiciTaglie_21").Value, rstESolver.Fields("DescrizioneTaglie_21").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_22").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_22").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 22, rstESolver.Fields("CodiciTaglie_22").Value, rstESolver.Fields("DescrizioneTaglie_22").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_23").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_23").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 23, rstESolver.Fields("CodiciTaglie_23").Value, rstESolver.Fields("DescrizioneTaglie_23").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_24").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_24").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 24, rstESolver.Fields("CodiciTaglie_24").Value, rstESolver.Fields("DescrizioneTaglie_24").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_25").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_25").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 25, rstESolver.Fields("CodiciTaglie_25").Value, rstESolver.Fields("DescrizioneTaglie_25").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_26").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_26").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 26, rstESolver.Fields("CodiciTaglie_26").Value, rstESolver.Fields("DescrizioneTaglie_26").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_27").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_27").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 27, rstESolver.Fields("CodiciTaglie_27").Value, rstESolver.Fields("DescrizioneTaglie_27").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_28").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_28").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 28, rstESolver.Fields("CodiciTaglie_28").Value, rstESolver.Fields("DescrizioneTaglie_28").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_29").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_29").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 29, rstESolver.Fields("CodiciTaglie_29").Value, rstESolver.Fields("DescrizioneTaglie_29").Value, connShare, cmdShare)
                End If

                If rstESolver.Fields("FlagTagliaValida_30").Value = 1 And Len(Trim(rstESolver.Fields("CodiciTaglie_30").Value)) > 0 Then
                    Call CaricaTaglia(rstESolver.Fields("CodiceArticolo").Value, 30, rstESolver.Fields("CodiciTaglie_30").Value, rstESolver.Fields("DescrizioneTaglie_30").Value, connShare, cmdShare)
                End If

            End If

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("Fine taglia")

        ' ==> Elimino ESOLVER_MATERIA_PRIMA_MISURE i materiali che non hanno record in tabella colore
        '     (come fatto per ESOLVCER_MATERIA_PRIMA_RADICE)
        SQL = "SELECT ESOLVER_MATERIA_PRIMA_MISURE.MATM_CODICE_MATERIALE, ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE " &
              "FROM ESOLVER_MATERIA_PRIMA_MISURE LEFT JOIN ESOLVER_MATERIA_PRIMA_COLORE ON ESOLVER_MATERIA_PRIMA_MISURE.MATM_CODICE_MATERIALE = ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE " &
              "GROUP BY ESOLVER_MATERIA_PRIMA_MISURE.MATM_CODICE_MATERIALE, ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE " &
              "HAVING (((ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_MATERIALE) Is Null));"

        rstShare.Open(SQL, connShare, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstShare.EOF
            cmdShare.CommandText = "DELETE FROM ESOLVER_MATERIA_PRIMA_MISURE " &
                                   "WHERE ESOLVER_MATERIA_PRIMA_MISURE.MATM_CODICE_MATERIALE='" & rstShare.Fields("MATM_CODICE_MATERIALE").Value & "';"
            cmdShare.Execute()

            'MsgBox("ESOLVER_MATERIA_PRIMA_MISURE")

            rstShare.MoveNext()
        Loop
        rstShare.Close()

        ' ==> Svuoto e ripopolo la tabella ESOLVER_FORNITORI
        SQL = "DELETE FROM ESOLVER_FORNITORI;"
        cmdShare.CommandText = SQL
        cmdShare.Execute()

        'MsgBox("ESOLVER_FORNITORI")

        SQL = "SELECT ArtAnagrafica.AcqCodForAbituale, AnagrGenIndirizzi.RagSoc1, RagSoc1, RagSoc2, AnagrGenIndirizzi.Indirizzo, Localita, Localita2, AnagrGenIndirizzi.Provincia, AnagrGenIndirizzi.CodStato, AnagrGenIndirizzi.NumTel, AnagrGenIndirizzi.NumFax, AnagrGenIndirizzi.IndirEmail, AnagTerzista " &
              "FROM (ArtAnagrafica LEFT JOIN ClientiFornitori ON (ArtAnagrafica.DBGruppo = ClientiFornitori.DBGruppo) AND (ArtAnagrafica.AcqTipoAnagrForAbit = ClientiFornitori.TipoAnagrafica) AND (ArtAnagrafica.AcqCodForAbituale = ClientiFornitori.CodCliFor)) LEFT JOIN AnagrGenIndirizzi ON (ClientiFornitori.NumRifAltroIndir = AnagrGenIndirizzi.NumProgr) AND (ClientiFornitori.IdAnagGen = AnagrGenIndirizzi.IdAnagGen) " &
              "WHERE ArtAnagrafica.DBGruppo='" & GruppoArchivi & "' AND ArtAnagrafica.TipoAnagr = 1 AND ArtAnagrafica.DesArt>' ' AND AnagrGenIndirizzi.FirmaUltVarData>='" & DataEsolver & "'" &
              "GROUP BY ArtAnagrafica.AcqCodForAbituale, AnagrGenIndirizzi.RagSoc1, RagSoc1, RagSoc2, AnagrGenIndirizzi.Indirizzo, Localita, Localita2, AnagrGenIndirizzi.Provincia, AnagrGenIndirizzi.CodStato, AnagrGenIndirizzi.NumTel, AnagrGenIndirizzi.NumFax, AnagrGenIndirizzi.IndirEmail, AnagTerzista " &
              "HAVING ArtAnagrafica.AcqCodForAbituale<>0;"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstESolver.EOF
            RagSociale = ""
            RagSocialeEstesa = ""
            Indirizzo = ""
            Localita = ""
            Tipo = ""

            If Not IsDBNull(rstESolver.Fields("RagSoc1").Value) Then
                RagSociale = Trim(rstESolver.Fields("RagSoc1").Value)
                RagSociale = Replace(RagSociale, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("RagSoc1").Value) Then
                RagSocialeEstesa = Trim(rstESolver.Fields("RagSoc1").Value) & " " & Trim(rstESolver.Fields("RagSoc2").Value)
                RagSocialeEstesa = Replace(RagSocialeEstesa, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("Indirizzo").Value) Then
                Indirizzo = Trim(rstESolver.Fields("Indirizzo").Value)
                Indirizzo = Replace(Indirizzo, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("Localita").Value) Then
                Localita = Left(Trim(rstESolver.Fields("Localita").Value) & " " & Trim(rstESolver.Fields("Localita2").Value), 40)
                Localita = Replace(Localita, "'", " ")
            End If

            If rstESolver.Fields("AnagTerzista").Value = 1 Then
                Tipo = 3
            Else
                Tipo = 1
            End If

            cmdShare.CommandText = "INSERT INTO ESOLVER_FORNITORI (Forn_Codice, Forn_RagSociale, Forn_RagSocialeEstesa, Forn_Indirizzo, Forn_Localita, Forn_Provincia, Forn_Nazione, Forn_Telefono, Forn_Fax, Forn_Email, Forn_Tipo ) " &
                                   "Values(" & rstESolver.Fields("AcqCodForAbituale").Value & ", '" & RagSociale & "', '" & RagSocialeEstesa & "', '" & Indirizzo & "', '" & Localita & "', '" & rstESolver.Fields("Provincia").Value & "', '" & rstESolver.Fields("CodStato").Value & "', '" & rstESolver.Fields("NumTel").Value & "', '" & rstESolver.Fields("numfax").Value & "', '" & rstESolver.Fields("IndirEmail").Value & "', '" & Tipo & "');"
            cmdShare.Execute()

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("FINE FORNITORI")

        '==> Svuoto e ripopolo la tabella dbo_ESOLVER_TAB_COLORI
        cmdShare.CommandText = "DELETE FROM ESOLVER_TAB_COLORI;"
        cmdShare.Execute()

        'MsgBox("ESOLVER_TAB_COLORI")

        ' Accodo da tabella anagrafica colori generale
        SQL = "SELECT ArtCodificaTipo.CodVariante, ArtCodificaTipo.Des " &
              "FROM ArtCodificaTipo " &
              "WHERE ArtCodificaTipo.DBGruppo='" & GruppoArchivi & "' AND ArtCodificaTipo.CodTipoVariante='TCOL';"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstESolver.EOF
            Descrizione = ""

            If Not IsDBNull(rstESolver.Fields("Des").Value) Then
                Descrizione = Replace(rstESolver.Fields("Des").Value, "'", " ")
            End If

            cmdShare.CommandText = "INSERT INTO ESOLVER_TAB_COLORI ( TABCOL_COD, TABCOL_DESCR )" &
                                   "Values('" & rstESolver.Fields("CodVariante").Value & "', '" & Descrizione & "');"
            cmdShare.Execute()

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("FINE TAB COLORI")

        ' Accodo colori non codificati ma presenti in anagrafica articoli
        cmdShare.CommandText = "INSERT INTO ESOLVER_TAB_COLORI ( TABCOL_COD, TABCOL_DESCR ) " &
                               "SELECT ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_COLORE, 'DESCRIZIONE DA ASSEGNARE' AS Des " &
                               "FROM ESOLVER_MATERIA_PRIMA_COLORE LEFT JOIN ESOLVER_TAB_COLORI ON ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_COLORE = ESOLVER_TAB_COLORI.TABCOL_COD " &
                               "WHERE ESOLVER_TAB_COLORI.TABCOL_COD Is Null " &
                               "GROUP BY ESOLVER_MATERIA_PRIMA_COLORE.MATC_CODICE_COLORE;"
        cmdShare.Execute()

        '==> Svuoto e ripopolo la tabella dbo_ESOLVER_TAB_COMPOSIZIONI
        cmdShare.CommandText = "DELETE FROM ESOLVER_TAB_COMPOSIZIONI;"
        cmdShare.Execute()

        'MsgBox("ESOLVER_TAB_COMPOSIZIONI")

        SQL = "SELECT ModaTabComposizioni.CodiceComposizione, ModaTabComposizioni.Descrizione_1, ModaTabComposizioni.DescrizioneEstesa " &
              "FROM ModaTabComposizioni " &
              "WHERE ModaTabComposizioni.DBGruppo='" & GruppoArchivi & "' AND ModaTabComposizioni.FirmaUltVarData>='" & DataEsolver & "'"

        rstESolver.Open(SQL, connESolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

        Do Until rstESolver.EOF
            Descrizione = ""
            DescrizioneEstesa = ""

            If Not IsDBNull(rstESolver.Fields("Descrizione_1").Value) Then
                Descrizione = Replace(rstESolver.Fields("Descrizione_1").Value, "'", " ")
            End If

            If Not IsDBNull(rstESolver.Fields("DescrizioneEstesa").Value) Then
                DescrizioneEstesa = Replace(rstESolver.Fields("DescrizioneEstesa").Value, "'", " ")
            End If

            cmdShare.CommandText = "INSERT INTO ESOLVER_TAB_COMPOSIZIONI ( TABCOM_CODICE, TABCOM_DESCRIZIONE, TABCOM_DESCRIZIONE_ESTESA ) " &
                                   "Values('" & rstESolver.Fields("CodiceComposizione").Value & "', '" & Descrizione & "', '" & DescrizioneEstesa & "');"
            cmdShare.Execute()

            rstESolver.MoveNext()
        Loop
        rstESolver.Close()

        'MsgBox("FINE TAB COMPOSIZIONI")

        ' Imposto il semaforo a verde per Lectra (=0) e aggiorno la data esolver
        cmdShare.CommandText = "UPDATE SEMAFORO_MATERIALI SET VALORE=0, DATA_ESOLVER= CURRENT_TIMESTAMP"
        cmdShare.Execute()

Chiusura_Connessione:
        rstESolver = Nothing
        rstShare = Nothing

        connShare.Close()
        connESolver.Close()

        ' MsgBox("FINE")

    End Sub


    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String,
            ByVal lpKeyName As String,
            ByVal lpDefault As String,
            ByVal lpReturnedString As System.Text.StringBuilder,
            ByVal nSize As Integer,
            ByVal lpFileName As String) As Integer

    Function TogliDec(ValoreCampo As Double, NumeroDecimali As Long) As Double
        If NumeroDecimali = 1 Then TogliDec = ValoreCampo * 10
        If NumeroDecimali = 2 Then TogliDec = ValoreCampo * 100
        If NumeroDecimali = 3 Then TogliDec = ValoreCampo * 1000
        If NumeroDecimali = 4 Then TogliDec = ValoreCampo * 10000
        If NumeroDecimali = 5 Then TogliDec = ValoreCampo * 100000
        If NumeroDecimali = 6 Then TogliDec = ValoreCampo * 1000000
        If NumeroDecimali = 7 Then TogliDec = ValoreCampo * 10000000
        If NumeroDecimali = 8 Then TogliDec = ValoreCampo * 100000000
        If NumeroDecimali = 9 Then TogliDec = ValoreCampo * 1000000000
        If NumeroDecimali = 10 Then TogliDec = ValoreCampo * 10000000000.0#
    End Function

    Function MettiDec(NomeCampo As String, NumDecimali As Long, Filtro As String, connessione As ADODB.Connection, comando As ADODB.Command)
        Dim Sql As String

        If NumDecimali = 1 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/10 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 2 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/100 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 3 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/1000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 4 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/10000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 5 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/100000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 6 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/1000000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 7 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/10000000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 8 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/100000000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 9 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/1000000000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 10 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_RADICE SET " & NomeCampo & "=" & NomeCampo & "/10000000000 where MATR_CODICE='" & Filtro & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

    End Function

    Function Colori_MettiDec(NomeCampo As String, NumDecimali As Long, Filtro1 As String, Filtro2 As String, connessione As ADODB.Connection, comando As ADODB.Command)
        Dim Sql As String

        If NumDecimali = 1 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/10 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 2 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/100 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 3 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/1000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 4 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/10000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 5 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/100000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 6 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/1000000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 7 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/10000000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 8 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/100000000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 9 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/1000000000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If

        If NumDecimali = 10 Then
            Sql = "UPDATE ESOLVER_MATERIA_PRIMA_COLORE SET " & NomeCampo & "=" & NomeCampo & "/10000000000 where MATC_CODICE_MATERIALE='" & Filtro1 & "' AND MATC_CODICE_COLORE='" & Filtro2 & "'"
            comando.CommandText = Sql
            comando.Execute()
        End If
    End Function

    Function CaricaTaglia(articolo As String, numeroTg As Int32, codTg As String, desTg As String, connessione As ADODB.Connection, comando As ADODB.Command)
        Dim Sql As String

        Sql = "INSERT INTO ESOLVER_MATERIA_PRIMA_MISURE ( MATM_CODICE_MATERIALE, MATM_CODICE_MISURA, MATM_DESCR_MISURA_FORN, MATM_SEQUENZA, MATM_COD_MISURA_FORN, MATM_DESCR_MISURA, MATM_ABILITATO ) " &
              "SELECT '" & articolo & "', '" & codTg & "', '" & desTg & "', " & numeroTg.ToString() & ", '" & codTg & "', '" & desTg & "', 1;"

        comando.CommandText = Sql
        comando.Execute()
    End Function

End Module

