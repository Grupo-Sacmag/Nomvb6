Attribute VB_Name = "modColumnas"
Option Explicit

Public C_NUM As Long, C_NOMBRE As Long, C_DIAS As Long
Public C_SUELDO As Long
Public C_CONC4 As Long, C_CONC5 As Long, C_CONC6 As Long
Public C_CONC7 As Long, C_CONC8 As Long
Public C_OTRAS As Long, C_EXENTA As Long, C_TOTING As Long
Public C_LQ_SDOVAC As Long, C_LQ_PRIMAVAC As Long, C_LQ_AGUI As Long
Public C_LQ_ANTIG As Long, C_LQ_COMPENS As Long, C_LQ_INDEMN As Long
Public C_LQ_BONO As Long
Public C_ISPT As Long, C_SUBEMP As Long, C_IMSS As Long
Public C_PRESTAMOS As Long, C_FONACOT As Long, C_PENSION As Long
Public C_INFONAVIT As Long, C_TOTDED As Long, C_NETO As Long
Public C_FLAG_ISR As Long, C_FLAG_IMSS As Long, C_FLAG_ESP As Long
Public C_BANAMEX As Long, C_TOTCOLS As Long
Public C_PERC_INI As Long, C_PERC_FIN As Long
Public C_DED_INI As Long, C_DED_FIN As Long

Public Const MARCA_SEPARACION_2026 As Long = 20261010

Public Const MARCA_AGUINALDO_2027 As Long = 20270121
Public Const MARCA_PTU_2027 As Long = 20270122
Public Const MARCA_BONO_2027 As Long = 20270123
Public Const MARCA_FINIQUITO_2027 As Long = 20270124
Public Const MARCA_LIQUIDACION_2027 As Long = 20270125

Public g_FormatoSeparacionNuevo As Boolean
Public g_EsFiniquito As Boolean
Public g_EspecialHistorica As Boolean
Public g_MarcaFormatoNomina As Currency

Public Function EsSeparacionNueva() As Boolean
    EsSeparacionNueva = (g_TipoNominaActiva = tnLiquidacionFiniquito And g_FormatoSeparacionNuevo)
End Function

Public Function MarcaFormatoNomina() As Currency
    MarcaFormatoNomina = g_MarcaFormatoNomina
End Function

Public Function EsMarcaEspecialNueva(ByVal marca As Currency) As Boolean
    Select Case marca
        Case CCur(MARCA_SEPARACION_2026), CCur(MARCA_AGUINALDO_2027), CCur(MARCA_PTU_2027), CCur(MARCA_BONO_2027), CCur(MARCA_FINIQUITO_2027), CCur(MARCA_LIQUIDACION_2027)
            EsMarcaEspecialNueva = True
    End Select
End Function

Public Sub InicializarColumnas(ByVal tipo As TipoNominaCFDI)
    C_NUM = 0: C_NOMBRE = 1: C_DIAS = 2: C_SUELDO = 3
    C_PERC_INI = C_SUELDO

    C_CONC4 = -1: C_CONC5 = -1: C_CONC6 = -1
    C_CONC7 = -1: C_CONC8 = -1
    C_LQ_SDOVAC = -1: C_LQ_PRIMAVAC = -1: C_LQ_AGUI = -1
    C_LQ_ANTIG = -1: C_LQ_COMPENS = -1: C_LQ_INDEMN = -1
    C_LQ_BONO = -1

    If tipo = tnLiquidacionFiniquito Then
        If g_FormatoSeparacionNuevo Then
            ' C_SUELDO (3) es auxiliar reservado y debe estar VACIO.
            ' Liquidacion: Comp, Agui, Antig, Indem, PV, Otras
            ' Finiquito: SdoVac, Agui, Antig, Bono, PV, Otras
            If g_EsFiniquito Then
                C_LQ_SDOVAC = 4
                C_LQ_AGUI = 5
                C_LQ_ANTIG = 6
                C_LQ_BONO = 7
                C_LQ_PRIMAVAC = 8
            Else
                C_LQ_COMPENS = 4
                C_LQ_AGUI = 5
                C_LQ_ANTIG = 6
                C_LQ_INDEMN = 7
                C_LQ_PRIMAVAC = 8
            End If
            C_OTRAS = 9: C_EXENTA = 10: C_TOTING = 11: C_ISPT = 12
        Else
            ' LEGACY: exactamente la estructura anterior de 27 columnas.
            C_LQ_SDOVAC = 4: C_LQ_PRIMAVAC = 5: C_LQ_AGUI = 6
            C_LQ_ANTIG = 7: C_LQ_COMPENS = 8: C_LQ_INDEMN = 9
            C_LQ_BONO = 10
            C_OTRAS = 11: C_EXENTA = 12: C_TOTING = 13: C_ISPT = 14
        End If
    Else
        ' Ordinaria, Aguinaldo, PTU y Bono: INTACTOS.
        C_CONC4 = 4: C_CONC5 = 5: C_CONC6 = 6
        C_CONC7 = 7: C_CONC8 = 8
        C_OTRAS = 9: C_EXENTA = 10: C_TOTING = 11: C_ISPT = 12
    End If

    C_SUBEMP = C_ISPT + 1
    C_IMSS = C_ISPT + 2
    C_PRESTAMOS = C_ISPT + 3
    C_FONACOT = C_ISPT + 4
    C_PENSION = C_ISPT + 5
    C_INFONAVIT = C_ISPT + 6
    C_TOTDED = C_ISPT + 7
    C_NETO = C_ISPT + 8
    C_FLAG_ISR = C_ISPT + 9
    C_FLAG_IMSS = C_ISPT + 10
    C_FLAG_ESP = C_ISPT + 11
    C_BANAMEX = C_ISPT + 12
    C_TOTCOLS = C_BANAMEX + 1

    C_PERC_FIN = C_OTRAS
    C_DED_INI = C_ISPT
    C_DED_FIN = C_INFONAVIT
End Sub

Public Sub AplicarTipoNomina(ByVal nombreArch As String, Optional ByVal tipoNuevo As Long = -1, Optional ByVal esFiniquitoNuevo As Boolean = False)

    Dim nArchivo As Integer
    Dim temporal As nom
    Dim existe As Boolean
    Dim nombreBase As String
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorLectura

    G_NombreArchivoNomina = UCase$(Trim$(nombreArch))
    nombreBase = NombreBaseNomina(G_NombreArchivoNomina)

    g_TipoNominaActiva = tnOrdinaria
    g_FormatoSeparacionNuevo = False
    g_EsFiniquito = False
    g_EspecialHistorica = False
    g_MarcaFormatoNomina = 0

    existe = (Len(Dir$(nombreArch)) > 0)

    If Not existe Then
        If tipoNuevo >= 0 Then
            Select Case tipoNuevo
                Case tnOrdinaria, tnAguinaldo, tnPTU, tnBonoPremio, tnLiquidacionFiniquito
                    g_TipoNominaActiva = CInt(tipoNuevo)

                Case Else
                    Err.Raise vbObjectError + 2030, "AplicarTipoNomina", "Tipo de nomina nuevo no valido."
            End Select
        Else

            g_TipoNominaActiva = ClasificarTipoNomina(nombreBase)

            If g_TipoNominaActiva = tnOrdinaria And Not EsNominaNormal(nombreBase) Then
                Err.Raise vbObjectError + 2031, "AplicarTipoNomina", "Seleccione el tipo de nomina especial."
            End If
        End If

        Select Case g_TipoNominaActiva
            Case tnAguinaldo
                g_MarcaFormatoNomina = CCur(MARCA_AGUINALDO_2027)

            Case tnPTU
                g_MarcaFormatoNomina = CCur(MARCA_PTU_2027)

            Case tnBonoPremio
                g_MarcaFormatoNomina = CCur(MARCA_BONO_2027)

            Case tnLiquidacionFiniquito
                g_FormatoSeparacionNuevo = True

                If tipoNuevo >= 0 Then
                    g_EsFiniquito = esFiniquitoNuevo
                Else
                    g_EsFiniquito = (InStr(1, nombreBase, "FINI", vbTextCompare) > 0)
                End If
                If g_EsFiniquito Then
                    g_MarcaFormatoNomina = CCur(MARCA_FINIQUITO_2027)
                Else
                    g_MarcaFormatoNomina = CCur(MARCA_LIQUIDACION_2027)
                End If
        End Select
        Else
            nArchivo = FreeFile

        Open nombreArch For Random Access Read As #nArchivo Len = Len(temporal)

        If LOF(nArchivo) = 0 Or LOF(nArchivo) Mod Len(temporal) <> 0 Then
            Err.Raise vbObjectError + 2032, "AplicarTipoNomina", "El archivo NOM no contiene registros completos."
        End If

        Get #nArchivo, 1, temporal

        Close #nArchivo
        nArchivo = 0

        Select Case temporal.hstri
            Case 0
                g_TipoNominaActiva = tnOrdinaria
                If Not EsNominaNormal(nombreBase) Then
                    g_EspecialHistorica = True
                End If
                
            Case CCur(MARCA_AGUINALDO_2027)
                g_TipoNominaActiva = tnAguinaldo

            Case CCur(MARCA_PTU_2027)
                g_TipoNominaActiva = tnPTU

            Case CCur(MARCA_BONO_2027)
                g_TipoNominaActiva = tnBonoPremio

            Case CCur(MARCA_FINIQUITO_2027)
                g_TipoNominaActiva = tnLiquidacionFiniquito
                g_FormatoSeparacionNuevo = True
                g_EsFiniquito = True

            Case CCur(MARCA_LIQUIDACION_2027)
                g_TipoNominaActiva = tnLiquidacionFiniquito
                g_FormatoSeparacionNuevo = True
                g_EsFiniquito = False

            Case CCur(MARCA_SEPARACION_2026)
                g_TipoNominaActiva = tnLiquidacionFiniquito
                g_FormatoSeparacionNuevo = True
                If InStr(1, nombreBase, "FINI", vbTextCompare) > 0 Then
                    g_EsFiniquito = True
                ElseIf InStr(1, nombreBase, "LIQ", vbTextCompare) > 0 Then
                    g_EsFiniquito = False
                Else
                    Err.Raise vbObjectError + 2033, "AplicarTipoNomina", "La separacion de prueba no identifica si es Finiquito o Liquidacion."
                End If

            Case Else
                Err.Raise vbObjectError + 2034, "AplicarTipoNomina", "El archivo contiene una marca de formato desconocida."
        End Select

        If Not g_EspecialHistorica Then
            g_MarcaFormatoNomina = temporal.hstri
        End If
    End If

    G_TipoNominaProcess = TipoNominaATexto(g_TipoNominaActiva)
    InicializarColumnas g_TipoNominaActiva

    If g_EspecialHistorica Then
        Debug.Print "Formato: ESPECIAL HISTORICA"
    ElseIf g_MarcaFormatoNomina <> 0 Then
        Debug.Print "Formato NOM: "; g_MarcaFormatoNomina
    Else
        Debug.Print "Formato: ORDINARIA"
    End If

    If existe And InStrRev(nombreArch, "\") > 0 Then
        If Not EsNominaNormal(nombreBase) Then
            Debug.Print "Generaciones FeCoNom: "; GeneracionesRegistradasFeCoNom(nombreArch)
        End If
    End If

    Exit Sub
    
ErrorLectura:
    numeroError = Err.Number
    descripcionError = Err.Description
    On Error Resume Next
    If nArchivo <> 0 Then Close #nArchivo
    On Error GoTo 0
    Err.Raise numeroError, "AplicarTipoNomina", descripcionError
End Sub

