Attribute VB_Name = "modColumnas"
Option Explicit
'============================================================
' modColumnas - Indices de ConNom1
' Normales: indices tradicionales INALTERADOS.
' Liq/Fini legacy: indices actuales INALTERADOS.
' Liq/Fini nuevo: seis percepciones en orden solicitado.
'============================================================
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

' Campo version en .hstri SOLO para Liq/Fini nuevos.
' Se respeta Type nom; Len(nomina) no cambia.
Public Const MARCA_SEPARACION_2026 As Long = 20261010
Public g_FormatoSeparacionNuevo As Boolean
Public g_EsFiniquito As Boolean

Public Function EsSeparacionNueva() As Boolean
    EsSeparacionNueva = (g_TipoNominaActiva = tnLiquidacionFiniquito _
                        And g_FormatoSeparacionNuevo)
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

Public Sub AplicarTipoNomina(ByVal nombreArch As String)
    Dim nArchivo As Integer
    Dim temporal As nom
    Dim existe As Boolean

    G_NombreArchivoNomina = UCase$(Trim$(nombreArch))
    g_TipoNominaActiva = ClasificarTipoNomina(G_NombreArchivoNomina)
    G_TipoNominaProcess = TipoNominaATexto(g_TipoNominaActiva)
    g_FormatoSeparacionNuevo = False
    g_EsFiniquito = False

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        g_EsFiniquito = (InStr(1, NombreBaseNomina(G_NombreArchivoNomina), "FINI", vbTextCompare) > 0)
        existe = (Len(Dir$(nombreArch)) > 0)
        If Not existe Then
            ' Fecha de inicio para creacion de nuevos archivos.
            g_FormatoSeparacionNuevo = (Date >= DateSerial(2026, 10, 10))
        Else
            ' En archivos existentes manda la marca interna, NO FileDateTime.
            nArchivo = FreeFile
            On Error GoTo ErrorLectura
            Open nombreArch For Random Access Read As #nArchivo Len = Len(temporal)
            If LOF(nArchivo) >= Len(temporal) Then
                Get #nArchivo, 1, temporal
                g_FormatoSeparacionNuevo = _
                    (temporal.hstri = CCur(MARCA_SEPARACION_2026))
            End If
            Close #nArchivo
            On Error GoTo 0
        End If
    End If

    InicializarColumnas g_TipoNominaActiva
    Exit Sub

ErrorLectura:
    On Error Resume Next
    Close #nArchivo
    On Error GoTo 0
    Err.Raise vbObjectError + 2026, "AplicarTipoNomina", "No fue posible identificar el formato del archivo " & nombreArch
End Sub


