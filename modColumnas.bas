Attribute VB_Name = "modColumnas"
Option Explicit
'======================================================================
' modColumnas - Indices de columnas de ConNom1 (Form8 y FormViewer)
'
' Son VARIABLES: se fijan segun el tipo de nomina con InicializarColumnas.
' Para los tipos que NO son Liquidacion/Finiquito los valores son los
' mismos de siempre. En Liq las columnas de concepto propias usan C_LQ_*.
' En los tipos no-Liq las C_LQ_* valen -1 y en Liq las C_CONC4..8 valen -1:
' NO se debe escribir ConNom1.col = -1 (da error); protege con esLiq.
'======================================================================

' --- Identificacion ---
Public C_NUM As Long, C_NOMBRE As Long, C_DIAS As Long

' --- Percepciones ---
Public C_SUELDO As Long
Public C_CONC4 As Long, C_CONC5 As Long, C_CONC6 As Long
Public C_CONC7 As Long, C_CONC8 As Long
Public C_OTRAS As Long, C_EXENTA As Long, C_TOTING As Long

' --- Solo Liquidacion/Finiquito (-1 en los demas tipos) ---
Public C_LQ_SDOVAC As Long, C_LQ_PRIMAVAC As Long, C_LQ_AGUI As Long
Public C_LQ_ANTIG As Long, C_LQ_COMPENS As Long, C_LQ_INDEMN As Long
Public C_LQ_BONO As Long

' --- Deducciones ---
Public C_ISPT As Long, C_SUBEMP As Long, C_IMSS As Long, C_PRESTAMOS As Long
Public C_FONACOT As Long, C_PENSION As Long, C_INFONAVIT As Long
Public C_TOTDED As Long, C_NETO As Long

' --- Banderas (columnas ocultas) y otras ---
Public C_FLAG_ISR As Long, C_FLAG_IMSS As Long, C_FLAG_ESP As Long
Public C_BANAMEX As Long, C_TOTCOLS As Long

' --- Rangos de suma (los usan sumah / sumavert) ---
Public C_PERC_INI As Long, C_PERC_FIN As Long
Public C_DED_INI As Long, C_DED_FIN As Long

Public Sub InicializarColumnas(ByVal tipo As TipoNominaCFDI)
    C_NUM = 0: C_NOMBRE = 1: C_DIAS = 2: C_SUELDO = 3
    C_PERC_INI = C_SUELDO

    If tipo = tnLiquidacionFiniquito Then
        C_CONC4 = -1: C_CONC5 = -1: C_CONC6 = -1: C_CONC7 = -1: C_CONC8 = -1
        C_LQ_SDOVAC = 4: C_LQ_PRIMAVAC = 5: C_LQ_AGUI = 6: C_LQ_ANTIG = 7
        C_LQ_COMPENS = 8: C_LQ_INDEMN = 9: C_LQ_BONO = 10
        C_OTRAS = 11: C_EXENTA = 12: C_TOTING = 13
        C_ISPT = 14
    Else
        C_CONC4 = 4: C_CONC5 = 5: C_CONC6 = 6: C_CONC7 = 7: C_CONC8 = 8
        C_LQ_SDOVAC = -1: C_LQ_PRIMAVAC = -1: C_LQ_AGUI = -1: C_LQ_ANTIG = -1
        C_LQ_COMPENS = -1: C_LQ_INDEMN = -1: C_LQ_BONO = -1
        C_OTRAS = 9: C_EXENTA = 10: C_TOTING = 11
        C_ISPT = 12
    End If

    ' Deducciones y cola: relativas a C_ISPT
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

    C_PERC_FIN = C_OTRAS            ' Sueldo..Otras (SIN exenta: no es base de ISR)
    C_DED_INI = C_ISPT
    C_DED_FIN = C_INFONAVIT
End Sub

' Un solo punto de entrada para fijar tipo + columnas.
' Usalo en nomcap_ini_Click (Form8) y en LoadPayroll (FormViewer).
Public Sub AplicarTipoNomina(ByVal nombreArch As String)
    G_NombreArchivoNomina = UCase$(Trim$(nombreArch))
    g_TipoNominaActiva = ClasificarTipoNomina(G_NombreArchivoNomina)
    G_TipoNominaProcess = TipoNominaATexto(g_TipoNominaActiva)
    InicializarColumnas g_TipoNominaActiva
End Sub

