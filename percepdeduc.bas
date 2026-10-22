Attribute VB_Name = "percepdeduc"
Public sue3 As Currency 'sueldos
Public via7 As Currency 'viaticos
'esta madre ni se ocupa lo capturan en otras
Public pva8 As Currency 'prima vacacional
Public otr9 As Currency 'otros ingresos
Public pee10 As Currency 'percepcion exenta
Public sub13 As Currency ' subsidio pagado
Public agui5 As Currency ' aguinaldo

Public subc13 As Currency 'subsidio caisado

Public isr12 As Currency 'isr
Public ims14 As Currency 'imss
Public ifv18 As Currency 'infonavit
Public pea17 As Currency 'pension alimenticia
Public fon16 As Currency 'fonacot
Public pre15 As Currency 'prestamos otros


Public t_oi As Currency 'total otros ingresos

Public t_per As Currency 'total de percepciones
Public t_ded As Currency 'total de deducciones
Public T_neto As Currency 'salario neto

Public t_grav As Currency   'total gravado percepciones
Public t_ext As Currency  'total  exento percepciones
Public t_gravded As Currency 'total gravado deducciones
Public t_extded As Currency 'total exento deducciones
Public t_oded As Currency 'total otras deducciones
Public ptu_1 As Currency, ptu_2 As Currency, ptu_3 As Currency
Public I7 As Long
Public OrigenCFDI As Integer
Public UltimaFilaCFDI As Long

' --- Módulo de Variables Globales ---
Public Enum TipoNominaCFDI
    tnOrdinaria = 0
    tnLiquidacionFiniquito = 1
    tnAguinaldo = 2
    tnPTU = 3
    tnBonoPremio = 4
End Enum

Public g_TipoNominaActiva As TipoNominaCFDI
Public g_ExentoAntig As Double, g_GravadoAntig As Double
Public g_ExentoIndem As Double, g_GravadoIndem As Double
Public g_ExentoAguiLiq As Double, g_GravadoAguiLiq As Double
Public g_ExentoPVacLiq As Double, g_GravadoPVacLiq As Double

Public Function DeterminarTipoNomina(ByVal NombreArchivo As String, Optional ByVal OpcionForm8 As Integer = -1) As TipoNominaCFDI
    Dim TipoPorNombre As TipoNominaCFDI
    TipoPorNombre = ClasificarTipoNomina(NombreArchivo)

    ' Si el nombre trae una palabra clave reconocida, esa manda.
    If TipoPorNombre <> tnOrdinaria Then
        DeterminarTipoNomina = TipoPorNombre
        Exit Function
    End If

    ' Si no se reconoció nada en el nombre, se respeta lo que venga de Form8 (si aplica).
    If OpcionForm8 >= 0 Then
        DeterminarTipoNomina = OpcionForm8
        Exit Function
    End If

    DeterminarTipoNomina = tnOrdinaria
End Function
Sub reng()

Dim GridOrigen As Object

    If OrigenCFDI = 1 Then
        Set GridOrigen = FormViewer.ConNom1
    Else
        Set GridOrigen = Form8.ConNom1
    End If
    

    ' FORMATO NUEVO: NO USAR POSICIONES FIJAS LEGACY.
    If EsSeparacionNueva() Then
        rengSeparacionNueva GridOrigen
        Exit Sub
    End If

    If GridOrigen Is Nothing Then
        MsgBox "GridOrigen es Nothing" & vbCrLf & _
               "OrigenCFDI = " & OrigenCFDI
        Exit Sub
    End If
    
    Dim c_per As Integer, c_ded As Integer
    
    t_per = 0: t_ded = 0: c_per = 0: c_ded = 0
    sue3 = 0: via7 = 0: pva8 = 0: otr9 = 0: pee10 = 0: sub13 = 0
    ptu_1 = 0: ptu_2 = 0: ptu_3 = 0
    isr12 = 0: ims14 = 0: pre15 = 0: fon16 = 0: pea17 = 0: ifv18 = 0
    t_oi = 0: T_neto = 0
    t_grav = 0: t_ext = 0: t_gravded = 0: t_extded = 0
    t_oded = 0
    
    '====================================================
    ' SUBSIDIO AL EMPLEO - COLUMNA DINAMICA
    '====================================================
    
    If IsNumeric(GridOrigen.TextMatrix(I7, C_SUBEMP)) Then
    
        sub13 = CCur(GridOrigen.TextMatrix(I7, C_SUBEMP)) * -1
    
    Else
    
        sub13 = 0
    
    End If
    
    If IsNumeric(GridOrigen.TextMatrix(I7, 0)) Then
        Dim pruebaArchivo As Long
        On Error Resume Next
        pruebaArchivo = LOF(14)
        If Err.Number <> 0 Then
            MsgBox "El archivo #14 NO está abierto." & vbCrLf & _
                   "Error: " & Err.Number & vbCrLf & Err.Description, vbCritical, "percepdeduc"
            Err.Clear: On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        Get 14, GridOrigen.TextMatrix(I7, 0), nom_com
        subc13 = nom_com.subdio
    End If
     
    'percepciones---------------------------------------------------
    If IsNumeric(GridOrigen.TextMatrix(I7, 3)) Then
        sue3 = GridOrigen.TextMatrix(I7, 3)
        c_per = c_per + 1: t_per = t_per + sue3
    Else
        sue3 = 0
    End If
    
    'Parche Aguinaldo 19/12/2017------------------------------------------
    If N_ormal = 1 Then
        If IsNumeric(GridOrigen.TextMatrix(I7, 5)) Then
            agui5 = GridOrigen.TextMatrix(I7, 5)
            c_per = c_per + 1: t_per = t_per + agui5
        Else
            agui5 = 0
        End If
    End If
    
    If IsNumeric(GridOrigen.TextMatrix(I7, 9)) Then
        otr9 = GridOrigen.TextMatrix(I7, 9)
        t_per = t_per + otr9
    Else
        otr9 = 0
    End If
    
    If IsNumeric(GridOrigen.TextMatrix(I7, 7)) Then
        via7 = GridOrigen.TextMatrix(I7, 7)
        t_per = t_per + via7
    Else
        via7 = 0
    End If
    
    '==========================================================================
    ' COLUMNA 8: en LIQ es "Sdo.Vac." (percepción ordinaria, SIN split de exención);
    ' en cualquier otro tipo es Prima Vacacional (con su tope de exención habitual).
    ' COLUMNA 10 en LIQ es el exento manual de Antig/Indem (Text3) — no debe tocarse aquí.
    '==========================================================================
    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
    
        Dim colSdoVacR As Double
        If IsNumeric(GridOrigen.TextMatrix(I7, 8)) Then
            colSdoVacR = Val(GridOrigen.TextMatrix(I7, 8))
        Else
            colSdoVacR = 0
        End If
        pva8 = 0                          ' anulado: MdAbr_1 no debe mandarlo a P021
        sue3 = sue3 + colSdoVacR          ' Sdo.Vac. -> percepción ordinaria (P001G)
        t_per = t_per + colSdoVacR
        If colSdoVacR <> 0 Then c_per = c_per + 1
    
        If IsNumeric(GridOrigen.TextMatrix(I7, 10)) Then
            pee10 = GridOrigen.TextMatrix(I7, 10)   ' se usa más abajo, NO se suma aquí
        Else
            pee10 = 0
        End If
    
    Else
    
        If IsNumeric(GridOrigen.TextMatrix(I7, 8)) Then
            pva8 = GridOrigen.TextMatrix(I7, 8)
            c_per = c_per + 1
        Else
            pva8 = 0
        End If
    
        If IsNumeric(GridOrigen.TextMatrix(I7, 10)) Then
            pee10 = GridOrigen.TextMatrix(I7, 10)
        Else
            pee10 = 0
        End If
    
        If pva8 > 0 Then
            If pee10 > 0 Then
                t_per = t_per + pva8 + pee10
            Else
                If pva8 > (empresa.sm * 15) Then
                    pee10 = (empresa.sm * 15)
                    pva8 = pva8 - pee10
                    t_per = t_per + pva8 + pee10
                Else
                    pee10 = (pva8 - 0.01)
                    pva8 = 0.01
                    t_per = t_per + pva8 + pee10
                End If
            End If
        Else
            t_per = t_per + pee10
        End If
    
    End If
    
    Rem ******************* PARCHE DE PTU DEL 6/6/17 **********************************
    If g_TipoNominaActiva = tnPTU Then
      If IsNumeric(GridOrigen.TextMatrix(I7, 6)) Then
        ptu_1 = GridOrigen.TextMatrix(I7, 6)
         If IsNumeric(GridOrigen.TextMatrix(I7, 10)) Then
            ptu_2 = GridOrigen.TextMatrix(I7, 10)
            Else
            ptu_2 = empresa.sm * 15
            If ptu_2 >= ptu_1 Then
                ptu_2 = ptu_1: ptu_1 = 0
               Else
                ptu_1 = ptu_1 - ptu_2
            End If
         End If
         ' PTU gravada + PTU exenta
        ptu_3 = ptu_1 + ptu_2
        
        ' La exenta ya fue acumulada antes en t_per.
        ' Retirarla temporalmente para no duplicarla
        ' al sumar la PTU completa.
        t_per = t_per - pee10 + ptu_3
        
        ' Conservar la parte exenta para el total exento CFDI
        pee10 = ptu_2
      End If
    End If
    
    t_grav = t_per - pee10
    t_ext = pee10
    
    '==========================================================================
    ' LIQUIDACIÓN / FINIQUITO: Compensación, Antigüedad, Indemnización,
    ' Aguinaldo y Prima Vacacional — todo en un solo bloque
    '==========================================================================
    g_ExentoAntig = 0: g_GravadoAntig = 0
    g_ExentoIndem = 0: g_GravadoIndem = 0
    g_ExentoAguiLiq = 0: g_GravadoAguiLiq = 0
    g_ExentoPVacLiq = 0: g_GravadoPVacLiq = 0
    
    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
    
        Dim colCompR As Double, colAntigR As Double, colIndemR As Double
        colCompR = Val(GridOrigen.TextMatrix(I7, 4))
        colAntigR = Val(GridOrigen.TextMatrix(I7, 5))
        colIndemR = Val(GridOrigen.TextMatrix(I7, 6))
    
        t_per = t_per + colCompR + colIndemR
    
        Dim totalExentoManual As Double, totalAntigIndemR As Double
        totalExentoManual = pee10
        totalAntigIndemR = colAntigR + colIndemR
    
        If totalExentoManual >= totalAntigIndemR Then
            g_ExentoAntig = colAntigR
            g_ExentoIndem = colIndemR
        ElseIf totalExentoManual <= colAntigR Then
            g_ExentoAntig = totalExentoManual
            g_ExentoIndem = 0
        Else
            g_ExentoAntig = colAntigR
            g_ExentoIndem = totalExentoManual - colAntigR
        End If
        g_GravadoAntig = colAntigR - g_ExentoAntig
        g_GravadoIndem = colIndemR - g_ExentoIndem
    
        ' --- Aguinaldo y Prima Vacacional de liquidación ---
        Dim colAguiLiqR As Double, colPVacLiqR As Double
        colAguiLiqR = Val(GridOrigen.TextMatrix(I7, 25))
        colPVacLiqR = Val(GridOrigen.TextMatrix(I7, 26))
    
        If colAguiLiqR > (empresa.sm * 30) Then
            g_ExentoAguiLiq = (empresa.sm * 30)
            g_GravadoAguiLiq = colAguiLiqR - g_ExentoAguiLiq
        Else
            g_ExentoAguiLiq = colAguiLiqR
        End If
    
        If colPVacLiqR > (empresa.sm * 15) Then
            g_ExentoPVacLiq = (empresa.sm * 15)
            g_GravadoPVacLiq = colPVacLiqR - g_ExentoPVacLiq
        Else
            g_ExentoPVacLiq = colPVacLiqR
        End If
    
        t_per = t_per + colAguiLiqR + colPVacLiqR
    
    End If
    
    t_grav = t_per - (pee10 + g_ExentoAguiLiq + g_ExentoPVacLiq)
    t_ext = pee10 + g_ExentoAguiLiq + g_ExentoPVacLiq
    
    'deducciones-----------------------------------------------------
    If IsNumeric(GridOrigen.TextMatrix(I7, C_ISPT)) Then
        isr12 = GridOrigen.TextMatrix(I7, C_ISPT)
        c_ded = c_ded + 1: t_ded = t_ded + isr12
    Else
        If sub13 > 0 Then
            sub13 = sub13 - 0.01: t_per = t_per + sub13
            isr12 = 0.01: t_ded = t_ded + isr12
        End If
    End If
    If isr12 = 0 Then isr12 = 0.01: t_ded = t_ded + isr12
    
    If IsNumeric(GridOrigen.TextMatrix(I7, C_IMSS)) Then
        ims14 = GridOrigen.TextMatrix(I7, C_IMSS)
        c_ded = c_ded + 1: t_ded = t_ded + ims14
    Else
        ims14 = 0
    End If
    If IsNumeric(GridOrigen.TextMatrix(I7, C_PRESTAMOS)) Then
        pre15 = GridOrigen.TextMatrix(I7, C_PRESTAMOS)
        c_ded = c_ded + 1: t_ded = t_ded + pre15
    Else
        pre15 = 0
    End If
    If IsNumeric(GridOrigen.TextMatrix(I7, C_FONACOT)) Then
        fon16 = GridOrigen.TextMatrix(I7, C_FONACOT)
        c_ded = c_ded + 1: t_ded = t_ded + fon16
    Else
        fon16 = 0
    End If
    If IsNumeric(GridOrigen.TextMatrix(I7, C_PENSION)) Then
        pea17 = GridOrigen.TextMatrix(I7, C_PENSION)
        c_ded = c_ded + 1: t_ded = t_ded + pea17
    Else
        pea17 = 0
    End If
    If IsNumeric(GridOrigen.TextMatrix(I7, C_INFONAVIT)) Then
        ifv18 = GridOrigen.TextMatrix(I7, C_INFONAVIT)
        c_ded = c_ded + 1: t_ded = t_ded + ifv18
    Else
        ifv18 = 0
    End If
    
    t_oded = t_ded - isr12
    T_neto = t_per - t_ded
    t_extded = t_ded - isr12
    
    
    Debug.Print "========== DESGLOSE CFDI =========="
    Debug.Print "Origen: "; OrigenCFDI
    Debug.Print "Empleado: "; GridOrigen.TextMatrix(I7, C_NUM)
    Debug.Print "Sueldo: "; sue3
    Debug.Print "Viaticos: "; via7
    Debug.Print "Prima vacacional: "; pva8
    Debug.Print "Otras: "; otr9
    Debug.Print "Exenta: "; pee10
    Debug.Print "Aguinaldo: "; agui5
    Debug.Print "PTU gravada: "; ptu_1
    Debug.Print "PTU exenta: "; ptu_2
    Debug.Print "Total percepciones: "; t_per
    Debug.Print "Total deducciones: "; t_ded
    Debug.Print "Neto CFDI: "; T_neto
    Debug.Print "=================================="
    
    '====================================================
    ' DIAGNOSTICO DE IMPORTES PARA CFDI
    '====================================================
    
    Debug.Print "=================================="
    Debug.Print "ORIGEN CFDI: "; OrigenCFDI
    Debug.Print "TIPO NOMINA: "; g_TipoNominaActiva
    Debug.Print "EMPLEADO: "; GridOrigen.TextMatrix(I7, C_NUM)
    
    Debug.Print "SUELDO: "; GridOrigen.TextMatrix(I7, C_SUELDO)
    Debug.Print "OTRAS: "; GridOrigen.TextMatrix(I7, C_OTRAS)
    Debug.Print "EXENTA: "; GridOrigen.TextMatrix(I7, C_EXENTA)
    
    Debug.Print "ISR: "; GridOrigen.TextMatrix(I7, C_ISPT)
    Debug.Print "NETO GRID: "; GridOrigen.TextMatrix(I7, C_NETO)
    
    Debug.Print "TOTAL PER CFDI: "; t_per
    Debug.Print "TOTAL DED CFDI: "; t_ded
    Debug.Print "NETO CFDI: "; T_neto
    
    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
    
        Debug.Print "S. VAC: "; GridOrigen.TextMatrix(I7, C_LQ_SDOVAC)
        Debug.Print "P. VAC: "; GridOrigen.TextMatrix(I7, C_LQ_PRIMAVAC)
        Debug.Print "AGUI: "; GridOrigen.TextMatrix(I7, C_LQ_AGUI)
        Debug.Print "ANTIG: "; GridOrigen.TextMatrix(I7, C_LQ_ANTIG)
        Debug.Print "COMPENS: "; GridOrigen.TextMatrix(I7, C_LQ_COMPENS)
        Debug.Print "INDEMN: "; GridOrigen.TextMatrix(I7, C_LQ_INDEMN)
        Debug.Print "BONO: "; GridOrigen.TextMatrix(I7, C_LQ_BONO)
    
    End If
    
    Debug.Print "=================================="
    
End Sub

Public Function ClasificarTipoNomina(ByVal nombreArch As String) As TipoNominaCFDI
    Dim n As String
    n = NombreBaseNomina(nombreArch)

    ' Orden importa: primero la normal, luego los mas especificos
    If EsNominaNormal(n) Then
        ClasificarTipoNomina = tnOrdinaria
    ElseIf InStr(n, "LIQ") > 0 Or InStr(n, "FINI") > 0 Then
        ClasificarTipoNomina = tnLiquidacionFiniquito
    ElseIf InStr(n, "AGUI") > 0 Then
        ClasificarTipoNomina = tnAguinaldo
    ElseIf InStr(n, "PTU") > 0 Then
        ClasificarTipoNomina = tnPTU
    ElseIf InStr(n, "BONO") > 0 Then
        ClasificarTipoNomina = tnBonoPremio
    Else
        ' Especial sin palabra clave: se trata como "ordinaria" en cuanto a tipo,
        ' pero N_ormal debe decidirse con EsNominaNormal (ver LoadPayroll).
        ClasificarTipoNomina = tnOrdinaria
    End If
End Function


Public Function AniosServicioRedondeado(ByVal fechaAlta As Date, ByVal fechaCorte As Date) As Integer
    Dim aniosCompletos As Integer
    Dim fechaAniversario As Date
    Dim fechaLimiteRedondeo As Date

    aniosCompletos = Year(fechaCorte) - Year(fechaAlta)
    fechaAniversario = DateSerial(Year(fechaAlta) + aniosCompletos, Month(fechaAlta), Day(fechaAlta))

    If fechaAniversario > fechaCorte Then
        aniosCompletos = aniosCompletos - 1
        fechaAniversario = DateSerial(Year(fechaAlta) + aniosCompletos, Month(fechaAlta), Day(fechaAlta))
    End If

    ' Límite: 6 meses + 1 día después del último aniversario cumplido
    fechaLimiteRedondeo = DateAdd("d", 1, DateAdd("m", 6, fechaAniversario))

    If fechaCorte >= fechaLimiteRedondeo Then
        AniosServicioRedondeado = aniosCompletos + 1
    Else
        AniosServicioRedondeado = aniosCompletos
    End If
End Function

Public Function TipoNominaATexto(ByVal t As TipoNominaCFDI) As String
    Select Case t
        Case tnLiquidacionFiniquito: TipoNominaATexto = "LIQ"
        Case tnAguinaldo, tnPTU, tnBonoPremio: TipoNominaATexto = "ESP"
        Case Else: TipoNominaATexto = "ORD"
    End Select
End Function

' Nombre sin ruta, en mayusculas y sin espacios sobrantes
Public Function NombreBaseNomina(ByVal nombre As String) As String
    Dim b As String
    Dim p As Long
    b = UCase$(Trim$(nombre))
    p = InStrRev(b, "\")
    If p > 0 Then b = Mid$(b, p + 1)
    NombreBaseNomina = b
End Function

' Nomina normal = MES(3 letras) + quincena (1 o 2) + ano (4 digitos), p.ej.
' OCT12026 o DIC22026 (con o sin .NOM)
Public Function EsNominaNormal(ByVal nombre As String) As Boolean
    Dim b As String
    b = NombreBaseNomina(nombre)
    If Right$(b, 4) = ".NOM" Then b = Left$(b, Len(b) - 4)
    If Not (b Like "???[12]####") Then Exit Function
    EsNominaNormal = (InStr("|ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC|", "|" & Left$(b, 3) & "|") > 0)
End Function

Public Sub SepararPTU(ByVal importePTU As Currency, ByVal salarioMinimoDiario As Currency, ByRef ptuGravada As Currency, ByRef ptuExenta As Currency)

    Dim limiteExento As Currency

    ptuGravada = 0
    ptuExenta = 0

    If importePTU <= 0 Then Exit Sub
    limiteExento = salarioMinimoDiario * 15
    If importePTU <= limiteExento Then
        ptuExenta = importePTU
    Else
        ptuExenta = limiteExento
        ptuGravada = importePTU - limiteExento
    End If
End Sub

Private Function ImporteCFDINuevo(ByVal gr As Object, ByVal fila As Long, _
                                 ByVal columna As Long) As Currency
    Dim dato As String
    If columna < 0 Then Exit Function
    If columna >= gr.Cols Then Exit Function
    If fila < 1 Or fila >= gr.Rows Then Exit Function
    dato = Trim$(gr.TextMatrix(fila, columna))
    If Len(dato) = 0 Then Exit Function
    If IsNumeric(dato) Then
        ImporteCFDINuevo = CCur(dato)
    Else
        Err.Raise vbObjectError + 301, "ImporteCFDINuevo", _
                  "Importe no numerico en fila " & CStr(fila) & _
                  ", columna " & CStr(columna) & ": " & dato
    End If
End Function

Private Sub rengSeparacionNueva(ByVal gr As Object)
    Dim importeIngreso As Currency
    Dim importeExento As Currency
    Dim sumaConceptos As Currency
    Dim deduccionGrid As Currency
    Dim creditoGrid As Currency

    On Error GoTo ErrorSeparacion

    ' Reiniciar las variables de trabajo heredadas.
    t_per = 0: t_ded = 0: t_oi = 0: T_neto = 0
    t_grav = 0: t_ext = 0: t_gravded = 0: t_extded = 0
    t_oded = 0: sub13 = 0: subc13 = 0
    sue3 = 0: via7 = 0: pva8 = 0: otr9 = 0: pee10 = 0
    agui5 = 0: ptu_1 = 0: ptu_2 = 0: ptu_3 = 0
    isr12 = 0: ims14 = 0: pre15 = 0: fon16 = 0
    pea17 = 0: ifv18 = 0
    g_ExentoAntig = 0: g_GravadoAntig = 0
    g_ExentoIndem = 0: g_GravadoIndem = 0
    g_ExentoAguiLiq = 0: g_GravadoAguiLiq = 0
    g_ExentoPVacLiq = 0: g_GravadoPVacLiq = 0

    sumaConceptos = ImporteCFDINuevo(gr, I7, C_SUELDO) + _
                    ImporteCFDINuevo(gr, I7, C_OTRAS) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_AGUI) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_ANTIG) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_PRIMAVAC) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_COMPENS) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_INDEMN) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_BONO) + _
                    ImporteCFDINuevo(gr, I7, C_LQ_SDOVAC)

    importeIngreso = ImporteCFDINuevo(gr, I7, C_TOTING)
    importeExento = ImporteCFDINuevo(gr, I7, C_EXENTA)

    If Abs(sumaConceptos - importeIngreso) > CCur(0.01) Then
        Err.Raise vbObjectError + 302, "rengSeparacionNueva", _
            "El total de conceptos no coincide con Tot.Ingr. del empleado " & _
            gr.TextMatrix(I7, C_NUM)
    End If
    If importeExento < 0 Or importeExento > importeIngreso Then
        Err.Raise vbObjectError + 303, "rengSeparacionNueva", _
            "Perc.exenta fuera del rango de ingresos."
    End If

    t_per = importeIngreso
    t_ext = importeExento
    t_grav = importeIngreso - importeExento
    pee10 = importeExento
    sue3 = ImporteCFDINuevo(gr, I7, C_SUELDO) + _
           ImporteCFDINuevo(gr, I7, C_LQ_SDOVAC)
    otr9 = ImporteCFDINuevo(gr, I7, C_OTRAS)
    agui5 = ImporteCFDINuevo(gr, I7, C_LQ_AGUI)

    isr12 = ImporteCFDINuevo(gr, I7, C_ISPT)
    ims14 = ImporteCFDINuevo(gr, I7, C_IMSS)
    pre15 = ImporteCFDINuevo(gr, I7, C_PRESTAMOS)
    fon16 = ImporteCFDINuevo(gr, I7, C_FONACOT)
    pea17 = ImporteCFDINuevo(gr, I7, C_PENSION)
    ifv18 = ImporteCFDINuevo(gr, I7, C_INFONAVIT)
    creditoGrid = ImporteCFDINuevo(gr, I7, C_SUBEMP)

    ' Un subsidio negativo en la cuadricula es un otro pago positivo.
    If creditoGrid < 0 Then
        sub13 = -creditoGrid
    End If

    t_ded = isr12 + ims14 + pre15 + fon16 + pea17 + ifv18
    If creditoGrid > 0 Then t_ded = t_ded + creditoGrid
    t_oded = t_ded - isr12
    t_extded = t_oded
    T_neto = t_per + sub13 - t_ded

    deduccionGrid = ImporteCFDINuevo(gr, I7, C_TOTDED)
    If Abs((t_ded - sub13) - deduccionGrid) > CCur(0.01) Then
        Err.Raise vbObjectError + 304, "rengSeparacionNueva", _
            "Las deducciones no coinciden con Tot.Deduc."
    End If
    If Abs(T_neto - ImporteCFDINuevo(gr, I7, C_NETO)) > CCur(0.01) Then
        Err.Raise vbObjectError + 305, "rengSeparacionNueva", _
            "El neto calculado no coincide con el neto del Generador."
    End If

    Exit Sub

ErrorSeparacion:
    Err.Raise Err.Number, "rengSeparacionNueva", Err.Description
End Sub


