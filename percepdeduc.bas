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

Public Function DeterminarTipoNomina(ByVal nombreArchivo As String, Optional ByVal OpcionForm8 As Integer = -1) As TipoNominaCFDI
    Dim TipoPorNombre As TipoNominaCFDI
    TipoPorNombre = ClasificarTipoNomina(nombreArchivo)

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
    
    '=========================================================
    ' CFDI - NOMINA ESPECIAL AGUINALDO
    '
    ' Se utiliza una rutina independiente para evitar:
    ' - Mezclar exento de aguinaldo con prima vacacional
    ' - Crear un ISR artificial de $0.01
    ' - Omitir o duplicar percepciones
    '
    ' No afecta nomina normal, PTU, Bono ni separaciones.
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Then
        If GridOrigen Is Nothing Then
            Err.Raise vbObjectError + 5700, "percepdeduc.reng", "No se encontro el grid de captura CFDI."
        End If
        rengAguinaldoCFDI GridOrigen
        Exit Sub
    End If
    
    
    '=========================================================
    ' CFDI - NOMINAS ESPECIALES PTU Y BONO
    '
    ' Cada modalidad utiliza solamente:
    '
    ' PTU:
    '   C_CONC6  = PTU gravada
    '   C_EXENTA = PTU exenta
    '
    ' BONO:
    '   C_CONC7  = Bono gravado
    '   C_EXENTA = sin importe para P038E en 2026
    '
    ' Se conservan las deducciones capturadas.
    ' No se modifica la nomina ordinaria ni AGUINALDO.
    '=========================================================

    Select Case g_TipoNominaActiva
        Case tnPTU, tnBonoPremio
            rengPTUBonoCFDI GridOrigen
            Exit Sub
    End Select
    
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
    
    If IsNumeric(GridOrigen.TextMatrix(I7, C_NUM)) Then
    
        Dim pruebaArchivo As Long
        Dim numeroErrorCMP As Long
        Dim descripcionErrorCMP As String
        Dim registroEmpleado As Long
    
        registroEmpleado = CLng(GridOrigen.TextMatrix(I7, C_NUM))
    
        'Comprobar que el archivo CMP este disponible.
        Err.Clear
        On Error Resume Next
    
        pruebaArchivo = LOF(14)
        numeroErrorCMP = Err.Number
        descripcionErrorCMP = Err.Description
    
        On Error GoTo 0
    
        If numeroErrorCMP <> 0 Then
            Err.Raise vbObjectError + 5350, "percepdeduc.reng", "No se pudo acceder al CMP del empleado " & CStr(registroEmpleado) & ". " & descripcionErrorCMP
    
        End If
    
        'Validar que exista el registro necesario.
        If registroEmpleado < 1 Or registroEmpleado > (pruebaArchivo \ Len(nom_com)) Then
            Err.Raise vbObjectError + 5351, "percepdeduc.reng", "No existe el registro CMP del empleado " & CStr(registroEmpleado)
    
        End If
    
        'Conservar la lectura original.
        Get #14, registroEmpleado, nom_com
        subc13 = nom_com.subdio
    
    End If
     
    'percepciones---------------------------------------------------
    If IsNumeric(GridOrigen.TextMatrix(I7, 3)) Then
        sue3 = GridOrigen.TextMatrix(I7, 3)
        c_per = c_per + 1: t_per = t_per + sue3
    Else
        sue3 = 0
    End If
    
    'Parche Aguinaldo 19/12/2017 Usar el tipo de periodo identificado para CFDI.
    If g_QuincenaCFDI2026 = 0 Then
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
    If p > 0 Then b = mID$(b, p + 1)
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

Public Sub rengAguinaldoCFDI(ByVal gr As Object)

    Dim brutoAguinaldoCFDI As Currency
    Dim horasSinClasificar As Currency
    Dim ptuSinClasificar As Currency
    Dim subsidioCapturado As Currency

    Dim totalConceptos As Currency
    Dim totalIngresosGrid As Currency
    Dim totalDeduccionesGrid As Currency
    Dim netoGrid As Currency

    Dim registroEmpleado As Long
    Dim tamanoCMP As Long
    Dim errorCMP As Long
    Dim descripcionCMP As String

    '=========================================================
    ' 1. VALIDAR LA FILA DE CAPTURA
    '=========================================================

    If I7 < 1 Or I7 >= gr.Rows Then
        Err.Raise vbObjectError + 5701, "rengAguinaldoCFDI", "Fila de empleado fuera de rango."

    End If

    If gr.Cols <= C_NETO Then
        Err.Raise vbObjectError + 5702, "rengAguinaldoCFDI", "El grid no contiene todas las columnas de AGUINALDO."

    End If

    If Not IsNumeric(gr.TextMatrix(I7, C_NUM)) Then
        Err.Raise vbObjectError + 5703, "rengAguinaldoCFDI", "No existe un numero de empleado valido."

    End If

    registroEmpleado = CLng(gr.TextMatrix(I7, C_NUM))

    If registroEmpleado <= 0 Then
        Err.Raise vbObjectError + 5704, "rengAguinaldoCFDI", "El numero de empleado debe ser positivo."

    End If

    '=========================================================
    ' 2. REINICIAR VARIABLES GLOBALES CFDI
    '
    ' Evitar conservar cantidades del empleado anterior.
    '=========================================================

    sue3 = 0
    agui5 = 0

    ptu_1 = 0
    ptu_2 = 0
    ptu_3 = 0

    via7 = 0
    pva8 = 0
    otr9 = 0
    pee10 = 0

    sub13 = 0
    subc13 = 0

    isr12 = 0
    ims14 = 0
    pre15 = 0
    fon16 = 0
    pea17 = 0
    ifv18 = 0

    t_per = 0
    t_ded = 0
    T_neto = 0

    t_grav = 0
    t_ext = 0

    t_gravded = 0
    t_extded = 0
    t_oded = 0
    t_oi = 0

    
    '=========================================================
    ' 3. LEER LAS PERCEPCIONES DE AGUINALDO
    '
    ' REGLA DE CAPTURA:
    '
    ' C_CONC5  = Aguinaldo total
    ' C_EXENTA = Parte exenta del aguinaldo
    '
    ' REGLA CFDI:
    '
    ' agui5 = Aguinaldo gravado
    ' pee10 = Aguinaldo exento
    '
    ' El exento forma parte del total capturado.
    '=========================================================

    sue3 = ImporteCFDINuevo(gr, I7, C_SUELDO)
    horasSinClasificar = ImporteCFDINuevo(gr, I7, C_CONC4)

    '---------------------------------------------------------
    ' AGUINALDO BRUTO
    '
    ' Leer el importe completo de la columna Aguinaldo.
    '---------------------------------------------------------

    brutoAguinaldoCFDI = ImporteCFDINuevo(gr, I7, C_CONC5)

    '---------------------------------------------------------
    ' PTU CAPTURADA EN AGUINALDO
    '
    ' Conservar la validacion posterior existente.
    '---------------------------------------------------------

    ptuSinClasificar = ImporteCFDINuevo(gr, I7, C_CONC6)

    '---------------------------------------------------------
    ' OTROS CONCEPTOS
    '
    ' Conservar las lecturas existentes.
    '---------------------------------------------------------
    
    via7 = ImporteCFDINuevo(gr, I7, C_CONC7)
    pva8 = ImporteCFDINuevo(gr, I7, C_CONC8)
    otr9 = ImporteCFDINuevo(gr, I7, C_OTRAS)

    '---------------------------------------------------------
    ' AGUINALDO EXENTO
    '
    ' No es una percepcion adicional al bruto.
    '---------------------------------------------------------

    pee10 = ImporteCFDINuevo(gr, I7, C_EXENTA)

    '=========================================================
    ' 4. VALIDAR EL DESGLOSE
    '=========================================================

    If brutoAguinaldoCFDI < 0 Then
        Err.Raise vbObjectError + 5902, "rengAguinaldoCFDI", "El aguinaldo total no puede ser negativo."
    End If

    If pee10 < 0 Then
        Err.Raise vbObjectError + 5903, "rengAguinaldoCFDI", "La percepcion exenta no puede ser negativa."

    End If

    If pee10 > brutoAguinaldoCFDI Then
        Err.Raise vbObjectError + 5904, "rengAguinaldoCFDI", "La percepcion exenta supera el aguinaldo total del empleado " & CStr(registroEmpleado)
    End If

    '=========================================================
    ' 5. DETERMINAR AGUINALDO GRAVADO
    '
    ' Ejemplo:
    '
    ' Bruto:   5000
    ' Exento:  1500
    ' Gravado: 3500
    '
    ' Este importe se utiliza para P002G.
    '=========================================================

    agui5 = brutoAguinaldoCFDI - pee10

    'No modificar la columna Aguinaldo del grid.
    'No modificar la percepcion exenta del grid.
    'No volver a restar la exencion de agui5.

    '=========================================================
    ' 4. VALIDAR CONCEPTOS SIN DESGLOSE FISCAL
    '
    ' No inventar destinos ni exenciones.
    '=========================================================

    If horasSinClasificar <> 0 Then
        Err.Raise vbObjectError + 5705, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": existe importe en hs.Norm." & vbCrLf & "No hay un desglose fiscal confirmado para timbrar esta columna en AGUINALDO."

    End If

    If ptuSinClasificar <> 0 Then

        Err.Raise vbObjectError + 5706, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": existe PTU dentro de AGUINALDO." & vbCrLf & "La columna Perc.exenta. se utiliza " & "para el aguinaldo." & vbCrLf & _
            "No es posible determinar que parte " & "de la PTU es gravada o exenta." & vbCrLf & "Debe identificarse su desglose fiscal."

    End If

    '=========================================================
    ' 5. VALIDAR IMPORTES NEGATIVOS
    '
    ' Las percepciones reportadas en estos destinos
    ' no deben generar importes negativos.
    '=========================================================

    If sue3 < 0 Or agui5 < 0 Or via7 < 0 _
       Or pva8 < 0 Or otr9 < 0 Or pee10 < 0 Then
        Err.Raise vbObjectError + 5707, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": existen percepciones negativas."

    End If

    '=========================================================
    ' 6. TOTAL DE PERCEPCIONES
    '
    ' La captura ya separo:
    '
    ' C5  = aguinaldo gravado
    ' C10 = aguinaldo exento
    '
    ' NO volver a calcular la exencion.
    ' NO utilizar el tope de prima vacacional.
    '=========================================================

    totalConceptos = sue3 + agui5 + via7 + pva8 + otr9 + pee10
    t_per = totalConceptos
    t_ext = pee10
    t_grav = t_per - t_ext

    totalIngresosGrid = _
        ImporteCFDINuevo(gr, I7, C_TOTING)

    If Abs(t_per - totalIngresosGrid) > CCur(0.01) Then
        Err.Raise vbObjectError + 5708, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": las percepciones no coinciden." & vbCrLf & "Captura: " & Format$(totalIngresosGrid, "#,##0.00") & vbCrLf & "CFDI: " & Format$(t_per, "#,##0.00")
    End If

    If t_ext > t_per Then
        Err.Raise vbObjectError + 5709, "rengAguinaldoCFDI", "El importe exento supera las percepciones."

    End If

    If t_per <= 0 Then
        Err.Raise vbObjectError + 5710, "rengAguinaldoCFDI", "No existen percepciones para el CFDI."

    End If

    '=========================================================
    ' 7. LEER DEDUCCIONES EXACTAMENTE COMO SE CAPTURARON
    '
    ' No crear una retencion de ISR de $0.01.
    '=========================================================

    isr12 = ImporteCFDINuevo(gr, I7, C_ISPT)
    ims14 = ImporteCFDINuevo(gr, I7, C_IMSS)
    pre15 = ImporteCFDINuevo(gr, I7, C_PRESTAMOS)
    fon16 = ImporteCFDINuevo(gr, I7, C_FONACOT)
    pea17 = ImporteCFDINuevo(gr, I7, C_PENSION)
    ifv18 = ImporteCFDINuevo(gr, I7, C_INFONAVIT)

    If isr12 < 0 Or ims14 < 0 Or pre15 < 0 Or fon16 < 0 Or pea17 < 0 Or ifv18 < 0 Then

        Err.Raise vbObjectError + 5711, "rengAguinaldoCFDI", "Existen deducciones negativas en el empleado " & CStr(registroEmpleado)

    End If

    t_ded = isr12 + ims14 + pre15 + fon16 + pea17 + ifv18
    t_oded = t_ded - isr12
    t_extded = t_oded

    '=========================================================
    ' 8. SUBSIDIO AL EMPLEO
    '
    ' Convencion existente del sistema:
    '
    ' C_SUBEMP negativo = subsidio a favor del empleado
    '
    ' El importe positivo se envia a OP002.
    '=========================================================

    subsidioCapturado = ImporteCFDINuevo(gr, I7, C_SUBEMP)

    If subsidioCapturado > 0 Then

        Err.Raise vbObjectError + 5712, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": Sub.P/Empl. contiene un importe positivo." & vbCrLf & "No se puede clasificar como OP002 sin revisar el signo y su tratamiento."

    End If

    If subsidioCapturado < 0 Then
        sub13 = -subsidioCapturado
    Else
        sub13 = 0
    End If

    'El subsidio no se suma a las percepciones.
    'El subsidio pagado incrementa el neto.

    T_neto = t_per + sub13 - t_ded

    '=========================================================
    ' 9. COMPARAR CONTRA TOTALES DE LA CAPTURA
    '
    ' La columna Tot.Deduc. incluye el signo que
    ' utiliza C_SUBEMP en el grid original.
    '=========================================================

    totalDeduccionesGrid = ImporteCFDINuevo(gr, I7, C_TOTDED)
    netoGrid = ImporteCFDINuevo(gr, I7, C_NETO)

    If Abs((t_ded - sub13) - totalDeduccionesGrid) > CCur(0.01) Then

        Err.Raise vbObjectError + 5713, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": Tot.Deduc. no coincide con el CFDI." & vbCrLf & "Captura: " & Format$(totalDeduccionesGrid, "#,##0.00") & _
            vbCrLf & "Calculado: " & Format$(t_ded - sub13, "#,##0.00")

    End If

    If Abs(T_neto - netoGrid) > CCur(0.01) Then
        Err.Raise vbObjectError + 5714, "rengAguinaldoCFDI", "Empleado " & CStr(registroEmpleado) & ": el neto del CFDI no coincide." & vbCrLf & "Captura: " & Format$(netoGrid, "#,##0.00") & vbCrLf & "CFDI: " & Format$(T_neto, "#,##0.00")

    End If

    '=========================================================
    ' 10. RECUPERAR SUBSIDIO CAUSADO DEL CMP
    '
    ' Conservar el mismo origen utilizado por reng().
    '=========================================================

    Err.Clear
    On Error Resume Next

    tamanoCMP = LOF(14)
    errorCMP = Err.Number
    descripcionCMP = Err.Description

    On Error GoTo 0

    If errorCMP <> 0 Then

        Err.Raise vbObjectError + 5715, "rengAguinaldoCFDI", "No fue posible consultar el CMP. " & descripcionCMP
    End If

    If registroEmpleado > (tamanoCMP \ Len(nom_com)) Then
        Err.Raise vbObjectError + 5716, "rengAguinaldoCFDI", "No existe el registro CMP del empleado " & CStr(registroEmpleado)
    End If

    Get #14, registroEmpleado, nom_com

    subc13 = nom_com.subdio

End Sub

Public Sub rengPTUBonoCFDI(ByVal gr As Object)

    Dim columnaPrincipal As Long
    Dim nombreConcepto As String
    Dim columna As Long

    Dim importeGravado As Currency
    Dim importeExento As Currency
    Dim importeIngresos As Currency

    Dim totalDeduccionesGrid As Currency
    Dim netoGrid As Currency
    Dim subsidioCapturado As Currency

    Dim registroEmpleado As Long
    Dim tamanoCMP As Long
    Dim numeroErrorCMP As Long
    Dim descripcionErrorCMP As String

    '=========================================================
    ' 1. IDENTIFICAR LA MODALIDAD
    '=========================================================

    Select Case g_TipoNominaActiva

        Case tnPTU
            columnaPrincipal = C_CONC6
            nombreConcepto = "PTU"

        Case tnBonoPremio
            columnaPrincipal = C_CONC7
            nombreConcepto = "BONO"

        Case Else
            Err.Raise vbObjectError + 5800, "rengPTUBonoCFDI", "La modalidad no es PTU ni BONO."

    End Select

    If gr Is Nothing Then
        Err.Raise vbObjectError + 5801, "rengPTUBonoCFDI", "No existe el grid de origen."
    End If

    If I7 < 1 Or I7 >= gr.Rows Then
        Err.Raise vbObjectError + 5802, "rengPTUBonoCFDI", "Fila de empleado fuera de rango."
    End If

    If gr.Cols <= C_NETO Then
        Err.Raise vbObjectError + 5803, "rengPTUBonoCFDI", "Faltan columnas en la captura."
    End If

    If Not IsNumeric(gr.TextMatrix(I7, C_NUM)) Then
        Err.Raise vbObjectError + 5804, "rengPTUBonoCFDI", "Numero de empleado no valido."
    End If

    registroEmpleado = CLng(gr.TextMatrix(I7, C_NUM))

    If registroEmpleado <= 0 Then
        Err.Raise vbObjectError + 5805, "rengPTUBonoCFDI", "Numero de empleado incorrecto."
    End If

    '=========================================================
    ' 2. REINICIAR VARIABLES DEL CFDI
    '
    ' Evitar cantidades heredadas del empleado anterior.
    '=========================================================

    sue3 = 0
    agui5 = 0

    ptu_1 = 0
    ptu_2 = 0
    ptu_3 = 0

    via7 = 0
    pva8 = 0
    otr9 = 0
    pee10 = 0

    isr12 = 0
    ims14 = 0
    pre15 = 0
    fon16 = 0
    pea17 = 0
    ifv18 = 0

    sub13 = 0
    subc13 = 0

    t_per = 0
    t_ded = 0
    T_neto = 0

    t_grav = 0
    t_ext = 0

    t_oded = 0
    t_oi = 0
    t_gravded = 0
    t_extded = 0

    '=========================================================
    ' 3. VALIDAR LAS COLUMNAS DE PERCEPCIONES
    '
    ' PTU: solamente columna 6 y columna 10.
    ' BONO: solamente columna 7 y columna 10.
    '
    ' No borrar importes de otras columnas.
    ' Si existen, impedir un CFDI con importes omitidos.
    '
    ' Las deducciones siguen permitidas.
    '=========================================================

    For columna = C_SUELDO To C_OTRAS

        If columna <> columnaPrincipal Then

            If ImporteCFDINuevo(gr, I7, columna) <> 0 Then

                Err.Raise vbObjectError + 5806, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & " - " & nombreConcepto & vbCrLf & "Existe un importe en la columna " & gr.TextMatrix(0, columna) & "." & vbCrLf & _
                    "Esta nomina especial solamente permite el concepto " & nombreConcepto & " y su percepcion exenta."
            End If
        End If
    Next columna

    '=========================================================
    ' 4. LEER CONCEPTO GRAVADO Y PERCEPCION EXENTA
    '
    ' No recalcular exenciones en el CFDI.
    ' La captura ya contiene los importes definitivos.
    '=========================================================

    importeGravado = ImporteCFDINuevo(gr, I7, columnaPrincipal)
    importeExento = ImporteCFDINuevo(gr, I7, C_EXENTA)

    If importeGravado < 0 Or importeExento < 0 Then
        Err.Raise vbObjectError + 5807, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": no se permiten percepciones negativas."

    End If

    '=========================================================
    ' 5. DISTRIBUCION POR MODALIDAD
    '=========================================================

    Select Case g_TipoNominaActiva

        Case tnPTU

            'La columna 6 ya contiene la PTU gravada.
            'La columna 10 contiene la PTU exenta.
            'No volver a aplicar el limite de exencion.

            ptu_1 = importeGravado
            ptu_2 = importeExento
            ptu_3 = ptu_1 + ptu_2

            pee10 = ptu_2

        Case tnBonoPremio

            'Para 2026, P038 es totalmente gravado.
            'No transformar ni perder una exencion
            'capturada sin clasificacion fiscal valida.

            If importeExento <> 0 Then

                Err.Raise vbObjectError + 5808, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": BONO contiene " & Format$(importeExento, "#,##0.00") & " en Perc.exenta." & vbCrLf & "La clave SAT 038 debe ser " & _
                    "totalmente gravada en 2026." & vbCrLf & "El contador debe confirmar la clasificacion de este importe."
            End If

            via7 = importeGravado
            pee10 = 0

    End Select

    '=========================================================
    ' 6. TOTALES DE PERCEPCIONES
    '=========================================================

    t_per = importeGravado + importeExento

    t_ext = importeExento
    t_grav = importeGravado

    If t_per <= 0 Then
        Err.Raise vbObjectError + 5809, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": no existe importe de " & nombreConcepto & " para generar su CFDI."
    End If

    importeIngresos = ImporteCFDINuevo(gr, I7, C_TOTING)

    If Abs(t_per - importeIngresos) > CCur(0.01) Then
        Err.Raise vbObjectError + 5810, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": Tot.Ingr. no coincide con " & nombreConcepto & "." & vbCrLf & "Captura: " & Format$(importeIngresos, "#,##0.00") & _
            vbCrLf & "CFDI: " & Format$(t_per, "#,##0.00")

    End If

    '=========================================================
    ' 7. LEER DEDUCCIONES DEL EMPLEADO
    '
    ' No inventar ISR de $0.01.
    ' Utilizar los importes existentes de la captura.
    '=========================================================

    isr12 = ImporteCFDINuevo(gr, I7, C_ISPT)
    ims14 = ImporteCFDINuevo(gr, I7, C_IMSS)

    pre15 = ImporteCFDINuevo(gr, I7, C_PRESTAMOS)
    fon16 = ImporteCFDINuevo(gr, I7, C_FONACOT)

    pea17 = ImporteCFDINuevo(gr, I7, C_PENSION)
    ifv18 = ImporteCFDINuevo(gr, I7, C_INFONAVIT)

    If isr12 < 0 Or ims14 < 0 Or pre15 < 0 _
       Or fon16 < 0 Or pea17 < 0 Or ifv18 < 0 Then
        Err.Raise vbObjectError + 5811, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": existen deducciones negativas."

    End If

    t_ded = isr12 + ims14 + pre15 + fon16 + pea17 + ifv18
    t_oded = t_ded - isr12
    t_extded = t_oded

    '=========================================================
    ' 8. SUBSIDIO AL EMPLEO
    '
    ' Se conserva la convencion existente:
    ' Subsidio negativo en grid = pago al trabajador.
    '=========================================================

    subsidioCapturado = ImporteCFDINuevo(gr, I7, C_SUBEMP)

    If subsidioCapturado > 0 Then
        Err.Raise vbObjectError + 5812, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": Sub.P/Empl. contiene un importe positivo." & vbCrLf & "Revise su tratamiento antes del CFDI."

    End If

    If subsidioCapturado < 0 Then
        sub13 = -subsidioCapturado
    Else
        sub13 = 0
    End If

    T_neto = t_per + sub13 - t_ded

    '=========================================================
    ' 9. VALIDAR DEDUCCIONES Y NETO
    '
    ' En el grid, el subsidio negativo reduce
    ' Tot.Deduc. e incrementa el neto.
    '=========================================================

    totalDeduccionesGrid = ImporteCFDINuevo(gr, I7, C_TOTDED)

    netoGrid = ImporteCFDINuevo(gr, I7, C_NETO)

    If Abs((t_ded - sub13) - totalDeduccionesGrid) > CCur(0.01) Then

        Err.Raise vbObjectError + 5813, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": las deducciones no coinciden." & vbCrLf & "Captura: " & Format$(totalDeduccionesGrid, "#,##0.00") & _
            vbCrLf & "CFDI: " & Format$(t_ded - sub13, "#,##0.00")

    End If

    If Abs(T_neto - netoGrid) > CCur(0.01) Then

        Err.Raise vbObjectError + 5814, "rengPTUBonoCFDI", "Empleado " & CStr(registroEmpleado) & ": el neto no coincide con la captura." & vbCrLf & "Captura: " & Format$(netoGrid, "#,##0.00") & vbCrLf & _
            "CFDI: " & Format$(T_neto, "#,##0.00")

    End If

    '=========================================================
    ' 10. OBTENER SUBSIDIO CAUSADO DEL CMP
    '
    ' Mantener el origen original de reng().
    '=========================================================

    Err.Clear
    On Error Resume Next

    tamanoCMP = LOF(14)
    numeroErrorCMP = Err.Number
    descripcionErrorCMP = Err.Description

    On Error GoTo 0

    If numeroErrorCMP <> 0 Then
        Err.Raise vbObjectError + 5815, "rengPTUBonoCFDI", "No se pudo leer el CMP del empleado " & CStr(registroEmpleado) & ". " & descripcionErrorCMP

    End If

    If registroEmpleado > (tamanoCMP \ Len(nom_com)) Then
        Err.Raise vbObjectError + 5816, "rengPTUBonoCFDI", "No existe el CMP del empleado " & CStr(registroEmpleado)

    End If

    Get #14, registroEmpleado, nom_com

    subc13 = nom_com.subdio

End Sub
