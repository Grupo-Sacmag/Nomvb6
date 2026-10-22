VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AJTECH 
   Caption         =   "Ajuste en MILIMETROS  impresión de cheques"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid AJTE 
      Height          =   3720
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   6562
      _Version        =   393216
      Rows            =   15
      Cols            =   3
      BackColorFixed  =   -2147483634
      FocusRect       =   2
      GridLinesFixed  =   1
   End
   Begin VB.Menu ach 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu achArch 
         Caption         =   "&Guardar"
         Index           =   1
      End
      Begin VB.Menu ArcSep1 
         Caption         =   "-"
      End
      Begin VB.Menu achImpr 
         Caption         =   "&Impresion datos"
         Index           =   2
      End
      Begin VB.Menu achPrue 
         Caption         =   "&Prueba Impresion"
         Index           =   3
      End
      Begin VB.Menu ArSep2 
         Caption         =   "-"
      End
      Begin VB.Menu achSal 
         Caption         =   "&Salida"
         Index           =   4
      End
   End
   Begin VB.Menu EdAj 
      Caption         =   "&Edición"
      Begin VB.Menu EdAjCop 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "AJTECH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim valcelant
Dim fila(50) As Integer, ubica(50) As Integer, rubro(50) As String * 30
' ==============================================================================
' SUBRUTINA depura MODIFICADA
' ==============================================================================
Sub depura(elemento)
    Dim elemUpper As String
    elemUpper = UCase$(Trim$(elemento))
    
    ' 1. Si no es totalmente numérico, evaluamos si es un nombre de archivo válido
    If Not IsNumeric(elemento) Then
        
        ' Permitir prefijos de nóminas especiales/liquidaciones (LIQ..., ESP..., FINIQ..., AGUI..., PTU...)
        If Left$(elemUpper, 3) = "LIQ" Or Left$(elemUpper, 3) = "ESP" _
           Or InStr(elemUpper, "FINIQ") > 0 Or InStr(elemUpper, "AGUI") > 0 _
           Or InStr(elemUpper, "PTU") > 0 Or InStr(elemUpper, "BONO") > 0 Then
            
            ' Es un nombre de archivo alfanumérico válido, lo dejamos pasar
            Exit Sub
            
        Else
            ' Si la celda era para importes numéricos y se ingresó texto no válido
            MsgBox "Solo se aceptan valores numéricos o nombres de nómina válidos (LIQ..., ESP..., FINIQ...)", vbExclamation, "Validación de Captura"
            Text1.Text = valcelant
        End If
        
    End If
End Sub

Private Sub achArch_Click(Index As Integer)
   archaj
End Sub

Private Sub achImpr_Click(Index As Integer)
   AJTECH.PrintForm
End Sub
Private Sub achSal_Click(Index As Integer)
  Close 8
  Unload AJTECH
End Sub

Private Sub AJTE_EnterCell()
   If AJTE.Row = 1 And AJTE.col = 2 Then AJTE.col = 1: AJTE_EnterCell
   If AJTE.Row = 7 And AJTE.col = 2 Then AJTE.col = 1: AJTE_EnterCell
   If AJTE.Row = 8 And AJTE.col = 2 Then AJTE.col = 1: AJTE_EnterCell
   If AJTE.Row > 8 And AJTE.col = 1 Then AJTE.col = 2: AJTE_EnterCell
    AJTE.CellBackColor = vbGreen
End Sub

Private Sub EdAjCop_Click()
   Clipboard.Clear
   Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(0, 0) & Chr(9)
   For f = 0 To AJTE.Cols - 1
         Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(0, f) & Chr(9)
   Next f
        Clipboard.SetText Clipboard.GetText + Chr(13)
        
   For i = 1 To AJTE.Rows - 1
           Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(i, 0) & Chr(9)
      For f = 1 To AJTE.Cols - 1
            Clipboard.SetText Clipboard.GetText + AJTE.TextMatrix(i, f) & Chr(9)
      Next f
      Clipboard.SetText Clipboard.GetText + Chr(13)
   Next i
   difer = AJTE.RowSel - AJTE.Row

End Sub

Private Sub Text1_Change()
    AJTE.Text = Text1.Text
End Sub

Private Sub Text1_GotFocus()
    SendKeys "{end}"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Rem AJTE.SetFocus
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        
        Case 13 ' --- AL PRESIONAR ENTER ---
            
            ' 1. Validar la entrada (permite números e identificadores LIQ, ESP, FINIQ)
            depura Text1.Text
            
            ' 2. Asignar el texto a la celda activa del grid
            AJTE.Text = Text1.Text
            
            ' 3. CAPTURA Y DETECCIÓN AUTOMÁTICA DEL TIPO DE NÓMINA
            NombreArchivoNomina = Trim$(Text1.Text)
            g_TipoNominaActiva = DeterminarTipoNomina(NombreArchivoNomina)
            
            ' Registro en la ventana Inmediato para confirmación del programador
            Debug.Print "Form8 -> Nombre Nómina Capturado: " & NombreArchivoNomina
            Debug.Print "Form8 -> Tipo Nómina Activa (Enum): " & g_TipoNominaActiva
            
            ' 4. Devolver foco a la grilla
            AJTE.SetFocus
            
        Case 27 ' --- AL PRESIONAR ESCAPE ---
            Text1.Text = valcelant
            AJTE.SetFocus
            
    End Select
End Sub

Private Sub AJTE_LeaveCell()
  AJTE.CellBackColor = vbWhite
End Sub

Private Sub AJTE_RowColChange()
    Rem AJTE.BackColor = vbWhite
    If AJTE.Text <> "" Then valcelant = AJTE.Text
    Text1.Text = AJTE.Text
End Sub
Private Sub AJTE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            AJTE.Text = ""
            Text1.Text = AJTE.Text
        Case vbKeyF2
            If AJTE.Text <> "" Then valcelant = AJTE.Text
            Text1.Text = AJTE.Text
            Text1.SetFocus
    End Select
End Sub

Private Sub AJTE_KeyPress(KeyAscii As Integer)
    Rem ajte.SelectionMode = flexSelectionFree
    If AJTE.Text <> "" Then valcelant = AJTE.Text
    Text1.Text = Chr$(KeyAscii)
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorGeneracionCFDI

    Dim GridOrigenCFDI As Object
    Dim ultimaFila As Long
    Dim x8 As Long
    Dim x9 As Long

    ' Comprobar la fecha de separacion nueva ANTES de crear una plantilla.
    ' Para nominas ordinarias / legacy se mantiene el codigo preexistente.
    If EsSeparacionNueva() Then
        If Len(Trim$(g_FechaPagoCFDI2026)) = 0 Then
            Err.Raise vbObjectError + 5210, "NOMCF2.Form_Load", _
                      "No se recibio la fecha de pago de la nueva nomina."
        End If
        Call FechaCFDI_InterpretarDMY(g_FechaPagoCFDI2026)
    End If

    IniCols

    Close 10
    Close 7

    Open "EMP_CFDI.DNO" For Random As 10 Len = Len(EmpCFDI)
    Open "Perscfdi.dno" For Random As 7 Len = Len(Empleado_1)
    Open "PerOtre.dno" For Random As 9 Len = Len(Otros_Rgtros)

    Get 10, 1, EmpCFDI

    Folio = InputBox("Dame el numero de Folio Anterior", "CFDI NOMINA", EmpCFDI.Folio)
    serie = InputBox("Dame la serie", "CFDI NOMINA", EmpCFDI.serie)
    ConNomina = InputBox("Numero de nomina consecutivo ", "CFDI NOMINA", (EmpCFDI.consecutivo + 1))

ERRORDEPAGO:
    Met_pago = InputBox("Metodo de Pago (utiliza la clave numerica)" & Chr(13) & "01" & Chr(9) & "EFECTIVO" & Chr(13) & "02" & Chr(9) & "CHEQUE NOMINATIVO" & Chr(13) & "03" & Chr(9) & "TRANSFERENCIA ELECTRONICA" & Chr(13) & "28" & Chr(9) & "TARJETA DE DEBITO", "CFDI NOMINA", "28")

    Select Case Met_pago
        Case "01", "02", "03", "28"
            If Met_pago = "01" Then Met_pagcon = "01 Eectivo"
            If Met_pago = "02" Then Met_pagcon = "02 Cheque nominativo"
            If Met_pago = "03" Then Met_pagcon = "03 TRANSFERENCIA"
            If Met_pago = "28" Then Met_pagcon = "28 Tarjeta de Debito"
        Case Else
            MsgBox "LA CLAVE NO EXISTE"
            GoTo ERRORDEPAGO
    End Select

    Reg_Patr = InputBox("Registro Patronal Imss ", "CFDI NOMINA", EmpCFDI.RegPatr)
    EntradaDATOS = "Riesgo Patronal(utiliza la clave numerica)" & Chr(13) & "1" & Chr(9) & "CLASE I" & Chr(13) & "2" & Chr(9) & "CLASE II" & Chr(13) & "3" & Chr(9) & "CLASE III" & Chr(13) & "4" & Chr(9) & "CLASE IV" & Chr(13) & "5" & Chr(9) & "CLASE V" & Chr(13)

DENUEVO:
    RiesgoImss = InputBox(EntradaDATOS, "CFDI NOMINA", 1)

    Select Case RiesgoImss
        Case 1 To 5
            If RiesgoImss = "1" Then RiesgoImss = "1-Clase I"
            If RiesgoImss = "2" Then RiesgoImss = "2-Clase II"
            If RiesgoImss = "3" Then RiesgoImss = "3-Clase III"
            If RiesgoImss = "4" Then RiesgoImss = "4-Clase IV"
            If RiesgoImss = "5" Then RiesgoImss = "5-Clase V"
        Case Else
            MsgBox "LA CLAVE NO EXISTE"
            GoTo DENUEVO
    End Select

    EmpCFDI.Folio = Folio
    EmpCFDI.consecutivo = ConNomina
    EmpCFDI.serie = serie
    EmpCFDI.RegPatr = Reg_Patr
    EmpCFDI.RiesgoImss = RiesgoImss
    Put 10, 1, EmpCFDI

    I8 = 0

    Get 10, 1, EmpCFDI
    EmpCFDI.Folio = Folio
    Put 10, 1, EmpCFDI
    Close 10

    SDI_Click (Index)
    fila = 0

    ' =========================================================
    ' GENERACION CFDI DESDE FORM8 O FORMVIEWER
    ' =========================================================
    If OrigenCFDI = 1 Then
        Set GridOrigenCFDI = FormViewer.ConNom1
    Else
        Set GridOrigenCFDI = Form8.ConNom1
    End If

    If GridOrigenCFDI.Rows < 3 Then
        MsgBox "No hay empleados para generar el CFDI.", _
               vbExclamation, "CFDI"
        GoTo FinGeneracionCFDI
    End If

    ultimaFila = GridOrigenCFDI.Rows - 2

    Debug.Print "=================================="
    Debug.Print "INICIO GENERACION CFDI"
    Debug.Print "Origen: "; OrigenCFDI
    Debug.Print "Rows origen: "; GridOrigenCFDI.Rows
    Debug.Print "Ultima fila: "; ultimaFila
    Debug.Print "Cols origen: "; GridOrigenCFDI.Cols

    For I7 = 1 To ultimaFila
        If Len(Trim$(GridOrigenCFDI.TextMatrix(I7, C_NUM))) > 0 Then
            If IsNumeric(GridOrigenCFDI.TextMatrix(I7, C_NUM)) Then
                NumerodePersonal = CLng(GridOrigenCFDI.TextMatrix(I7, C_NUM))
                If NumerodePersonal > 0 Then
                    Debug.Print "Empleado: "; NumerodePersonal
                    Debug.Print "Fila origen: "; I7

                    reng
                    MdAbr_1

                    ' Sobrescribir SOLO las 5 posiciones generales
                    ' para las NUEVAS Liquidaciones y Finiquitos.
                    ' El resto de las 296 posiciones se conserva igual.
                    If EsSeparacionNueva() Then
                        CompletarCamposGeneralesCFDI2026
                    End If
                End If
            End If
        End If
    Next I7

FinGeneracionCFDI:
    Close 7, 9

    ' =========================================================
    ' VALIDACION DE SALARIOS EN EL GRID CFDI (CODIGO EXISTENTE)
    ' =========================================================
    For fila = 1 To NomCf.Rows - 2
        If IsNumeric(NomCf.TextMatrix(fila, 56)) And _
           IsNumeric(NomCf.TextMatrix(fila, 58)) Then
            If CCur(NomCf.TextMatrix(fila, 56)) >= _
               CCur(NomCf.TextMatrix(fila, 58)) Then
                pintarRojo fila
            End If
        End If
    Next fila

    Exit Sub

ErrorGeneracionCFDI:
    Dim numeroError As Long
    Dim descripcionError As String
    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next
    Close 7, 9, 10
    On Error GoTo 0

    MsgBox "No se pudo completar la plantilla de CFDI." & vbCrLf & _
           "Error " & CStr(numeroError) & ": " & descripcionError, _
           vbCritical, "NOMCF2"
End Sub

' Este procedimiento solo se usa DESPUES de MdAbr_1 en formatos nuevos.
' Conserva el criterio ANTERIOR para las fechas y antiguedad:
'   Fecha inicial: dia 1 o 16; fecha final: la misma fecha de pago.
'   Antiguedad: diferencia entre anos calendario, como MdAbr_1 antigua.
Private Sub CompletarCamposGeneralesCFDI2026()
    Dim fechaPago As Date
    Dim fechaInicio As Date
    Dim fechaAlta As Date
    Dim anosAntiguedad As Long

    fechaPago = FechaCFDI_InterpretarDMY(g_FechaPagoCFDI2026)

    Get #2, NumerodePersonal, personal
    fechaAlta = FechaCFDI_InterpretarDMY(Trim$(personal.fal))

    If fechaAlta > fechaPago Then
        Err.Raise vbObjectError + 5211, "CompletarCamposGeneralesCFDI2026", _
                  "La fecha de alta del empleado " & CStr(NumerodePersonal) & _
                  " es posterior a la fecha de pago de la nomina."
    End If

    If Day(fechaPago) < 16 Then
        fechaInicio = DateSerial(Year(fechaPago), Month(fechaPago), 1)
    Else
        fechaInicio = DateSerial(Year(fechaPago), Month(fechaPago), 16)
    End If

    ' Fechas en ISO 8601 para las columnas 43, 44 y 45 del Excel.
    NomCf.TextMatrix(I7, 43) = FechaCFDI_FormatoISO(fechaPago)
    NomCf.TextMatrix(I7, 44) = FechaCFDI_FormatoISO(fechaInicio)
    NomCf.TextMatrix(I7, 45) = FechaCFDI_FormatoISO(fechaPago)

    ' Misma regla de antiguedad de la version ANTERIOR: diferencia de anos.
    anosAntiguedad = Year(fechaPago) - Year(fechaAlta)
    NomCf.TextMatrix(I7, 51) = CStr(anosAntiguedad)

    ' Dato REAL de personal.dno; nunca sustituir por sueldo diario.
    NomCf.TextMatrix(I7, 58) = Format$(personal.integrado, "#,##0.00")

    Debug.Print "CFDI 2026 empleado "; NumerodePersonal
    Debug.Print "  Fecha pago: "; NomCf.TextMatrix(I7, 43)
    Debug.Print "  Fecha inicio: "; NomCf.TextMatrix(I7, 44)
    Debug.Print "  Antiguedad (anos calendario): "; anosAntiguedad
    Debug.Print "  SDI: "; personal.integrado
End Sub

' comentario
