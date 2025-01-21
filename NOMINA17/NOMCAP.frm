VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form NOMCF2 
   Caption         =   "nomxf"
   ClientHeight    =   8940
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11760
   LinkTopic       =   "Form9"
   ScaleHeight     =   8940
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox LTXT1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   10815
   End
   Begin MSFlexGridLib.MSFlexGrid NOMCF 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6376
      _Version        =   393216
      FixedRows       =   0
      AllowUserResizing=   3
   End
   Begin VB.Menu NCFED 
      Caption         =   "&Edicion"
      Begin VB.Menu NCFEDCOP 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdSep1 
         Caption         =   "-"
      End
      Begin VB.Menu NCFEDPEG 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu NCFEDSEP1 
         Caption         =   "&CAPTION"
      End
      Begin VB.Menu NCFEDSEL 
         Caption         =   "&Selecionar todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu EdSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdQui 
         Caption         =   "&Quitar Titulos"
         Shortcut        =   ^Q
      End
      Begin VB.Menu EdFij 
         Caption         =   "&FijarTitulos"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu NCFEspr 
      Caption         =   "&Eliminar"
   End
   Begin VB.Menu Excel 
      Caption         =   "Generar Excel"
   End
End
Attribute VB_Name = "NOMCF2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c_p As Integer
Dim EntradaDATOS
Dim NumerodePersonal As Long
Dim Folio As Long, serie As String, Nombrey As String
    Dim Sub_total As Currency, Des_cto As Currency, Gravado As Currency, ConNomina As Currency
    Dim Ex_to As Currency, T_ingreso As Currency, T_desctos As Currency, T_isr As Currency, T_gravado As Currency
    Dim T_exento As Currency, T_imss As Currency, Reg_Patr As String
    Dim Met_pago As String, MiFechaAlta, RiesgoImss, VMiFechaAlta, VMiFecha, VMiFecha1, VFal, ViaTico As Currency
    Dim Met_pagcon As String
    Dim RFC_EXTERNO As String
    
    Dim prestamos As Double
    Dim fonacot As Double
    Dim pension As Double
    Dim infonavit As Double
    Dim impuesto As Double
    Dim imss As Double
    
    
Sub IniCols()
   Dim I_7 As Integer
    Open "DatCFDI.dno" For Random As 7 Len = Len(DATcf)
    cm = LOF(7) / Len(DATcf)
    'Close 7
   NOMCF.Clear: NOMCF.Rows = 2: NOMCF.Cols = cm: NOMCF.FixedCols = 0
   NOMCF.FixedRows = 1
   NOMCF.Rows = Form8.ConNom1.Rows
   NOMCF.Row = 0
   For I_7 = 0 To cm - 1: Get 7, (I_7 + 1), DATcf
      
   NOMCF.Col = I_7: NOMCF.CellAlignment = 4: NOMCF.ColWidth(I7) = 1200: NOMCF.Text = Trim(DATcf.Concepto)
        'NOMCF.Col = 1: NOMCF.CellAlignment = 4: NOMCF.ColWidth(1) = 3600: NOMCF.Text = "Nombre"
        'NOMCF.Col = 2: NOMCF.CellAlignment = 4: NOMCF.ColWidth(2) = 2400: NOMCF.Text = "Direccion"
        'NOMCF.Col = 3: NOMCF.CellAlignment = 4: NOMCF.ColWidth(3) = 2400: NOMCF.Text = "Colonia"
        'NOMCF.Col = 4: NOMCF.CellAlignment = 4: NOMCF.ColWidth(4) = 2400: NOMCF.Text = "Ciudad"
        'NOMCF.Col = 5: NOMCF.CellAlignment = 4: NOMCF.ColWidth(5) = 2400: NOMCF.Text = "Estado"
        'NOMCF.Col = 6: NOMCF.CellAlignment = 4: NOMCF.ColWidth(6) = 2400: NOMCF.Text = "Delegacion"
        'NOMCF.Col = 7: NOMCF.CellAlignment = 4: NOMCF.ColWidth(7) = 800: NOMCF.Text = "Codigo"
        'NOMCF.Col = 8: NOMCF.CellAlignment = 4: NOMCF.ColWidth(8) = 2400: NOMCF.Text = "Correo "
        'NOMCF.Col = 9: NOMCF.CellAlignment = 4: NOMCF.ColWidth(9) = 600: NOMCF.Text = "Cons."
  Next I_7
  Close 7
End Sub



Private Sub EdFij_Click()
  NOMCF.FixedRows = 1
End Sub

Private Sub EdQui_Click()
    NOMCF.FixedRows = 0
End Sub

Private Sub Excel_Click()
    If Exportar_Excel(App.Path & "\libro1.xls", NOMCF) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
End Sub

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
 On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(fila, Columna + 1).Value = .TextMatrix(fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Private Sub Form_Load()
    'Dim Folio As Long, serie As String, Nombrey As String
    'Dim Sub_total As Currency, Des_cto As Currency, Gravado As Currency, ConNomina As Currency
    'Dim Ex_to As Currency, T_ingreso As Currency, T_desctos As Currency, T_isr As Currency, T_gravado As Currency
    'Dim t_neto As Currency, T_exento As Currency, T_imss As Currency, Reg_Patr As String
    'Dim Met_pago As String, MiFechaAlta, RiesgoImss, VMiFechaAlta, VMiFecha, VMiFecha1, VFal, ViaTico As Currency
    'Dim Met_pagcon As String
    
    'Dim prestamos As Double
    'Dim fonacot As Double
    'Dim pension As Double
    'Dim infonavit As Double
    'Dim impuesto As Double
    'Dim imss As Double
    
    IniCols
    Close 10
    
    Open "EMP_CFDI.DNO" For Random As 10 Len = Len(EmpCFDI)
    Get 10, 1, EmpCFDI
    Open "Perscfdi.dno" For Random As 7 Len = Len(Empleado_1)
    Open "PerOtre.dno" For Random As 9 Len = Len(Otros_Rgtros)
    
    Folio = InputBox("Dame el numero de Folio Anterior", "CFDI NOMINA", EmpCFDI.Folio)
    serie = InputBox("Dame la serie", "CFDI NOMINA", EmpCFDI.serie)
    ConNomina = InputBox("Numero de nomina consecutivo ", "CFDI NOMINA", (EmpCFDI.Consecutivo + 1))
    
ERRORDEPAGO:
    Met_pago = InputBox("Metodo de Pago (utiliza la clave numerica)" & Chr(13) & "01" & Chr(9) & "EFECTIVO" & Chr(13) & "02" & Chr(9) & "CHEQUE NOMINATIVO" & Chr(13) & "03" & Chr(9) & "TRANSFERENCIA ELECTRONICA" & Chr(13) & "28" & Chr(9) & "TARJETA DE DEBITO", "CFDI NOMINA", "28")
    
    Select Case Met_pago
        Case "01", "02", "03", "28"
            If Met_pago = "01" Then Met_pagcon = "01 Eectivo"
            If Met_pago = "02" Then Met_pagcon = "02 Cheque nominativo"
            If Met_pago = "03" Then Met_pagcon = "03 TRANSFERENCIA"
            If Met_pago = "28" Then Met_pagcon = "28 Tarjeta de Débito"
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
    
    EmpCFDI.Folio = Folio: EmpCFDI.Consecutivo = ConNomina: EmpCFDI.serie = serie:  EmpCFDI.RegPatr = Reg_Patr: EmpCFDI.RiesgoImss = RiesgoImss
    
    
    pregunta = MsgBox("La empresa es OutSorcing?", vbYesNo)
    If pregunta = vbYes Then
        RFC_EXTERNO = InputBox("INTRODUCE EL RFC DE LA EMPRESA DE LA EMPRESA DONDE SE PRESTA EL SERVICIO")
    End If
    
    
    Put 10, 1, EmpCFDI
    'Dim i8 As Long
    Dim x8 As Long
    Dim x9 As Long
    I8 = 0
    
   
    For I7 = 1 To Form8.ConNom1.Rows - 3
    Folio = Folio + 1
    
             If IsNumeric(Form8.ConNom1.TextMatrix(I7, 0)) Then NumerodePersonal = Form8.ConNom1.TextMatrix(I7, 0)
             'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 20)) Then t_neto = Form8.ConNom1.TextMatrix(I7, 20) Else t_neto = 0
             'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 11)) Then T_ingreso = Form8.ConNom1.TextMatrix(I7, 11) Else T_ingreso = 0
             'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 10)) Then T_exento = Form8.ConNom1.TextMatrix(I7, 10) Else T_exento = 0
             'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 12)) Then T_isr = Form8.ConNom1.TextMatrix(I7, 12) Else T_isr = 0
             'If T_isr < 0 Then T_isr = 0
             'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 14)) Then T_imss = Form8.ConNom1.TextMatrix(I7, 14) Else T_imss = 0
  

    
'*****************************************************************************************************************************************
    reng
'*****************************************************************************************************************************************
 
        'x8 = I8 + 5
        'x9 = 0
        'For i9 = I8 To x8
            'I8 = I8 + 1
            'x9 = x9 + 1
            'If I8 = NOMCF.Rows - 1 Then NOMCF.Rows = NOMCF.Rows + 1
    
             
  
'***************************************************************************************************************************************
'MdAbr
MdAbr_1

GoTo SALTALO
'***************************************************************************************************************************************
             NOMCF.TextMatrix(I7, 0) = 1:                             '1 SUCURSAL
              NOMCF.TextMatrix(I7, 1) = Folio       '2 FOLIO
             NOMCF.TextMatrix(I7, 2) = serie:                         '3 SERIE
             Get 2, NumerodePersonal, personal:
             MiFechaAlta = Trim(personal.fal)
             Nombrey = Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2)
             NOMCF.TextMatrix(I7, 3) = Nombrey:                       '4 NOMBRE
             Get 7, NumerodePersonal, Empleado_1
             NOMCF.TextMatrix(I7, 4) = Trim(Empleado_1.Direccion):     '5 DIRECCION
             NOMCF.TextMatrix(I7, 5) = Trim(Empleado_1.Colonia):       '6 COLONIA
             NOMCF.TextMatrix(I7, 6) = Trim(Empleado_1.Ciudad):        '7 CIUDAD
             NOMCF.TextMatrix(I7, 7) = Trim(Empleado_1.Estado):        '8 ESTADO
             NOMCF.TextMatrix(I7, 8) = Trim(Empleado_1.Delegacion):    '9 DELEGACION
             NOMCF.TextMatrix(I7, 9) = Trim(Empleado_1.Cpostal):       '10 CP
             NOMCF.TextMatrix(I7, 10) = Trim(personal.rfc):            '11 RFC
             NOMCF.TextMatrix(I7, 11) = "MEXICO":                      '12 PAIS
             NOMCF.TextMatrix(I7, 12) = Trim(Empleado_1.correo):       '13 CORREO
             NOMCF.TextMatrix(I7, 13) = ""                             '14 OBSERVACIONES
             NOMCF.TextMatrix(I7, 14) = "PESOS"                        '15 MONEDA
             NOMCF.TextMatrix(I7, 15) = 1                              '16 TIPOCAMBIO
     
             reng
            
             NOMCF.TextMatrix(I8, 16) = Format(t_per - t_ded, "#,##0.00")                        '17 TOTAL
             NOMCF.TextMatrix(I8, 17) = Format(t_per, "#,##0.00")                     '18 SUBTOTAL
             
             If Form8.ConNom1.TextMatrix(I7, 13) <> "" Then
                NOMCF.TextMatrix(I8, 18) = Format(t_ded, "#,##0.00")
                Else
                NOMCF.TextMatrix(I8, 18) = Format(t_ded, "#,##0.00")            '19 DESCUENTO
            End If
             
             Des_cto = T_ingreso - Sub_total
             
             NOMCF.TextMatrix(I8, 19) = "Deducciones Nomina":            '20 MOTIVODESCUENTO
             Gravado = T_ingreso - T_exento:
             NOMCF.TextMatrix(I8, 20) = ConNomina:                       '21 PARTIDA
             NOMCF.TextMatrix(I8, 21) = "Pago de nómina":  '22 DESCRIPCION
             NOMCF.TextMatrix(I8, 22) = 1:                               '23 CANTIDAD
             NOMCF.TextMatrix(I8, 23) = "ACT":                      '24 UNIDAD
             NOMCF.TextMatrix(I8, 24) = Format(t_per, "#,##0.00")                    '25 VALORUNITARIO
             NOMCF.TextMatrix(I8, 25) = Format(t_per, "#,##0.00")                    '26 IMPORTE
             Aletra (t_per - t_ded)
             NOMCF.TextMatrix(I8, 26) = LetrVal                              '27 TOTALCONLETRA
             NOMCF.TextMatrix(I8, 27) = Met_pagcon:                        '28 METODOPAGO
             NOMCF.TextMatrix(I8, 28) = "DIF":                  '29 LUGAREXPEDICION
             NOMCF.TextMatrix(I8, 29) = "601":            '30 REGIMEN
             '---------------------------------------------------------------------------------------------TEXTO------------------------------
             NOMCF.TextMatrix(I8, 30) = CStr(Form8.ConNom1.TextMatrix(I7, 24))   '36 numctapag solo los ultimos 4 digitos
             '---------------------------------------------------------------------------------------------TEXTO------------------------------
             NOMCF.TextMatrix(I8, 31) = "O":                      '32 tipode nomina
             
            DDINI = Left(MiFecha, 2)
             If DDINI < 16 Then
                  DDINI = 1
                  MiFecha1 = "01" + Mid(MiFecha, 3)
                  Else
                  DDINI = 16
                  MiFecha1 = "16" + Mid(MiFecha, 3)
             End If
            VMiFecha = Mid(Trim(MiFecha), 7, 4) + "-" + Mid(Trim(MiFecha), 4, 2) + "-" + Mid(Trim(MiFecha), 1, 2)
            VMiFecha1 = Mid(Trim(MiFecha1), 7, 4) + "-" + Mid(Trim(MiFecha1), 4, 2) + "-" + Mid(Trim(MiFecha1), 1, 2)

             NOMCF.TextMatrix(I8, 32) = VMiFecha:                         '33 FECHAPAGO
                
             NOMCF.TextMatrix(I8, 33) = VMiFecha1:                         '34 FECHAINICIALPAGO
             NOMCF.TextMatrix(I8, 34) = VMiFecha:                          '45 FECHAFINALPAGO
             NOMCF.TextMatrix(I8, 35) = Form8.ConNom1.TextMatrix(I7, 2):   '36 NUMDIASPAGADOS
             
            NOMCF.TextMatrix(I8, 36) = Format(t_per, "#,##0.00")
             
              NOMCF.TextMatrix(I8, 37) = Format(t_ded, "#,##0.00")
            
            If sub13 > 0 Then
                NOMCF.TextMatrix(I8, 38) = Format(sub13, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 38) = 0
            End If
            
             'NOMCF.TextMatrix(i8, 38) = ""      'totalotrospagos
             NOMCF.TextMatrix(I8, 39) = ""               '40 CURP
             NOMCF.TextMatrix(I8, 40) = Reg_Patr:                        '41 REGISTROPATRONAL
             '---------------------------------------------------------------------------------------------------------------------------
                NOMCF.TextMatrix(I8, 41) = ""
                '-------------------------------------------------------------------------------------------------------------------------
                NOMCF.TextMatrix(I8, 42) = "":             '43 EMISOR_ORIGENRECURSO
                '-------------------------------------------------------------------------------------------------------------------------
                             If Form1.Label1.Caption = "Empresa : " + "SOLUCIONES ADMINISTRATIVAS Y DE ORGANIZACION, S.A. DE C.V." Then
                NOMCF.TextMatrix(I8, 43) = "":             '44 EMISOR_MONTORECURSOPROPIO
             End If
             If Form1.Label1.Caption = "Empresa : " + "NVA ADMINISTRACION, S.A. DE C.V." Then
                NOMCF.TextMatrix(I8, 43) = "":             '44 EMISOR_MONTORECURSOPROPIO
             End If
             If Form1.Label1.Caption = "Empresa : " + "SUP CONSULTORES, S.A. DE C.V." Then
                NOMCF.TextMatrix(I8, 43) = "":             '44 EMISOR_MONTORECURSOPROPIO
                
             Else
                NOMCF.TextMatrix(I8, 43) = "0":
             End If
                '-------------------------------------------------------------------------------------------------------------------------
            Get 9, NumerodePersonal, Otros_Rgtros
             NOMCF.TextMatrix(I8, 44) = Trim(Otros_Rgtros.curp):               '45 CURP
             NOMCF.TextMatrix(I8, 45) = Trim(personal.imss):             '46 NUMSEGURIDADSOCIAL
             
            VFal = Mid(Trim(personal.fal), 7, 4) + "-" + Mid(Trim(personal.fal), 4, 2) + "-" + Mid(Trim(personal.fal), 1, 2)
             NOMCF.TextMatrix(I8, 46) = VFal                      '47 FECHAINICIORELLABORAL
             
             NOMCF.TextMatrix(I8, 47) = Year(MiFecha) - Year(MiFechaAlta) '48 ANTIGUEDAD
             '---------------------------------------------------------------------------------------------------------------------------------
             NOMCF.TextMatrix(I8, 48) = "01 CONTRATO DE TRABAJO POR TIEMPO INDETERMINADO"                       '49 TIPOCONTRATO"
             NOMCF.TextMatrix(I8, 49) = "No"                      '50 RECEPTOR SINDICALIZADO
             NOMCF.TextMatrix(I8, 50) = "01 DIURNO"                       '54 TIPOJORNADA 01 DIURNA
             NOMCF.TextMatrix(I8, 51) = "02 SUELDOS"                               '41 TIPOREGIMEN 02 SUELDOS
             NOMCF.TextMatrix(I8, 52) = NumerodePersonal:                '39 NUMEMPLEADO
             NOMCF.TextMatrix(I8, 53) = "ADMINISTRACION"                         '47 DEPARTAMENTO
             NOMCF.TextMatrix(I8, 54) = "ADMINISTRACION"                       '52 PUESTO
             NOMCF.TextMatrix(I8, 55) = "1"                        '56 RIESGOPUESTO
                          
           If N_ormal = 1 Then
                    NOMCF.TextMatrix(I8, 56) = "Anual"
                    Else
                    NOMCF.TextMatrix(I8, 56) = "04 Quincenal":                      '57 PERIODICIDADPAGO
                
            End If
             
            NOMCF.TextMatrix(I8, 57) = "002":                       '58 BANCO
            'NOMCF.TextMatrix(i8, 58) = Form8.ConNom1.TextMatrix(I7, 24) '59 RECEPTOR_CUENTABANCARIA liz lo dejaba en blanco
            NOMCF.TextMatrix(I8, 58) = ""
            NOMCF.TextMatrix(I8, 59) = Format(personal.ingr, "#,##0.00")  '60 SALARIOBASECOTAPOR
            
            'NOMCF.TextMatrix(i8, 60) = Format(personal.integrado, "#,##0.00")               '61 SALARIODIARIOINTEGRADO
            NOMCF.TextMatrix(I8, 60) = Format(personal.ingr, "#,##0.00")  '60 SALARIOBASECOTAPOR
            NOMCF.TextMatrix(I8, 61) = "DIF":                  '62 LUGAREXPEDICION
            NOMCF.TextMatrix(I8, 62) = RFC_EXTERNO      '63 SUBCON_RFCLABORA EMPRESA CONTRATISTA
            If RFC_EXTERNO <> "" Then
                NOMCF.TextMatrix(I8, 63) = 100
                Else
                NOMCF.TextMatrix(I8, 63) = "0"           '64  SUBCON_PORCENTAJETIEMPO
            End If
            NOMCF.TextMatrix(I8, 64) = Form8.ConNom1.TextMatrix(I7, 11) '65 TOTALSUELDOS 11-10
            NOMCF.TextMatrix(I8, 65) = "0"   '66 TOTALSEPARACIONINDEMINIZACION
            NOMCF.TextMatrix(I8, 66) = "0"   '67 TOTALJUBILACIONPENSIONRETIRO
            
            If (x9 = 1) Then
            
                If Form8.ConNom1.TextMatrix(I7, 10) = "" Then
                    NOMCF.TextMatrix(I8, 67) = Form8.ConNom1.TextMatrix(I7, 11)
                Else
                    NOMCF.TextMatrix(I8, 67) = Format(CDbl(Form8.ConNom1.TextMatrix(I7, 11)) - CDbl(Form8.ConNom1.TextMatrix(I7, 10)), "#,##0.00")
                End If
                
            Else
                NOMCF.TextMatrix(I8, 67) = 0
            End If
            
                    Dim p_v6 As Currency
                    Dim p_v7 As Currency
                    Dim p_v8 As Currency
                    p_v8 = empresa.sm * 15
                    If IsNumeric(Form8.ConNom1.TextMatrix(I7, 8)) Then p_v6 = Form8.ConNom1.TextMatrix(I7, 8) Else p_v6 = 0
                    If IsNumeric(Form8.ConNom1.TextMatrix(I7, 10)) Then p_v7 = Form8.ConNom1.TextMatrix(I7, 10) Else p_v7 = 0
                    If p_v6 > 0 And p_v7 = 0 Then
                       If p_v6 < p_v8 Then
                           p_v7 = p_v6
                           p_v6 = 0
                         Else
                            p_v7 = p_v8
                            p_v6 = p_v6 - p_v8
                         End If
                    End If
                    
            '69 TOTALEXENTO
            If pee10 > 0 Then
              pee10 = pee10 + sub13
            Else
               pee10 = sub13
            End If

                NOMCF.TextMatrix(I8, 68) = 0

            '------------------------------------------------------------------------------------------------------------------------------
            'PERCEPCIONES
            '------------------------------------------------------------------------------------------------------------------------------
            
            Select Case x9
                Case 1
                If Form8.ConNom1.TextMatrix(I7, 3) = "" Then
                    NOMCF.TextMatrix(I8, 69) = "" 'TIPOPERCEPCION Sueldos, Salarios  Rayas y Jornales
                    NOMCF.TextMatrix(I8, 70) = ""
                    NOMCF.TextMatrix(I8, 71) = ""
                    NOMCF.TextMatrix(I8, 72) = ""
                    NOMCF.TextMatrix(I8, 73) = ""
                Else
                    NOMCF.TextMatrix(I8, 69) = "001" 'TIPOPERCEPCION Sueldos, Salarios  Rayas y Jornales
                    NOMCF.TextMatrix(I8, 70) = "001" 'ESTA TAMBIEN ES TEXTO
                    NOMCF.TextMatrix(I8, 71) = "Sueldos, Salarios  Rayas y Jornales"
                    NOMCF.TextMatrix(I8, 72) = Format(sue3, "#,##0.00")
                    NOMCF.TextMatrix(I8, 73) = ""
                End If
                
                Case 2
                If Form8.ConNom1.TextMatrix(I7, 8) = "" Then
                    NOMCF.TextMatrix(I8, 69) = ""
                    NOMCF.TextMatrix(I8, 70) = ""
                    NOMCF.TextMatrix(I8, 71) = ""
                    NOMCF.TextMatrix(I8, 72) = ""
                    NOMCF.TextMatrix(I8, 73) = ""

                Else
                    NOMCF.TextMatrix(I8, 69) = "021"
                    NOMCF.TextMatrix(I8, 70) = "021"
                    NOMCF.TextMatrix(I8, 71) = "Prima vacacional"
                    NOMCF.TextMatrix(I8, 72) = Format(pva8, "#,##0.00")
                    NOMCF.TextMatrix(I8, 73) = Format(pee10, "#,##0.00")
                End If
                
                Case 3
                    'Dim ox1 As Currency
                    'Dim ox2 As Currency
                    'Dim ox3 As Currency
                    'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 7)) Then ox1 = Form8.ConNom1.TextMatrix(I7, 7) Else ox1 = " "
                    'If IsNumeric(Form8.ConNom1.TextMatrix(I7, 9)) Then ox2 = Form8.ConNom1.TextMatrix(I7, 9) Else ox2 = " "
                If Form8.ConNom1.TextMatrix(I7, 9) = "" Then
                    NOMCF.TextMatrix(I8, 69) = ""
                    NOMCF.TextMatrix(I8, 70) = ""
                    NOMCF.TextMatrix(I8, 71) = ""
                    NOMCF.TextMatrix(I8, 72) = ""
                    NOMCF.TextMatrix(I8, 73) = ""
                Else
                    NOMCF.TextMatrix(I8, 69) = "038"
                    NOMCF.TextMatrix(I8, 70) = "038"
                    NOMCF.TextMatrix(I8, 71) = "Otros ingresos por salarios"
                    NOMCF.TextMatrix(I8, 72) = Format(t_oi, "#,##0.00")
                If sub13 > 0 Then
                    NOMCF.TextMatrix(I8, 72) = Format((t_oi + 0.01), "#,##0.00")
                    NOMCF.TextMatrix(I8, 73) = Format((sub13 - 0.01), "#,##0.00")
                    
                Else
                    NOMCF.TextMatrix(I8, 73) = Format(pee10, "#,##0.00")
                End If
                End If
                
            End Select
                
            '---------------------------------------------------------------------------------------------------------------------------------
            NOMCF.TextMatrix(I8, 74) = 0
            NOMCF.TextMatrix(I8, 75) = 0
            NOMCF.TextMatrix(I8, 76) = 0
            NOMCF.TextMatrix(I8, 77) = ""
            NOMCF.TextMatrix(I8, 78) = 0
            NOMCF.TextMatrix(I8, 79) = 0
            NOMCF.TextMatrix(I8, 80) = 0
            NOMCF.TextMatrix(I8, 81) = 0
            NOMCF.TextMatrix(I8, 82) = 0
            NOMCF.TextMatrix(I8, 83) = 0
            NOMCF.TextMatrix(I8, 84) = 0
            NOMCF.TextMatrix(I8, 85) = 0
            NOMCF.TextMatrix(I8, 86) = 0
            NOMCF.TextMatrix(I8, 87) = 0
            NOMCF.TextMatrix(I8, 88) = 0
            NOMCF.TextMatrix(I8, 89) = 0
'----------------------------------------------------------------------------------------------------------------------------------
            
            If Form8.ConNom1.TextMatrix(I7, 15) = "" Then
                prestamos = 0
            Else
                prestamos = Form8.ConNom1.TextMatrix(I7, 15)
            End If
 '--------------------------------------------------------------------------------------------------------------------------------
            If Form8.ConNom1.TextMatrix(I7, 16) = "" Then
              fonacot = 0
            Else
                fonacot = Form8.ConNom1.TextMatrix(I7, 16)
             End If
 '--------------------------------------------------------------------------------------------------------------------------------
            If Form8.ConNom1.TextMatrix(I7, 17) = "" Then
                pension = 0
            Else
                pension = Form8.ConNom1.TextMatrix(I7, 17)
            End If
 '---------------------------------------------------------------------------------------------------------------------------------
            If Form8.ConNom1.TextMatrix(I7, 18) = "" Then
                infonavit = 0
            Else
               infonavit = Form8.ConNom1.TextMatrix(I7, 18)
            End If
'----------------------------------------------------------------------------------------------------------------------------------
            If Form8.ConNom1.TextMatrix(I7, 12) = "" Then
                impuesto = 0
            Else
                impuesto = Form8.ConNom1.TextMatrix(I7, 12)
            End If
 '--------------------------------------------------------------------------------------------------------------------------------
            If Form8.ConNom1.TextMatrix(I7, 14) = "" Then
                imss = 0
            Else
               imss = Form8.ConNom1.TextMatrix(I7, 14)
            End If

            'Format(pre15 + fon16 + pea17 + ifv18 + ims14+ isr12, "#,##0.00")
            NOMCF.TextMatrix(I8, 90) = Format(prestamos + fonacot + pension + infonavit + imss, "#,##0.00")
            NOMCF.TextMatrix(I8, 91) = Format(isr12, "#,##0.00")
            

            '----------------------------------------------------------------------------------------------------------------------------------
            'DEDUCCIONES
            '----------------------------------------------------------------------------------------------------------------------------------
    Select Case x9
         Case 1
            NOMCF.TextMatrix(I8, 92) = "001"
            NOMCF.TextMatrix(I8, 93) = "001"
            NOMCF.TextMatrix(I8, 94) = "Seguridad social"
            If ims14 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(ims14, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
        Case 2
            NOMCF.TextMatrix(I8, 92) = "002"
            NOMCF.TextMatrix(I8, 93) = "002"
            NOMCF.TextMatrix(I8, 94) = "ISR"
            If isr12 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(isr12, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
            
        Case 3
        If pre15 > 0 Then
            NOMCF.TextMatrix(I8, 92) = "004"
            NOMCF.TextMatrix(I8, 93) = "004"
            NOMCF.TextMatrix(I8, 94) = "OTROS"
            If pre15 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(pre15, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
        Else
            NOMCF.TextMatrix(I8, 92) = ""
            NOMCF.TextMatrix(I8, 93) = ""
            NOMCF.TextMatrix(I8, 94) = ""
            NOMCF.TextMatrix(I8, 95) = ""
        End If
            
                

        Case 4
            NOMCF.TextMatrix(I8, 92) = "007"
            NOMCF.TextMatrix(I8, 93) = "007"
            NOMCF.TextMatrix(I8, 94) = "PENSION ALIMENTICIA"
            If pea17 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(pea17, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
        
        Case 5
            NOMCF.TextMatrix(I8, 92) = "010"
            NOMCF.TextMatrix(I8, 93) = "010"
            NOMCF.TextMatrix(I8, 94) = "Pago por crédito de vivienda"
            If ifv18 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(ifv18, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
            
        Case 6
            NOMCF.TextMatrix(I8, 92) = "011"
            NOMCF.TextMatrix(I8, 93) = "011"
            NOMCF.TextMatrix(I8, 94) = "Pago de abonos INFONACOT"
            If fon16 > 0 Then
                NOMCF.TextMatrix(I8, 95) = Format(fon16, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 95) = ""
            End If
        End Select
            '-----------------------------------------------------------------------------------------------------------
    Select Case x9
        Case 1
            Get 14, Form8.ConNom1.TextMatrix(I7, 0), nom_com
            'subc13 = nom_com.subdio
            If nom_com.subdio <> 0 Then
                NOMCF.TextMatrix(I8, 96) = "002"
                NOMCF.TextMatrix(I8, 97) = "002"
                NOMCF.TextMatrix(I8, 98) = "Subsidio para el empleo"
                'Get 14, Form8.connom1.TextMatrix(I7, 0), nom_com
                'Debug.Print nom_com.CredNe; nom_com.subdio
                NOMCF.TextMatrix(I8, 99) = Format(sub13, "#,##0.00")
                NOMCF.TextMatrix(I8, 100) = Format(nom_com.subdio, "#,##0.00")
            Else
                NOMCF.TextMatrix(I8, 96) = ""
                NOMCF.TextMatrix(I8, 97) = ""
                NOMCF.TextMatrix(I8, 98) = ""
                NOMCF.TextMatrix(I8, 99) = ""
                NOMCF.TextMatrix(I8, 100) = ""
            End If
        Case Else
    End Select
            '-----------------------------------------------------------------------------------------------------------
             NOMCF.TextMatrix(I8, 106) = "0"
            NOMCF.TextMatrix(I8, 107) = "0"
        
         
        'Next i9
'******************************************************************************************************************************************
SALTALO:
'******************************************************************************************************************************************
    Next I7
    Get 10, 1, EmpCFDI: EmpCFDI.Folio = Folio: Put 10, 1, EmpCFDI
    Close 7, 9, 10
End Sub

Sub cont_r()
'contador de renglones
 
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then
   NOMCF.Height = ScaleHeight - 1000
   NOMCF.Width = ScaleWidth - 600
 End If
End Sub

Private Sub NCfEdCop_Click()
      Dim Temporal1
 Clipboard.Clear
   
   difer = NOMCF.RowSel - NOMCF.Row
   For i = NOMCF.Row To NOMCF.RowSel
      
      For f = NOMCF.Col To NOMCF.ColSel
            Temporal1 = Temporal1 + NOMCF.TextMatrix(i, f)
            If f < NOMCF.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
      Next f
      Temporal1 = Temporal1 & Chr(13) & Chr(10)
      
   Next i
    Clipboard.SetText Temporal1
End Sub
Private Sub LTXT1_Change()
    NOMCF.Text = LTXT1.Text
End Sub


Private Sub NCfEdSel_Click()
     Dim limite As Long
    Clipboard.Clear
    NOMCF.Row = 1: NOMCF.Col = 0
   For limite = 1 To NOMCF.Rows - 1
       renglon = limite
    If IsNumeric(NOMCF.TextMatrix(renglon, 0)) Then
           NOMCF.RowSel = renglon
    End If
   Next limite
    NOMCF.ColSel = NOMCF.Cols - 1
End Sub


Private Sub NCFEspr_Click()
Dim conteocolumnas  As Integer
Dim totaldecolumnas As Integer
    totaldecolumnas = NOMCF.Rows - 1
    conteocolumnas = 1
    While conteocolumnas < totaldecolumnas
        If (NOMCF.TextMatrix(conteocolumnas, 95) = "") And (NOMCF.TextMatrix(conteocolumnas, 70) = "") Then
                    NOMCF.RowSel = conteocolumnas
                    NOMCF.RemoveItem conteocolumnas
                    conteocolumnas = 0
                    totaldecolumnas = (totaldecolumnas - 1)
        Else
                    conteocolumnas = (conteocolumnas + 1)
        End If
    Wend
End Sub


Private Sub NOMCF_EnterCell()
  LTXT1.Text = NOMCF.Text
End Sub

Private Sub NomCf_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyDelete
               
               For Q = NOMCF.Row To NOMCF.RowSel
                 
                 For W = NOMCF.Col To NOMCF.ColSel
                    
                    NOMCF.TextMatrix(Q, W) = ""
                 Next W
               Next Q
               
                LTXT1.Text = NOMCF.Text
            Case vbKeyF2
                If NOMCF.Text <> "" Then valcelant = NOMCF.Text
                LTXT1.Text = LTrim(RTrim(NOMCF.Text))
                LTXT1.SetFocus
               
       End Select

End Sub

Private Sub NomCf_KeyPress(KeyAscii As Integer)
    valcelant = NOMCF.Text
    LTXT1.Text = Chr(KeyAscii)
    LTXT1.SetFocus
End Sub

Private Sub NomCf_LeaveCell()
  If NOMCF.Row > 0 Then
   NOMCF.CellBackColor = vbWhite
  End If
End Sub

Sub MdAbr_1()

    '0 1 SUCURSAL
        NOMCF2.NOMCF.TextMatrix(I7, 0) = 1:
    '1 2 FOLIO
        NOMCF2.NOMCF.TextMatrix(I7, 1) = Folio
    '2 3 SERIE
        NOMCF2.NOMCF.TextMatrix(I7, 2) = serie:
    '3 4 NOMBRE
        Get 2, NumerodePersonal, personal:
        MiFechaAlta = Trim(personal.fal)
        Nombrey = Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2)
        NOMCF2.NOMCF.TextMatrix(I7, 3) = Nombrey:
        Get 7, NumerodePersonal, Empleado_1
    '4 5 DIRECCION
        NOMCF2.NOMCF.TextMatrix(I7, 4) = Trim(Empleado_1.Direccion):
    '5 6 COLONIA
        NOMCF2.NOMCF.TextMatrix(I7, 5) = Trim(Empleado_1.Colonia):
    '6 7 CIUDAD
        NOMCF2.NOMCF.TextMatrix(I7, 6) = Trim(Empleado_1.Ciudad):
    '7 8 ESTADO
        NOMCF2.NOMCF.TextMatrix(I7, 7) = Trim(Empleado_1.Estado):
    '8 9 DELEGACION
        NOMCF2.NOMCF.TextMatrix(I7, 8) = Trim(Empleado_1.Delegacion):
    '9 10 CP
        NOMCF2.NOMCF.TextMatrix(I7, 9) = Trim(Empleado_1.Cpostal):
    '10 11 RFC
        NOMCF2.NOMCF.TextMatrix(I7, 10) = Trim(personal.rfc):
    '11 12 PAIS
        NOMCF2.NOMCF.TextMatrix(I7, 11) = "MEXICO":
    '12 13 CORREO
        NOMCF2.NOMCF.TextMatrix(I7, 12) = Trim(Empleado_1.correo):
    '13 14 OBSERVACIONES
        NOMCF2.NOMCF.TextMatrix(I7, 13) = ""
    '14 15 MONEDA
        NOMCF2.NOMCF.TextMatrix(I7, 14) = "PESOS"
    '15 16 TIPOCAMBIO
        NOMCF2.NOMCF.TextMatrix(I7, 15) = 1
'******************************************LLAMADO AL MODULO ***********************************************
    '16 17 TOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 16) = Format(T_neto, "#,##0.00")
    '17 18 SUBTOTAL
        NOMCF2.NOMCF.TextMatrix(I7, 17) = Format(t_per, "#,##0.00")
    '18 19 DESCUENTO
        NOMCF2.NOMCF.TextMatrix(I7, 18) = Format(t_ded, "#,##0.00")
    '19 20 TOTALGRAVADOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 19) = Format(t_grav, "#,##0.00")
    '20 21 TOTALEXENTOPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 20) = Format(t_ext, "#,##0.00")
    '21 22 TOTALPERCEPCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 21) = Format(t_per - sub13, "#,##0.00")
    '22 23 TOTALDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 22) = Format(t_ded, "#,##0.00")
    '23 24 TOTALOTROSPAGOS
        NOMCF2.NOMCF.TextMatrix(I7, 23) = Format(sub13, "#,##0.00")
    
    '24 25 TOTALSUELDOS
     NOMCF2.NOMCF.TextMatrix(I7, 24) = Format(t_per - sub13, "#,##0.00")
    
    '25 26 TOTALSEPARACIONINDEMNIZACION
        NOMCF2.NOMCF.TextMatrix(I7, 25) = Format(0, "#,##0.00")
    '26 27 TOTALJUBILACIONPENSIONRETIRO
        NOMCF2.NOMCF.TextMatrix(I7, 26) = Format(0, "#,##0.00")
    '27 28 TOTALOTRASDEDUCCIONES
        NOMCF2.NOMCF.TextMatrix(I7, 27) = Format(t_oded, "#,##0.00")
    '28 29 TOTALIMPUESTOSRETENIDOS
        NOMCF2.NOMCF.TextMatrix(I7, 28) = Format(isr12, "#,##0.00")
    '29 30 VALORUNITARIO
        NOMCF2.NOMCF.TextMatrix(I7, 29) = Format(t_per, "#,##0.00")
    '30 31 IMPORTE
        NOMCF2.NOMCF.TextMatrix(I7, 30) = Format(t_per, "#,##0.00")
    Aletra (T_neto)
    '31 32 TOTALCONLETRA
        NOMCF2.NOMCF.TextMatrix(I7, 31) = LetrVal
    '32 33 TIPODENOMINA
      If N_ormal = 0 Then
        NOMCF2.NOMCF.TextMatrix(I7, 32) = "ORDINARIA"
        Else
        NOMCF2.NOMCF.TextMatrix(I7, 32) = "EXTRAORDINARIA"
      End If
    '33 34 SINDICALIZADO
        NOMCF2.NOMCF.TextMatrix(I7, 33) = "No"
    '34 35 METODOPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 34) = Met_pagcon
    '35 36 LUGAREXPEDICION
        NOMCF2.NOMCF.TextMatrix(I7, 35) = "DIF"
    '36 37 REGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 36) = "601"
    '37 38 NUMCTAPAG
        'NOMCF2.NOMCF.TextMatrix(I7, 37) = CStr(Form8.ConNom1.TextMatrix(I7, 24))
        NOMCF2.NOMCF.TextMatrix(I7, 37) = Right(Trim(Form8.ConNom1.TextMatrix(I7, 24)), 4)
    '38 39 REGISTROPATRONAL
        NOMCF2.NOMCF.TextMatrix(I7, 38) = Reg_Patr
    '39 40 NUMEMPLEADO
        NOMCF2.NOMCF.TextMatrix(I7, 39) = NumerodePersonal
    '40 41 CURP
        Get 9, NumerodePersonal, Otros_Rgtros
        NOMCF2.NOMCF.TextMatrix(I7, 40) = Trim(Otros_Rgtros.curp)
    '41 42 TIPOREGIMEN
        NOMCF2.NOMCF.TextMatrix(I7, 41) = "02 Sueldos"
    '42 43 NUMSEGURIDADSOCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 42) = Trim(personal.imss)
        
        DDINI = Left(MiFecha, 2)
             If DDINI < 16 Then
                  DDINI = 1
                  MiFecha1 = "01" + Mid(MiFecha, 3)
                  Else
                  DDINI = 16
                  MiFecha1 = "16" + Mid(MiFecha, 3)
             End If
        VMiFecha = Mid(Trim(MiFecha), 7, 4) + "-" + Mid(Trim(MiFecha), 4, 2) + "-" + Mid(Trim(MiFecha), 1, 2)
        VMiFecha1 = Mid(Trim(MiFecha1), 7, 4) + "-" + Mid(Trim(MiFecha1), 4, 2) + "-" + Mid(Trim(MiFecha1), 1, 2)
        
    '43 44 FECHAPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 43) = VMiFecha
    '44 45 FECHAINICIALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 44) = VMiFecha1
    '45 46 FECHAFINALPAGO
        NOMCF2.NOMCF.TextMatrix(I7, 45) = VMiFecha
    Rem *****************   MODIFICADO CON PARCHE DEL 6/6/17
    '46 47 NUMDIASPAGADOS
       If N_ormal = 1 Then
        NOMCF2.NOMCF.TextMatrix(I7, 46) = 1
        Else
        NOMCF2.NOMCF.TextMatrix(I7, 46) = Form8.ConNom1.TextMatrix(I7, 2)
       End If
    '47 48 DEPARTAMENTO
        NOMCF2.NOMCF.TextMatrix(I7, 47) = "ADMINISTRACION"
    '48 49 CLABE
        NOMCF2.NOMCF.TextMatrix(I7, 48) = ""
    '49 50 BANCO
        NOMCF2.NOMCF.TextMatrix(I7, 49) = "002"
    '50 51 FECHAINICIORELLABORAL
        VFal = Mid(Trim(personal.fal), 7, 4) + "-" + Mid(Trim(personal.fal), 4, 2) + "-" + Mid(Trim(personal.fal), 1, 2)
        NOMCF2.NOMCF.TextMatrix(I7, 50) = VFal
    '51 52 ANTIGUEDAD
        NOMCF2.NOMCF.TextMatrix(I7, 51) = Year(MiFecha) - Year(MiFechaAlta)
    '52 53 PUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 52) = "Administracion"
    '53 54 TIPOCONTRATO
        NOMCF2.NOMCF.TextMatrix(I7, 53) = "01 Contrato de trabajo por tiempo indeterminado"
    '54 55 TIPOJORNADA
        NOMCF2.NOMCF.TextMatrix(I7, 54) = "01 Diurna"
    '55 56 PERIODICIDADPAGO
        If N_ormal = 1 Then
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "99 OTRA PERIODICIDAD"
                    Else
                    NOMCF2.NOMCF.TextMatrix(I7, 55) = "04 QUINCENAL":
        End If
    '56 57 SALARIOBASECOTAPOR
        NOMCF2.NOMCF.TextMatrix(I7, 56) = Format(personal.ingr, "#,##0.00")
    '57 58 RIESGOPUESTO
        NOMCF2.NOMCF.TextMatrix(I7, 57) = RiesgoImss
    '58 59 SALARIODIARIOINTEGRADO
        NOMCF2.NOMCF.TextMatrix(I7, 58) = Format(personal.ingr, "#,##0.00")
    '59 60 ENTIDADFEDERATIVA
        NOMCF2.NOMCF.TextMatrix(I7, 59) = "DIF"
    '60 61 RFCLABORA
    '61 62 PORCENTAJETIEMPO
        If RFC_EXTERNO <> "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 60) = RFC_EXTERNO
            NOMCF2.NOMCF.TextMatrix(I7, 61) = 100
        Else
            NOMCF2.NOMCF.TextMatrix(I7, 61) = Format(0, "#,##0.00")
        End If
'***********************************************************PERCEPCIONES***************************************
    '62 63 P001
        NOMCF2.NOMCF.TextMatrix(I7, 62) = Format(sue3, "#,##0.00")
    '63 64 P019HED
        NOMCF2.NOMCF.TextMatrix(I7, 63) = Format(0, "#,##0.00")
    '64 65 P019HET
        NOMCF2.NOMCF.TextMatrix(I7, 64) = Format(0, "#,##0.00")
    '65 66 P019HS
        NOMCF2.NOMCF.TextMatrix(I7, 65) = Format(0, "#,##0.00")
    '66 67 P020
        NOMCF2.NOMCF.TextMatrix(I7, 66) = Format(0, "#,##0.00")
    '67 68 P021G
        NOMCF2.NOMCF.TextMatrix(I7, 67) = Format(pva8, "#,##0.00")
    '68 69 P021E
    If N_ormal = 0 Then
        NOMCF2.NOMCF.TextMatrix(I7, 68) = Format(pee10, "#,##0.00")
        Else
        NOMCF2.NOMCF.TextMatrix(I7, 68) = Format(0, "#,##0.00")
    End If
    '69 70 P002G AGUINALDO
    '70 71 P002E AGUINALDO
    If N_ormal = 1 Then
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 5)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(Form8.ConNom1.TextMatrix(I7, 5), "###0.00"):  Rem 64 P002 GGratificación Anual (Aguinaldo)
           If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
             NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(0, "###0.00")
            Else
            NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(Form8.ConNom1.TextMatrix(I7, 10), "###0.00"): Rem 65 P002 GGratificación Anual (Aguinaldo exento)
           End If
        ElseIf Form8.ConNom1.TextMatrix(I7, 5) = "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(0, "#,##0.00")
        End If
    Else
            NOMCF2.NOMCF.TextMatrix(I7, 69) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 70) = Format(0, "#,##0.00")
    End If
   
    '71 72 P003G PTU
    '72 73 P003E PTU
    Rem **************** MODIFICADO CON EL PARCHE DEL 6/6/17 *************************************
    If N_ormal = 1 Then
         
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 6)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(ptu_1, "###0.00"):          Rem 66 P003G PTU
            NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(ptu_2, "###0.00"):          Rem 67 P003G PTU
        ElseIf Form8.ConNom1.TextMatrix(I7, 6) = "" Then
            NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(0, "#,##0.00")
        End If
    Else
            NOMCF2.NOMCF.TextMatrix(I7, 71) = Format(0, "#,##0.00")
            NOMCF2.NOMCF.TextMatrix(I7, 72) = Format(0, "#,##0.00")
    End If
    
    '73 74 P004GTOS.MEDICOS
        NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(0, "#,##0.00")
    '74 75 P005 FONDO DE AHORRO
        NOMCF2.NOMCF.TextMatrix(I7, 74) = Format(0, "#,##0.00")
    '76 P006 CAJA DE AHORRO
        NOMCF2.NOMCF.TextMatrix(I7, 75) = Format(0, "#,##0.00")
    '77 P009 CONT.A CARGO DEL TRAB.
        NOMCF2.NOMCF.TextMatrix(I7, 76) = Format(0, "#,##0.00")
    '78 P010 PREMIO DE PUNTUALIDAD
    
    If N_ormal = 1 Then
        
        If IsNumeric(Form8.ConNom1.TextMatrix(I7, 7)) Then
            NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(Form8.ConNom1.TextMatrix(I7, 7) + Form8.ConNom1.TextMatrix(I7, 10), "###0.00"):           Rem 74 P003GPREMIO DE PUNTUALIDAD
        End If
        
    Else
        NOMCF2.NOMCF.TextMatrix(I7, 73) = Format(0, "#,##0.00")
    End If
    
    '78 79 P011 SEGURO DE VIDA
        NOMCF2.NOMCF.TextMatrix(I7, 78) = Format(0, "#,##0.00")
    '80 P012 Seguro de Gastos Medicos Mayores
        NOMCF2.NOMCF.TextMatrix(I7, 79) = Format(0, "#,##0.00")
    '81 P013 Cuotas Sindicales Pagadas por el Patrón
        NOMCF2.NOMCF.TextMatrix(I7, 80) = Format(0, "#,##0.00")
    '82 P014I Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 81) = Format(0, "#,##0.00")
    '83 P014R Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 82) = Format(0, "#,##0.00")
    '84 P014M Subsidios por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 83) = Format(0, "#,##0.00")
    '85 P015 Becas para trabajadores y/o hijos
        NOMCF2.NOMCF.TextMatrix(I7, 84) = Format(0, "#,##0.00")
    '86 P022G Prima por antigüedad
        NOMCF2.NOMCF.TextMatrix(I7, 85) = Format(0, "#,##0.00")
    '87 P022E Prima por antigüedad
        NOMCF2.NOMCF.TextMatrix(I7, 86) = Format(0, "#,##0.00")
    '88 P023G Pagos por separación
        NOMCF2.NOMCF.TextMatrix(I7, 87) = Format(0, "#,##0.00")
    '89 P023E Pagos por separación
        NOMCF2.NOMCF.TextMatrix(I7, 88) = Format(0, "#,##0.00")
    '90 P024G Seguro de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 89) = Format(0, "#,##0.00")
    '91 P024E Seguro de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 90) = Format(0, "#,##0.00")
    '92 P025G Indeminizaciones
        NOMCF2.NOMCF.TextMatrix(I7, 91) = Format(0, "#,##0.00")
    '93 P025E Indeminizaciones
        NOMCF2.NOMCF.TextMatrix(I7, 92) = Format(0, "#,##0.00")
    '94 P026 Reembolso por funeral
        NOMCF2.NOMCF.TextMatrix(I7, 93) = Format(0, "#,##0.00")
    '95 P027 Cuotas de seguridad social pagadas por el patrón
        NOMCF2.NOMCF.TextMatrix(I7, 94) = Format(0, "#,##0.00")
    '96 P028 Comisiones
        NOMCF2.NOMCF.TextMatrix(I7, 95) = Format(0, "#,##0.00")
    '97 P029G Vales de despensa
        NOMCF2.NOMCF.TextMatrix(I7, 96) = Format(0, "#,##0.00")
    '98 P029E Vales de despensa
        NOMCF2.NOMCF.TextMatrix(I7, 97) = Format(0, "#,##0.00")
    '99 P030 Vales de restaurante
        NOMCF2.NOMCF.TextMatrix(I7, 98) = Format(0, "#,##0.00")
    '100 P031 Vales de gasolina
        NOMCF2.NOMCF.TextMatrix(I7, 99) = Format(0, "#,##0.00")
    '101 P032 Vales de ropa
        NOMCF2.NOMCF.TextMatrix(I7, 100) = Format(0, "#,##0.00")
    '102 P033 Ayuda para renta
        NOMCF2.NOMCF.TextMatrix(I7, 101) = Format(0, "#,##0.00")
    '103 P034 Ayuda para artículos escolares
        NOMCF2.NOMCF.TextMatrix(I7, 102) = Format(0, "#,##0.00")
    '104 P035 Ayuda para anteojos
        NOMCF2.NOMCF.TextMatrix(I7, 103) = Format(0, "#,##0.00")
    '105 P036 Ayuda para transporte
        NOMCF2.NOMCF.TextMatrix(I7, 104) = Format(0, "#,##0.00")
    '106 P037 Ayuda para gastos de funeral
        NOMCF2.NOMCF.TextMatrix(I7, 105) = Format(0, "#,##0.00")
    '107 P038 Otros ingresos por salarios
        NOMCF2.NOMCF.TextMatrix(I7, 106) = Format(t_oi, "#,##0.00")
    '108 P039 Jubilaciones, pensiones o haberes de retiro
        NOMCF2.NOMCF.TextMatrix(I7, 107) = Format(0, "#,##0.00")
    '109 P044 Jubilaciones, pensiones o haberes de retiro en parcialidades
        NOMCF2.NOMCF.TextMatrix(I7, 108) = Format(0, "#,##0.00")
    '110 P045 Ingresos en acciones o títulos valor que representan bienes
        NOMCF2.NOMCF.TextMatrix(I7, 109) = Format(0, "#,##0.00")
    '111 P046 Ingresos asimilados a salarios
        NOMCF2.NOMCF.TextMatrix(I7, 110) = Format(0, "#,##0.00")
    '112 P047 Alimentación
        NOMCF2.NOMCF.TextMatrix(I7, 111) = Format(0, "#,##0.00")
    '113 P048 Habitación
        NOMCF2.NOMCF.TextMatrix(I7, 112) = Format(0, "#,##0.00")
    '114 P049 Premios por asistencia
        NOMCF2.NOMCF.TextMatrix(I7, 113) = Format(0, "#,##0.00")
    '115 P050 Viáticos
        NOMCF2.NOMCF.TextMatrix(I7, 114) = Format(0, "#,##0.00")

'**************************************************DEDUCCIONES******************************************
    '116 D001 Seguridad social
         NOMCF2.NOMCF.TextMatrix(I7, 115) = Format(ims14, "#,##0.00")
    '117 D002 ISR
        NOMCF2.NOMCF.TextMatrix(I7, 116) = Format(isr12, "#,##0.00")
    '118 D003 Aportaciones a retiro, cesantía en edad avanzada y vejez
        NOMCF2.NOMCF.TextMatrix(I7, 117) = Format(0, "#,##0.00")
    '119 D004 Otros
        NOMCF2.NOMCF.TextMatrix(I7, 118) = Format(pre15, "#,##0.00")
    '120 D005 Aportaciones a Fondo de vivienda
        NOMCF2.NOMCF.TextMatrix(I7, 119) = Format(0, "#,##0.00")
    '121 D006M Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 120) = Format(0, "#,##0.00")
    '122 D006I Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 121) = Format(0, "#,##0.00")
    '123 D006R Descuento por incapacidad
        NOMCF2.NOMCF.TextMatrix(I7, 122) = Format(0, "#,##0.00")
    '124 D007 Pensión alimenticia
        NOMCF2.NOMCF.TextMatrix(I7, 123) = Format(pea17, "#,##0.00")
    '125 D008 Renta
        NOMCF2.NOMCF.TextMatrix(I7, 124) = Format(0, "#,##0.00")
    '126 D009 Préstamos provenientes del Fondo Nacional de la vivienda para los trab
        NOMCF2.NOMCF.TextMatrix(I7, 125) = Format(0, "#,##0.00")
    '127 D010 Pago por crédito de vivienda
        NOMCF2.NOMCF.TextMatrix(I7, 126) = Format(ifv18, "#,##0.00")
    '128 D011 Pago de abonos INFONACOT
        NOMCF2.NOMCF.TextMatrix(I7, 127) = Format(fon16, "#,##0.00")
    '129 D012 Anticipo de salarios
        NOMCF2.NOMCF.TextMatrix(I7, 128) = Format(0, "#,##0.00")
    '130 D013 Pagos hechos con exceso al trabajador
        NOMCF2.NOMCF.TextMatrix(I7, 129) = Format(0, "#,##0.00")
    '131 D014 Errores
        NOMCF2.NOMCF.TextMatrix(I7, 130) = Format(0, "#,##0.00")
    '132 D015 Pérdidas
        NOMCF2.NOMCF.TextMatrix(I7, 131) = Format(0, "#,##0.00")
    '133 D016 Averías
        NOMCF2.NOMCF.TextMatrix(I7, 132) = Format(0, "#,##0.00")
    '134 D017 Adquisición de artículos producidos por la empresa o establecimiento
        NOMCF2.NOMCF.TextMatrix(I7, 133) = Format(0, "#,##0.00")
    '135 D018 Cuotas para la constitución y fomento de sociedades cooperativas y de
        NOMCF2.NOMCF.TextMatrix(I7, 134) = Format(0, "#,##0.00")
    '136 D019 Cuotas sindicales
        NOMCF2.NOMCF.TextMatrix(I7, 135) = Format(0, "#,##0.00")
    '137 D020 Ausencia (Ausentismo)
        NOMCF2.NOMCF.TextMatrix(I7, 136) = Format(0, "#,##0.00")
    '138 D021 Cuotas obrero patronales
        NOMCF2.NOMCF.TextMatrix(I7, 137) = Format(0, "#,##0.00")
    '139 D022 Impuestos Locales
        NOMCF2.NOMCF.TextMatrix(I7, 138) = Format(0, "#,##0.00")
    '140 D023 Aportaciones voluntarias
        NOMCF2.NOMCF.TextMatrix(I7, 139) = Format(0, "#,##0.00")
    '141 D024 Ajuste en Gratificación Anual (Aguinaldo) Exento
        NOMCF2.NOMCF.TextMatrix(I7, 140) = Format(0, "#,##0.00")
    '142 D025 Ajuste en Gratificación Anual (Aguinaldo) Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 141) = Format(0, "#,##0.00")
    '143 D026 Ajuste en Participación de los Trabajadores en las Utilidades PTU Exento
        NOMCF2.NOMCF.TextMatrix(I7, 142) = Format(0, "#,##0.00")
    '144 D027 Ajuste en Participación de los Trabajadores en las Utilidades PTU Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 143) = Format(0, "#,##0.00")
    '145 D028 Ajuste en Reembolso de Gastos Médicos Dentales y Hospitalarios Exento
        NOMCF2.NOMCF.TextMatrix(I7, 144) = Format(0, "#,##0.00")
    '146 D029 Ajuste en Fondo de ahorro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 145) = Format(0, "#,##0.00")
    '147 D030 Ajuste en Caja de ahorro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 146) = Format(0, "#,##0.00")
    '148 D031 Ajuste en Contribuciones a Cargo del Trabajador Pagadas por el Patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 147) = Format(0, "#,##0.00")
    '149 D032 Ajuste en Premios por puntualidad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 148) = Format(0, "#,##0.00")
    '150 D033 Ajuste en Prima de Seguro de vida Exento
        NOMCF2.NOMCF.TextMatrix(I7, 149) = Format(0, "#,##0.00")
    '151 D034 Ajuste en Seguro de Gastos Médicos Mayores Exento
        NOMCF2.NOMCF.TextMatrix(I7, 150) = Format(0, "#,##0.00")
    '152 D035 Ajuste en Cuotas Sindicales Pagadas por el Patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 151) = Format(0, "#,##0.00")
    '153 D036 Ajuste en Subsidios por incapacidad Exento
        NOMCF2.NOMCF.TextMatrix(I7, 152) = Format(0, "#,##0.00")
    '154 D037 Ajuste en Becas para trabajadores y/o hijos Exento
        NOMCF2.NOMCF.TextMatrix(I7, 153) = Format(0, "#,##0.00")
    '155 D038 Ajuste en Horas extra Exento
        NOMCF2.NOMCF.TextMatrix(I7, 154) = Format(0, "#,##0.00")
    '156 D039 Ajuste en Horas extra Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 155) = Format(0, "#,##0.00")
    '157 D040 Ajuste en Prima dominical Exento
        NOMCF2.NOMCF.TextMatrix(I7, 156) = Format(0, "#,##0.00")
    '158 D041 Ajuste en Prima dominical Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 157) = Format(0, "#,##0.00")
    '159 D042 Ajuste en Prima vacacional Exento
        NOMCF2.NOMCF.TextMatrix(I7, 158) = Format(0, "#,##0.00")
    '160 D043 Ajuste en Prima vacacional Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 159) = Format(0, "#,##0.00")
    '161 D044 Ajuste en Prima por antigüedad Exento
        NOMCF2.NOMCF.TextMatrix(I7, 160) = Format(0, "#,##0.00")
    '162 D045 Ajuste en Prima por antigüedad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 161) = Format(0, "#,##0.00")
    '163 D046 Ajuste en Pagos por separación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 162) = Format(0, "#,##0.00")
    '164 D047 Ajuste en Pagos por separación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 163) = Format(0, "#,##0.00")
    '165 D048 Ajuste en Seguro de retiro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 164) = Format(0, "#,##0.00")
    '166 D049 Ajuste en Indemnizaciones Exento
        NOMCF2.NOMCF.TextMatrix(I7, 165) = Format(0, "#,##0.00")
    '167 D050 Ajuste en Indemnizaciones Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 166) = Format(0, "#,##0.00")
    '168 D051 Ajuste en Reembolso por funeral Exento
        NOMCF2.NOMCF.TextMatrix(I7, 167) = Format(0, "#,##0.00")
    '169 D052 Ajuste en Cuotas de seguridad social pagadas por el patrón Exento
        NOMCF2.NOMCF.TextMatrix(I7, 168) = Format(0, "#,##0.00")
    '170 D053 Ajuste en Comisiones Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 169) = Format(0, "#,##0.00")
    '171 D054 Ajuste en Vales de despensa Exento
        NOMCF2.NOMCF.TextMatrix(I7, 170) = Format(0, "#,##0.00")
    '172 D055 Ajuste en Vales de restaurante Exento
        NOMCF2.NOMCF.TextMatrix(I7, 171) = Format(0, "#,##0.00")
    '173 D056 Ajuste en Vales de gasolina Exento
        NOMCF2.NOMCF.TextMatrix(I7, 172) = Format(0, "#,##0.00")
    '174 D057 Ajuste en Vales de ropa Exento
        NOMCF2.NOMCF.TextMatrix(I7, 173) = Format(0, "#,##0.00")
    '175 D058 Ajuste en Ayuda para renta Exento
        NOMCF2.NOMCF.TextMatrix(I7, 174) = Format(0, "#,##0.00")
    '176 D059 Ajuste en Ayuda para artículos escolares Exento
        NOMCF2.NOMCF.TextMatrix(I7, 175) = Format(0, "#,##0.00")
    '177 D060 Ajuste en Ayuda para anteojos Exento
        NOMCF2.NOMCF.TextMatrix(I7, 176) = Format(0, "#,##0.00")
    '178 D061 Ajuste en Ayuda para transporte Exento
        NOMCF2.NOMCF.TextMatrix(I7, 177) = Format(0, "#,##0.00")
    '179 D062 Ajuste en Ayuda para gastos de funeral Exento
        NOMCF2.NOMCF.TextMatrix(I7, 178) = Format(0, "#,##0.00")
    '180 D063 Ajuste en Otros ingresos por salarios Exento
        NOMCF2.NOMCF.TextMatrix(I7, 179) = Format(0, "#,##0.00")
    '181 D064 Ajuste en Otros ingresos por salarios Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 180) = Format(0, "#,##0.00")
    '182 D065 Ajuste en Jubilaciones, pensiones o haberes de retiro Exento
        NOMCF2.NOMCF.TextMatrix(I7, 181) = Format(0, "#,##0.00")
    '183 D066 Ajuste en Jubilaciones, pensiones o haberes de retiro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 182) = Format(0, "#,##0.00")
    '184 D067 Ajuste en Pagos por separación Acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 183) = Format(0, "#,##0.00")
    '185 D068 Ajuste en Pagos por separación No acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 184) = Format(0, "#,##0.00")
    '186 D069 Ajuste en Jubilaciones, pensiones o haberes de retiro Acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 185) = Format(0, "#,##0.00")
    '187 D070 Ajuste en Jubilaciones, pensiones o haberes de retiro No acumulable
        NOMCF2.NOMCF.TextMatrix(I7, 186) = Format(0, "#,##0.00")
    '188 D071 Ajuste en Subsidio para el empleo (efectivamente entregado al trabajador)
        NOMCF2.NOMCF.TextMatrix(I7, 187) = Format(0, "#,##0.00")
    '189 D072 Ajuste en Ingresos en acciones o títulos valor que representan bienes Exento
        NOMCF2.NOMCF.TextMatrix(I7, 188) = Format(0, "#,##0.00")
    '190 D073 Ajuste en Ingresos en acciones o títulos valor que representan bienes Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 189) = Format(0, "#,##0.00")
    '191 D074 Ajuste en Alimentación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 190) = Format(0, "#,##0.00")
    '192 D075 Ajuste en Alimentación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 191) = Format(0, "#,##0.00")
    '193 D076 Ajuste en Habitación Exento
        NOMCF2.NOMCF.TextMatrix(I7, 192) = Format(0, "#,##0.00")
    '194 D077 Ajuste en Habitación Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 193) = Format(0, "#,##0.00")
    '195 D078 Ajuste en Premios por asistencia
        NOMCF2.NOMCF.TextMatrix(I7, 194) = Format(0, "#,##0.00")
    '196 D079 Ajuste en Pagos distintos a los listados y que no deben considerarse como ingreso por sueldos, salarios o ingresos asimilados.
        NOMCF2.NOMCF.TextMatrix(I7, 195) = Format(0, "#,##0.00")
    '197 D080 Ajuste en Viáticos gravados
        NOMCF2.NOMCF.TextMatrix(I7, 196) = Format(0, "#,##0.00")
    '198 D081 Ajuste en Viáticos (entregados al trabajador)
        NOMCF2.NOMCF.TextMatrix(I7, 197) = Format(0, "#,##0.00")
    '199 D082 Ajuste en Fondo de ahorro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 198) = Format(0, "#,##0.00")
    '200 D083 Ajuste en Caja de ahorro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 199) = Format(0, "#,##0.00")
    '201 D084 Ajuste en Prima de Seguro de vida Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 200) = Format(0, "#,##0.00")
    '202 D085 Ajuste en Seguro de Gastos Médicos Mayores Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 201) = Format(0, "#,##0.00")
    '203 D086 Ajuste en Subsidios por incapacidad Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 202) = Format(0, "#,##0.00")
    '204 D087 Ajuste en Becas para trabajadores y/o hijos Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 203) = Format(0, "#,##0.00")
    '205 D088 Ajuste en Seguro de retiro Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 204) = Format(0, "#,##0.00")
    '206 D089 Ajuste en Vales de despensa Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 205) = Format(0, "#,##0.00")
    '207 D090 Ajuste en Vales de restaurante Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 206) = Format(0, "#,##0.00")
    '208 D091 Ajuste en Vales de gasolina Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 207) = Format(0, "#,##0.00")
    '209 D092 Ajuste en Vales de ropa Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 208) = Format(0, "#,##0.00")
    '210 D093 Ajuste en Ayuda para renta Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 209) = Format(0, "#,##0.00")
    '211 D094 Ajuste en Ayuda para artículos escolares Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 210) = Format(0, "#,##0.00")
    '212 D095 Ajuste en Ayuda para anteojos Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 211) = Format(0, "#,##0.00")
    '213 D096 Ajuste en Ayuda para transporte Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 212) = Format(0, "#,##0.00")
    '214 D097 Ajuste en Ayuda para gastos de funeral Gravado
        NOMCF2.NOMCF.TextMatrix(I7, 213) = Format(0, "#,##0.00")
    '215 D098 Ajuste a ingresos asimilados a salarios gravados
        NOMCF2.NOMCF.TextMatrix(I7, 214) = Format(0, "#,##0.00")
    '216 D099 Ajuste a ingresos por sueldos y salarios gravados
        NOMCF2.NOMCF.TextMatrix(I7, 215) = Format(0, "#,##0.00")
    '217 D100 Ajuste en Viáticos exentos
        NOMCF2.NOMCF.TextMatrix(I7, 216) = Format(0, "#,##0.00")


    '218 OP001 Reintegro de ISR pagado en exceso (siempre que no haya sido enterado al SAT).
        NOMCF2.NOMCF.TextMatrix(I7, 217) = Format(0, "#,##0.00")
    '219 OP002 Subsidio para el empleo (efectivamente entregado al trabajador).
        NOMCF2.NOMCF.TextMatrix(I7, 218) = Format(sub13, "#,##0.00")
    '220 OP003 Viáticos (entregados al trabajador).
        NOMCF2.NOMCF.TextMatrix(I7, 219) = Format(0, "#,##0.00")
    '221 OP004 Aplicación de saldo a favor por compensación anual.
        NOMCF2.NOMCF.TextMatrix(I7, 220) = Format(0, "#,##0.00")
    '222 OP999 Pagos distintos a los listados y que no deben considerarse como ingreso por sueldos, salarios o ingresos asimilados.
        NOMCF2.NOMCF.TextMatrix(I7, 221) = Format(0, "#,##0.00")
    '223 DINCAPACIDADM
        NOMCF2.NOMCF.TextMatrix(I7, 222) = Format(0, "#,##0.00")
    '224 DINCAPACIDADI
        NOMCF2.NOMCF.TextMatrix(I7, 223) = Format(0, "#,##0.00")
    '225 DINCAPACIDADR
        NOMCF2.NOMCF.TextMatrix(I7, 224) = Format(0, "#,##0.00")
    '226 P014DINCAM
        NOMCF2.NOMCF.TextMatrix(I7, 225) = Format(0, "#,##0.00")
    '227 P014DINCAI
        NOMCF2.NOMCF.TextMatrix(I7, 226) = Format(0, "#,##0.00")
    '228 P014DINCAR
        NOMCF2.NOMCF.TextMatrix(I7, 227) = Format(0, "#,##0.00")
    '229 DIASHE
        NOMCF2.NOMCF.TextMatrix(I7, 228) = Format(0, "#,##0.00")
    '230 HORASDOBLES
        NOMCF2.NOMCF.TextMatrix(I7, 229) = Format(0, "#,##0.00")
    '231 DIASHETRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 230) = Format(0, "#,##0.00")
    '232 HORASTRIPLES
        NOMCF2.NOMCF.TextMatrix(I7, 231) = Format(0, "#,##0.00")
    '233 DIASHESIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 232) = Format(0, "#,##0.00")
    '234 HORASSIMPLES
        NOMCF2.NOMCF.TextMatrix(I7, 233) = Format(0, "#,##0.00")
    '235 AT_VMERCADO
        NOMCF2.NOMCF.TextMatrix(I7, 234) = Format(0, "#,##0.00")
    '236 AT_POTORGARSE
        NOMCF2.NOMCF.TextMatrix(I7, 235) = Format(0, "#,##0.00")
    '237 JB_TEXHIBICION
        NOMCF2.NOMCF.TextMatrix(I7, 236) = Format(0, "#,##0.00")
    '238 JB_TPARCIAL
        NOMCF2.NOMCF.TextMatrix(I7, 237) = Format(0, "#,##0.00")
    '239 JB_MDIARIO
        NOMCF2.NOMCF.TextMatrix(I7, 238) = Format(0, "#,##0.00")
    '240 JB_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 239) = Format(0, "#,##0.00")
    '241 JB_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 240) = Format(0, "#,##0.00")
    '242 SI_TPAGADO
        NOMCF2.NOMCF.TextMatrix(I7, 241) = Format(0, "#,##0.00")
    '243 SI_ASERVICIO
        NOMCF2.NOMCF.TextMatrix(I7, 242) = Format(0, "#,##0.00")
    '244 SI_USUELDO
        NOMCF2.NOMCF.TextMatrix(I7, 243) = Format(0, "#,##0.00")
    '245 SI_IACUM
        NOMCF2.NOMCF.TextMatrix(I7, 244) = Format(0, "#,##0.00")
    '246 SI_INOACUM
        NOMCF2.NOMCF.TextMatrix(I7, 245) = Format(0, "#,##0.00")
    '247 CSF_SALFAV
        NOMCF2.NOMCF.TextMatrix(I7, 246) = Format(0, "#,##0.00")
    '248 CSF_ANIO
        NOMCF2.NOMCF.TextMatrix(I7, 247) = Format(0, "#,##0.00")
    '249 CSF_RSFAV
        NOMCF2.NOMCF.TextMatrix(I7, 248) = Format(0, "#,##0.00")
    '250 BANDERA
        NOMCF2.NOMCF.TextMatrix(I7, 249) = 1
End Sub


