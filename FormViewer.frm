VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormViewer 
   Caption         =   "Visualizador de Nomina Histórica"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   17490
   LinkTopic       =   "FormViewer"
   ScaleHeight     =   8220
   ScaleWidth      =   17490
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdLimpiarFiltro 
      Caption         =   "Limpiar"
      Height          =   375
      Left            =   16320
      TabIndex        =   8
      Top             =   120
      Width           =   1000
   End
   Begin VB.TextBox TxtBuscar 
      Height          =   315
      Left            =   13200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   120
      Width           =   3000
   End
   Begin VB.CommandButton CmdCopyAll 
      Caption         =   "Copiar Todo"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdCopySel 
      Caption         =   "Copiar Selección"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Regresar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   7200
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   17265
      _ExtentX        =   30454
      _ExtentY        =   12700
      _Version        =   393216
      AllowUserResizing=   3
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7845
      Visible         =   0   'False
      Width           =   17265
      _ExtentX        =   30454
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      Height          =   255
      Left            =   12480
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Label LblFile 
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   180
      Width           =   6195
   End
   Begin VB.Menu mnuOrdenar 
      Caption         =   "Ordenar"
      Begin VB.Menu mnuOrdenarID 
         Caption         =   "Por ID"
         Begin VB.Menu mnuIDDesc 
            Caption         =   "Mayor a Menor"
         End
         Begin VB.Menu mnuIDAsc 
            Caption         =   "Menor a Mayor"
         End
      End
      Begin VB.Menu mnuOrdenarNombre 
         Caption         =   "Por Nombre"
         Begin VB.Menu mnuNombreAsc 
            Caption         =   "A - Z"
         End
         Begin VB.Menu mnuNombreDesc 
            Caption         =   "Z - A"
         End
      End
   End
   Begin VB.Menu mnuExcel 
      Caption         =   "Excel"
      Begin VB.Menu mnuExportarExcel 
         Caption         =   "Exportar a Excel"
      End
   End
   Begin VB.Menu mnuCFDI 
      Caption         =   "CFDI"
   End
   Begin VB.Menu mnuImpresion 
      Caption         =   "Impresión"
      Begin VB.Menu mnuNomina 
         Caption         =   "Nomina"
         Begin VB.Menu mnuNominaNominaTotal 
            Caption         =   "Nomina Total"
         End
         Begin VB.Menu mnuNominaTodasLasObras 
            Caption         =   "Todas las Obras"
         End
         Begin VB.Menu mnuNominaNominaOficina 
            Caption         =   "Nomina Oficina"
         End
      End
      Begin VB.Menu mnuRecibos 
         Caption         =   "Recibos"
         Begin VB.Menu mnuReciboTodos 
            Caption         =   "Todos"
         End
         Begin VB.Menu mnuReciboIndividuales 
            Caption         =   "Individuales"
         End
      End
      Begin VB.Menu mnuCheques 
         Caption         =   "Cheques"
         Begin VB.Menu mnuChequesAjustarImpresion 
            Caption         =   "Ajustar Impresion"
         End
         Begin VB.Menu mnuChequesIndividuales 
            Caption         =   "Individuales"
         End
         Begin VB.Menu mnuChequesTotales 
            Caption         =   "Totales"
         End
         Begin VB.Menu mnuChequesDirectorioDeObra 
            Caption         =   "Directorio de Obra"
         End
      End
      Begin VB.Menu mnuDetalleCalculoImpuestos 
         Caption         =   "Detalle Calculo Imptos."
      End
   End
   Begin VB.Menu mnuPortapapeles 
      Caption         =   "Portapapeles"
      Begin VB.Menu mnuCopiarTodo 
         Caption         =   "Copiar todo"
      End
      Begin VB.Menu mnuCopiarSeleccion 
         Caption         =   "Copiar selección"
      End
   End
End
Attribute VB_Name = "FormViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N_ormal As Integer
Dim diat As Double
Dim Arch As String
Dim Arch1 As String

Dim Dm As Long
Dim nm As Long
Dim fi_nm As Long
Dim limite As Long
Dim renglon As Long

Private cargandoNomina As Boolean

Public CargaCorrecta As Boolean

Private Const FILA_INICIO_DATOS As Long = 1
Private Const COL_BUSQUEDA As Long = 1

Private ordenAscendente As Boolean
Private textoFiltro As String

Private Const COL_ID As Long = 0
Private Const COL_NOMBRE As Long = 1

Private Sub CmdLimpiarFiltro_Click()

    Dim i As Long
    Dim ultimaFila As Long

    If cargandoNomina Then Exit Sub

    TxtBuscar.Text = ""
    textoFiltro = ""

    If ConNom1.Rows <= 1 Then Exit Sub

    ultimaFila = limite

    If ultimaFila > ConNom1.Rows - 2 Then
        ultimaFila = ConNom1.Rows - 2
    End If

    If ultimaFila >= 1 Then

        For i = FILA_INICIO_DATOS To ultimaFila

            If Trim$(ConNom1.TextMatrix(i, COL_NOMBRE)) <> "" Then

                ConNom1.RowHeight(i) = -1

            Else

                ConNom1.RowHeight(i) = 0

            End If

        Next i

    End If

    If limite + 1 >= 0 And limite + 1 < ConNom1.Rows Then
        ConNom1.RowHeight(limite + 1) = -1
    End If

    If ConNom1.Rows > 1 Then
        ConNom1.Row = 1
        ConNom1.col = 0
    End If

End Sub

Private Sub Form_Load()
    
    InicializarColumnas tnOrdinaria
    cargandoNomina = True
    CargaCorrecta = False

    If z1$ = "" Then
        z1$ = "###,###,##0.00"
    End If

    limite = 0
    renglon = 0
    nm = 0
    fi_nm = 0
    Dm = 0

    textoFiltro = ""

    TxtBuscar.Text = ""

    ConNom1.SelectionMode = flexSelectionByRange
    ConNom1.AllowUserResizing = flexResizeColumns

    cargandoNomina = False

End Sub

Public Sub LoadPayroll(ByVal fName As String)

    On Error GoTo ErrorHandler

    Dim nameOnly As String
    Dim pos As Long
    Dim rutaBase As String
    Dim r As Long
    Dim yavas As Integer
    Dim totalMaestro As Long
    Dim totalCmp As Long
    Dim totalBnx As Long
    Dim numeroError As Long
    Dim descripcionError As String

    ' Variables para diagnostico
    Dim etapaCarga As String
    Dim registroActual As Long

    etapaCarga = "Inicio"
    registroActual = 0

    '========================================================
    ' 1. INICIO DE CARGA
    '========================================================

    etapaCarga = "Cerrar archivos anteriores"

    cargandoNomina = True
    CargaCorrecta = False

    CerrarArchivosVisor

    '========================================================
    ' 2. REINICIAR ESTADO
    '========================================================

    etapaCarga = "Reiniciar variables"

    limite = 0
    renglon = 0
    nm = 0
    fi_nm = 0
    Dm = 0

    textoFiltro = ""
    TxtBuscar.Text = ""

    '========================================================
    ' 3. OBTENER NOMBRE Y RUTA
    '========================================================

    etapaCarga = "Obtener nombre y ruta"

    nameOnly = fName
    pos = InStrRev(fName, "\")

    If pos > 0 Then

        nameOnly = Mid$(fName, pos + 1)
        rutaBase = Left$(fName, pos)

    Else

        rutaBase = CurDir$

        If Right$(rutaBase, 1) <> "\" Then
            rutaBase = rutaBase & "\"
        End If

    End If

    LblFile.Caption = "Archivo: " & nameOnly

    Debug.Print "========================================"
    Debug.Print "INICIANDO CARGA DE NOMINA"
    Debug.Print "Archivo: "; nameOnly
    Debug.Print "Ruta: "; rutaBase
    Debug.Print "========================================"

    '========================================================
    ' 4. VALIDAR NOMBRE
    '========================================================

    etapaCarga = "Validar nombre del archivo"

    If Len(nameOnly) < 5 Then

        MsgBox "El nombre del archivo de n mina no es v lido.", _
               vbExclamation, "Visor de n mina"

        GoTo Salida

    End If

    
    '========================================================
    ' VALIDAR ARCHIVOS ANTES DE ABRIRLOS
    '========================================================
    
    Arch = rutaBase & nameOnly
    
    Arch1 = rutaBase & _
            Left$(nameOnly, Len(nameOnly) - 4) & ".cmp"
    
    If Len(Dir$(Arch)) = 0 Then
        MsgBox "No existe el archivo de nomina:" & vbCrLf & Arch, vbExclamation, "Visor"
        GoTo Salida
    End If
    
    If Len(Dir$(rutaBase & "personal.dno")) = 0 Then
        MsgBox "No existe personal.dno.", vbExclamation, "Visor"
        GoTo Salida
    End If
    
    If Len(Dir$(rutaBase & "maestro.dno")) = 0 Then
        MsgBox "No existe maestro.dno.", vbExclamation, "Visor"
        GoTo Salida
    End If
    
    If Len(Dir$(rutaBase & "bnxcla.dno")) = 0 Then
        MsgBox "No existe bnxcla.dno.", vbExclamation, "Visor"
        GoTo Salida
    End If
    
    If Len(Dir$(Arch1)) = 0 Then
        MsgBox "No existe el complemento:" & vbCrLf & Arch1, vbExclamation, "Visor"
        GoTo Salida
    End If


    '========================================================
    ' 5. ABRIR PERSONAL.DNO
    '========================================================

    etapaCarga = "Abrir personal.dno"

    Open rutaBase & "personal.dno" For Random As #2 _
        Len = Len(personal)

    Dm = LOF(2) / Len(personal)

    Debug.Print "Registros personal.dno: "; Dm

    If Dm <= 0 Then

        MsgBox "No existe informaci n v lida en personal.dno." & _
               vbCrLf & "No se puede abrir la n mina.", _
               vbCritical, "Error"

        GoTo Salida

    End If

    '========================================================
    ' 6. IDENTIFICAR TIPO DE NOMINA
    '========================================================

    etapaCarga = "AplicarTipoNomina"

    ' Usar la ruta completa para comprobar marca interna 20261010.
    AplicarTipoNomina Arch

    If EsNominaNormal(nameOnly) Then

        N_ormal = 0
        diat = 15

    Else

        N_ormal = 1
        diat = 0

    End If

    Debug.Print "Tipo de nomina: "; g_TipoNominaActiva
    Debug.Print "N_ormal: "; N_ormal
    Debug.Print "C_EXENTA: "; C_EXENTA
    Debug.Print "C_NETO: "; C_NETO
    Debug.Print "C_TOTCOLS: "; C_TOTCOLS

    '========================================================
    ' 7. CONSTRUIR RUTAS
    '========================================================

    etapaCarga = "Construir rutas de archivos"

    Arch = fName

    Arch1 = rutaBase & _
            Left$(nameOnly, Len(nameOnly) - 4) & ".cmp"

    Debug.Print "NOM: "; Arch
    Debug.Print "CMP: "; Arch1

    '========================================================
    ' 8. ABRIR NOMINA
    '========================================================

    etapaCarga = "Abrir archivo NOM"

    Open Arch For Random As #6 Len = Len(nomina)

    nm = LOF(6) / Len(nomina)

    Debug.Print "Registros NOM: "; nm

    If nm <= 0 Then

        MsgBox "El archivo de n mina est  vac o o " & _
               "no contiene registros.", _
               vbInformation, "Visor de N mina"

        GoTo Salida

    End If

    '========================================================
    ' 9. ABRIR ARCHIVOS AUXILIARES
    '========================================================

    etapaCarga = "Abrir archivos auxiliares"

    Open Arch1 For Random As #14 Len = Len(nom_com)

    Open rutaBase & "bnxcla.dno" For Random As #12 _
        Len = Len(Clbnx)

    Open rutaBase & "maestro.dno" For Random As #8 _
        Len = Len(maestro)

    totalCmp = LOF(14) / Len(nom_com)
    totalBnx = LOF(12) / Len(Clbnx)
    totalMaestro = LOF(8) / Len(maestro)

    Debug.Print "Registros CMP: "; totalCmp
    Debug.Print "Registros BNX: "; totalBnx
    Debug.Print "Registros Maestro: "; totalMaestro

    '========================================================
    ' 10. DETERMINAR NUMERO DE REGISTROS
    '========================================================

    etapaCarga = "Determinar numero de registros"

    If nm < Dm Then
        fi_nm = nm
    Else
        fi_nm = Dm
    End If

    Debug.Print "Registros a procesar: "; fi_nm

    If fi_nm <= 0 Then

        MsgBox "No existen registros compatibles entre " & _
               "la n mina y personal.dno.", _
               vbInformation, "Visor de N mina"

        GoTo Salida

    End If

    '========================================================
    ' 11. PREPARAR GRID
    '========================================================

    etapaCarga = "Preparar grid"

    ConNom1.Redraw = False
    ConNom1.Clear

    ConNom1.Cols = C_TOTCOLS
    ConNom1.Rows = fi_nm + 2

    ConNom1.Row = 0
    ConNom1.RowSel = 0
    ConNom1.col = 0
    ConNom1.ColSel = 0

    Debug.Print "ANTES DE DEFINE"
    Debug.Print "Rows: "; ConNom1.Rows
    Debug.Print "Cols: "; ConNom1.Cols

    '========================================================
    ' 12. DEFINIR ENCABEZADOS
    '========================================================

    etapaCarga = "Ejecutar define"

    define

    Debug.Print "DESPUES DE DEFINE"
    Debug.Print "Rows: "; ConNom1.Rows
    Debug.Print "Cols: "; ConNom1.Cols
    Debug.Print "Cols esperadas: "; C_TOTCOLS

    limite = 0
    renglon = 0

    '========================================================
    ' 13. PROGRESSBAR
    '========================================================

    etapaCarga = "Inicializar ProgressBar"

    ProgressBar1.Visible = True
    ProgressBar1.Min = 0
    ProgressBar1.Max = fi_nm
    ProgressBar1.Value = 0

    '========================================================
    ' 14. LEER REGISTROS
    '========================================================

    For r = 1 To fi_nm

        registroActual = r
        etapaCarga = "Actualizar ProgressBar"

        ProgressBar1.Value = r

        '--------------------------------------------
        ' NOMINA
        '--------------------------------------------

        etapaCarga = "Leer registro NOM"

        Get #6, r, nomina

        '--------------------------------------------
        ' PERSONAL
        '--------------------------------------------

        etapaCarga = "Leer registro PERSONAL"

        Get #2, r, personal

        '--------------------------------------------
        ' MAESTRO
        '--------------------------------------------

        etapaCarga = "Leer registro MAESTRO"

        If r <= totalMaestro Then
            Get #8, r, maestro
        End If

        '--------------------------------------------
        ' CMP
        '--------------------------------------------

        etapaCarga = "Leer registro CMP"

        If r <= totalCmp Then
            Get #14, r, nom_com
        End If

        '--------------------------------------------
        ' VERIFICAR INFORMACION
        '--------------------------------------------

        etapaCarga = "Ejecutar verifica"

        yavas = 0
        verifica yavas

        If yavas > 0 Then

            '----------------------------------------
            ' AUMENTAR RENGLON
            '----------------------------------------

            etapaCarga = "Preparar renglon"

            renglon = renglon + 1

            If renglon >= ConNom1.Rows Then
                ConNom1.Rows = renglon + 2
            End If

            '----------------------------------------
            ' SELECCIONAR FILA
            '----------------------------------------

            etapaCarga = "Seleccionar fila del grid"

            ConNom1.Row = renglon
            ConNom1.col = C_NUM

            '----------------------------------------
            ' NUMERO DE EMPLEADO
            '----------------------------------------

            etapaCarga = "Escribir numero de empleado"

            ConNom1.Text = Format$(r, "#####")

            limite = limite + 1
            regtro = r

            '----------------------------------------
            ' CARGAR DATOS DEL EMPLEADO
            '----------------------------------------

            etapaCarga = "Ejecutar carganom"

            Debug.Print "CARGANOM - Registro: "; r; _
                        " Fila: "; renglon; _
                        " Cols: "; ConNom1.Cols

            carganom

        End If

    Next r

    '========================================================
    ' 15. ELIMINAR REGISTROS SIN INFORMACION
    '========================================================
    
    etapaCarga = "Ejecutar eliminacion"
    
    Debug.Print "ANTES DE ELIMINACION"
    Debug.Print "Limite: "; limite
    Debug.Print "Rows: "; ConNom1.Rows
    Debug.Print "Cols: "; ConNom1.Cols
    
    eliminacion
    
    Debug.Print "DESPUES DE ELIMINACION"
    Debug.Print "Limite: "; limite
    Debug.Print "Rows: "; ConNom1.Rows
    Debug.Print "Cols: "; ConNom1.Cols

    '========================================================
    ' 16. RESTABLECER SELECCION
    '========================================================

    etapaCarga = "Restablecer seleccion"

    ConNom1.Row = 0
    ConNom1.RowSel = 0

    ConNom1.col = 0
    ConNom1.ColSel = 0

    '========================================================
    ' 17. AJUSTAR FILAS
    '========================================================

    etapaCarga = "Ajustar filas finales"

    If limite < 0 Then limite = 0

    Debug.Print "ANTES DE AJUSTAR FILAS"
    Debug.Print "Limite: "; limite
    Debug.Print "Rows actuales: "; ConNom1.Rows
    Debug.Print "Rows nuevas: "; limite + 2

    ConNom1.Rows = limite + 2

    Debug.Print "DESPUES DE AJUSTAR FILAS"
    Debug.Print "Rows: "; ConNom1.Rows

    '========================================================
    ' 18. CALCULAR SUMAS
    '========================================================

    etapaCarga = "Ejecutar sumavert"

    Debug.Print "ANTES DE SUMAVERT"
    Debug.Print "Limite: "; limite
    Debug.Print "Rows: "; ConNom1.Rows
    Debug.Print "Cols: "; ConNom1.Cols
    Debug.Print "C_NETO: "; C_NETO

    sumavert

    Debug.Print "DESPUES DE SUMAVERT"

    '========================================================
    ' 19. DEJAR SELECCION EN FILA VALIDA
    '========================================================

    etapaCarga = "Seleccionar fila final"

    If limite > 0 Then

        ConNom1.Row = 1
        ConNom1.RowSel = 1

    Else

        ConNom1.Row = 0
        ConNom1.RowSel = 0

    End If

    ConNom1.col = 0
    ConNom1.ColSel = 0

    '========================================================
    ' 20. FINALIZAR
    '========================================================

    etapaCarga = "Finalizar carga"

    ConNom1.Redraw = True
    ProgressBar1.Visible = False

    CargaCorrecta = True

    Debug.Print "========================================"
    Debug.Print "NOMINA CARGADA CORRECTAMENTE"
    Debug.Print "Archivo: "; nameOnly
    Debug.Print "Empleados: "; limite
    Debug.Print "========================================"

    GoTo Salida

    '========================================================
    ' MANEJO DE ERRORES
    '========================================================

ErrorHandler:

    numeroError = Err.Number
    descripcionError = Err.Description

    CargaCorrecta = False

    ' Capturar informacion antes de limpiar el visor

    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "ERROR EN FORMVIEWER"
    Debug.Print "========================================"

    Debug.Print "Etapa: "; etapaCarga
    Debug.Print "Error: "; numeroError
    Debug.Print "Descripcion: "; descripcionError

    Debug.Print "Archivo: "; nameOnly
    Debug.Print "Tipo: "; g_TipoNominaActiva

    Debug.Print "Registro actual: "; registroActual

    Debug.Print "Filas del Grid: "; ConNom1.Rows
    Debug.Print "Columnas del Grid: "; ConNom1.Cols

    Debug.Print "Fila seleccionada: "; ConNom1.Row
    Debug.Print "Columna seleccionada: "; ConNom1.col

    Debug.Print "Limite: "; limite
    Debug.Print "Renglon: "; renglon

    Debug.Print "C_EXENTA: "; C_EXENTA
    Debug.Print "C_NETO: "; C_NETO
    Debug.Print "C_TOTCOLS: "; C_TOTCOLS

    Debug.Print "========================================"

    On Error Resume Next

    ConNom1.Redraw = True
    ProgressBar1.Visible = False

    On Error GoTo 0

    MsgBox "Error al cargar la n mina." & vbCrLf & _
           vbCrLf & _
           "Etapa: " & etapaCarga & vbCrLf & _
           "Registro: " & CStr(registroActual) & vbCrLf & _
           "N mero: " & CStr(numeroError) & vbCrLf & _
           "Descripci n: " & descripcionError, _
           vbCritical, "Error FormViewer"

    '========================================================
    ' SALIDA COMUN
    '========================================================

Salida:

    CerrarArchivosVisor
    cargandoNomina = False

End Sub
Private Function IsNormalPayrollName(ByVal fName As String) As Boolean
    Dim namePart As String
    If Len(fName) = 12 Then
        namePart = UCase(Left$(fName, 3))
        If namePart = "ENE" Or namePart = "FEB" Or namePart = "MAR" Or namePart = "ABR" Or _
           namePart = "MAY" Or namePart = "JUN" Or namePart = "JUL" Or namePart = "AGO" Or _
           namePart = "SEP" Or namePart = "OCT" Or namePart = "NOV" Or namePart = "DIC" Then
            
            Dim q As String
            q = Mid$(fName, 4, 1)
            If q = "1" Or q = "2" Then
                If IsNumeric(Mid$(fName, 5, 4)) Then
                    IsNormalPayrollName = True
                    Exit Function
                End If
            End If
        End If
    End If
    IsNormalPayrollName = False
End Function
Private Sub PonCel(ByVal f As Long, ByVal c As Long, ByVal v As Currency, Optional ByVal fmt As String = "")
    If fmt = "" Then fmt = z1$
    If v <> 0 Then
        ConNom1.TextMatrix(f, c) = Format(v, fmt)
    Else
        ConNom1.TextMatrix(f, c) = ""
    End If
End Sub

Sub carganom()
    Dim liLocal As Long
    Dim ingresosLocal As Currency
    Dim deduccionesLocal As Currency
    Dim netoLocal As Currency

    liLocal = ConNom1.Row
    ingresosLocal = nomina.sueldo

    ConNom1.TextMatrix(liLocal, C_NOMBRE) = _
        RTrim$(personal.ape1) & " " & RTrim$(personal.ape2) & " " & RTrim$(personal.nom)

    PonCel liLocal, C_DIAS, nomina.dias, "##0.00"
    PonCel liLocal, C_SUELDO, nomina.sueldo

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        ' Los indices no habilitados valen -1. Nunca escribir en -1.
        ' Acumular los conceptos visibles segun el tipo y la version.
        If C_LQ_SDOVAC >= 0 Then
            PonCel liLocal, C_LQ_SDOVAC, nomina.viaticos
            ingresosLocal = ingresosLocal + nomina.viaticos
        End If
        If C_LQ_AGUI >= 0 Then
            PonCel liLocal, C_LQ_AGUI, nomina.aguin
            ingresosLocal = ingresosLocal + nomina.aguin
        End If
        If C_LQ_ANTIG >= 0 Then
            PonCel liLocal, C_LQ_ANTIG, nomina.hs_dbl
            ingresosLocal = ingresosLocal + nomina.hs_dbl
        End If
        If C_LQ_COMPENS >= 0 Then
            PonCel liLocal, C_LQ_COMPENS, nomina.hs_nor
            ingresosLocal = ingresosLocal + nomina.hs_nor
        End If
        If C_LQ_INDEMN >= 0 Then
            PonCel liLocal, C_LQ_INDEMN, nomina.hs_tri
            ingresosLocal = ingresosLocal + nomina.hs_tri
        End If
        If C_LQ_BONO >= 0 Then
            PonCel liLocal, C_LQ_BONO, nomina.hsnor
            ingresosLocal = ingresosLocal + nomina.hsnor
        End If
        If C_LQ_PRIMAVAC >= 0 Then
            PonCel liLocal, C_LQ_PRIMAVAC, nomina.pvac
            ingresosLocal = ingresosLocal + nomina.pvac
        End If
    Else
        ' Logica anterior de ordinarias y especiales no-separacion.
        PonCel liLocal, C_CONC4, nomina.hs_nor
        ingresosLocal = ingresosLocal + nomina.hs_nor
        If N_ormal = 1 Then
            PonCel liLocal, C_CONC5, nomina.aguin
            ingresosLocal = ingresosLocal + nomina.aguin
        Else
            PonCel liLocal, C_CONC5, nomina.hs_dbl
            ingresosLocal = ingresosLocal + nomina.hs_dbl
        End If
        If N_ormal = 1 Then
            PonCel liLocal, C_CONC6, nomina.ptu
            ingresosLocal = ingresosLocal + nomina.ptu
        Else
            PonCel liLocal, C_CONC6, nomina.hs_tri
            ingresosLocal = ingresosLocal + nomina.hs_tri
        End If
        PonCel liLocal, C_CONC7, nomina.viaticos
        PonCel liLocal, C_CONC8, nomina.pvac
        ' IMPORTANTE: formula original sin cambios (incluye campos
        ' historicos aun cuando no sean visibles en este tipo).
        ingresosLocal = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + _
                        nomina.hs_tri + nomina.aguin + nomina.ptu + _
                        nomina.viaticos + nomina.pvac
    End If

    PonCel liLocal, C_OTRAS, nomina.otras
    PonCel liLocal, C_EXENTA, nomina.exentos
    ingresosLocal = ingresosLocal + nomina.otras

    ' En LIQ/FIN nuevo, exenta es PARTE del ingreso ya capturado.
    ' En normales y legacy conservar el comportamiento previo.
    If Not EsSeparacionNueva() Then
        ingresosLocal = ingresosLocal + nomina.exentos
    End If
    PonCel liLocal, C_TOTING, ingresosLocal

    PonCel liLocal, C_ISPT, nomina.ispt
    PonCel liLocal, C_SUBEMP, nomina.crdsal
    PonCel liLocal, C_IMSS, nomina.imss
    PonCel liLocal, C_PRESTAMOS, nomina.prestamos
    PonCel liLocal, C_FONACOT, nomina.fonacot
    PonCel liLocal, C_PENSION, nomina.telefono
    PonCel liLocal, C_INFONAVIT, nomina.otraded

    deduccionesLocal = nomina.crdsal + nomina.ispt + nomina.imss + _
                       nomina.prestamos + nomina.fonacot + _
                       nomina.telefono + nomina.otraded
    PonCel liLocal, C_TOTDED, deduccionesLocal

    netoLocal = ingresosLocal - deduccionesLocal
    PonCel liLocal, C_NETO, netoLocal

    ' Mantener globales actualizadas como hacia el carganom original.
    ingresos = ingresosLocal
    deducciones = deduccionesLocal
    neto = netoLocal

    ConNom1.TextMatrix(liLocal, C_FLAG_ISR) = "0"
    ConNom1.TextMatrix(liLocal, C_FLAG_IMSS) = "0"
    If N_ormal = 1 Then
        ConNom1.TextMatrix(liLocal, C_FLAG_ESP) = "1"
    Else
        ConNom1.TextMatrix(liLocal, C_FLAG_ESP) = "0"
    End If

    ' El numero de empleado corresponde al numero del registro NOM.
    If regtro > 0 And regtro <= LOF(12) / Len(Clbnx) Then
        Get #12, regtro, Clbnx
        ConNom1.TextMatrix(liLocal, C_BANAMEX) = " " & Clbnx.Q1
    Else
        ConNom1.TextMatrix(liLocal, C_BANAMEX) = ""
    End If
End Sub

Sub define()
    ' FORMVIEWER - los indices siempre proceden de modColumnas.
    ConNom1.Cols = C_TOTCOLS
    ConNom1.Font = "Arial"
    ConNom1.Font.Size = 8
    ConNom1.Font.Bold = True
    ConNom1.Row = 0

    Encab C_NUM, "No.", 600
    Encab C_NOMBRE, "Nombre", 3500
    Encab C_DIAS, "dias T.", 1200
    Encab C_SUELDO, "Sueldo", 1200

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        If Not EsSeparacionNueva() Then
            ' LEGACY: conservar exactamente los textos y anchos anteriores.
            Encab C_LQ_SDOVAC, "Sdo.Vac.", 1200
            Encab C_LQ_PRIMAVAC, "P.Vacac.", 1200
            Encab C_LQ_AGUI, "Aguinaldo", 1200
            Encab C_LQ_ANTIG, "P.Antig.", 1200
            Encab C_LQ_COMPENS, "Compens.", 1200
            Encab C_LQ_INDEMN, "Indemn.", 1200
            Encab C_LQ_BONO, "BONO", 1200
        Else
            ' Nuevo: solo los conceptos definidos para FINI o LIQ.
            If C_LQ_SDOVAC >= 0 Then
                Encab C_LQ_SDOVAC, "Sdo.Vacaciones", 1450
            End If
            If C_LQ_AGUI >= 0 Then
                Encab C_LQ_AGUI, "Aguinaldo", 1200
            End If
            If C_LQ_ANTIG >= 0 Then
                Encab C_LQ_ANTIG, "Prima Antiguedad", 1450
            End If
            If C_LQ_COMPENS >= 0 Then
                Encab C_LQ_COMPENS, "Compensaciones", 1450
            End If
            If C_LQ_INDEMN >= 0 Then
                Encab C_LQ_INDEMN, "Indemnizacion", 1450
            End If
            If C_LQ_BONO >= 0 Then
                Encab C_LQ_BONO, "Bono", 1200
            End If
            If C_LQ_PRIMAVAC >= 0 Then
                Encab C_LQ_PRIMAVAC, "Prima Vacacional", 1450
            End If
        End If
    Else
        ' MANTENER las columnas de normales, aguinaldo, PTU y bono.
        Encab C_CONC4, "hs.Norm.", 1200

        If g_TipoNominaActiva = tnAguinaldo Or N_ormal <> 0 Then
            Encab C_CONC5, "Aguinaldo", 1200
        Else
            Encab C_CONC5, "hs.Dobles", 1200
        End If

        If g_TipoNominaActiva = tnPTU Then
            Encab C_CONC6, "PTU", 1200
        ElseIf N_ormal = 0 Then
            Encab C_CONC6, "hs.Triples", 1200
        Else
            Encab C_CONC6, "Ptu", 1200
        End If

        If g_TipoNominaActiva = tnBonoPremio Or N_ormal <> 0 Then
            Encab C_CONC7, "BONO", 1200
        Else
            Encab C_CONC7, "O.F.", 1200
        End If
        Encab C_CONC8, "P.Vacac.", 1200
    End If

    Encab C_OTRAS, "Otras", 1200
    Encab C_EXENTA, "Perc.exenta.", 1200
    Encab C_TOTING, "Tot.Ingr.", 1200
    Encab C_ISPT, "Ispt", 1200
    Encab C_SUBEMP, "Sub.P/Empl.", 1200
    Encab C_IMSS, "Imss", 1200
    Encab C_PRESTAMOS, "Prestamos", 1200
    Encab C_FONACOT, "Fonacot", 1200
    Encab C_PENSION, "Pension Alimenticia", 1200
    Encab C_INFONAVIT, "Infonavit", 1200
    Encab C_TOTDED, "Tot.Deduc", 1200
    Encab C_NETO, "Neto", 1200

    ConNom1.ColWidth(C_FLAG_ISR) = 0
    ConNom1.ColWidth(C_FLAG_IMSS) = 0
    ConNom1.ColWidth(C_FLAG_ESP) = 0
    Encab C_BANAMEX, "Banamex", 2400
End Sub


Sub verifica(yavas)
    ' FORMVIEWER: no contar la marca 20261010 como un pago.
    If nomina.dias <> 0 Then yavas = 1
    If nomina.hs_nor <> 0 Then yavas = 1
    If nomina.hs_dbl <> 0 Then yavas = 1
    If nomina.hs_tri <> 0 Then yavas = 1
    If nomina.ispt <> 0 Then yavas = 1
    If nomina.crdsal <> 0 Then yavas = 1
    If nomina.imss <> 0 Then yavas = 1
    If nomina.sueldo <> 0 Then yavas = 1
    If nomina.viaticos <> 0 Then yavas = 1
    If nomina.pvac <> 0 Then yavas = 1
    If nomina.otras <> 0 Then yavas = 1
    If nomina.aguin <> 0 Then yavas = 1
    If nomina.ptu <> 0 Then yavas = 1
    If nomina.exentos <> 0 Then yavas = 1
    If nomina.prestamos <> 0 Then yavas = 1
    If nomina.fonacot <> 0 Then yavas = 1
    If nomina.telefono <> 0 Then yavas = 1
    If nomina.otraded <> 0 Then yavas = 1
    If nomina.hsnor <> 0 Then yavas = 1

    If nomina.hstri <> 0 Then
        If Not (EsSeparacionNueva() And _
                nomina.hstri = CCur(MARCA_SEPARACION_2026)) Then
            yavas = 1
        End If
    End If
End Sub
Sub eliminacion()

    Dim re As Long

    re = 1

    Do While re <= limite

        ' Verificar que la fila exista
        If re >= ConNom1.Rows Then Exit Do

        ' Revisar el TOTAL DE INGRESOS
        If Trim$(ConNom1.TextMatrix(re, C_TOTING)) = "" Then

            ' Evitar una seleccion invalida al quitar la fila
            ConNom1.Row = 0
            ConNom1.RowSel = 0
            ConNom1.col = 0
            ConNom1.ColSel = 0

            ConNom1.RemoveItem re

            limite = limite - 1

        Else

            re = re + 1

        End If

    Loop

End Sub

Sub sumavert()

    Dim Vw As Integer
    Dim late As Integer
    Dim li As Long
    Dim ii As Integer
    Dim sumv() As Currency
    ReDim sumv(0 To C_NETO)

    ProgressBar1.Min = 0
    ProgressBar1.Max = limite
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True

    Vw = 0

    For late = C_SUELDO To C_NETO
        sumv(late) = 0
    Next late

    '========================================================
    ' SUMAR ÚNICAMENTE LOS EMPLEADOS
    '========================================================
    For li = FILA_INICIO_DATOS To limite

        Vw = Vw + 1
        ProgressBar1.Value = li

        For late = C_SUELDO To C_NETO

            ii = late

            If Trim$(ConNom1.TextMatrix(li, ii)) <> "" Then

                If IsNumeric(ConNom1.TextMatrix(li, ii)) Then
                    sumv(late) = sumv(late) + _
                                 CCur(ConNom1.TextMatrix(li, ii))
                End If

            End If

        Next late

    Next li

    '========================================================
    ' FILA DE TOTALES
    '========================================================

    li = limite + 1

    'Limpiar la fila antes de escribir nuevamente los totales
    For ii = 0 To ConNom1.Cols - 1
        ConNom1.TextMatrix(li, ii) = ""
    Next ii

    ConNom1.TextMatrix(li, 1) = _
        "Empleados... " & Str$(Vw) & " S u m a s ....."

    For late = C_SUELDO To C_NETO

        ii = late

        If sumv(late) <> 0 Then
            ConNom1.TextMatrix(li, ii) = _
                Format$(sumv(late), z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If

    Next late

    ProgressBar1.Visible = False

End Sub

Private Sub CmdCopyAll_Click()
 On Error GoTo CopyAllError

    Dim r As Long
    Dim c As Long
    Dim txt As String
    Dim filaTexto As String

    txt = ""

    For r = 0 To limite + 1

        filaTexto = ""

        For c = 0 To ConNom1.Cols - 1

            If c > 0 Then
                filaTexto = filaTexto & vbTab
            End If

            filaTexto = filaTexto & ConNom1.TextMatrix(r, c)

        Next c

        If r > 0 Then
            txt = txt & vbCrLf
        End If

        txt = txt & filaTexto

    Next r

    Clipboard.Clear
    Clipboard.SetText txt, vbCFText

    Exit Sub

CopyAllError:

    MsgBox "Error al copiar: " & Err.Description, _
           vbCritical, "Copiar información"
End Sub

Private Sub CmdCopySel_Click()
 On Error GoTo CopySelError

    Dim rStart As Long
    Dim rEnd As Long
    Dim cStart As Long
    Dim cEnd As Long
    Dim r As Long
    Dim c As Long

    Dim txt As String
    Dim filaTexto As String

    If ConNom1.Row < ConNom1.RowSel Then
        rStart = ConNom1.Row
        rEnd = ConNom1.RowSel
    Else
        rStart = ConNom1.RowSel
        rEnd = ConNom1.Row
    End If

    If ConNom1.col < ConNom1.ColSel Then
        cStart = ConNom1.col
        cEnd = ConNom1.ColSel
    Else
        cStart = ConNom1.ColSel
        cEnd = ConNom1.col
    End If

    ' Nunca copiar la fila de totales si se seleccionó junto con ella
    If rEnd > limite Then
        rEnd = limite
    End If

    If rStart < 1 Then
        rStart = 1
    End If

    If rStart > rEnd Then Exit Sub

    txt = ""

    For r = rStart To rEnd

        filaTexto = ""

        For c = cStart To cEnd

            If c > cStart Then
                filaTexto = filaTexto & vbTab
            End If

            filaTexto = filaTexto & ConNom1.TextMatrix(r, c)

        Next c

        If r > rStart Then
            txt = txt & vbCrLf
        End If

        txt = txt & filaTexto

    Next r

    Clipboard.Clear
    Clipboard.SetText txt, vbCFText

    Exit Sub

CopySelError:

    MsgBox "Error al copiar selección: " & Err.Description, _
           vbCritical, "Copiar selección"
End Sub

Private Sub CmdExit_Click()

    Unload Me

End Sub
Private Sub Form_Unload(Cancel As Integer)

    cargandoNomina = True
    CargaCorrecta = False

    CerrarArchivosVisor

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub

    If Me.ScaleWidth > 240 Then

        ConNom1.Width = Me.ScaleWidth - 240
        ProgressBar1.Width = Me.ScaleWidth - 240

    End If

    If Me.ScaleHeight > 1500 Then

        ConNom1.Height = Me.ScaleHeight - 1500

    End If

    If Me.ScaleHeight > 350 Then

        ProgressBar1.Top = Me.ScaleHeight - 350

    End If

End Sub

Private Sub mnuCFDI_Click()
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorCFDI

    OrigenCFDI = 1
    g_FechaPagoCFDI2026 = ""

    ' Solo las nuevas Liquidaciones y Finiquitos usan la ficha persistida.
    ' No dependemos de la variable MiFecha que dejo Form8.
    If EsSeparacionNueva() Then
        g_FechaPagoCFDI2026 = FechaCFDI_LeerFicha(Arch)
    End If

    UltimaFilaCFDI = limite
    AbrirArchivosParaCFDI
    NOMCF2.Show vbModal
    CerrarArchivosCFDI
    Exit Sub

ErrorCFDI:
    numeroError = Err.Number
    descripcionError = Err.Description
    CerrarArchivosCFDI
    MsgBox "Error al generar el CFDI." & vbCrLf & "Error " & CStr(numeroError) & ": " & descripcionError, vbCritical, "CFDI"
End Sub


Private Sub mnuCopiarSeleccion_Click()
    CmdCopySel_Click
End Sub

Private Sub mnuCopiarTodo_Click()
    CmdCopyAll_Click
End Sub

Private Sub mnuExportarExcel_Click()
    ExportarNominaExcel
End Sub

Private Sub ExportarNominaExcel()

    On Error GoTo ErrorExcel

    Dim objXL As Object
    Dim wbXL As Object
    Dim wsXL As Object

    Dim r As Long
    Dim c As Long

    Set objXL = CreateObject("Excel.Application")

    If objXL Is Nothing Then

        MsgBox "Necesitas tener instalado Microsoft Excel.", _
               vbExclamation, "Exportar a Excel"

        Exit Sub

    End If

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = wbXL.Worksheets(1)

    objXL.Visible = True

    '------------------------------------------
    ' ENCABEZADOS
    '------------------------------------------

    For c = 0 To ConNom1.Cols - 1

        If ConNom1.ColWidth(c) > 0 Then

            wsXL.Cells(1, c + 1).Value = _
                ConNom1.TextMatrix(0, c)

        End If

    Next c

    '------------------------------------------
    ' DATOS
    '------------------------------------------

    For r = 1 To limite

        For c = 0 To ConNom1.Cols - 1

            If ConNom1.ColWidth(c) > 0 Then

                wsXL.Cells(r + 1, c + 1).Value = _
                    ConNom1.TextMatrix(r, c)

            End If

        Next c

    Next r

    '------------------------------------------
    ' FILA DE TOTALES
    '------------------------------------------

    For c = 0 To ConNom1.Cols - 1

        If ConNom1.ColWidth(c) > 0 Then

            wsXL.Cells(limite + 2, c + 1).Value = _
                ConNom1.TextMatrix(limite + 1, c)

        End If

    Next c

    '------------------------------------------
    ' FORMATO
    '------------------------------------------

    With wsXL.Range( _
        wsXL.Cells(1, 1), _
        wsXL.Cells(1, ConNom1.Cols))

        .Font.Bold = True

    End With

    With wsXL.Range( _
        wsXL.Cells(limite + 2, 1), _
        wsXL.Cells(limite + 2, ConNom1.Cols))

        .Font.Bold = True

    End With

    wsXL.Columns.AutoFit

    wsXL.Activate

    Set wsXL = Nothing
    Set wbXL = Nothing
    Set objXL = Nothing

    Exit Sub

ErrorExcel:

    On Error Resume Next

    If Not objXL Is Nothing Then
        objXL.Visible = True
    End If

    MsgBox "No fue posible exportar la nómina a Excel." & vbCrLf & _
           "Error: " & Err.Description, _
           vbCritical, "Exportar a Excel"

End Sub

Private Sub mnuIDAsc_Click()

    OrdenarConNom1 COL_ID, 3
    
End Sub

Private Sub mnuIDDesc_Click()

    OrdenarConNom1 COL_ID, 4

End Sub

Private Sub mnuNombreAsc_Click()
    
    OrdenarConNom1 COL_NOMBRE, 5
    
End Sub

Private Sub mnuNombreDesc_Click()

     OrdenarConNom1 COL_NOMBRE, 6

End Sub

Private Sub TxtBuscar_Change()

    Dim i As Long
    Dim filtro As String
    Dim nombre As String
    Dim ultimaFila As Long

    'No filtrar mientras se está cargando otra nómina
    If cargandoNomina Then Exit Sub

    'El grid debe tener al menos encabezado + una fila
    If ConNom1.Rows <= 1 Then Exit Sub

    filtro = UCase$(Trim$(TxtBuscar.Text))
    textoFiltro = filtro

    'Nunca permitir recorrer más filas de las que existen
    ultimaFila = limite

    If ultimaFila > ConNom1.Rows - 2 Then
        ultimaFila = ConNom1.Rows - 2
    End If

    If ultimaFila < 1 Then Exit Sub

    For i = FILA_INICIO_DATOS To ultimaFila

        nombre = Trim$(ConNom1.TextMatrix(i, COL_NOMBRE))

        If nombre = "" Then

            ConNom1.RowHeight(i) = 0

        ElseIf filtro = "" Then

            ConNom1.RowHeight(i) = -1

        ElseIf InStr(1, UCase$(nombre), filtro, vbTextCompare) > 0 Then

            ConNom1.RowHeight(i) = -1

        Else

            ConNom1.RowHeight(i) = 0

        End If

    Next i

    'Fila de totales
    If limite + 1 >= 0 And limite + 1 < ConNom1.Rows Then
        ConNom1.RowHeight(limite + 1) = -1
    End If

End Sub

Private Sub ConNom1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyC And (Shift And vbCtrlMask) <> 0 Then

        CmdCopySel_Click
        KeyCode = 0

    End If

End Sub

Private Sub CerrarArchivosVisor()

    On Error Resume Next

    Close #2
    Close #6
    Close #8
    Close #12
    Close #14

    On Error GoTo 0

End Sub

' comentario

Private Sub AbrirArchivosParaCFDI()

    Dim rutaCFDI As String

    On Error GoTo ErrorAbrir

    'La ruta de la nómina se obtiene a partir de Arch
    rutaCFDI = Left$(Arch, InStrRev(Arch, "\"))

    'Por seguridad, cerrar primero estos números de archivo
    On Error Resume Next
    Close #2
    Close #6
    Close #8
    Close #12
    Close #14
    On Error GoTo ErrorAbrir

    '========================================================
    ' ARCHIVOS QUE NECESITA EL PROCESO CFDI
    '========================================================
    
    'Personal
    Open rutaCFDI & "personal.dno" _
        For Random As #2 _
        Len = Len(personal)

    'Nómina que está visualizando FormViewer
    Open Arch _
        For Random As #6 _
        Len = Len(nomina)

    'Maestro
    Open rutaCFDI & "maestro.dno" _
        For Random As #8 _
        Len = Len(maestro)

    'Bnxcla
    Open rutaCFDI & "bnxcla.dno" _
        For Random As #12 _
        Len = Len(Clbnx)

    'Complemento .cmp de la nómina
    Open Arch1 _
        For Random As #14 _
        Len = Len(nom_com)

    Exit Sub

ErrorAbrir:

    MsgBox "No se pudieron abrir los archivos necesarios para generar el CFDI." & vbCrLf & _
           "Error: " & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "CFDI"

End Sub
Private Sub CerrarArchivosCFDI()

    On Error Resume Next

    Close #2
    Close #6
    Close #8
    Close #12
    Close #14

    On Error GoTo 0

End Sub

Private Sub OrdenarConNom1(ByVal columna As Long, ByVal tipoOrden As Integer)

    Dim filaActual As Long
    Dim colActual As Long
    Dim filasOriginales As Long

    On Error GoTo ErrorOrden

    If limite <= FILA_INICIO_DATOS Then Exit Sub

    filaActual = ConNom1.Row
    colActual = ConNom1.col

    'Guardar cantidad original de filas
    filasOriginales = ConNom1.Rows

    ConNom1.Redraw = False

    '========================================================
    ' QUITAR TEMPORALMENTE LA FILA DE TOTALES
    '========================================================

    'La fila de totales está en limite + 1
    ConNom1.Rows = limite + 1

    '========================================================
    ' ORDENAR SOLAMENTE LOS EMPLEADOS
    '========================================================

    ConNom1.Row = FILA_INICIO_DATOS
    ConNom1.RowSel = limite

    ConNom1.col = columna
    ConNom1.ColSel = columna

    ConNom1.Sort = tipoOrden

    '========================================================
    ' VOLVER A CREAR LA FILA DE TOTALES
    '========================================================

    ConNom1.Rows = limite + 2

    'Recalcular los totales y volver a colocar la fila
    sumavert

    'Restaurar posición
    If filaActual < ConNom1.Rows Then
        ConNom1.Row = filaActual
    Else
        ConNom1.Row = ConNom1.Rows - 1
    End If

    If colActual < ConNom1.Cols Then
        ConNom1.col = colActual
    Else
        ConNom1.col = 0
    End If

    ConNom1.Redraw = True

    Exit Sub

ErrorOrden:

    'Intentar recuperar la fila de totales
    On Error Resume Next

    If ConNom1.Rows < limite + 2 Then
        ConNom1.Rows = limite + 2
    End If

    sumavert

    ConNom1.Redraw = True

    MsgBox "Error al ordenar la nómina." & vbCrLf & _
           "Error: " & Err.Number & vbCrLf & _
           Err.Description, _
           vbExclamation, "Ordenamiento"

End Sub

Private Sub Encab(ByVal c As Long, ByVal txt As String, ByVal ancho As Long)
    ConNom1.col = c
    ConNom1.CellAlignment = 4
    ConNom1.ColWidth(c) = ancho
    ConNom1.Text = txt
End Sub
