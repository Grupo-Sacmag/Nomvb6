VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PeCfdi 
   Caption         =   "Datos complementarios de Personal"
   ClientHeight    =   5940
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   13140
   LinkTopic       =   "Form9"
   ScaleHeight     =   5940
   ScaleWidth      =   13140
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid dat 
      Height          =   2175
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   3836
      _Version        =   393216
   End
   Begin VB.TextBox LTXT 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7092
   End
   Begin MSFlexGridLib.MSFlexGrid PCpCfdi 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.Menu ArCf 
      Caption         =   "&Archivo"
      Begin VB.Menu DB 
         Caption         =   "&DB"
      End
      Begin VB.Menu ArCfAbr 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu ArCfGuarda 
         Caption         =   "&Guardar"
      End
      Begin VB.Menu ArCfsep1 
         Caption         =   "-"
      End
      Begin VB.Menu ArCfSale 
         Caption         =   "&Salida"
      End
   End
   Begin VB.Menu EdCf 
      Caption         =   "&Editar"
      Begin VB.Menu EdCfPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdCfCopia 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu EdCfSep1 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfInsertar 
         Caption         =   "&Insertar"
      End
      Begin VB.Menu EdCfEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu EdCfSep2 
         Caption         =   "-"
      End
      Begin VB.Menu EdCfSelT 
         Caption         =   "&Seleccionar Todo"
      End
   End
End
Attribute VB_Name = "PeCfdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Nueva Versión

Private temporal As String
Private Temporal1 As String
Private valcelant As String
Private mSincronizacionInicial As Boolean
Private mEditandoCFDI As Boolean
Private mLeyendoGrid As Boolean
Private mFilasCargadas As Long
Private mRutaCarga As String
Private mIdsFila() As Long
Private mNombresFila() As String
Private mIdentidadesFila() As String
Private mOriginalCFDI() As String


'True: en empresas con sincronizacion SQL, el grid solo muestra a los empleados
'que ya tienen algun dato en Perscfdi.dno. False: muestra a todos los activos.
'(En empresas en modo local siempre se muestran todos, porque ahi se capturan a mano.)
Private Const OCULTAR_SIN_DATOS As Boolean = False


'============================================================
' VALIDAR EMPRESA AUTORIZADA PARA SQL
'============================================================

Private Function EsSACMAGExacta() As Boolean
    'Autorizacion central: empresa.name EXACTAMENTE como razon social.
    'Ver modSincronizacion.bas. El codigo SQL de datosSat sigue siendo SACMAG.
    EsSACMAGExacta = SincronizacionRemotaPermitida()
End Function

' ============================================================
' CONFIGURACION DEL GRID
' ============================================================

Sub Col_Def()
    PCpCfdi.Clear
    PCpCfdi.Rows = 200
    PCpCfdi.Cols = 9
    PCpCfdi.FixedCols = 0
    PCpCfdi.FixedRows = 1
    PCpCfdi.Row = 0

    PCpCfdi.col = 0
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(0) = 800
    PCpCfdi.Text = "Num"

    PCpCfdi.col = 1
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(1) = 3600
    PCpCfdi.Text = "Nombre"

    PCpCfdi.col = 2
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(2) = 2400
    PCpCfdi.Text = "Direccion"

    PCpCfdi.col = 3
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(3) = 2400
    PCpCfdi.Text = "Colonia"

    PCpCfdi.col = 4
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(4) = 2400
    PCpCfdi.Text = "Ciudad"

    PCpCfdi.col = 5
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(5) = 2400
    PCpCfdi.Text = "Estado"

    PCpCfdi.col = 6
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(6) = 2400
    PCpCfdi.Text = "Delegacion"

    PCpCfdi.col = 7
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(7) = 800
    PCpCfdi.Text = "Codigo"

    PCpCfdi.col = 8
    PCpCfdi.CellAlignment = 4
    PCpCfdi.ColWidth(8) = 2400
    PCpCfdi.Text = "Correo "
End Sub

Private Sub Form_Resize()
    PCpCfdi.Height = PeCfdi.Height * 0.85
    PCpCfdi.Width = PeCfdi.Width * 0.95
End Sub

' ============================================================
' CARGA DEL FORMULARIO / SINCRONIZACION INICIAL
' ============================================================


Private Sub Form_Load()

    mEditandoCFDI = False
    mSincronizacionInicial = False
    mLeyendoGrid = False
    mFilasCargadas = 0

    'Preparar solamente los encabezados.
    Col_Def

    'No mostrar empleados hasta presionar Abrir o DB.
    PCpCfdi.Rows = 1

    ActualizarModoCFDI

    If EsSACMAGExacta() Then
        Me.Caption = "Datos complementarios - SACMAG"
    Else
        Me.Caption = "Datos complementarios - Modo local"
    End If

End Sub

Private Function ObtenerEmpresaSAT() As String

    ObtenerEmpresaSAT = vbNullString

    If EsSACMAGExacta() Then
        ObtenerEmpresaSAT = "SACMAG"
    End If

End Function


' ============================================================
' ARCHIVO LOCAL CFDI
' ============================================================

Private Sub LimpiarEmpleadoCFDI()
    Empleado_1.Direccion = ""
    Empleado_1.Colonia = ""
    Empleado_1.Ciudad = ""
    Empleado_1.Estado = ""
    Empleado_1.Delegacion = ""
    Empleado_1.Cpostal = ""
    Empleado_1.correo = ""
End Sub

'Quita espacios y los caracteres nulos (Chr(0)) que quedan en registros
'que existen en el archivo pero nunca se han escrito.
Private Function CampoArchivoCFDI(ByVal valor As String) As String
    CampoArchivoCFDI = Trim$(Replace(valor, vbNullChar, ""))
End Function

'True si el registro cargado en el buffer Empleado_1 tiene algun dato guardado.
Private Function TieneDatosCFDI() As Boolean
    TieneDatosCFDI = (CampoArchivoCFDI(Empleado_1.Direccion) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.Colonia) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.Ciudad) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.Estado) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.Delegacion) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.Cpostal) <> "") Or _
                     (CampoArchivoCFDI(Empleado_1.correo) <> "")
End Function

'True si el empleado cargado en el buffer "personal" tiene fecha de baja.
'Siguen existiendo en Personal.dno, pero no se muestran ni se sincronizan.
Private Function EmpleadoDeBaja() As Boolean
    EmpleadoDeBaja = IsDate(Trim$(Replace(personal.fab, vbNullChar, "")))
End Function

'Llena el grid leyendo Personal.dno y Perscfdi.dno.
'Lo usan el boton Abrir y, al terminar, el boton DB, para que el grid siempre muestre lo que realmente quedo guardado en el archivo.
Private Sub CargarGridDesdeArchivos()
    Dim hPers As Integer
    Dim hCfdi As Integer
    Dim totalPers As Long
    Dim totalCfdi As Long
    Dim i As Long
    Dim nombre As String
    Dim cadena As String
    Dim modoLocal As Boolean

    On Error GoTo ErrCarga
    If Not RutaTrabajoSegura() Then Exit Sub

    mLeyendoGrid = True
    mEditandoCFDI = False
    mFilasCargadas = 0
    mRutaCarga = CurDir$
    modoLocal = Not EsSACMAGExacta()
    PCpCfdi.Rows = 1

    If Dir$("personal.dno") = "" Then
        Err.Raise vbObjectError + 1101, , "No existe personal.dno en la carpeta actual."
    End If

    hPers = FreeFile
    Open "personal.dno" For Random Access Read As #hPers Len = Len(personal)
    If LOF(hPers) Mod Len(personal) <> 0 Then
        Err.Raise vbObjectError + 1102, , "personal.dno tiene estructura incompatible."
    End If
    totalPers = LOF(hPers) \ Len(personal)

    If totalPers > 0 Then
        ReDim mIdsFila(1 To totalPers)
        ReDim mNombresFila(1 To totalPers)
        ReDim mIdentidadesFila(1 To totalPers)
        ReDim mOriginalCFDI(1 To totalPers)
    End If

    If Dir$("Perscfdi.dno") <> "" Then
        hCfdi = FreeFile
        Open "Perscfdi.dno" For Random Access Read As #hCfdi Len = Len(Empleado_1)
        If LOF(hCfdi) Mod Len(Empleado_1) <> 0 Then
            Err.Raise vbObjectError + 1103, , "Perscfdi.dno tiene estructura incompatible."
        End If
        totalCfdi = LOF(hCfdi) \ Len(Empleado_1)
    End If

    For i = 1 To totalPers
        Get #hPers, i, personal
        If CampoArchivoCFDI(personal.nom) <> "" And Not EmpleadoDeBaja() Then
            LimpiarEmpleadoCFDI
            If i <= totalCfdi Then Get #hCfdi, i, Empleado_1

            If TieneDatosCFDI() Or modoLocal Or Not OCULTAR_SIN_DATOS Then
                nombre = Trim$(personal.nom) & " " & _
                         Trim$(personal.ape1) & " " & Trim$(personal.ape2)
                cadena = Format$(i, "####0") & vbTab & nombre & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Direccion) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Colonia) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Ciudad) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Estado) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Delegacion) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Cpostal) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.correo)
                PCpCfdi.AddItem cadena
                mFilasCargadas = mFilasCargadas + 1
                mIdsFila(mFilasCargadas) = i
                mNombresFila(mFilasCargadas) = nombre
                mIdentidadesFila(mFilasCargadas) = FirmaIdentidadCFDI()
                mOriginalCFDI(mFilasCargadas) = FirmaRegistroCFDI()
            End If
        End If
    Next i

Salir:
    On Error Resume Next
    If hCfdi <> 0 Then Close #hCfdi
    If hPers <> 0 Then Close #hPers
    mLeyendoGrid = False
    ActualizarModoCFDI
    On Error GoTo 0
    Exit Sub

ErrCarga:
    MsgBox "No fue posible cargar datos CFDI: " & Err.Description, _
           vbExclamation, "Datos complementarios"
    Resume Salir
End Sub




Private Sub ArCfAbr_Click()

    If Not RutaTrabajoSegura() Then Exit Sub

    MsgBox "Se mostraran los registros actualmente guardados en los archivos locales (.dno)." & vbCrLf & vbCrLf & "Esta consulta es de solo lectura.", vbInformation, "Archivos locales"

    'No consulta SQL ni escribe archivos.
    CargarGridDesdeArchivos

End Sub

Private Sub DB_Click()

    Dim nuevos As Long
    Dim completados As Long
    Dim omitidos As Long
    Dim detalleError As String

    On Error GoTo ErrorDB

    If Not RutaTrabajoSegura() Then Exit Sub

    '====================================================
    ' 1. AUTORIZACION DE EMPRESA
    '====================================================

    If Not EsSACMAGExacta() Then

        MsgBox "La sincronizacion SQL esta disponible unicamente para SACMAG DE MEXICO, S.A. DE C.V." & vbCrLf & "No se consulto la base de datos.", vbInformation, "Sincronizacion no disponible"

        Exit Sub

    End If

    '====================================================
    ' 2. AVISO INFORMATIVO
    '====================================================

    MsgBox "Se consultara la base de datos de SACMAG para completar los datos faltantes de los archivos locales (.dno)." & vbCrLf & vbCrLf & "Los datos locales existentes se conservaran." & vbCrLf & _
           "Los empleados nuevos se incorporaran en su posicion correspondiente.", vbInformation, "Sincronizacion de personal"

    '====================================================
    ' 3. CONSULTAR SQL
    '====================================================

    'Solo lectura SQL.
    'La rutina importadora escribe unicamente en .dno.

    ImportarPersonalDesdeSQL "SACMAG", nuevos, completados, omitidos

    '====================================================
    ' 4. MOSTRAR RESULTADO DESDE ARCHIVOS LOCALES
    '====================================================

    CargarGridDesdeArchivos

    Me.Caption = "Datos complementarios - SACMAG"

    '====================================================
    ' 5. NOTIFICAR EMPLEADOS NUEVOS
    '====================================================

    If nuevos > 0 Then

        MsgBox "Se encontraron " & CStr(nuevos) & " empleados nuevos en SQL." & vbCrLf & "Se incorporaron a los archivos locales para su consulta y manipulacion desde los formularios correspondientes.", vbInformation, "Personal importado"

    End If

    Debug.Print "SACMAG - Nuevos: " & nuevos & " / Completados: " & completados & " / Omitidos: " & omitidos

    Exit Sub

ErrorDB:

    detalleError = Err.Description

    MsgBox "No fue posible completar la sincronizacion." & vbCrLf & detalleError, vbExclamation, "Datos complementarios"

    'Permitir visualizar los archivos anteriores,
    'sin realizar una segunda importacion.
    CargarGridDesdeArchivos

End Sub



' ============================================================
' COPIAR / PEGAR / INSERTAR / ELIMINAR
' ============================================================

Private Sub EdCfCopia_Click()
    Dim Temporal1 As String
    Dim i As Long
    Dim f As Long

    Clipboard.Clear

    For i = PCpCfdi.Row To PCpCfdi.RowSel
        For f = PCpCfdi.col To PCpCfdi.ColSel
            Temporal1 = Temporal1 + PCpCfdi.TextMatrix(i, f)
            If f < PCpCfdi.ColSel Then
                Temporal1 = Temporal1 & Chr(9)
            End If
        Next f
        Temporal1 = Temporal1 & Chr(13) & Chr(10)
    Next i
    Clipboard.SetText Temporal1
End Sub


Private Sub EdCfSelT_Click()

    If PCpCfdi.Rows <= 1 Then Exit Sub

    PCpCfdi.Row = 1
    PCpCfdi.col = 0

    PCpCfdi.RowSel = PCpCfdi.Rows - 1
    PCpCfdi.ColSel = PCpCfdi.Cols - 1

End Sub

Private Sub pcpcfdi_EnterCell()
    'Resaltar visualmente la celda seleccionada.
    'No activar cuadros de edicion.
    PCpCfdi.CellBackColor = vbYellow
End Sub


Private Sub pcpcfdi_KeyDown(KeyCode As Integer, Shift As Integer)

    'Permitir copiar la seleccion con Ctrl + C.
    If (Shift And vbCtrlMask) <> 0 Then

        If KeyCode = vbKeyC Then
            EdCfCopia_Click
            KeyCode = 0
            Exit Sub
        End If

        'Bloquear Ctrl + V.
        If KeyCode = vbKeyV Then
            KeyCode = 0
            Exit Sub
        End If

    End If

    Select Case KeyCode

        Case vbKeyDelete, vbKeyBack, vbKeyF2
            'No permitir borrar ni editar.
            KeyCode = 0

        Case vbKeyReturn
            'Enter no edita ni muestra mensajes.
            KeyCode = 0

    End Select

End Sub

Private Sub pcpcfdi_KeyPress(KeyAscii As Integer)

    'Impedir escribir directamente en las celdas.
    'La seleccion y la navegacion siguen disponibles.
    KeyAscii = 0

End Sub


Private Sub pcpcfdi_LeaveCell()
    PCpCfdi.CellBackColor = vbWhite
End Sub

' ============================================================
' FUNCIONES AUXILIARES DE LA SINCRONIZACION
' ============================================================

Private Function CampoSQLCFDI(ByVal valor As Variant) As String

    If IsNull(valor) Or IsEmpty(valor) Then
        CampoSQLCFDI = ""
    Else
        CampoSQLCFDI = Trim$(CStr(valor))
    End If

End Function

Private Function NombreComparableCFDI(ByVal texto As String) As String

    texto = UCase$(Trim$(texto))

    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ü", "U")

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    'Un apellido "N/A" en SQL equivale a un apellido vacio en el archivo local.
    If texto = "N/A" Then texto = ""

    NombreComparableCFDI = texto

End Function

'Devuelve el idNomina como numero (1 a maximo), o 0 si no es valido.
Private Function ObtenerIdCFDI(ByVal idTexto As String, ByVal maximo As Long) As Long

    ObtenerIdCFDI = 0

    If idTexto = "" Then Exit Function
    If idTexto Like "*[!0-9]*" Then Exit Function
    If Len(idTexto) > 9 Then Exit Function

    If CLng(idTexto) < 1 Or CLng(idTexto) > maximo Then Exit Function

    ObtenerIdCFDI = CLng(idTexto)

End Function

'Devuelve "" si el valor cabe y se puede guardar sin perder caracteres;
'si no, devuelve el motivo.
Private Function RevisarCampoCFDI(ByVal etiqueta As String, ByVal valor As String, ByVal largoCampo As Long) As String

    If Len(valor) > largoCampo Then
        RevisarCampoCFDI = etiqueta & " excede " & largoCampo & " caracteres"
    ElseIf StrConv(StrConv(valor, vbFromUnicode), vbUnicode) <> valor Then
        RevisarCampoCFDI = etiqueta & " tiene caracteres que no se pueden guardar"
    Else
        RevisarCampoCFDI = ""
    End If

End Function

'True si SQL trae vacio CUALQUIERA de los 7 datos del registro actual.
'Si algun dato puede quedar vacio legitimamente (por ejemplo la colonia),
'basta con quitar su linea de aqui.
Private Function SQLIncompleto(ByVal oRS As ADODB.Recordset) As Boolean

    SQLIncompleto = True

    If CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreColonia").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("entidadFederativa").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("cp").Value) = "" Then Exit Function
    If CampoSQLCFDI(oRS.Fields("correoElectronico").Value) = "" Then Exit Function

    SQLIncompleto = False

End Function

'True si Perscfdi.dno ya tiene algun dato guardado en la posicion id.
'Solo lee; abre y cierra el archivo por su cuenta.
'Si no puede leerlo responde True, para no arriesgarse a sobrescribir.
Private Function LocalCFDITieneDatos(ByVal ruta As String, ByVal id As Long) As Boolean
    Dim h As Integer
    Dim total As Long

    LocalCFDITieneDatos = False

    If Dir$(ruta) = "" Then Exit Function

    On Error GoTo Fallo

    h = FreeFile
    Open ruta For Random Access Read As #h Len = Len(Empleado_1)
    total = LOF(h) \ Len(Empleado_1)

    'Mas alla del final del archivo no hay nada guardado todavia.
    If id <= total Then
        Get #h, id, Empleado_1

        LocalCFDITieneDatos = TieneDatosCFDI()
    End If

Salir:
    On Error Resume Next
    If h <> 0 Then Close #h
    Exit Function

Fallo:
    LocalCFDITieneDatos = True
    Resume Salir
End Function

Private Function VerificarRegistroCFDI(ByVal oRS As ADODB.Recordset) As Boolean

    VerificarRegistroCFDI = False

    If Trim$(Empleado_1.Direccion) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreDeVialidad").Value)) Then Exit Function
    If Trim$(Empleado_1.Colonia) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreColonia").Value)) Then Exit Function
    If Trim$(Empleado_1.Ciudad) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreLocalidad").Value)) Then Exit Function
    If Trim$(Empleado_1.Estado) <> UCase$(CampoSQLCFDI(oRS.Fields("entidadFederativa").Value)) Then Exit Function
    If Trim$(Empleado_1.Delegacion) <> UCase$(CampoSQLCFDI(oRS.Fields("nombreMunicipio").Value)) Then Exit Function
    If Trim$(Empleado_1.Cpostal) <> UCase$(CampoSQLCFDI(oRS.Fields("cp").Value)) Then Exit Function
    If Trim$(Empleado_1.correo) <> UCase$(CampoSQLCFDI(oRS.Fields("correoElectronico").Value)) Then Exit Function

    VerificarRegistroCFDI = True

End Function

'============================================================
' EDICION EXPLICITA Y PROTECCION DE IDENTIFICADORES
'============================================================


Private Sub ActualizarModoCFDI()

    'PeCfdi es exclusivamente de consulta.
    'No permitir habilitar edicion.

    mEditandoCFDI = False

    LTXT.Locked = True
    LTXT.Visible = False

End Sub

Private Function PuedeEditarCeldaCFDI(Optional ByVal aviso As Boolean = True) As Boolean
    'PeCfdi es exclusivamente de consulta.
    'No mostrar mensajes de habilitar edicion.
    PuedeEditarCeldaCFDI = False
End Function


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mEditandoCFDI Then
        If MsgBox("Hay cambios potenciales sin guardar. Desea cerrar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Cerrar CFDI") <> vbYes Then Cancel = 1
    End If
End Sub

Private Function FirmaCampoCFDI(ByVal valor As String) As String
    Dim limpio As String
    limpio = CampoArchivoCFDI(valor)
    FirmaCampoCFDI = CStr(Len(limpio)) & ":" & limpio
End Function

Private Function FirmaRegistroCFDI() As String
    FirmaRegistroCFDI = _
        FirmaCampoCFDI(Empleado_1.Direccion) & _
        FirmaCampoCFDI(Empleado_1.Colonia) & _
        FirmaCampoCFDI(Empleado_1.Ciudad) & _
        FirmaCampoCFDI(Empleado_1.Estado) & _
        FirmaCampoCFDI(Empleado_1.Delegacion) & _
        FirmaCampoCFDI(Empleado_1.Cpostal) & _
        FirmaCampoCFDI(Empleado_1.correo)
End Function

Private Function FirmaIdentidadCFDI() As String
    FirmaIdentidadCFDI = FirmaCampoCFDI(personal.nom) & FirmaCampoCFDI(personal.ape1) & FirmaCampoCFDI(personal.ape2) & FirmaCampoCFDI(personal.RFC)
End Function

Private Function FirmaGridCFDI(ByVal fila As Long) As String
    Dim c As Long
    Dim resultado As String
    For c = 2 To 8
        resultado = resultado & FirmaCampoCFDI(PCpCfdi.TextMatrix(fila, c))
    Next c
    FirmaGridCFDI = resultado
End Function

Private Sub ValidarValorCFDI(ByVal valor As String, ByVal ancho As Long, _
                              ByVal fila As Long, ByVal columna As Long)
    If Len(valor) > ancho Then
        Err.Raise vbObjectError + 1210, , "Fila " & fila & ", columna " & columna & ": " & "supera " & ancho & " caracteres."
    End If
    If StrConv(StrConv(valor, vbFromUnicode), vbUnicode) <> valor Then
        Err.Raise vbObjectError + 1211, , _
             "Fila " & fila & ": caracteres no compatibles con el .dno."
    End If
End Sub

Private Sub ValidarCamposGridCFDI(ByVal fila As Long)
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 2), Len(Empleado_1.Direccion), fila, 2
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 3), Len(Empleado_1.Colonia), fila, 3
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 4), Len(Empleado_1.Ciudad), fila, 4
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 5), Len(Empleado_1.Estado), fila, 5
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 6), Len(Empleado_1.Delegacion), fila, 6
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 7), Len(Empleado_1.Cpostal), fila, 7
    ValidarValorCFDI PCpCfdi.TextMatrix(fila, 8), Len(Empleado_1.correo), fila, 8
End Sub

Private Sub CopiarGridARegistroCFDI(ByVal fila As Long)
    Empleado_1.Direccion = PCpCfdi.TextMatrix(fila, 2)
    Empleado_1.Colonia = PCpCfdi.TextMatrix(fila, 3)
    Empleado_1.Ciudad = PCpCfdi.TextMatrix(fila, 4)
    Empleado_1.Estado = PCpCfdi.TextMatrix(fila, 5)
    Empleado_1.Delegacion = PCpCfdi.TextMatrix(fila, 6)
    Empleado_1.Cpostal = PCpCfdi.TextMatrix(fila, 7)
    Empleado_1.correo = PCpCfdi.TextMatrix(fila, 8)
End Sub


