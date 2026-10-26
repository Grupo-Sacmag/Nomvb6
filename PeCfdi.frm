VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PeCfdi 
   Caption         =   "Datos complementarios de Personal"
   ClientHeight    =   5940
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9555
   LinkTopic       =   "Form9"
   ScaleHeight     =   5940
   ScaleWidth      =   9555
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

'True: en empresas con sincronizacion SQL, el grid solo muestra a los empleados
'que ya tienen algun dato en Perscfdi.dno. False: muestra a todos los activos.
'(En empresas en modo local siempre se muestran todos, porque ahi se capturan a mano.)
Private Const OCULTAR_SIN_DATOS As Boolean = False

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

    Col_Def

    If EsEmpresaLocal(emp) Then
        PeCfdi.Caption = "Datos complementarios de Personal - Modo local"
    Else
        PeCfdi.Caption = "Datos complementarios de Personal - Cargando archivos locales"
    End If

End Sub

'Importacion automatica la primera vez que se muestra el formulario.
'El boton DB queda disponible para sincronizar manualmente otra vez.
Private Sub Form_Activate()
    If mSincronizacionInicial Then Exit Sub
    mSincronizacionInicial = True
    DB_Click
End Sub

Private Function ObtenerEmpresaSAT() As String
    Dim abrEmpresa As String

    abrEmpresa = Left(Trim(emp), 4)

    If (UCase(abrEmpresa) = "SACM") Then abrEmpresa = "SACMAG"
    If (UCase(abrEmpresa) = "SUPT") Then abrEmpresa = "SUPTER"
    If (UCase(abrEmpresa) = "COOR") Then abrEmpresa = "CORDINA"
    If (UCase(abrEmpresa) = "EPES") Then abrEmpresa = "EPESA"
    If (UCase(abrEmpresa) = "SUPE") Then abrEmpresa = "SUPERVISA"
    If (UCase(abrEmpresa) = "CONS") Then abrEmpresa = "CONSULTE"
    If (UCase(abrEmpresa) = "GEOA") Then abrEmpresa = "GEOAMBIENTE"

    ObtenerEmpresaSAT = abrEmpresa
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
'Lo usan el boton Abrir y, al terminar, el boton DB, para que el grid
'siempre muestre lo que realmente quedo guardado en el archivo.
Private Sub CargarGridDesdeArchivos()
    Dim hPers As Integer
    Dim hCfdi As Integer
    Dim totalPers As Long
    Dim totalCfdi As Long
    Dim i As Long
    Dim nombre As String
    Dim cadena As String

    On Error GoTo ErrCarga

    PCpCfdi.Rows = 1

    If Dir$("Personal.dno") = "" Then
        Err.Raise vbObjectError + 1101, , "No existe Personal.dno en la carpeta actual."
    End If

    hPers = FreeFile
    Open "Personal.dno" For Random Access Read As #hPers Len = Len(personal)
    totalPers = LOF(hPers) \ Len(personal)

    'Si Perscfdi.dno aun no existe, se muestran los empleados sin datos complementarios.
    If Dir$("Perscfdi.dno") <> "" Then
        hCfdi = FreeFile
        Open "Perscfdi.dno" For Random Access Read As #hCfdi Len = Len(Empleado_1)
        totalCfdi = LOF(hCfdi) \ Len(Empleado_1)
    End If

    For i = 1 To totalPers
        Get #hPers, i, personal

        'Los empleados de baja no se muestran en el grid.
        If CampoArchivoCFDI(personal.nom) <> "" And Not EmpleadoDeBaja() Then

            'Los empleados mas nuevos pueden quedar despues del final de Perscfdi.dno.
            'Get no marca error en ese caso y dejaria los datos del empleado anterior.
            If i <= totalCfdi Then
                Get #hCfdi, i, Empleado_1
            Else
                LimpiarEmpleadoCFDI
            End If

            'Un registro sin ningun dato no se muestra (salvo en modo local o si
            'OCULTAR_SIN_DATOS = False).
            If TieneDatosCFDI() Or EsEmpresaLocal(emp) Or Not OCULTAR_SIN_DATOS Then

                nombre = Trim$(personal.nom) & " " & Trim$(personal.ape1) & " " & Trim$(personal.ape2)

                cadena = Format$(i, "####0") & vbTab & nombre & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Direccion) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Colonia) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Ciudad) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Estado) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Delegacion) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.Cpostal) & vbTab & _
                         CampoArchivoCFDI(Empleado_1.correo)

                PCpCfdi.AddItem cadena

            End If
        End If
    Next i

Salir:
    On Error Resume Next
    If hCfdi <> 0 Then Close #hCfdi
    If hPers <> 0 Then Close #hPers
    Exit Sub

ErrCarga:
    MsgBox "No fue posible cargar los datos." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Datos complementarios"
    Resume Salir
End Sub

Private Sub ArCfAbr_Click()
    CargarGridDesdeArchivos
End Sub

Private Sub ArCfGuarda_Click()
    Dim h As Integer
    Dim EmpNum As Long
    Dim renglon As Long

    h = FreeFile
    Open "Perscfdi.dno" For Random As #h Len = Len(Empleado_1)

    For renglon = 1 To PCpCfdi.Rows - 1
        If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
            EmpNum = CLng(PCpCfdi.TextMatrix(renglon, 0))

            If EmpNum >= 1 Then
                'Leer antes de modificar: conserva cualquier otro campo del tipo
                'y evita arrastrar valores del renglon anterior.
                Get #h, EmpNum, Empleado_1

                Empleado_1.Direccion = PCpCfdi.TextMatrix(renglon, 2)
                Empleado_1.Colonia = PCpCfdi.TextMatrix(renglon, 3)
                Empleado_1.Ciudad = PCpCfdi.TextMatrix(renglon, 4)
                Empleado_1.Estado = PCpCfdi.TextMatrix(renglon, 5)
                Empleado_1.Delegacion = PCpCfdi.TextMatrix(renglon, 6)
                Empleado_1.Cpostal = PCpCfdi.TextMatrix(renglon, 7)
                Empleado_1.correo = PCpCfdi.TextMatrix(renglon, 8)

                Put #h, EmpNum, Empleado_1
            End If
        End If
    Next renglon

    Close #h
End Sub

Private Sub ArCfSale_Click()
    Close
    Unload PeCfdi
End Sub

' ============================================================
' CONSULTA REMOTA
' ============================================================

Private Sub DB_Click()
    Dim nuevos As Long
    Dim completados As Long
    Dim omitidos As Long

    On Error GoTo ErrorDB

    'La importacion SQL -> DNO existe solo para empresas conectadas.
    If Not EsEmpresaLocal(emp) Then
        ImportarPersonalDesdeSQL ObtenerEmpresaSAT(), nuevos, completados, omitidos
    End If

    'Mostrar SIEMPRE lo que quedo en archivos locales.
    Col_Def
    CargarGridDesdeArchivos

    Debug.Print "PeCfdi: " & nuevos & " altas locales; " & _
                completados & " empleados completados; " & _
                omitidos & " omitidos."

    If EsEmpresaLocal(emp) Then
        PeCfdi.Caption = "Datos complementarios - Modo local"
    Else
        PeCfdi.Caption = "Datos complementarios - Archivos locales"
    End If
    Exit Sub

ErrorDB:
    Debug.Print "Error importando SQL: " & Err.Number & " / " & Err.Description
    MsgBox "No se pudo completar la importacion local." & vbCrLf & _
           Err.Description, vbExclamation, "Datos complementarios"

    'Mostrar los datos locales previos incluso si fallo la conexion.
    Col_Def
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

Private Sub EdCfEliminar_Click()
    Dim w1 As Long
    For w1 = PCpCfdi.Row To PCpCfdi.RowSel
        PCpCfdi.RemoveItem PCpCfdi.Row
    Next w1
End Sub

Private Sub EdCfInsertar_Click()
    Dim W As Long

    For W = PCpCfdi.Row To PCpCfdi.RowSel
        PCpCfdi.AddItem (""), PCpCfdi.Row
    Next W
End Sub

Private Sub EdCfPegar_Click()
    Dim DeAqui As Long
    Dim RetornoCarro As Long
    Dim InicioCopia As Long
    Dim i As Long
    Dim InIPGr As Long
    Dim InIPGc As Long

    Temporal1 = ""
    temporal = Clipboard.GetText(vbCFText)

    RetornoCarro = PCpCfdi.col
    InicioCopia = PCpCfdi.Row
    DeAqui = InicioCopia
    InIPGr = InicioCopia
    InIPGc = RetornoCarro

    If temporal <> "" Then
        Clipboard.Clear
        DeAqui = 1

        For i = 1 To Len(temporal)
            Select Case mID$(temporal, i, 1)
                Case Chr(9)
                    LTXT.Text = mID$(temporal, DeAqui, (i - DeAqui))
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.col) & Chr(9)
                    PCpCfdi.Text = mID$(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.col = PCpCfdi.col + 1
                    DeAqui = i + 1

                Case Chr(13)
                    Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.col) & Chr(13)
                    LTXT.Text = mID$(temporal, DeAqui, (i - DeAqui))
                    PCpCfdi.Text = mID$(temporal, DeAqui, (i - DeAqui))

                    If (PCpCfdi.Rows - 1) <= PCpCfdi.Row Then PCpCfdi.Rows = PCpCfdi.Rows + 1

                    PCpCfdi.Row = PCpCfdi.Row + 1
                    PCpCfdi.TopRow = 1
                    DeAqui = i + 1

                Case Chr(10)
                    Temporal1 = Temporal1 & Chr(10)
                    PCpCfdi.col = RetornoCarro
                    DeAqui = i + 1

                Case Else
                    Rem nada
            End Select
        Next i

        PCpCfdi.col = RetornoCarro
    End If

    Clipboard.Clear
    Clipboard.SetText Clipboard.GetText + Temporal1
End Sub

Private Sub EdCfSelT_Click()
    Dim limite As Long
    Dim renglon As Long

    Clipboard.Clear

    PCpCfdi.Row = 1
    PCpCfdi.col = 0

    For limite = 1 To PCpCfdi.Rows - 1
        renglon = limite

        If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
            PCpCfdi.RowSel = PCpCfdi.Row
        End If
    Next limite

    PCpCfdi.ColSel = PCpCfdi.Cols - 1
End Sub

' ============================================================
' EDICION DE CELDAS
' ============================================================

Private Sub LTXT_Change()
    PCpCfdi.Text = LTXT.Text
End Sub

Private Sub pcpcfdi_EnterCell()
    PCpCfdi.CellBackColor = vbYellow
    valcelant = PCpCfdi.Text
    LTXT.Text = valcelant
End Sub

Private Sub pcpcfdi_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim q As Long
    Dim W As Long

    Select Case KeyCode
        Case vbKeyDelete
            For q = PCpCfdi.Row To PCpCfdi.RowSel
                For W = PCpCfdi.col To PCpCfdi.ColSel
                    PCpCfdi.TextMatrix(q, W) = ""
                Next W
            Next q

            LTXT.Text = PCpCfdi.Text

        Case vbKeyF2
            If PCpCfdi.Text <> "" Then valcelant = PCpCfdi.Text

            LTXT.Text = LTrim(RTrim(PCpCfdi.Text))
            LTXT.SetFocus
    End Select
End Sub

Private Sub pcpcfdi_KeyPress(KeyAscii As Integer)
    valcelant = PCpCfdi.Text
    LTXT.Text = Chr(KeyAscii)
    LTXT.SetFocus
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



