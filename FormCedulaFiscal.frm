VERSION 5.00
Begin VB.Form FormCedulaFiscal 
   Appearance      =   0  'Flat
   Caption         =   "Captura de Cédula de Indentificación FIscal"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16305
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form10"
   ScaleHeight     =   7710
   ScaleWidth      =   16305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar"
      Height          =   555
      Left            =   8520
      TabIndex        =   66
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Limpiar"
      Height          =   555
      Left            =   11160
      TabIndex        =   65
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   555
      Left            =   13800
      TabIndex        =   64
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text34 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   62
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fiscales"
      Height          =   2895
      Left            =   8640
      TabIndex        =   5
      Top             =   3840
      Width           =   7455
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Top             =   2040
         Width           =   3735
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   32
         Top             =   1320
         Width           =   6855
      End
      Begin VB.TextBox Text27 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   6855
      End
      Begin VB.Label Label29 
         Caption         =   "Empresa"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label28 
         Caption         =   "Régimen"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Actividad Económica"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dirección"
      Height          =   2895
      Left            =   8640
      TabIndex        =   2
      Top             =   840
      Width           =   7455
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   25
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   24
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   21
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   20
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "Y Calle"
         Height          =   375
         Left            =   6240
         TabIndex        =   51
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Entre Calle"
         Height          =   255
         Left            =   4800
         TabIndex        =   50
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label19 
         Caption         =   "Entidad Federativa"
         Height          =   255
         Left            =   2520
         TabIndex        =   49
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Municipio"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Localidad"
         Height          =   255
         Left            =   4560
         TabIndex        =   47
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Colonia"
         Height          =   375
         Left            =   1440
         TabIndex        =   46
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Número Interior"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Número Exterior"
         Height          =   255
         Left            =   6120
         TabIndex        =   44
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Nombre de Vialidad"
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Vialidad"
         Height          =   255
         Left            =   1200
         TabIndex        =   42
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "CP"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Contacto"
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   8415
      Begin VB.TextBox Text26 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   30
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text25 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text24 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   28
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text23 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label26 
         Caption         =   "Estado Cont Domiciolio"
         Height          =   255
         Left            =   4080
         TabIndex        =   56
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label25 
         Caption         =   "Estado Domicilio"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "Teléfono"
         Height          =   255
         Left            =   5160
         TabIndex        =   54
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "LADA"
         Height          =   255
         Left            =   4080
         TabIndex        =   53
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "Email"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos personales"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   8415
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   60
         Top             =   2400
         Width           =   3495
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   13
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Top             =   1560
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   8
         Top             =   720
         Width           =   3735
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Nombre comercial"
         Height          =   255
         Left            =   4680
         TabIndex        =   61
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Ultimo Cambio"
         Height          =   255
         Left            =   2160
         TabIndex        =   40
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Estatus"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Inicio Operaciones"
         Height          =   375
         Left            =   5760
         TabIndex        =   38
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "A. Materno"
         Height          =   255
         Left            =   3120
         TabIndex        =   37
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "A. Paterno"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre (s)"
         Height          =   255
         Left            =   4440
         TabIndex        =   35
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "CURP"
         Height          =   255
         Left            =   1800
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "RFC"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label32 
      Caption         =   "Registra"
      Height          =   375
      Left            =   240
      TabIndex        =   63
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "IDNomina:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "FormCedulaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'============================================================
' FORM CEDULA FISCAL - CONSULTA SEGURA
'
' Los .dno son la fuente principal.
' SQL se consulta SOLO para campos SAT sin equivalencia .dno.
' NUNCA ejecuta INSERT, UPDATE ni DELETE en SQL.
' Esta version es SOLO CONSULTA: la edicion local se agregara
' una vez definida la persistencia de los campos SAT restantes.
'============================================================

Private mID As Long
Private mEmpresa As String
Private mPreparado As Boolean

'============================================================
' INICIALIZACION
'============================================================
Private Sub Form_Load()
    mPreparado = False
    BloquearCedula
    Command1.Visible = False     'Antes insertaba en SQL: deshabilitado
    Command1.Enabled = False
    Command3.Visible = False     'Antes actualizaba SQL: deshabilitado
    Command3.Enabled = False
    Command2.Caption = "Cerrar"
    Command2.Visible = True
End Sub

'Form2 debe llamar a esta funcion pasando el ID seleccionado.
Public Function PrepararEmpleado(ByVal idEmpleado As Long) As Boolean
    On Error GoTo ErrorPreparar

    PrepararEmpleado = False
    mPreparado = False

    If idEmpleado < 1 Or idEmpleado > 32767 Then
        Err.Raise vbObjectError + 301, , "ID de empleado fuera de rango."
    End If

    mID = idEmpleado
    mEmpresa = Trim$(ObtenerEmpresaSAT())
    If mEmpresa = "" Then
        Err.Raise vbObjectError + 302, , "No se pudo determinar la empresa actual."
    End If

    LimpiarCampos
    Text1.Text = CStr(mID)
    Text29.Text = mEmpresa

    'Primero leer exclusivamente los datos locales del ID correcto.
    CargarDatosLocales mID

    'Solo enriquecer campos SAT que NO tienen mapeo local conocido.
    'El fallo de SQL nunca elimina los datos ya cargados desde .dno.
    ConsultarCamposSATAdicionales

    BloquearCedula
    mPreparado = True
    PrepararEmpleado = True
    Exit Function

ErrorPreparar:
    Debug.Print "Cedula fiscal - ID " & CStr(idEmpleado) & ": " & Err.Description
    mPreparado = False
    LimpiarCampos
    Text1.Text = vbNullString
    Text29.Text = vbNullString
    BloquearCedula
    MsgBox "No fue posible cargar la cedula del empleado " & _
           CStr(idEmpleado) & "." & vbCrLf & Err.Description, _
           vbExclamation, "Cedula Fiscal"
End Function

'============================================================
' BOTONES
'============================================================
Private Sub Command1_Click()
    'Prohibir operaciones de escritura desde esta version.
    'No hay equivalencia local confirmada para todos los campos SAT.
    MsgBox "La edicion de esta cedula aun no esta habilitada.", _
           vbExclamation, "Cedula Fiscal"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    MsgBox "La edicion de esta cedula aun no esta habilitada.", _
           vbExclamation, "Cedula Fiscal"
End Sub

'Ya no se permite consultar otros IDs escribiendo en Text1.
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

'============================================================
' DATOS LOCALES: SIEMPRE PREVALECEN
'============================================================
Private Sub CargarDatosLocales(ByVal id As Long)
    Dim h As Integer
    Dim ruta As String
    Dim cantidad As Long
    Dim ancho As Long
    Dim descripcion As String

    On Error GoTo ErrorLocal

    '---- 1. Personal.dno: identidad principal ----
    ruta = "personal.dno"
    If Dir$(ruta) = "" Then
        Err.Raise vbObjectError + 310, , "No existe personal.dno."
    End If

    ancho = Len(personal)
    If ancho <= 0 Then Err.Raise vbObjectError + 311, , "Estructura personal no valida."
    If (FileLen(ruta) Mod ancho) <> 0 Then
        Err.Raise vbObjectError + 311, , "personal.dno tiene un tamano incompatible."
    End If

    h = FreeFile
    Open ruta For Random Access Read As #h Len = ancho
    cantidad = LOF(h) \ ancho

    If id > cantidad Then
        Err.Raise vbObjectError + 312, , _
                  "El ID " & CStr(id) & " no existe en personal.dno."
    End If

    Get #h, id, personal
    Close #h
    h = 0

    If CampoArchivo(personal.nom) = "" Then
        Err.Raise vbObjectError + 313, , _
                  "El registro " & CStr(id) & " no tiene nombre local."
    End If

    Text2.Text = CampoArchivo(personal.RFC)
    Text4.Text = CampoArchivo(personal.nom)
    Text5.Text = CampoArchivo(personal.ape1)
    Text6.Text = CampoArchivo(personal.ape2)

    '---- 2. PerOtre.dno: CURP ----
    ruta = "PerOtre.dno"
    If Dir$(ruta) <> "" Then
        ancho = Len(Otros_Rgtros)
        If ancho <= 0 Then Err.Raise vbObjectError + 314, , "Estructura PerOtre no valida."
        If (FileLen(ruta) Mod ancho) <> 0 Then
            Err.Raise vbObjectError + 314, , "PerOtre.dno tiene un tamano incompatible."
        End If

        h = FreeFile
        Open ruta For Random Access Read As #h Len = ancho
        cantidad = LOF(h) \ ancho

        If id <= cantidad Then
            Get #h, id, Otros_Rgtros
            Text3.Text = CampoArchivo(Otros_Rgtros.curp)
        End If

        Close #h
        h = 0
    End If

    '---- 3. Perscfdi.dno: direccion y contacto ----
    ruta = "Perscfdi.dno"
    If Dir$(ruta) <> "" Then
        ancho = Len(Empleado_1)
        If ancho <= 0 Then Err.Raise vbObjectError + 315, , "Estructura Perscfdi no valida."
        If (FileLen(ruta) Mod ancho) <> 0 Then
            Err.Raise vbObjectError + 315, , "Perscfdi.dno tiene un tamano incompatible."
        End If

        h = FreeFile
        Open ruta For Random Access Read As #h Len = ancho
        cantidad = LOF(h) \ ancho

        'IMPORTANTE: no llamar Get si el ID supera la cantidad.
        'Asi no se arrastran datos del empleado anterior.
        If id <= cantidad Then
            Get #h, id, Empleado_1

            Text11.Text = CampoArchivo(Empleado_1.Cpostal)
            Text13.Text = CampoArchivo(Empleado_1.Direccion)
            Text16.Text = CampoArchivo(Empleado_1.Colonia)
            Text17.Text = CampoArchivo(Empleado_1.Ciudad)
            Text18.Text = CampoArchivo(Empleado_1.Delegacion)
            Text19.Text = CampoArchivo(Empleado_1.Estado)
            Text22.Text = CampoArchivo(Empleado_1.correo)
        End If

        Close #h
        h = 0
    End If

    Exit Sub

ErrorLocal:
    descripcion = Err.Description
    On Error Resume Next
    If h > 0 Then Close #h
    On Error GoTo 0
    Err.Raise vbObjectError + 319, "CargarDatosLocales", descripcion
End Sub

'============================================================
' CAMPOS SAT SIN ALMACENAMIENTO LOCAL CONFIRMADO
'============================================================
Private Sub ConsultarCamposSATAdicionales()
    Dim rs As ADODB.Recordset
    Dim sSQL As String

    On Error GoTo ErrorConsulta

    If EsEmpresaLocal(emp) Then Exit Sub
    If con Is Nothing Then Exit Sub
    If con.State <> adStateOpen Then Exit Sub

    'La identidad, direccion, CP y correo NO se copian de SQL
    'porque YA proceden de los .dno.
    sSQL = "SELECT * FROM datosSat WHERE idNomina = " & _
           CStr(mID) & " AND empresa = '" & _
           Replace(mEmpresa, "'", "''") & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sSQL, con, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Then GoTo Salida

    'Si hay varias filas para el mismo ID/empresa, no escoger una
    'aleatoriamente; solo se mostraran los valores locales.
    rs.MoveNext
    If Not rs.EOF Then
        Debug.Print "Cedula fiscal: SQL duplicado en empresa=" & _
                    mEmpresa & ", ID=" & CStr(mID)
        GoTo Salida
    End If

    rs.MoveFirst

    'Evitar mezclar campos SAT de una persona diferente.
    'Si hay discrepancias de identidad, no usar SQL.
    If Not IdentidadCompatibleConSQL(rs) Then
        Debug.Print "Cedula: identidad SQL distinta. Empresa=" & _
                    mEmpresa & "; ID=" & CStr(mID)
        GoTo Salida
    End If

    Text7.Text = CampoSQL(rs, "fInicioOper")
    Text8.Text = CampoSQL(rs, "estatusPadron")
    Text9.Text = CampoSQL(rs, "fechaUltCamEstado")
    Text10.Text = CampoSQL(rs, "nombreComercial")
    Text12.Text = CampoSQL(rs, "tipoVialidad")
    Text14.Text = CampoSQL(rs, "numeroExterior")
    Text15.Text = CampoSQL(rs, "numeroInterior")
    Text20.Text = CampoSQL(rs, "entreCalle")
    Text21.Text = CampoSQL(rs, "yCalle")
    Text23.Text = CampoSQL(rs, "ladaTel")
    Text24.Text = CampoSQL(rs, "numeroTel")
    Text25.Text = CampoSQL(rs, "estadoDomicilio")
    Text26.Text = CampoSQL(rs, "estadoContDomicilio")
    Text27.Text = CampoSQL(rs, "actividadEconomica")
    Text28.Text = CampoSQL(rs, "regimen")

Salida:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    Exit Sub

ErrorConsulta:
    Debug.Print "Cedula fiscal: campos SAT no disponibles. " & Err.Description
    Resume Salida
End Sub


'Valida la identidad ANTES de leer campos SAT auxiliares.
'El ID y la empresa deben corresponder al empleado local.
Private Function IdentidadCompatibleConSQL( _
    ByVal rs As ADODB.Recordset) As Boolean

    Dim coincidencia As Boolean
    IdentidadCompatibleConSQL = False

    If IdentidadEnConflicto(Text4.Text, CampoSQL(rs, "nombre")) Then Exit Function
    If IdentidadEnConflicto(Text5.Text, CampoSQL(rs, "apellidoP")) Then Exit Function
    If IdentidadEnConflicto(Text6.Text, CampoSQL(rs, "apellidoM")) Then Exit Function
    If IdentidadEnConflicto(Text2.Text, CampoSQL(rs, "rfc")) Then Exit Function
    If IdentidadEnConflicto(Text3.Text, CampoSQL(rs, "curp")) Then Exit Function

    'Exigir al menos una coincidencia para no asociar registros
    'en los que SQL no dispone de informacion de identidad.
    coincidencia = ParCoincide(Text4.Text, CampoSQL(rs, "nombre")) Or _
                   ParCoincide(Text5.Text, CampoSQL(rs, "apellidoP")) Or _
                   ParCoincide(Text6.Text, CampoSQL(rs, "apellidoM")) Or _
                   ParCoincide(Text2.Text, CampoSQL(rs, "rfc")) Or _
                   ParCoincide(Text3.Text, CampoSQL(rs, "curp"))

    IdentidadCompatibleConSQL = coincidencia
End Function

Private Function IdentidadEnConflicto( _
    ByVal valorLocal As String, ByVal valorSQL As String) As Boolean

    valorLocal = NormalizarIdentidad(valorLocal)
    valorSQL = NormalizarIdentidad(valorSQL)

    If Len(valorLocal) > 0 And Len(valorSQL) > 0 Then
        IdentidadEnConflicto = (valorLocal <> valorSQL)
    End If
End Function

Private Function ParCoincide( _
    ByVal valorLocal As String, ByVal valorSQL As String) As Boolean

    valorLocal = NormalizarIdentidad(valorLocal)
    valorSQL = NormalizarIdentidad(valorSQL)

    If Len(valorLocal) > 0 And Len(valorSQL) > 0 Then
        ParCoincide = (valorLocal = valorSQL)
    End If
End Function

Private Function NormalizarIdentidad(ByVal s As String) As String
    s = UCase$(Trim$(Replace(s, vbNullChar, "")))
    s = Replace(s, "?", "A")
    s = Replace(s, "?", "E")
    s = Replace(s, "?", "I")
    s = Replace(s, "?", "O")
    s = Replace(s, "?", "U")
    s = Replace(s, "?", "U")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    If s = "N/A" Then s = ""
    NormalizarIdentidad = s
End Function

Private Function CampoSQL(ByVal rs As ADODB.Recordset, _
                          ByVal nombreCampo As String) As String
    Dim v As Variant
    v = rs.Fields(nombreCampo).Value
    If IsNull(v) Or IsEmpty(v) Then
        CampoSQL = vbNullString
    Else
        CampoSQL = Trim$(CStr(v))
    End If
End Function

Private Function CampoArchivo(ByVal valor As String) As String
    CampoArchivo = Trim$(Replace(valor, vbNullChar, ""))
End Function

'============================================================
' VISUALIZACION
'============================================================
Private Sub LimpiarCampos()
    Dim i As Long
    For i = 1 To 29
        Me.Controls("Text" & CStr(i)).Text = vbNullString
    Next i
    Text34.Text = vbNullString
End Sub

Private Sub BloquearCedula()
    Dim i As Long
    For i = 1 To 29
        Me.Controls("Text" & CStr(i)).Locked = True
    Next i
    Text34.Locked = True
End Sub

'Se conserva el mapeo de empresas usado en PeCfdi.
Private Function ObtenerEmpresaSAT() As String
    Dim abrEmpresa As String
    abrEmpresa = UCase$(Left$(Trim$(emp), 4))
    Select Case abrEmpresa
        Case "SACM": abrEmpresa = "SACMAG"
        Case "SUPT": abrEmpresa = "SUPTER"
        Case "COOR": abrEmpresa = "CORDINA"
        Case "EPES": abrEmpresa = "EPESA"
        Case "SUPE": abrEmpresa = "SUPERVISA"
        Case "CONS": abrEmpresa = "CONSULTE"
        Case "GEOA": abrEmpresa = "GEOAMBIENTE"
    End Select
    ObtenerEmpresaSAT = abrEmpresa
End Function

