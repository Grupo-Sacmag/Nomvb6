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
Private Sub Command1_Click()
' Boton para guardar una CIF
insertarEmpleado
   ' MsgBox ("Lo siento aun no funciono :C")
End Sub
Private Sub reviarEmpleado()
    If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28.Text = "" Or Text29.Text = "" Or Text34.Text = "" Then
        MsgBox ("Campos Obligatorios")
        colorearCamposVacios
    Else
    
        MsgBox "La cédula de indentificación fiscal para " & Text1.Text & " se guardo con éxito!"
    End If

End Sub

Private Sub colorearCamposVacios()
    Dim Control As Control

    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.Text = "" Then
                Control.BackColor = RGB(255, 0, 0) ' Rojo
            End If
        End If
    Next Control
End Sub

Private Sub Command2_Click()
    limpiarCampos
End Sub
Private Sub limpiarCampos()
' Limpia todos los campos
    Text2.Text = Empty
    Text3.Text = Empty
    Text4.Text = Empty
    Text5.Text = Empty
    Text6.Text = Empty
    Text7.Text = Empty
    Text8.Text = Empty
    Text9.Text = Empty
    Text10.Text = Empty
    Text11.Text = Empty
    Text12.Text = Empty
    Text13.Text = Empty
    Text14.Text = Empty
    Text15.Text = Empty
    Text16.Text = Empty
    Text17.Text = Empty
    Text18.Text = Empty
    Text19.Text = Empty
    Text20.Text = Empty
    Text21.Text = Empty
    Text22.Text = Empty
    Text23.Text = Empty
    Text24.Text = Empty
    Text25.Text = Empty
    Text26.Text = Empty
    Text27.Text = Empty
    Text28.Text = Empty
    Text29.Text = Empty
    Text34.Text = Empty
End Sub
Private Sub insertarEmpleado()
    Dim oRS As New ADODB.Recordset
    Dim sSQL As String

    ' Construir la consulta SQL con los valores de los controles de texto
    sSQL = "insert into datosSat values (" & Text1.Text & ",'" & Text2.Text & "','" & _
    Text3.Text & "','" & Text4.Text & "','" & _
    Text5.Text & "','" & Text6.Text & "','" & _
    Text7.Text & "','" & Text8.Text & "','" & _
    Text9.Text & "','" & Text10.Text & "','" & _
    Text11.Text & "','" & Text12.Text & "','" & _
    Text13.Text & "','" & Text14.Text & "','" & _
    Text15.Text & "','" & Text16.Text & "','" & _
    Text17.Text & "','" & Text18.Text & "','" & _
    Text19.Text & "','" & Text20.Text & "','" & _
    Text21.Text & "','" & Text22.Text & "','" & _
    Text23.Text & "','" & Text24.Text & "','" & _
    Text25.Text & "','" & Text26.Text & "','" & _
    Text27.Text & "','" & Text28.Text & "','" & _
    Text29.Text & "', null , null , null , null , null )"

    
    ' Abre el recordset y ejecuta la consulta SQL
    oRS.CursorLocation = adUseClient
    oRS.Open sSQL, con, adOpenStatic, adLockBatchOptimistic, adCmdText


    ' Disconnect the Recordset.
    Set oRS.ActiveConnection = Nothing

    ' Cierra el Recordset
    Set oRS = Nothing
    
    MsgBox "Se añadio con éxito"
    
    limpiarCampos

End Sub



Private Sub consultarCIF()
On Error GoTo manejador

Dim id As Integer
Dim abrEmpresa As String

    id = Text1.Text
    
    abrEmpresa = Left(Trim(emp), 4)
  
    Select Case UCase(abrEmpresa)
        Case "SACM"
            abrEmpresa = "SACMAG"
        Case "COOR"
            abrEmpresa = "CORDINA"
        Case "EPES"
            abrEmpresa = "EPESA"
        Case "SUPE"
            abrEmpresa = "SUPERVISA"
        Case "CONS"
            abrEmpresa = "CONSULTE"
        ' Agregar más casos según sea necesario
    End Select
    
    Text29.Text = abrEmpresa

    sSQL = "SELECT * FROM datosSat WHERE idNomina = '" & id & "' AND empresa = '" & abrEmpresa & "'"
            
    Set oRS = New ADODB.Recordset
    oRS.CursorLocation = adUseClient
    oRS.Open sSQL, con, adOpenStatic, adLockBatchOptimistic, adCmdText
                
    oRS.MarshalOptions = adMarshalModifiedOnly
    
    If Not oRS.EOF Then
        ' Mostrar cada campo en textboxes diferentes (sustituye con los nombres reales de tus campos)
        Text2.Text = oRS.Fields("rfc").Value
        Text3.Text = oRS.Fields("curp").Value
        Text4.Text = oRS.Fields("nombre").Value
        Text5.Text = oRS.Fields("apellidoP").Value
        Text6.Text = oRS.Fields("apellidoM").Value
        Text7.Text = oRS.Fields("fInicioOper").Value
        Text8.Text = oRS.Fields("estatusPadron").Value
        Text9.Text = oRS.Fields("fechaUltCamEstado").Value
        Text10.Text = oRS.Fields("nombreComercial").Value
        Text11.Text = oRS.Fields("cp").Value
        Text12.Text = oRS.Fields("tipoVialidad").Value
        Text13.Text = oRS.Fields("nombreDeVialidad").Value
        Text14.Text = oRS.Fields("numeroExterior").Value
        Text15.Text = oRS.Fields("numeroInterior").Value
        Text16.Text = oRS.Fields("nombreColonia").Value
        Text17.Text = oRS.Fields("nombreLocalidad").Value
        Text18.Text = oRS.Fields("nombreMunicipio").Value
        Text19.Text = oRS.Fields("entidadFederativa").Value
        Text20.Text = oRS.Fields("entreCalle").Value
        Text21.Text = oRS.Fields("yCalle").Value
        Text22.Text = oRS.Fields("correoElectronico").Value
        Text23.Text = oRS.Fields("ladaTel").Value
        Text24.Text = oRS.Fields("numeroTel").Value
        Text25.Text = oRS.Fields("estadoDomicilio").Value
        Text26.Text = oRS.Fields("estadoContDomicilio").Value
        Text27.Text = oRS.Fields("actividadEconomica").Value
        Text28.Text = oRS.Fields("regimen").Value
        Text29.Text = oRS.Fields("empresa").Value
    Else
        MsgBox "No se encontraron registros con ese ID"
    End If
    
    ' Disconnect the Recordset.
    Set oRS.ActiveConnection = Nothing
    oRS.Close
    Set oRS = Nothing
    
Exit Sub

manejador:
    Resume Next
    

End Sub



Private Sub Form_Load()
    Text1.Text = Form2.Text16.Text
    Text1_KeyPress 13
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ' Primero limpia los datos
        limpiarCampos
        ' Manda llamar la CIF del empleado
        consultarCIF
    End If
End Sub

Private Sub Command3_Click()
    actualizarEmpleado
End Sub

Private Sub actualizarEmpleado()
    Dim idNomina As String
    Dim oRS As New ADODB.Recordset
    Dim sSQL As String
    
    idNomina = Text1.Text
    
    ' Construye la consulta SQL con todas las actualizaciones
    sSQL = "UPDATE datosSat SET " & _
           "rfc = '" & Text2.Text & "', " & "curp = '" & Text3.Text & "', " & "nombre = '" & Text4.Text & "', " & "apellidoP = '" & Text5.Text & "', " & "apellidoM = '" & Text6.Text & "', " & _
           "fInicioOper = '" & Text7.Text & "', " & "estatusPadron = '" & Text8.Text & "', " & "fechaUltCamEstado = '" & Text9.Text & "', " & "nombreComercial = '" & Text10.Text & "', " & "cp = '" & Text11.Text & "', " & _
           "tipoVialidad = '" & Text12.Text & "', " & "nombreDeVialidad = '" & Text13.Text & "', " & "numeroExterior = '" & Text14.Text & "', " & "numeroInterior = '" & Text15.Text & "', " & "nombreColonia = '" & Text16.Text & "', " & _
           "nombreLocalidad = '" & Text17.Text & "', " & "nombreMunicipio = '" & Text18.Text & "', " & "entidadFederativa = '" & Text19.Text & "', " & "entreCalle = '" & Text20.Text & "', " & "yCalle = '" & Text21.Text & "', " & _
           "correoElectronico = '" & Text22.Text & "', " & "ladaTel = '" & Text23.Text & "', " & "numeroTel = '" & Text24.Text & "', " & "estadoDomicilio = '" & Text25.Text & "', " & "estadoContDomicilio = '" & Text26.Text & "', " & _
           "actividadEconomica = '" & Text27.Text & "', " & "regimen = '" & Text28.Text & "', " & "empresa = '" & Text29.Text & "' " & "WHERE idNomina = '" & idNomina & "'"
    
    ' Abre el recordset y ejecuta la consulta SQL
    oRS.CursorLocation = adUseClient
    oRS.Open sSQL, con, adOpenStatic, adLockBatchOptimistic, adCmdText


    ' Disconnect the Recordset.
    Set oRS.ActiveConnection = Nothing

    ' Cierra el Recordset
    Set oRS = Nothing
    
    MsgBox "Los campos han sido actualizados!!!"
    
    limpiarCampos
    
End Sub


