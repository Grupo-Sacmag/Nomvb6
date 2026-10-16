VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   Caption         =   "Nomina Captura"
   ClientHeight    =   7995
   ClientLeft      =   7680
   ClientTop       =   2640
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "capnom.frx":0000
   LinkTopic       =   "Form8"
   ScaleHeight     =   7995
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPre 
      Height          =   375
      Left            =   3360
      MaxLength       =   4
      TabIndex        =   20
      Text            =   "Text5"
      Top             =   960
      Width           =   2415
   End
   Begin VB.ComboBox cboEspecial 
      Height          =   360
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   16
      Top             =   960
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   7680
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog excel 
      Left            =   9480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   13
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   4005
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7064
      _Version        =   393216
      Rows            =   24
      Cols            =   25
      FixedCols       =   2
      BackColorBkg    =   -2147483644
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "capnom.frx":0442
      Left            =   3480
      List            =   "capnom.frx":0444
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Quincena:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton Option4 
         Caption         =   "Segunda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Primera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nomina:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
      Begin VB.OptionButton Option2 
         Caption         =   "Especial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdRestoreFilter 
      Caption         =   "Restaurar filtros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9240
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   3360
      TabIndex        =   21
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "   Usando tarifas de impuestos : "
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   2400
      Width           =   6495
   End
   Begin VB.Label Label8 
      Caption         =   "Cantidad exenta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   1920
      Width           =   9375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Tipo de la nomina especial:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.Menu capnoar 
      Caption         =   "&Archivo"
      Index           =   0
      Begin VB.Menu archnom 
         Caption         =   "&Archivar nomina"
         Index           =   2
         Shortcut        =   ^G
      End
      Begin VB.Menu edicionPersonalCaptura 
         Caption         =   "Edición personal"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu capnosal 
         Caption         =   "&Salida"
         Index           =   1
      End
   End
   Begin VB.Menu NomEd 
      Caption         =   "&Edicion"
      Begin VB.Menu EdiNomSel 
         Caption         =   "&Seleccionar Todo"
         Shortcut        =   ^S
      End
      Begin VB.Menu EdNomSep1 
         Caption         =   "-"
      End
      Begin VB.Menu funcionPegar 
         Caption         =   "Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu EdNomCopiar 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu nomcapini 
      Caption         =   "&Captura"
      Index           =   0
      Begin VB.Menu nomcap_ini 
         Caption         =   "&Iniciar captura"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu capturaGeneral 
         Caption         =   "&Iniciar captura general"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu CapArch 
         Caption         =   "&Archivo banamex"
         Index           =   2
      End
      Begin VB.Menu csep3 
         Caption         =   "-"
      End
      Begin VB.Menu CapDis 
         Caption         =   "&Distribucion Nomina"
      End
      Begin VB.Menu capsep4 
         Caption         =   "-"
      End
      Begin VB.Menu CapInrt 
         Caption         =   "&Internet"
      End
      Begin VB.Menu capsep5 
         Caption         =   "-"
      End
      Begin VB.Menu CFDI2 
         Caption         =   "C&FDI"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu capact 
      Caption         =   "Ac&tualizacion"
      Index           =   0
      Begin VB.Menu capacimp 
         Caption         =   "&Impuestos"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   {F8}
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu cap_act 
         Caption         =   "&Sumas"
         Index           =   1
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu capnomord 
      Caption         =   "&Ordenar"
      Index           =   0
      Begin VB.Menu capnoalf 
         Caption         =   "&Alfabeticamente"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu capnomnum 
         Caption         =   "&Numericamente"
         Index           =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu minToMax 
         Caption         =   "De menor a mayor"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu nomordeli 
         Caption         =   "&Eliminar"
         Index           =   3
      End
   End
   Begin VB.Menu capnoimp 
      Caption         =   "&Impresion"
      Index           =   0
      Begin VB.Menu capimpno 
         Caption         =   "&Nomina"
         Index           =   1
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra0"
            Index           =   0
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra1"
            Index           =   1
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra2"
            Index           =   2
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "-"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra8"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra9"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra10"
            Index           =   10
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra11"
            Index           =   11
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra12"
            Index           =   12
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra13"
            Index           =   13
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra14"
            Index           =   14
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra15"
            Index           =   15
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra16"
            Index           =   16
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra17"
            Index           =   17
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra18"
            Index           =   18
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra19"
            Index           =   19
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra20"
            Index           =   20
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra21"
            Index           =   21
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra22"
            Index           =   22
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra23"
            Index           =   23
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra24"
            Index           =   24
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra25"
            Index           =   25
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra26"
            Index           =   26
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra27"
            Index           =   27
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra28"
            Index           =   28
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra29"
            Index           =   29
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra30"
            Index           =   30
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra31"
            Index           =   31
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra32"
            Index           =   32
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra33"
            Index           =   33
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra34"
            Index           =   34
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra35"
            Index           =   35
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra36"
            Index           =   36
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra37"
            Index           =   37
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra38"
            Index           =   38
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra39"
            Index           =   39
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra40"
            Index           =   40
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra41"
            Index           =   41
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra42"
            Index           =   42
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra43"
            Index           =   43
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra44"
            Index           =   44
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra45"
            Index           =   45
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra46"
            Index           =   46
            Visible         =   0   'False
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra47"
            Index           =   47
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra48"
            Index           =   48
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra49"
            Index           =   49
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra50"
            Index           =   50
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra51"
            Index           =   51
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra52"
            Index           =   52
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra53"
            Index           =   53
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra54"
            Index           =   54
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra55"
            Index           =   55
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra56"
            Index           =   56
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra57"
            Index           =   57
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra58"
            Index           =   58
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra59"
            Index           =   59
         End
         Begin VB.Menu Mnunomobr 
            Caption         =   "obra60"
            Index           =   60
         End
      End
      Begin VB.Menu nomcaprec 
         Caption         =   "&Recibos"
         Index           =   2
         Begin VB.Menu caprectd 
            Caption         =   "&Todos"
            Index           =   1
         End
         Begin VB.Menu caprecind 
            Caption         =   "&Individuales"
            Index           =   2
         End
      End
      Begin VB.Menu imprche 
         Caption         =   "&Cheques"
         Index           =   0
         Begin VB.Menu cheajte 
            Caption         =   "&Ajustar impresion"
            Index           =   4
         End
         Begin VB.Menu sep5 
            Caption         =   "-"
         End
         Begin VB.Menu cheind 
            Caption         =   "&Individuales"
            Index           =   1
         End
         Begin VB.Menu chetot 
            Caption         =   "&Totales"
            Index           =   2
         End
         Begin VB.Menu sep6 
            Caption         =   "-"
         End
         Begin VB.Menu cheobra 
            Caption         =   "&Directorio de obra"
            Index           =   3
         End
      End
      Begin VB.Menu CapDetImp 
         Caption         =   "&Detalle Calculo Imptos"
      End
   End
   Begin VB.Menu exportar 
      Caption         =   "Exportar a Excel"
      Index           =   0
   End
   Begin VB.Menu Importar 
      Caption         =   "&Importar Datos"
   End
   Begin VB.Menu Borrar 
      Caption         =   "&Borrar"
      Begin VB.Menu Borrarblancos 
         Caption         =   "&BorrarBlancos"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sum(23) As Currency, sumv(29) As Currency, meselegido As Integer, diat As Currency
Dim integrado As Currency, seguro As Currency, diaseg As Currency
Dim limite As Integer, valcelant, hs_ext As Currency
Dim dato_ent, dato_sal As Currency, ruta As Integer
Dim si_impto As Integer, si_imss As Integer, compa$, mientras$, compa_re As Integer, comp As Integer, mientra As Integer
Dim sec(20) As Integer, retorno As Integer, actual As Integer
Dim conta As Integer, conta1 As Integer, conta2 As Integer, conta3 As Integer
Dim hoja As Integer, valtit As Integer, aumento(3) As Integer
Dim empieza As Integer, termina As Integer, fpago As Integer
Dim MiDía, qna As Integer, dia_pago As Integer
Dim impto1 As Integer, imss1 As Integer, f As Long, ii As Long, ii1 As Long
Dim exento As Currency, archiva_o, cuentaobra As Integer
Dim requisito As Integer, comp_ti As String, Clave As Integer
Dim si_hay As Integer, operacion, lim_conta As Integer
Dim fi_nm As Integer, anchopapel As Long, largopapel As Long
Dim Entrada As String * 32, abre, caso, Salida, numerillo As Integer, lincam As Currency
Dim Slinimpte As Currency, Slincam As Currency, Sabono As Integer
Dim porc_apli As Currency, p_vacacional As Currency, Detener As Integer, Detener1 As Integer
Dim PrincIngr As Long, PrincDesc As Long, Color_gris As Long, Nu_mero As Double, M_aximo As Long
Public subsidioParaElEmpleo As Currency
Private ActualizandoNombre As Boolean

Private Sub Borrarblancos_Click(Index As Integer)
    Dim i As Integer
    Dim vacia As Boolean

    ' Validar que ConNom1 tiene filas suficientes
    If ConNom1.Rows <= 1 Then
        MsgBox "No hay filas suficientes para procesar.", vbExclamation, "Error"
        Exit Sub
    End If

    ' Recorremos las filas de abajo hacia arriba
    For i = ConNom1.Rows - 1 To 1 Step -1
        vacia = True

        ' Verificar las columnas 4, 5, 6 y 7 (Índices 3, 4, 5, 6 porque en VB6 son base 0)
        Rem If Trim(ConNom1.TextMatrix(i, 3)) <> "" Then vacia = False
        Rem If Trim(ConNom1.TextMatrix(i, 4)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 5)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 6)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 7)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 8)) <> "" Then vacia = False
        Rem If Trim(ConNom1.TextMatrix(i, 9)) <> "" Then vacia = False
    

        ' Si está vacía, eliminar la fila
        If vacia Then
            ConNom1.RemoveItem i
        End If
    Next i

End Sub

Private Sub cboEspecial_Click()

    ActualizarNombreEspecial

    If cboEspecial.ListIndex >= 0 Then
        txtPre.SetFocus
        txtPre.SelStart = Len(txtPre.Text)
    End If

End Sub

Private Sub Form_Load()

    InicializarColumnas tnOrdinaria
    Mnunomobr.Item(0).Visible = True
    Mnunomobr.Item(0).Caption = "Nomina Total"
    Mnunomobr.Item(1).Visible = True
    Mnunomobr.Item(1).Caption = "Todas las Obras"
    Mnunomobr.Item(2).Visible = True
    Mnunomobr.Item(2).Caption = "Nomina Oficina"
    Mnunomobr.Item(3).Visible = True
    Label9.Caption = Label9.Caption + Dir_imptos
    cuentaobra = 3
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Importar.Enabled = False
    ConNom1.Clear
    ConNom1.Height = Form8.Height * 1.9
    ConNom1.Width = Form8.Width * 0.975
    cal_anual = 0
    meselegido = 1
    ProgressBar1.Left = 1500
    ProgressBar1.Top = Form8.Height * 1.3
    ProgressBar1.Width = Form8.ScaleWidth * 2

    Close 2
    Close 8
     
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    ddm = LOF(8) / Len(maestro)
     
    Get #8, 1, maestro
    Get #2, 1, personal
     
    If Dm > 0 Then
    ConNom1.Cols = 25
    ConNom1.Rows = Dm + 2

    For i = 1 To 12
        Combo1.AddItem RTrim$(mm(i))
    Next i

    Combo1.Text = Combo1.List(0)
    miMesesito = Month(Date)
    Combo1.ListIndex = miMesesito - 1

    Label7.Caption = "Nomina de la 1a.quincena de" + Combo1.Text + Str$(empresa.ao)
    Else
        Close 2
        MsgBox "No existe archivo de personal no es posible capturar la nomina"
    End If

    cboEspecial.Clear
    cboEspecial.AddItem "AGUINALDO"
    cboEspecial.AddItem "PTU"
    cboEspecial.AddItem "BONO"
    cboEspecial.AddItem "FINIQUITO"
    cboEspecial.AddItem "LIQUIDACION"

    Text1.Visible = False
    MostrarControlesEspecial False

    GoTo sale8

Manejo8:
    capnosal_Click 1

sale8:
End Sub

Private Sub nomcap_ini_Click(Index As Integer)

    On Error GoTo ErrorManejador

    Dim Arch As String
    Dim fullArchPath As String
    Dim rutaNominas As String
    Dim tamArchivo As Long
    Dim registrosArchivo As Long

    '=======================================================
    ' 1. CONSTRUIR EL NOMBRE DE LA NOMINA
    '=======================================================

    If Option1.Value = True Then

        ' NOMINA NORMAL

        If Option3.Value = True Then

            Arch = UCase$(Mid$(Combo1.Text, 1, 3)) & _
                   "1" & LTrim$(Str$(empresa.ao)) & ".NOM"

        ElseIf Option4.Value = True Then

            Arch = UCase$(Mid$(Combo1.Text, 1, 3)) & _
                   "2" & LTrim$(Str$(empresa.ao)) & ".NOM"

        Else

            MsgBox "Seleccione Primera o Segunda quincena.", _
                   vbExclamation, "Nomina"

            Exit Sub

        End If

    ElseIf Option2.Value = True Then

        ' NOMINA ESPECIAL

        If cboEspecial.ListIndex < 0 Then

            MsgBox "Seleccione el tipo de nomina especial.", _
                   vbExclamation, "Nomina"

            Exit Sub

        End If

        ActualizarNombreEspecial

        If Len(Trim$(Text1.Text)) = 0 Then

            MsgBox "Indique el nombre de la nomina especial.", _
                   vbExclamation, "Nomina"

            Exit Sub

        End If

        Arch = Trim$(Text1.Text) & ".NOM"

    Else

        MsgBox "Seleccione el tipo de nomina.", _
               vbExclamation, "Nomina"

        Exit Sub

    End If

    '=======================================================
    ' 2. OBTENER RUTA COMPLETA
    '=======================================================

    rutaNominas = Trim$(CStr(Form1.Dir1))

    If Len(rutaNominas) = 0 Then

        MsgBox "No se ha seleccionado un directorio de nominas.", _
               vbExclamation, "Nomina"

        Exit Sub

    End If

    If Right$(rutaNominas, 1) <> "\" Then
        rutaNominas = rutaNominas & "\"
    End If

    fullArchPath = rutaNominas & Arch

    Debug.Print "===================================="
    Debug.Print "VALIDACION PREVIA DE NOMINA"
    Debug.Print "Archivo: "; Arch
    Debug.Print "Ruta: "; fullArchPath

    '=======================================================
    ' 3. SI YA EXISTE, NO VOLVER A GENERARLA
    '=======================================================

    If Len(Dir$(fullArchPath)) > 0 Then

        tamArchivo = FileLen(fullArchPath)

        Debug.Print "Archivo existente"
        Debug.Print "Bytes: "; tamArchivo

        ' Archivo vacio
        If tamArchivo = 0 Then

            MsgBox "Ya existe un archivo con ese nombre, " & _
                   "pero esta vacio." & vbCrLf & vbCrLf & _
                   Arch & vbCrLf & vbCrLf & _
                   "No se iniciara otra captura para evitar " & _
                   "sobrescribirlo." & vbCrLf & _
                   "Revise el archivo antes de continuar.", _
                   vbExclamation, "Archivo vacio"

            Exit Sub

        End If

        ' Archivo con longitud incompatible
        If tamArchivo Mod Len(nomina) <> 0 Then

            MsgBox "El archivo de nomina existe, pero su " & _
                   "tamano no corresponde a registros completos." & _
                   vbCrLf & vbCrLf & _
                   Arch & vbCrLf & _
                   "No se modificara el archivo.", _
                   vbCritical, "Nomina"

            Exit Sub

        End If

        registrosArchivo = tamArchivo \ Len(nomina)

        Debug.Print "Registros existentes: "; registrosArchivo

        '===================================================
        ' 4. ABRIR LA NOMINA EXISTENTE EN FORMVIEWER
        '===================================================

        MsgBox "Esta nomina ya existe:" & vbCrLf & _
               Arch & vbCrLf & vbCrLf & _
               "Se abrira en el visualizador." & vbCrLf & _
               "No se iniciara una nueva captura.", _
               vbInformation, "Nomina existente"

        Load FormViewer

        FormViewer.LoadPayroll fullArchPath

        If Not FormViewer.CargaCorrecta Then

            Unload FormViewer

            MsgBox "No fue posible cargar la nomina existente." & _
                   vbCrLf & "No se modifico el archivo.", _
                   vbExclamation, "Nomina"

            Exit Sub

        End If

        Me.Hide

        FormViewer.Show vbModal

        Me.Show

        Exit Sub

    End If

    '=======================================================
    ' 5. NOMINA NUEVA
    '=======================================================

    Debug.Print "La nomina aun no existe."

    ' generarNominas usa nombres relativos.
    ' No iniciar si la ruta activa es diferente.

    If StrComp(CurDir$, _
               Left$(rutaNominas, Len(rutaNominas) - 1), _
               vbTextCompare) <> 0 Then

        MsgBox "El directorio de trabajo no coincide " & _
               "con el directorio seleccionado." & vbCrLf & _
               vbCrLf & _
               "Directorio activo: " & CurDir$ & vbCrLf & _
               "Directorio seleccionado: " & rutaNominas & _
               vbCrLf & vbCrLf & _
               "No se iniciara la captura para evitar " & _
               "crear archivos en una ubicacion incorrecta.", _
               vbExclamation, "Verificar directorio"

        Exit Sub

    End If

    '=======================================================
    ' 6. CLASIFICAR E INICIALIZAR COLUMNAS
    '=======================================================

    AplicarTipoNomina Arch

    Debug.Print "Tipo: "; g_TipoNominaActiva
    Debug.Print "C_EXENTA: "; C_EXENTA
    Debug.Print "C_TOTING: "; C_TOTING
    Debug.Print "C_NETO: "; C_NETO

    '=======================================================
    ' 7. CONTINUAR CON LA CAPTURA ORIGINAL
    '=======================================================

    generarNominas
    FuncionImpuesto

    Exit Sub

    '=======================================================
    ' ERROR
    '=======================================================

ErrorManejador:

    MsgBox "No fue posible iniciar o consultar la nomina." & _
           vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Nomina"

End Sub
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
 On Error GoTo Error_Handler
  
    Dim o_Excel As Object
    Dim o_Libro As Object
    Dim o_Hoja As Object
    Dim fila As Long
    Dim columna As Long
    
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For fila = 1 To .Rows - 1
            'For Columna = 0 To .Cols - 1
            For columna = 0 To 1
                o_Hoja.Cells(fila, columna + 1).Value = .TextMatrix(fila, columna)
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

Sub DibujaLinea()
   
   For r1 = 3 To 10
       If ConNom1.TextMatrix(ConNom1.Row, r1) <> "" Then
           Primera = r1
           r1 = 10
       End If
   Next r1
   
   For r1 = 15 To 19
      If ConNom1.TextMatrix(ConNom1.Row, r1) <> "" Then
           Primera1 = r1
           r1 = 19
       End If
   Next r1

End Sub
Sub depura()
   caso = ""
    For t = 1 To Len(Entrada)
        If (Mid(Entrada, t, 1) >= Chr(48) And Mid(Entrada, t, 1) <= Chr(57)) And (t < 4) Then
            Rem nada
        Else
            caso = caso + Mid(Entrada, t, 1)
        End If
    Next
   Entrada = RTrim(caso)
End Sub

Sub leer_cat()
    Close 9
    sal_cat = 0: Entrada = ""
    abre = dir_obras + "CATAUX"
   
    Open abre For Random As 9 Len = Len(CATAUX)
    cat_fin = LOF(9) / Len(CATAUX)
   
    For rr = 11 To 410: Get 9, rr, CATAUX
        If Val(CATAUX.C1) = maestro.O_1 Then
            Entrada = RTrim(CATAUX.C2)
            sal_cat = 1
        End If
    Next rr
   
SALE_LEER:
    If sal_cat = 0 Then
        Entrada = "Capturar"
    End If
    
    depura
     
   Close 9
End Sub


Private Sub CapArch_Click(Index As Integer)
    Dim Impte_Tr As Currency, Siguiente As Integer
    
    Close 13, 12, 11
    TREXC.Show
    
    Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
    Open "quin.txt" For Output As #11
    Open "Empcomp.dno" For Random As 13 Len = Len(Dat_ide)
    Get 13, 1, Dat_ide
    
    Slinimpte = 0: Sabono = 0: Siguiente = 0
    For r1 = 1 To limite
        li = r1
        rgtro = ConNom1.TextMatrix(li, 0)
    
        Get #2, rgtro, personal
        Get #8, rgtro, maestro
        Get #12, rgtro, Clbnx
    
        If Val(Clbnx.Q1) > 1 Then
        If IsNumeric(ConNom1.TextMatrix(r1, 20)) Then
            vales = ConNom1.TextMatrix(r1, 10)
        Else
            vales = 0
        End If
        If vales = "" Then vales = 0
            linimpte = ConNom1.TextMatrix(r1, 20)
            Slinimpte = Slinimpte + linimpte
            Sabono = Sabono + 1
        End If
    Next r1
    
    If Slinimpte = 0 Then GoTo terminuevo
        Slincam = Slinimpte - Int(Slinimpte)
        slinca$ = Mid(Str(Slincam), 3, 2)
        Slinimpte = Int(Slinimpte)
        slinimp$ = String(16 - Len(LTrim(Str(Slinimpte))), "0") + LTrim(Str(Slinimpte))
        If Len(slinca$) < 2 Then
            slinca$ = slinca$ + "0"
        End If
        cte = LTrim(RTrim(Dat_ide.clte))
        cte = String(12 - Len(cte), "0") + RTrim(Dat_ide.clte)
        LINARCH$ = "1" + cte
    
        If meselegido > 9 Then
            LINFECHA$ = LTrim(Str(dia_pago)) + LTrim(Str(meselegido)) + Right(RTrim(Str(empresa.ao)), 2)
        Else
            LINFECHA$ = LTrim(Str(dia_pago)) + "0" + LTrim(Str(meselegido)) + Right(RTrim(Str(empresa.ao)), 2)
        End If
      
        If dia_pago > 20 Then
            linQNA = meselegido * 2
        Else
            linQNA = meselegido * 2 - 1
        End If
        If linQNA > 9 Then
            linqn_a$ = "00" + LTrim(Str(linQNA))
        Else
            linqn_a$ = "000" + LTrim(Str(linQNA))
    End If
            
    LINARCH$ = LINARCH$ + LINFECHA$ + linqn_a$
    linemp$ = Left(empresa.name, 35)
      
    If Len(linemp$) < 36 Then
        linemp$ = linemp$ + String$(36 - Len(linemp$), " ")
    End If
      
    LINARCH$ = LINARCH$ + linemp$ + "QUIN" + String(16, " ") + "05" + String(40, " ") + "B" + "00"
    Print #11, LINARCH$
    LINARCH$ = "21001"
      
    If Len(RTrim(Dat_ide.suc)) < 4 Then
        susal = "0" + RTrim(Dat_ide.suc)
    Else
         susal = RTrim(Dat_ide.suc)
    End If
      
    ctaban = String(20 - Len(RTrim(Dat_ide.cta)), "0") + RTrim(Dat_ide.cta)
    LINARCH$ = LINARCH$ + slinimp$ + slinca$ + "01" + susal + ctaban + String(20, " ")
    Print #11, LINARCH$
    Slinimpte = 0
    
    For r1 = 1 To limite
        li = r1
        rgtro = ConNom1.TextMatrix(li, 0)
        Get #2, rgtro, personal
        Get #8, rgtro, maestro
        Get #12, rgtro, Clbnx
        referencia$ = String(8, " ") + "01"
        If Val(Clbnx.Q1) < 1 Then GoTo nuevo
            If ConNom1.TextMatrix(r1, 20) > "" Then
                vales = ConNom1.TextMatrix(r1, 10)
            Else
                vales = 0
        End If
        
        If vales = "" Then vales = 0
            linimpte = ConNom1.TextMatrix(r1, 20)
            Slinimpte = Slinimpte + linimpte
            lincam = linimpte - Int(linimpte)
            If lincam = 0 Then
                linca$ = "00"
            Else
                linca$ = Mid(Str(lincam), 3, 2)
            End If
        If Len(linca$) < 2 Then linca$ = linca$ + "0"
            Impte_Tr = linimpte
            linimpte = Int(linimpte)
            linimp$ = String(16 - Len(LTrim(Str(linimpte))), "0") + LTrim(Str(linimpte))
            plastico = ""
        For i = 1 To Len(Clbnx.Q1)
            If Mid(Clbnx.Q1, i, 1) = "-" Then
                plastico = "01"
                punto = i - 1
                Exit For
            End If
        Next i
      
    If plastico <> "01" Then
        plastico = "03"
        claveban = Clbnx.Q1
        claveban = String(20 - Len(Trim(claveban)), "0") + claveban
    Else
        numsuc = Left(Clbnx.Q1, punto)
        numsuc = String(13 - Len(LTrim(numsuc)), "0") + numsuc
        claveban = Mid(Clbnx.Q1, (punto + 2))
        claveban = String(7 - Len(Trim(claveban)), "0") + claveban
    End If
    
    LINARCH$ = "30001" + linimp$ + linca$ + plastico + claveban + String(30, " ") + referencia$
    nomban$ = Trim(personal.nom) & Chr(9) & Trim(personal.ape1) & Chr(9) & Trim(personal.ape2)
    nomban$ = Mid$(nomban$, 1, 55)
    
    If Len(nomban$) < 55 Then
        nomban$ = nomban$ + String(55 - Len(nomban$), " ")
    End If
    If plastico = "01" Then
        LINARCH$ = LINARCH$ + nomban$ + "NOMINA" + String(34, " ") + "NOMINA" + String(18, " ") + String(4, "0") + String(6, "0") + "1" + "00"
    Else
        LINARCH$ = LINARCH$ + nomban$ + "NOMINA" + String(58, " ") + String(10, "0") + "1" + "00"
    End If
      
    Print #11, LINARCH$
    Siguiente = Siguiente + 1
        TREXC.Ex.AddItem _
            Clbnx.Q1 & Chr(9) & _
            Trim(personal.nom) & Chr(9) & _
            Trim(personal.ape1) & Chr(9) & _
            Trim(personal.ape2) & Chr(9) & _
            Format(Impte_Tr, "#,##0.00") & Chr(9) & _
            Format(rgtro, "####0") & Chr(9) & _
            " " & Chr(9) & _
            Mid(Label7.Caption, 1, 7) & Chr(9) & _
            "Mismo dia" & Chr(9) & _
            Format(Siguiente, "####0")
           
nuevo:
     
    Next r1
    
    slinca$ = "":
    LINARCH$ = "4001"
    Slincam = Int((Slinimpte - Int(Slinimpte)) * 100) ' Se determina la fraccion
    slinca$ = Trim(Str(Slincam)) ' Se considera el simbolo
    
    If Slincam <= 0 Then slinca$ = "00"
    
    sabon$ = Mid(Str(Sabono), 2)
    psabon$ = String(6 - Len(sabon$), "0") + sabon$
    Slinimpte = Int(Slinimpte)
    slinimp$ = String(16 - Len(LTrim(Str(Slinimpte))), "0") + LTrim(Str(Slinimpte))
    LINARCH$ = LINARCH$ + psabon$ + slinimp$ + slinca$ + "000001" + slinimp$ + slinca$
    
    Print #11, LINARCH$
    
    LINARCH$ = "": Sabono = 0
    
terminuevo:
    Close 12, 13, 11
End Sub
Sub TituloCompl()
  Printer.CurrentX = Printer.Width / 2 - (Printer.TextWidth(RTrim(empresa.name)) / 2)
   Printer.Print empresa.name
   Encabezado = "Detalle Calculo de Impuestos correspondiente a la " + Label7.Caption
   Printer.CurrentX = Printer.Width / 2 - (Printer.TextWidth(RTrim(Encabezado)) / 2)
   Printer.Print Encabezado
   Printer.Print
   Printer.Print Tab(8); String(230, "-")
   Printer.Print
   Printer.Print Tab(8); "Num.";
   Printer.Print Tab(30); "N O M B R E";
   Printer.Print Tab(58); "INGRESO";
   Printer.Print Tab(73); "Impuesto";
   Printer.Print Tab(88); "Subsidio";
   If empresa.ao < 2008 Then
        Printer.Print Tab(103); "Cr.Salario";
        Else
        Printer.Print Tab(103); "Sbdio.p/empleo";
   End If
   Printer.Print Tab(118); "Impuesto";
   If empresa.ao < 2008 Then
        Printer.Print Tab(131); "Cr.Pagado"
        Else
        Printer.Print Tab(131); "Subdio.Pagado"
   End If
   Printer.Print Tab(75); "Total";
   Printer.Print Tab(87); "Acreditable";
   Printer.Print Tab(103); "Calculado";
   Printer.Print Tab(118); "Retenido";

   Printer.Print
   Printer.Print Tab(8); String(230, "-")
   Printer.Print

End Sub

Private Sub Capcf_Click()

    OrigenCFDI = 0

    NOMCF2.Show

End Sub

Private Sub CapDetImp_Click()
    For i = 2 To 7: sumv(i) = 0: Next i
      Printer.FontBold = False
      nomb_e$ = Printer.FontName
      tama_o = Printer.FontSize
      Printer.FontName = "Ms sans serif"
      Printer.FontSize = 8
      
      TituloCompl
      r1 = 1
      For r = 1 To ConNom1.Rows - 2
       If ConNom1.TextMatrix(r, C_NUM) <> "" Then
            rgtro = ConNom1.TextMatrix(r, C_NUM)
            If rgtro > 0 Then
              Get 14, rgtro, nom_com
              pone = 0: colocar pone, Format(rgtro, "####0"), "####0"
              Printer.Print Tab(8);
              Printer.CurrentX = Printer.CurrentX + pone
              Printer.Print Format(rgtro, "####0");
              Printer.Print Tab(15); Left(ConNom1.TextMatrix(r, C_NOMBRE), 28);
              pone = 0: colocar pone, ConNom1.TextMatrix(r, C_TOTING), z1$
              sumv(2) = sumv(2) + ConNom1.TextMatrix(r, C_TOTING)
              Printer.Print Tab(55);
              Printer.CurrentX = Printer.CurrentX + pone
              Printer.Print Format(ConNom1.TextMatrix(r, C_TOTING), z1$);
              pone = 0: colocar pone, Format(nom_com.ImpTot, z1), z1$
              sumv(3) = sumv(3) + nom_com.ImpTot
              Printer.Print Tab(70);
              Printer.CurrentX = Printer.CurrentX + pone
              Printer.Print Format(nom_com.ImpTot, z1$);
              pone = 0: colocar pone, Format(nom_com.subapl, z1$), z1$
              sumv(4) = sumv(4) + nom_com.subapl
              Printer.Print Tab(85);
              Printer.CurrentX = Printer.CurrentX + pone
              Printer.Print Format(nom_com.subapl, z1$);
              pone = 0: colocar pone, Format(nom_com.CreTot, z1$), z1$
              sumv(5) = sumv(5) + nom_com.CreTot
              Printer.Print Tab(100);
              Printer.CurrentX = Printer.CurrentX + pone
              Printer.Print Format(nom_com.CreTot, z1$);
              DescuentoIsr = nom_com.ImpTot - nom_com.subapl - nom_com.CreTot
              pone = 0: colocar pone, Format(DescuentoIsr, z1$), z1$
              Printer.Print Tab(115);
              If DescuentoIsr > 0 Then
                        Printer.CurrentX = Printer.CurrentX + pone
                        Printer.Print Format(DescuentoIsr, z1$)
                        sumv(6) = sumv(6) + DescuentoIsr
                        Else
                        Printer.Print Tab(130);
                        Printer.CurrentX = Printer.CurrentX + pone
                        Printer.Print Format(DescuentoIsr, z1$)
                        sumv(7) = sumv(7) + DescuentoIsr
              End If
              r1 = r1 + 1
            End If
       End If
     If r1 > 58 Then
        Printer.Print Tab(55); String(170, "-")
        Printer.Print Tab(30); "S U B - T O T A L E S";
        SumasCompl
        Printer.NewPage
        TituloCompl
        r1 = 1
     End If
   Next r
    Printer.Print Tab(55); String(170, "-")
    Printer.Print Tab(30); "T O T A L E S";
    SumasCompl
    Printer.EndDoc
End Sub
Sub SumasCompl()
   For g = 2 To 7
        pone = 0: colocar pone, Format(sumv(g), z1$), z1$
        Printer.Print Tab(25 + (15 * g));
        Printer.CurrentX = Printer.CurrentX + pone
        Printer.Print Format(sumv(g), z1$);
    Next g
    
    Printer.Print
    Printer.Print Tab(55); String(100, "=")
    Printer.CurrentY = Printer.Height - 1600
    Printer.CurrentX = Printer.Width / 2  '
    Printer.Print "-  " & Printer.Page & "  -"   ' Imprimir.

End Sub
Private Sub CapDis_Click()
    final = limite
    Poliza.Show
End Sub

Private Sub CapInrt_Click()
    CapArch_Click 2
                'If IsNumeric(ConNom1.TextMatrix(1, 0)) Then
                    'CnvNom.Show
                    'Else
                    'MsgBox "SELECCIONE LA NOMINA QUE SE CAPTURA"
                'End If
End Sub

Private Sub capturaGeneral_Click()
    generarNominaGeneral
End Sub


Private Sub CFDI2_Click()

    OrigenCFDI = 0

    NOMCF2.Show

End Sub

Private Sub cheajte_Click(Index As Integer)
   Load AJTECH
   AJTECH.Show 1
   
End Sub

Private Sub cheind_Click(Index As Integer)
    colanti = ConNom1.Col
    renati = ConNom1.Row
    If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
    If rgtro > 0 Then
        lin_che ConNom1.Row, ConNom1.RowSel
        ConNom1.SelectionMode = flexSelectionFree
        ConNom1.Col = colanti
        ConNom1.Row = renati
        ConNom1.SetFocus
    Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If

End Sub

Private Sub cheobra_Click(Index As Integer)
    
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
     CommonDialog1.FileName = "CATAUX*.*"
     CommonDialog1.ShowOpen
     Direc_torio = Mid(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
     Open "C:\GconTA\perma.dno" For Random As #7 Len = Len(basico)
     fin_basico = LOF(7) / Len(basico)
     basico.datoarch = Direc_torio
     Put #7, 2, basico
     Close #7
    Exit Sub
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub chetot_Click(Index As Integer)
   colanti = ConNom1.Col
   renati = ConNom1.Row
   ConNom1.Row = 1
    ConNom1.Col = 0
   If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
   If rgtro > 0 Then
            
            lin_che 1, limite
            ConNom1.Col = colanti
            ConNom1.Row = renati
            ConNom1.SetFocus
     Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If

End Sub

Private Sub edicionPersonalCaptura_Click()
    Form4.Show
End Sub

Private Sub EdiNomSel_Click()
    Clipboard.Clear
    ConNom1.Row = 0
    ConNom1.Col = 0
    ConNom1.RowSel = ConNom1.Rows - 1
    ConNom1.ColSel = ConNom1.Cols - 1
End Sub

Private Sub EdNomCopiar_Click()
    Dim textoSeleccionado As String
    Dim i As Integer, j As Integer

    ' Limpiar el portapapeles
    Clipboard.Clear

    ' Inicializar la cadena que contendrá el texto seleccionado
    textoSeleccionado = ""

    ' Iterar sobre las filas seleccionadas
    For i = ConNom1.Row To ConNom1.RowSel
        ' Iterar sobre las columnas seleccionadas
        For j = ConNom1.Col To ConNom1.ColSel
            ' Añadir el texto de la celda a la cadena
            textoSeleccionado = textoSeleccionado & ConNom1.TextMatrix(i, j)
            
            ' Añadir un tabulador si no es la última columna seleccionada
            If j < ConNom1.ColSel Then
                textoSeleccionado = textoSeleccionado & vbTab
            End If
        Next j
        
        ' Añadir una nueva línea si no es la última fila seleccionada
        If i < ConNom1.RowSel Then
            textoSeleccionado = textoSeleccionado & vbCrLf
        End If
    Next i

    ' Colocar el texto seleccionado en el portapapeles
    Clipboard.SetText textoSeleccionado
    
End Sub

Private Sub exportar_Click(Index As Integer)
    FlexGrid_To_Excel ConNom1
End Sub

Public Sub FlexGrid_To_Excel(TheFlexgrid As MSFlexGrid, Optional GridStyle As Integer = 1, Optional WorkSheetName As String)

    Dim objXL As Object
    Dim wbXL As Object
    Dim wsXL As Object
    Dim intRow As Long
    Dim intCol As Long
    Dim j As Long

    On Error GoTo ErrHandler

    'Crear la instancia de Excel
    Set objXL = CreateObject("Excel.Application")

    objXL.Visible = True

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet

    'Nombre de la hoja
    If WorkSheetName <> "" Then
        wsXL.name = WorkSheetName
    End If

    'Exportar datos
    For intRow = 1 To TheFlexgrid.Rows - 1
        j = 1
        For intCol = 0 To 1
            wsXL.Cells(intRow, j).Value = TheFlexgrid.TextMatrix(intRow, intCol)
            j = j + 1
        Next intCol
    Next intRow

    Exit Sub

ErrHandler:
    MsgBox "Necesitas tener instalado Microsoft Excel en tu equipo." & vbCrLf & _
           "Detalle: " & Err.Description, _
           vbExclamation, _
           "Exportar a Excel"

    Set wsXL = Nothing
    Set wbXL = Nothing

    If Not objXL Is Nothing Then
        objXL.Quit
    End If

    Set objXL = Nothing

End Sub
Private Sub Form_Resize()
    ConNom1.Height = Form8.Height * 0.645
    ConNom1.Width = Form8.Width * 0.975
End Sub

Private Sub Form_Unload(Cancel As Integer)
    capnosal_Click 1
    Form1.Show
End Sub

Private Sub GenBan_Click(Index As Integer)
    Form9.Show
End Sub

Private Sub funcionPegar_Click()
    Dim DeAqui As Integer
    Dim RetornoCarro As Long
    Dim InicioCopia As Long
    Dim Temporal1 As String
    Dim temporal As String

    Temporal1 = ""
    temporal = Clipboard.GetText()

    RetornoCarro = ConNom1.Col
    InicioCopia = ConNom1.Row
    DeAqui = InicioCopia
    InIPGr = InicioCopia
    InIPGc = RetornoCarro

    If ConNom1.Rows = ConNom1.Rows - 3 Then
        Exit Sub
    End If

    If temporal <> "" Then
        Clipboard.Clear
        DeAqui = 1

        For i = 1 To Len(temporal)
            Select Case Mid(temporal, i, 1)
                Case Chr(13)
                    ' Verificar si el texto es un número antes de asignarlo
                    If IsNumeric(Mid(temporal, DeAqui, (i - DeAqui))) Then
                        Temporal1 = Temporal1 + ConNom1.TextMatrix(ConNom1.Row, ConNom1.Col) & Chr(13)
                        ConNom1.Text = Mid(temporal, DeAqui, (i - DeAqui))
                        Text2.Text = Mid(temporal, DeAqui, (i - DeAqui))
                        Text2_KeyPress (13)
                        ConNom1.Row = ConNom1.Row + 1
                        
                        If ConNom1.Row >= ConNom1.Rows - 3 Then
                            ConNom1.Row = ConNom1.Rows - 3
                        End If
                        ConNom1.TopRow = 1
                    Else
                        ' Si no es un número, pegar espacio vacío
                        Temporal1 = Temporal1 & vbCrLf
                        ConNom1.Row = ConNom1.Row + 1
                        If ConNom1.Row >= ConNom1.Rows - 3 Then
                            ConNom1.Row = ConNom1.Rows - 3
                        End If
                        ConNom1.TopRow = 1
                    End If
                    DeAqui = i + 1
            End Select
        Next i
        ' Restaurar posición del cursor
        ConNom1.Col = RetornoCarro
    End If

    Clipboard.Clear
    ConNom1.SetFocus

End Sub

Private Sub Importar_Click()
Dim oRS As New ADODB.Recordset
Dim oConn As New ADODB.Connection

    excel.CancelError = True
    On Error GoTo ErrHandler
    excel.ShowOpen
    
    If (excel.FileName <> "") Then
        oConn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "data source=  '" + excel.FileName + "' ; " & "Extended Properties= ""Excel 8.0;HDR=Yes"""
        Set oRS = New ADODB.Recordset
            oRS.CursorLocation = adUseClient
            oRS.CursorType = adOpenStatic
            oRS.LockType = adLockOptimistic
            oRS.Open "SELECT * FROM [Hoja1$]", oConn, , , adCmdText
        Call Llenar_FlexGrid(oConn, oRS)
    End If
    
ErrHandler:
End Sub

Private Sub Llenar_FlexGrid(cn As Connection, Rs As Recordset)
  
On Error GoTo ErrSub
  
    Screen.MousePointer = vbHourglass
    ' Deshabilita el repintado del Flexgrid
    ConNom1.Redraw = False
    ' Mueve el recordset al primer registro
    Rs.MoveFirst
    ' Agrega las filas necesarias en el FlexGRid
    ConNom1.Rows = Rs.RecordCount + 1
    ' Agrega las columnas necesarias
    ConNom1.Cols = Rs.Fields.Count
      
    Dim c As Integer
    ' Selecciona
    ConNom1.Row = 1
    ConNom1.Col = 0
      
    ConNom1.RowSel = ConNom1.Rows - 1
    ConNom1.ColSel = ConNom1.Cols - 1
      
    'Devuelve o establece el contenido de las celdas en _
     una región de FlexGrid seleccionada. No está disponible en tiempo de diseño. _
     ( Esta linea de código es la que carga los registros )
    ConNom1.Clip = Rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    ConNom1.Row = 1
      
    ' habilita nuevamente el Redraw en el control
    ConNom1.Redraw = True
      
    ' Cierra y elimina las referencias ( recordset y la conexión )
    Rs.Close
    Set Rs = Nothing
    cn.Close
    Set cn = Nothing
  
    Screen.MousePointer = vbDefault
  
Exit Sub
  
ErrSub:
' Mensaje de error
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
End Sub

Private Sub minToMax_Click()
    colanti = ConNom1.Col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.Col = 3
    ConNom1.RowSel = limite
    ConNom1.Sort = 3
    ConNom1.Col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus
End Sub

Sub Mnunomobr_Click(Index As Integer)
    Select Case Index
         Case 0
            requisito = 10001
            comp_ti = " "
            impr_nomina requisito
         Case 1
            requisito = 10002
            comp_ti = "de todas las Obras "
            impr_nomina requisito
         Case 2
            requisito = 9000
            comp_ti = "de Oficina Central"
            impr_nomina requisito
         Case 4 To 46
                       
            requisito = Val(Mid(Mnunomobr.Item(Index).Caption, 1, 5))
            Entrada = Mid(Mnunomobr.Item(Index).Caption, 7)
            If RTrim(Entrada) = "Capturar" Then
                    
                    comp_ti = InputBox("Teclee el nombre de la obra " + Str(requisito), "Impresion nomina obra")
                    Else
                    comp_ti = (Mnunomobr.Item(Index).Caption)
            End If
            impr_nomina requisito
     End Select
End Sub
Sub impr_nomina(Clave As Integer)
    For i = 1 To 10
        sum(i) = 0
    Next i
        
        hoja = 0: conta = 0: conta3 = 0: conta1 = 0: conta2 = 0
        nomb_e$ = Printer.FontName
        tama_o = Printer.FontSize
        Printer.FontName = "arial"
        Printer.FontSize = 7
        Printer.FontBold = True
        
        If si_hay = 1 Then
            Printer.Orientation = 2
            lim_conta = Printer.ScaleHeight - 3000
            Aumento1 = 2000
        Else
            Printer.Orientation = 1
            lim_conta = Printer.ScaleHeight - 3000
            Aumento1 = 0
        End If
        
        If ConNom1.TextMatrix(1, 0) <> "" Then rgtro = ConNom1.TextMatrix(1, 0) Else rgtro = 0
        
        If rgtro > 0 Then
            impr_tit
            For r = 1 To limite
                li = r
                rgtro = ConNom1.TextMatrix(li, 0)
                Get #8, rgtro, maestro
                
                Select Case Clave
                    Case 10001
                        GoSub imprim_e
                    Case 10002
                        If maestro.O_1 > 0 Then
                            If maestro.O_1 <> 9000 Then
                                GoSub imprim_e
                            End If
                        End If
                    Case 9000
                        ya_estuvo = 0
                        If maestro.O_2 > 0 Then
                            GoSub imprim_e
                            ya_estuvo = 1
                        End If
                        If maestro.O_1 = 9000 Then
                            If ya_estuvo = 0 Then
                                GoSub imprim_e
                            End If
                        End If
                    Case Else
                        If maestro.O_2 > 0 Then
                        Else
                            If maestro.O_1 = Clave Then
                                If maestro.por_1 = 100 Then
                                    GoSub imprim_e
                                End If
                            End If
                        End If
                End Select
            GoTo salid_a
imprim_e:
            If Printer.CurrentY >= lim_conta Then
                Printer.Print
                tot_sub 0
                Printer.NewPage
                conta = 0
                Printer.CurrentX = 0
                Printer.CurrentY = 0
                actual = 0
                impr_tit
            End If
                  
            valor = rgtro
            uso$ = "####0."
            ancho2 = 0
            colocar ancho2, valor, uso$
            Printer.CurrentX = 0 + ancho2 + Aumento1
            Printer.Print rgtro;
            M_aximo = Printer.CurrentY
            lin_nom
            Return
salid_a:
        
        Next r
        
        Printer.Print: Printer.Print
        tot_sub 1
        Printer.EndDoc
        Printer.Orientation = 1
    Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If
    
    Printer.FontName = nomb_e$
    Printer.FontSize = tama_o
End Sub
Sub locobra()
    ya = 0
    
    For i = 4 To cuentaobra
        If i > 46 Then maestro.O_1 = 9000
        If maestro.O_1 > 0 Then
            If maestro.O_1 = Val(Mid$(Mnunomobr(i).Caption, 1, 5)) Then ya = 1
            Exit Sub
        Else
            maestro.O_1 = 9000
        End If
    Next i
        
    If maestro.O_1 <> 9000 Then
        If maestro.O_1 > 0 Then
            If (ya = 0) And (maestro.por_1 = 100) Then
                cuentaobra = cuentaobra + 1
                aa = String(5 - Len(LTrim(maestro.O_1)), " ")
                Mnunomobr.Item(cuentaobra).Visible = True
                leer_cat
                Mnunomobr.Item(cuentaobra).Caption = aa + LTrim(Str(maestro.O_1)) + " " + Entrada
            End If
        End If
    End If
    
End Sub
Sub checa_fecha(MiFecha, MiDía, Resta)
   MiDía = Weekday(MiFecha)
   If MiDía = 7 Then dia_pago = dia_pago - 1
   If MiDía = 1 Then dia_pago = dia_pago - 2
End Sub

    Sub lin_che(empieza, termina)
        antes = Printer.FontSize
        anntes$ = Printer.FontName
        anchopapel = Printer.Height
        largopapel = Printer.Width
        Printer.Height = 10118.57
        Printer.Width = 12247.2
        Printer.Orientation = 1
        Printer.Font = "sans serif"
        Printer.FontSize = 10
        respuesta = MsgBox("Desea hacer pausa despues de imprimir cada cheque?", vbCritical + vbYesNo, "Impresion Cheques")
       
        If respuesta = vbYes Then pausa = 1 Else pausa = 0
       
        apeajte
        Get 8, 1, formajte
       
        For r = empieza To termina
            If ConNom1.TextMatrix(r, 0) > "" Then
                numerillo = ConNom1.TextMatrix(r, 0)
                Get 2, numerillo, personal
                nombrecillo = RTrim(personal.nom) + " " + RTrim(personal.ape1) + " " + RTrim(personal.ape2)
            Else
                nombrecillo = ConNom1.TextMatrix(r, 1)
            End If
        
            For ppe = 1 To formajte.totr
                Select Case (ppe)
                    Case formajte.fechar
                        Printer.Print Tab(formajte.fechac);
                        Printer.Print dia_pago; " DE "; RTrim$(mm(meselegido)); " de "; empresa.ao;
                    Case formajte.benefr
                        Printer.Print Tab(formajte.benefc);
                        Printer.Print nombrecillo;
                        If ConNom1.TextMatrix(r, C_EXENTA) <> "" Then sacale = ConNom1.TextMatrix(r, C_EXENTA) Else sacale = 0
                        operacion = ConNom1.TextMatrix(r, C_NETO) - sacale
                        Printer.Print Tab(formajte.impnumc); Format(operacion, z2$)
                    Case formajte.impletr
                        Printer.Print Tab(formajte.impletc);
                        gt_bi# = operacion
                        Dinero
                        Printer.Print Feria;
                    Case formajte.concepr
                        Printer.Print Tab(formajte.concepc);
                        Printer.Print Label7.Caption; " "; comp_ti;
                    Case formajte.inicopr
                        Printer.Print Tab(formajte.name);
                        Printer.Print "GASTOS A COMPROBAR";
                        Printer.Print Tab(formajte.debec);
                        Printer.Print Format(operacion, z2$)
                        Printer.Print Tab(formajte.name);
                        Printer.Print ConNom1.TextMatrix(r, 1)
                        Printer.Print Tab(formajte.name);
                        Printer.Print "BANCOS";
                        Printer.Print Tab(formajte.haberc); Format(operacion, z2$)
                        Printer.Print Tab(formajte.name); "CITIBANK"
                    Case formajte.sumasr
                        Printer.Print Tab(formajte.debec);
                        Printer.Print Format(operacion, z2$);
                        Printer.Print Tab(formajte.haberc); Format(operacion, z2$);
                    End Select
                        Printer.Print
            Next ppe
           
            If r < termina Then
                Rem Printer.NewPage
                If pausa = 1 Then
                         respuesta = MsgBox("Acomode papel antes de continuar", vbOKCancel + vbCritical + vbDefaultButton2, "Impresion cheques")
                    If respuesta = vbOK Then
                            Rem nada
                        Else
                            r = termina
                            GoTo saltoimpre
                
                     End If
                  End If
             End If
        Next r
         
saltoimpre:
       
       Close 8
       Printer.EndDoc
       Printer.Height = anchopapel
       Printer.Width = largopapel
       Printer.FontSize = antes
       Printer.FontName = anntes$
       Printer.FontBold = False
    End Sub
Sub lin_rbo(empieza, termina)
On Error GoTo manejadorError:
    nomb_e$ = Printer.FontName
    tama_o = Printer.FontSize

    For r = empieza To termina
        li = r
        Printer.Orientation = 1
        For g = 1 To 2
            If ConNom1.TextMatrix(li, 0) <> "" Then regtro = ConNom1.TextMatrix(li, 0) Else regtro = 0
            If regtro > 0 Then
            recibo
            Printer.CurrentX = 0
            Printer.CurrentY = 0
            Get #2, regtro, personal
            Printer.ForeColor = 0
            Printer.FontName = "Arial"
            Printer.FontSize = 9
            Printer.FontBold = True
            ConNom1.Col = 1
            Printer.CurrentY = 380 + aumento(g)
            Printer.CurrentX = 30
            Printer.Print personal.RFC;
            Printer.CurrentX = 1900
            Printer.Print regtro; " "; ConNom1.TextMatrix(li, 1);
            Printer.CurrentX = 9820
            Printer.Print personal.imss
            Printer.CurrentY = 950 + aumento(g)
        
    For i = 2 To 20
        If i = 12 Then Printer.CurrentY = 950 + aumento(g): Printer.Print: Printer.Print
        ii = i
        If ConNom1.TextMatrix(li, ii) <> "" Then
            Select Case i
               Case 2
                    Printer.CurrentX = 200
                    Printer.Print "Sueldo diario $"; Format(personal.ingr, z1$)
                    Printer.CurrentX = 200
                    Printer.Print "Dias trabajados "; ConNom1.TextMatrix(li, ii)
               Case 3
                    Printer.CurrentX = 200
                    Printer.Print "Salario ";:
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4250
                    Printer.Print "$";
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 4
                    Printer.CurrentX = 200
                    Printer.Print "Hs.Extras Normales ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 5
                    Printer.CurrentX = 200
                    Printer.Print "Aguinaldo ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 6
                    Printer.CurrentX = 200
                    Printer.Print "Participacion de utilidades ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 7
                    Printer.CurrentX = 200
                    Printer.Print "O.F. ";
                    'Printer.Print "Compensacion ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 8
                    Printer.CurrentX = 200
                    Printer.Print "Prima vacacional ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 9
                    Printer.CurrentX = 200
                    Printer.Print "Otras ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 10
                    Printer.CurrentX = 200
                    Rem Printer.Print "Vales Desp.";
                    Printer.Print "Perc.exenta.";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 11
                    Printer.CurrentY = 4400 + aumento(g)
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 4290
                    Printer.Print "$";
                    Printer.CurrentX = 4300 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 12
                    Printer.CurrentX = 5650
                    Printer.Print "I.S.P.T. ";:
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9940
                    Printer.Print "$";
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 13
               
               Printer.CurrentX = 5650
               If empresa.ao < 2008 Then
                    Printer.Print "Cr. al salario ";
                    Else
                    Printer.Print "Sbdio.al Empleo ";
               End If
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9940
                    Printer.Print "$";
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 14
                    Printer.CurrentX = 5650
                    Printer.Print "I.M.S.S.";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 15
                    Printer.CurrentX = 5650
                    Printer.Print "Prestamos ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 16
                    Printer.CurrentX = 5650
                    Printer.Print "Fonacot ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 17
                    Printer.CurrentX = 5650
                    Printer.Print "Pension Alimenticia ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 18
                    Printer.CurrentX = 5650
                    Printer.Print "Otras(infonavit) ";
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 19
                    Printer.CurrentX = 5650
                    Printer.CurrentY = 4400 + aumento(g)
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9940
                    Printer.Print "$";
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
               Case 20
                    Printer.CurrentY = 4750 + aumento(g)
                    Printer.CurrentX = 10
                    Printer.Print LTrim$(Mid$(Label7.Caption, 1, 20));
                    Printer.CurrentX = 5500
                    Printer.Print dia_pago; " de "; RTrim$(mm(meselegido)); " de "; empresa.ao;
                    Printer.CurrentX = 5650
                    colocar ancho2, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = 9940
                    Printer.Print "$";
                    Printer.CurrentX = 9950 + ancho2
                    Printer.Print ConNom1.TextMatrix(li, ii)
            End Select
        End If
    Next i
    
    End If
    Printer.CurrentY = 4400 + aumento(g) - 1500
    Printer.FontUnderline = True
    Printer.CurrentX = 5650
    Get 14, regtro, nom_com
    Printer.Print "Datos Informativos : "
    Printer.FontUnderline = False
    Printer.CurrentX = 5650
    Printer.Print "ISR Retenido..";
    valor$ = Format((nom_com.ImpTot - nom_com.subapl), z1$)
    colocar ancho2, valor$, z1$
    Printer.CurrentX = 5650
    Printer.CurrentX = 8550 + ancho2
    Printer.Print Format((nom_com.ImpTot - nom_com.subapl), z1$)
    Printer.CurrentX = 5650
    If empresa.ao < 2008 Then
        Printer.Print "Cr.Sal.Aplic..";
        Else
        Printer.Print "Subdio E.Aplic..";
    End If
    valor$ = Format(nom_com.CreTot, z1$)
    colocar ancho2, valor$, z1$
    Printer.CurrentX = 8550 + ancho2
    Printer.Print Format(nom_com.CreTot, z1$)
    Printer.CurrentX = 5650
    If empresa.ao < 2008 Then
        Printer.Print "Cr.Sal.Pagado.";
        Else
        Printer.Print "Subdio.Pagado.";
    End If
    valor$ = Format(nom_com.CredNe, z1$)
    colocar ancho2, valor$, z1$
    Printer.CurrentX = 8550 + ancho2
    Printer.Print Format(nom_com.CredNe, z1$)
    
  Next g
    Printer.EndDoc

  Next r
   Printer.FontName = nomb_e
   Printer.FontSize = tama_o

Exit Sub
manejadorError:
    MsgBox (Err.Number & " " & Err.Description)

End Sub
Sub recibo()
On Error GoTo manejadorError
    nomb_e$ = Printer.FontName
    tama_o = Printer.FontSize
    Printer.FontName = "bodoni black"
    aumento(1) = 1000
    aumento(2) = 8500
 For i = 1 To 2
    Printer.ForeColor = 8421376
    Printer.FontTransparent = True
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.DrawWidth = 10
    Printer.Line (0, 0 + aumento(i))-(11200, 5000 + aumento(i)), , B: Rem cuadro general
    Printer.Line (0, 0 + aumento(i))-(11200, 314 + aumento(i)), , B: Rem cuadro rfc,nombre,imss
    Printer.Line (0, 315 + aumento(i))-(11200, 628 + aumento(i)), , B: Rem cuadro espacio rfc,nombre,imss
    Printer.Line (2300, 0 + aumento(i))-(2310, 628 + aumento(i)), , BF: Rem cuadro titulos percepciones, deducciones
    Printer.Line (9800, 0 + aumento(i))-(9810, 628 + aumento(i)), , BF:
    Printer.Line (0, 629 + aumento(i))-(11200, 942 + aumento(i)), , B:
    Printer.Line (4200, 630 + aumento(i))-(4210, 4685 + aumento(i)), , BF: Rem importe percepciones
    Printer.Line (5600, 630 + aumento(i))-(5610, 4685 + aumento(i)), , BF: Rem mitad
    Printer.Line (9800, 630 + aumento(i))-(9810, 5000 + aumento(i)), , BF: Rem importe deducciones
    Printer.Line (0, 4372 + aumento(i))-(11200, 4685 + aumento(i)), , B
    Printer.Line (0, 4686 + aumento(i))-(11200, 5000 + aumento(i)), , B
    Printer.CurrentX = 100
    Printer.CurrentY = 80 + aumento(i)
    Printer.Print " ";
    Rem Printer.ForeColor = 8421376
    Printer.ForeColor = 32896
    Printer.Print "Reg.Fed.Caus/CURP";
    Printer.CurrentX = 4000
    Printer.Print "N  o  m  b  r  e";
    Printer.CurrentX = 10000
    Printer.Print "Reg.IMSS"
    Printer.CurrentY = 650 + aumento(i)
    Printer.CurrentX = 1200
    Printer.Print "P e r c e p c i o n e s";
    Printer.CurrentX = 4500
    Printer.Print "Importe";
    Printer.CurrentX = 6600
    Printer.Print "D e d u c c i o n e s";
    Printer.CurrentX = 10000
    Printer.Print "Importe"
    Printer.CurrentY = 4700 + aumento(i)
    Printer.CurrentX = 3500
    Printer.Print "Fecha de pago: ";
    Printer.CurrentX = 8500
    Printer.Print "N  e  t  o"
    Printer.CurrentY = 5050 + aumento(i)
    antes = Printer.FontSize
    anntes$ = Printer.FontName
    Printer.ForeColor = 32896
    Printer.FontTransparent = True
    Printer.CurrentY = 2000 + aumento(i)
    Printer.FontBold = False
    Printer.FontSize = 40
    MideLetrero = Printer.TextWidth(RTrim(empresa.name))
    MaximoEspacio = Printer.ScaleWidth - 400
    If MideLetrero > (MaximoEspacio) Then Printer.FontSize = Int(MaximoEspacio * 40 / MideLetrero)
    Printer.CurrentX = 1
    pone = (MaximoEspacio) / 2 - (Printer.TextWidth(RTrim(empresa.name)) / 2)
    Printer.CurrentX = pone
    Printer.Print ; RTrim$(empresa.name);
    Printer.FontBold = True
    Printer.FontSize = 8
    Printer.CurrentY = 5050 + aumento(i)
    Printer.CurrentX = 50
    Printer.ForeColor = 8421376
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.CurrentX = 50
    Printer.Print "Recibi de :"; empresa.name
    Printer.CurrentX = 50
    Printer.Print "la  cantidad  indicada  que  cubre  a  la   fecha el importe  de  mi  salario"
    Printer.CurrentX = 50
    Printer.Print "tiempo extra septimo dia y todas las percepciones a que tengo derecho"
    Printer.CurrentX = 50
    Printer.Print "sin que se me adeude alguna cantidad por otro concepto.";
    Printer.CurrentX = 7800
    Printer.FontSize = 10: Printer.Print "F i r m a"
    Printer.Line (4200, 630 + aumento(i))-(4210, 4685 + aumento(i)), , BF: Rem importe percepciones
    Printer.Line (5600, 630 + aumento(i))-(5610, 4685 + aumento(i)), , BF: Rem mitad
    Printer.Line (0, 5000 + aumento(i))-(11200, 5800 + aumento(i)), , B
    Printer.FontSize = antes
    Printer.FontName = anntes$
Next i
    GoTo Salida
Salida:
  Printer.FontName = nomb_e
   Printer.FontSize = tama_o

Exit Sub
manejadorError:
   
   MsgBox (Err.Number & Err.Description)
End Sub

Sub verifica(yavas)
' verifica si existe un monto en los ingresos
    If nomina.dias <> 0 Then yavas = 1
    If nomina.hs_nor <> 0 Then yavas = 1
    If nomina.hs_dbl <> 0 Then yavas = 1
    If nomina.hs_tri <> 0 Then yavas = 1
    If nomina.ispt <> 0 Then yavas = 1
    If nomina.crdsal <> 0 Then yavas = 1
    If nomina.imss <> 0 Then yavas = 1
    If nomina.sueldo <> 0 Then yavas = 1
    If nomina.hs_nor <> 0 Then yavas = 1
    If nomina.hs_dbl <> 0 Then yavas = 1
    If nomina.hs_tri <> 0 Then yavas = 1
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
End Sub


Sub eliminacion()
    Dim antecol As Long
    Dim anteren As Long
    Dim re As Long

    antecol = ConNom1.Col
    anteren = ConNom1.Row

    re = 1

    Do While re <= limite

        ' Revisar la columna Banamex, que sí es un índice válido.
        If Trim$(ConNom1.TextMatrix(re, C_BANAMEX)) = "" Then
            ConNom1.RemoveItem re
            limite = limite - 1
        Else
            re = re + 1
        End If

    Loop

    If limite < 1 Then
        ConNom1.Row = 1
    ElseIf anteren <= limite Then
        ConNom1.Row = anteren
    Else
        ConNom1.Row = limite
    End If

    If antecol >= 0 And antecol < ConNom1.Cols Then
        ConNom1.Col = antecol
    End If

    sumavert
End Sub

Sub tot_sub(fin)
     
     actual = Printer.CurrentY + 220
    If si_hay = 1 Then
            Printer.Line (0 + Aumento1, valtit)-(12200 + Aumento1, actual + 220), , B
            Printer.Line (0 + Aumento1, actual - 80)-(12200 + Aumento1, actual + 220), , B
    Else
            Printer.Line (0 + Aumento1, valtit)-(11400 + Aumento1, actual + 220), , B
            Printer.Line (0 + Aumento1, actual - 80)-(11400 + Aumento1, actual + 220), , B
    End If
     
     Printer.CurrentY = actual
     Printer.CurrentX = (120 * 5) + Aumento1
     Printer.Print Format(conta3, "###0");
     Printer.CurrentX = (120 * 10) + Aumento1
    
    If fin = 1 Then
        Printer.Print "Totales ";
    Else
        Printer.Print "Sub-totales ";
    End If
     retorno = 2040
     If si_hay = 1 Then
           fin_A = 10
           Else
           fin_A = 9
     End If
     For i1 = 1 To fin_A
      If (i1 = 4) Or (i1 = 9) Then retorno = retorno + 360
      Printer.CurrentX = retorno + (i1 * 7 * 120) + Aumento1
      If i1 > 1 Then
      If sum(i1) <> 0 Then
         valor$ = Format(Str$(sum(i1)), z1$)
         pone = 0: colocar pone, valor$, z1$
         Printer.CurrentX = Printer.CurrentX + pone
         Printer.Print valor$;
      End If
      End If
      Next i1
      Printer.Print: Printer.Print: Printer.Print
      If fin = 1 Then
        Printer.CurrentX = 420 * 5
        If Option2 = False Then
            Printer.Print "Percepciones : 1.Hs.Extras 2.O.F. 3.P.vacacional 4.Otras 5.Perc.exenta."
        Else
            Printer.Print "Percepciones :"; RTrim(Label7.Caption); " 2.O.F. 3.P.vacacional 4.Otras 5.Perc.exenta."
        End If
        Printer.CurrentX = 420 * 5
        Printer.Print "Deducciones  : 1.Prestamos 2.Fonacot 3.Pension Alimenticia 4.Otras "
      End If
End Sub
Sub hsextras()
    dato_ent = 0
    For he = 4 To 6:
        ii1 = he
        If ConNom1.TextMatrix(li, ii1) <> "" Then
            dato_ent = dato_ent + ConNom1.TextMatrix(li, ii1)
        End If
    Next he
    
    If dato_ent > 0 Then
        valor$ = Format(Str$(dato_ent), z1$)
        pone = 0: colocar pone, valor$, z1$
        Printer.CurrentX = 4560 + pone + Aumento1
        Printer.Print valor$; "(1)";
        sum(3) = sum(3) + dato_ent
        Detener = 1
    End If
       
End Sub
Sub lin_nom()
        
    sec(1) = 2: sec(2) = 3
    sec(3) = 4: sec(4) = 11
    sec(5) = 12: sec(6) = 13
    sec(7) = 14: sec(8) = 15
    sec(9) = 20
        
    Printer.CurrentX = (120 * 7) + Aumento1
    Printer.Print RTrim$(Mid$(ConNom1.TextMatrix(li, 1), 1, 27));
        
    conta3 = conta3 + 1
    conta = conta + 1
    retorno = 2040: gi = 1
    Detener = 0: PrincIngr = Printer.CurrentY
    Detener1 = 0: PrincDesc = Printer.CurrentY
      
    For i = 1 To 9
        ii = sec(i)
        If (i = 4) Or (i = 9) Then retorno = retorno + 360
            Printer.CurrentX = retorno + (i * 7 * 120) + Aumento1
            If i = 3 Then
                hsextras
            Else
                If ConNom1.TextMatrix(li, ii) <> "" Then
                    pone = 0: colocar pone, ConNom1.TextMatrix(li, ii), z1$
                    Printer.CurrentX = Printer.CurrentX + pone: Rem
                    Printer.Print ConNom1.TextMatrix(li, ii);
                    If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
                    If i = 8 Then Printer.Print "(1)";: Detener1 = 1
                    If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
                    If i = 9 Then
                        If si_hay = 1 Then
                          If IsNumeric(ConNom1.TextMatrix(li, 10)) Then
                             operacion = ConNom1.TextMatrix(li, ii) - ConNom1.TextMatrix(li, 10)
                             sum(10) = sum(10) + operacion
                             opera$ = Format(operacion, z2$)
                             pone = 0: colocar pone, opera$, z1$
                             Printer.CurrentX = Printer.CurrentX + pone
                             Printer.Print Format(operacion, z1$);
                             If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
                             Else
                             operacion = ConNom1.TextMatrix(li, ii)
                             sum(10) = sum(10) + operacion
                             opera$ = Format(operacion, z2$)
                             pone = 0: colocar pone, opera$, z1$
                             Printer.CurrentX = Printer.CurrentX + pone
                             Printer.Print Format(operacion, z1$);
                             If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY

                           End If
                        End If
                    End If
                    sum(i) = sum(i) + ConNom1.TextMatrix(li, ii)
                End If
         End If
      Next i
        
        If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
        If ((Detener = 1) And (Detener1 = 1)) Then
                Printer.Print: Detener1 = 0:
                If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
                Detener = 0: PrincIngr = Printer.CurrentY
                PrincDesc = Printer.CurrentY
        End If
           
        If ((Detener = 1) And (Detener1 = 0)) Then
                Detener = 0: PrincDesc = Printer.CurrentY
                Printer.Print: PrincIngr = Printer.CurrentY
            
                If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY

        End If
        If ((Detener = 0) And (Detener1 = 1)) Then
                Detener1 = 0: PrincIngr = Printer.CurrentY
                Printer.Print: PrincDesc = Printer.CurrentY
                
                If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
        End If
        
        retorno = PrincIngr
        Printer.CurrentY = retorno
        
        conta1 = 0: Detener = 1
      For i = 7 To 10
         ii = i
         If ConNom1.TextMatrix(li, ii) <> "" Then
            Detener = 0
            Printer.CurrentX = 4560 + Aumento1
            pone = 0: colocar pone, ConNom1.TextMatrix(li, ii), z1$
            Printer.CurrentX = Printer.CurrentX + pone
            Printer.Print ConNom1.TextMatrix(li, ii); "("; LTrim$(Str$(i - 5)); ")"
            
            If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
            sum(3) = sum(3) + ConNom1.TextMatrix(li, ii)
            conta1 = conta1 + 1
        End If
       Next i
       actual = Printer.CurrentY
       If actual > M_aximo Then M_aximo = actual
       Printer.CurrentY = PrincDesc
       conta2 = 0: gi = 0
       
       For i = 16 To 18
         ii = gi + i
         
         If ConNom1.TextMatrix(li, ii) <> "" Then
            Detener = 0
            Printer.CurrentX = 9120 + Aumento1
            pone = 0: colocar pone, ConNom1.TextMatrix(li, ii), z1$
            Printer.CurrentX = Printer.CurrentX + pone
            Printer.Print ConNom1.TextMatrix(li, ii); "("; LTrim$(Str$(i - 14)); ")"
            
            If M_aximo < Printer.CurrentY Then M_aximo = Printer.CurrentY
            
            sum(8) = sum(8) + ConNom1.TextMatrix(li, ii)
            conta2 = conta2 + 1
        End If
        
        If conta2 > conta1 Then
                conta = conta + conta2
                conta1 = 0: conta2 = 0
                Else
                conta = conta + conta1
                conta1 = 0: conta2 = 0
        End If
        
        If M_aximo > Printer.CurrentY Then Printer.CurrentY = M_aximo
        
       Next i
       If Detener = 1 Then Detener = 0: Printer.Print
       If actual > Printer.CurrentY Then Printer.CurrentY = actual Else actual = Printer.CurrentY
       
End Sub
Sub impr_tit()
    fontviejo = Printer.FontSize
    Printer.FontSize = 10
    Printer.Print
    Printer.Print
    Printer.Print
    ancho1 = Int(Printer.TextWidth(RTrim$(empresa.name)) / 2)
    Printer.CurrentX = (45 * 120) - ancho1 + Aumento1
    Printer.Print empresa.name;
    Printer.Print
    ancho1 = Int(Printer.TextWidth(LTrim$(Label7.Caption) + comp_ti) / 2)
    ancho2 = (45 * 120) - ancho1
    Printer.CurrentX = ancho2 + Aumento1
    Printer.Print Label7.Caption; " "; comp_ti;
    Printer.CurrentX = (120 * 78) + Aumento1: hoja = hoja + 1
    Printer.Print "Hoja .... "; Format(hoja, "####0")
    
    If si_hay = 1 Then
        Printer.Line (0 + Aumento1, Printer.CurrentY)-(12200 + Aumento1, Printer.CurrentY + 50), , BF
        Else
        Printer.Line (0 + Aumento1, Printer.CurrentY)-(11400 + Aumento1, Printer.CurrentY + 50), , BF
    End If
    retorno = Printer.CurrentY
    valtit = Printer.CurrentY
    Printer.CurrentY = retorno
    Printer.FontSize = fontviejo
    Printer.Print
    Printer.CurrentX = (120 * 3) + Aumento1
    Printer.Print "No.";
    Printer.CurrentX = (120 * 10) + Aumento1
    Printer.Print " N o m b r e";
    Printer.CurrentX = (120 * 30) + Aumento1
    Printer.Print "P  E  R  C  E  P  C  I  O  N  E  S";
    Printer.CurrentX = (120 * 58) + Aumento1
    Printer.Print "D  E  D  U  C  C  I  O  N  E S"
    retorno = Printer.CurrentY
    Printer.CurrentY = retorno
    Printer.Print
    Printer.CurrentX = (120 * 28) + Aumento1
    Printer.Print "Dias T.";
    Printer.CurrentX = (120 * 34) + Aumento1
    Printer.Print " Sueldo";
    Printer.CurrentX = (120 * 42) + Aumento1
    Printer.Print "Otras Cl.";
    Printer.CurrentX = (120 * 51) + Aumento1
    Printer.Print "Total";
    Printer.CurrentX = (120 * 58) + Aumento1
    Printer.Print "I.S.P.T.";
    Printer.CurrentX = (120 * 66) + Aumento1
    If empresa.ao < 2008 Then
        Printer.Print "Cr.Sal.";
        Else
        Printer.Print "Subdio.";
    End If
    Printer.CurrentX = (120 * 73) + Aumento1
    Printer.Print "I.M.S.S.";
    Printer.CurrentX = (120 * 80) + Aumento1
    Printer.Print "Otras Cl.";
    Printer.CurrentX = (120 * 86) + Aumento1
    Printer.Print "Pago Neto";
    If si_hay = 1 Then
       Printer.CurrentX = (120 * 94) + Aumento1
       Printer.Print "  Cheque";
    End If
    Printer.Print
    If si_hay = 1 Then
        Printer.Line (0 + Aumento1, Printer.CurrentY)-(12200 + Aumento1, Printer.CurrentY + 50), , BF
        Else
        Printer.Line (0 + Aumento1, Printer.CurrentY)-(11400 + Aumento1, Printer.CurrentY + 50), , BF
    End If
    Printer.Print
End Sub

Sub ida()
    For i = 2 To 22
        ConNom1.Col = i
        If ConNom1.Text <> "" Then
            sum(i) = ConNom1.Text
        Else
            sum(i) = 0
        End If
    Next i
    
    ConNom1.Col = 0
    sum(0) = ConNom1.Text

End Sub
Sub ida_1()
    For i = 2 To 22: ConNom1.Col = i
        If ConNom1.Text <> "" Then
            sumv(i) = ConNom1.Text
        Else
            sumv(i) = 0
        End If
    Next i
    
ConNom1.Col = 0
sumv(0) = ConNom1.Text
End Sub
Sub regre()
         ConNom1.Col = 0
         ConNom1.Text = Format(sumv(0), "###0")
         ConNom1.Col = 1
         ConNom1.Text = compa$
         For i = 2 To 20: ConNom1.Col = i
            If sumv(i) <> 0 Then
              ConNom1.Text = Format(sumv(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i
         For i = 21 To 22: ConNom1.Col = i
            If sumv(i) <> 0 Then
              ConNom1.Text = sumv(i)
              Else
              ConNom1.Text = ""
            End If
         Next i
End Sub
Sub regre_1()
    ConNom1.Col = 0
    ConNom1.Text = Format(sum(0), "###0")
         For i = 2 To 20: ConNom1.Col = i
              If sum(i) <> 0 Then
              ConNom1.Text = Format(sum(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i
         For i = 21 To 22: ConNom1.Col = i
              If sum(i) <> 0 Then
              ConNom1.Text = Format(sum(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i

End Sub
Sub datoper()
On Error GoTo handler
      ruta = ConNom1.Col

    If ConNom1.Text = "" Then
        dato_ent = 0
    Else
        dato_ent = ConNom1.Text
    End If

      ConNom1.Col = C_FLAG_ISR: If ConNom1.Text <> "" Then si_impto = ConNom1.Text Else si_impto = 0
      ConNom1.Col = C_FLAG_IMSS: If ConNom1.Text <> "" Then si_imss = ConNom1.Text Else si_imss = 0
      dato_sal = 0: rgtro = 0
      ConNom1.Col = C_NUM

      If ConNom1.Text <> "" Then
          regtro = ConNom1.Text
          Get 2, regtro, personal
          Get 8, regtro, maestro
      End If
      ConNom1.Col = ruta

Exit Sub

handler:
Close 2: Close 8
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    ddm = LOF(8) / Len(maestro)
End Sub


Sub checanum(camcol, datoret As Currency)
   
   antecol = ConNom1.Col
   ConNom1.Col = camcol
   If ConNom1.Text = "" Then
      datoret = 0
   Else: datoret = ConNom1.Text
   End If
   
   ConNom1.Col = antecol
End Sub
Sub checar()
    Dim esLiq As Boolean
    Dim utilidad As Currency

    esLiq = (g_TipoNominaActiva = tnLiquidacionFiniquito)

    datoper
    If regtro > 0 Then
      Select Case ConNom1.Col

         '---------------------------------------------------------- dias
         Case C_DIAS
              diat = dato_ent
              If diat > 0 Then
                  ConNom1.Text = Format(diat, "##0.00")
                  ConNom1.Col = C_SUELDO
                  ConNom1.Text = Format(diat * personal.ingr, z1$)
              Else
                  ConNom1.Text = ""
                  ConNom1.Col = C_SUELDO
                  ConNom1.Text = ""
                  If Not esLiq Then
                      ConNom1.Col = C_CONC7
                      ConNom1.Text = ""
                      ConNom1.Col = C_CONC8
                      ConNom1.Text = ""
                  End If
              End If
              If Not esLiq Then
                  If personal.viat > 0 Then
                      ConNom1.Col = C_CONC7
                      ConNom1.Text = Format(diat * personal.viat, z1$)
                  End If
                  If personal.otras > 0 Then
                      ConNom1.Col = C_OTRAS
                      ConNom1.Text = Format(diat * personal.otras, z1$)
                  End If
              End If
              ConNom1.Col = ruta

         '-------------------------------------------------------- sueldo
         Case C_SUELDO
              MsgBox "El sueldo se modifica editando su archivo de personal", vbCritical, "Captura de Nomina"
              Text2.Text = valcelant

         '------------------------------ hs.Norm / hs.Dobles / hs.Triples
         ' (en Liq estas constantes valen -1 y el Case nunca se activa)
         Case C_CONC4 To C_CONC6
              hs_ext = personal.ingr / 8

              If N_ormal = 0 Then
                  dato_sal = dato_ent * hs_ext
                  ConNom1.Col = ruta
                  ConNom1.Text = Format(dato_sal, z1$)
              Else
                  Select Case ruta
                     Case C_CONC4
                        dato_sal = dato_ent
                        ConNom1.Col = ruta
                        ConNom1.Text = Format(dato_sal, z1$)
                     Case C_CONC5
                        If N_ormal = 1 And ConNom1.Col = C_CONC5 Then
                            If dato_ent > exento Then
                                dato_sal = dato_ent - exento
                                ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format(exento, z1)
                                Text2.Text = dato_sal
                                ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = Format(dato_sal, z1)
                            Else
                                ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format(dato_ent, z1)
                                Text2.Text = 0
                                ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = Format(0, z1)
                            End If
                        End If
                     Case C_CONC6
                        dato_sal = dato_ent
                        ConNom1.Col = ruta
                        ConNom1.Text = Format(dato_sal, z1$)

                        utilidad = ConNom1.TextMatrix(ConNom1.Row, C_CONC6)
                        utilidad = utilidad - exento
                        ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = Format(utilidad, z1)
                        ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format(exento, z1)

                        If utilidad < 0 Then
                            ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = Format(0.01, z1)
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format(dato_sal, z1)
                        End If
                  End Select
              End If

         '------------------------------------- O.F./BONO  y  Otras
         Case C_CONC7, C_OTRAS
              If esLiq Then
                  ' Liq: captura libre, sin avisos y SIN tocar personal.dno
                  ConNom1.Col = ruta
                  If dato_ent <> 0 Then
                      ConNom1.Text = Format(dato_ent, z1$)
                  Else
                      ConNom1.Text = ""
                  End If
              Else
                  aoingr = Val(Mid$(personal.fal, 7, 4))
                  If aoingr < 1900 Then
                      antig = 1
                  Else
                      antig = CalcularAntiguedad(personal.fal)
                  End If
                  facto = 0
                  factor antig, facto

                  dato_sal = 0
                  checanum C_DIAS, dato_sal
                  Select Case ruta
                     Case C_CONC7
                        If dato_ent = 0 Or dato_sal = 0 Then
                            personal.viat = 0
                        Else
                            personal.viat = dato_ent / dato_sal
                        End If
                     Case C_OTRAS
                        If dato_ent = 0 Or dato_sal = 0 Then
                            personal.otras = 0
                        Else
                            personal.otras = dato_ent / dato_sal
                        End If
                  End Select

                  personal.integrado = (personal.ingr * facto) + personal.viat + personal.otras
                  If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
                  MsgBox "El salario integrado en el archivo de personal fue cambiado", vbExclamation
                  Put 2, regtro, personal
                  ConNom1.Col = ruta: ConNom1.Text = Format(dato_ent, z1$)
              End If

         '------------------------------------------------------ ISPT
         Case C_ISPT
              respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
              If respuesta = vbYes Then
                  si_impto = 1
                  ConNom1.Text = Format(dato_ent, z1$)
                  ConNom1.Col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.Col = C_SUBEMP: ConNom1.Text = ""
                  ConNom1.Col = ruta
              Else
                  si_impto = 0
                  ConNom1.Col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.Col = ruta
                  ConNom1.Text = Format(valcelant, z1$)
              End If

         '------------------------------------------------ Sub.P/Empl.
         Case C_SUBEMP
              respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
              If respuesta = vbYes Then
                  si_impto = 1
                  ConNom1.Text = Format(dato_ent, z1$)
                  ConNom1.Col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.Col = C_ISPT: ConNom1.Text = ""
                  ConNom1.Col = ruta
              Else
                  si_impto = 0
                  ConNom1.Col = C_FLAG_ISR: ConNom1.Text = 0
                  ConNom1.Col = ruta
                  ConNom1.Text = Format(valcelant, z1$)
              End If

         '------------------------------------------------------- IMSS
         Case C_IMSS
              respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
              If respuesta = vbYes Then
                  si_imss = 1
                  ConNom1.Text = Format(dato_ent, z1$)
                  ConNom1.Col = C_FLAG_IMSS: ConNom1.Text = 1
                  ConNom1.Col = ruta
              Else
                  si_imss = 0
                  ConNom1.Col = C_FLAG_IMSS: ConNom1.Text = 0
                  ConNom1.Col = ruta
                  ConNom1.Text = Format(valcelant, z1$)
              End If

         '----------------------------------------------------- Banamex
         Case C_BANAMEX
              Close 12

              Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
              fincl = LOF(12) / Len(Clbnx)
              If fincl < 1 Then

                  For W = 1 To Dm: Get 12, W, Clbnx
                      Clbnx.Q1 = "0"
                      Put 12, W, Clbnx
                  Next W
              End If
              If ConNom1.Text = "" Then
                  Clbnx.Q1 = 0
              Else
                  Clbnx.Q1 = LTrim(ConNom1.Text)
              End If
              regtro = ConNom1.TextMatrix(ConNom1.Row, C_NUM)
              Put 12, regtro, Clbnx
              Close 12

         '---------------------- resto de columnas (incluye los conceptos
         '---------------------- de Liq, la Exenta y los totales)
         Case Else
              ConNom1.Col = ruta
              If dato_ent <> 0 Then
                  ConNom1.Text = Format(dato_ent, z1$)
              Else
                  ConNom1.Text = ""
              End If
      End Select

      If ConNom1.TextMatrix(ConNom1.Row, C_FLAG_ESP) = "1" Then si_imss = 1
      li = ConNom1.Row

      sumah si_impto, si_imss

    End If

    Exit Sub
End Sub

Sub sumavert()
Dim Vw As Integer
    ProgressBar1.Min = 0: ProgressBar1.Max = ConNom1.Rows: ProgressBar1.Value = 0: ProgressBar1.Visible = True
    colanti = ConNom1.Col: renati = ConNom1.Row
    sum(1) = 0: sum(2) = 0: sum(3) = 0: Vw = 0

    For late = 3 To C_NETO
        sumv(late) = 0
    Next late

    For li = 1 To limite: Vw = Vw + 1
        ProgressBar1.Value = li
        ConNom1.Row = li
        For late = 3 To C_NETO
            ii = late
            If ConNom1.TextMatrix(li, ii) <> "" Then
                sumv(late) = sumv(late) + ConNom1.TextMatrix(li, ii)
            End If
        Next late
    Next li

    li = limite + 1
    ConNom1.TextMatrix(li, C_NOMBRE) = "Empleados... " + Str(Vw) + " S u m a s ....."

    For late = 3 To C_NETO
        ii = late
        If sumv(late) <> 0 Then
            ConNom1.TextMatrix(li, ii) = Format(sumv(late), z1$)
        Else
            ConNom1.TextMatrix(li, ii) = ""
        End If
    Next late

    ConNom1.Col = colanti: ConNom1.Row = renati
    ProgressBar1.Visible = False
End Sub


Sub sumah(si_impto1 As Integer, si_imss1 As Integer)
    Dim cPV As Long

    sum(1) = 0
    sum(2) = 0
    sum(3) = 0
    colanti = ConNom1.Col

    ' Columna de prima vacacional (para integrar el IMSS)
    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        cPV = C_LQ_PRIMAVAC
    Else
        cPV = C_CONC8
    End If

    ' Base de ISR: Sueldo .. Otras (la exenta NO entra)
    For late = C_PERC_INI To C_PERC_FIN
        If ConNom1.TextMatrix(li, late) <> "" Then
            sum(1) = sum(1) + ConNom1.TextMatrix(li, late)
        End If
    Next late

    rgtro = ConNom1.TextMatrix(li, C_NUM)

    Acum_Doble

    If si_impto1 = 0 Then
        impto = 0
        cr_sal = 0
        If cal_anual = 0 Then
            If sum(1) > 0 Then
                impto = 0
                calculoRetencion sum(1), impto, 1, regtro
            End If
        Else
            aoalta = Val(Mid(personal.fal, 7, 4))
            If (aoalta = empresa.ao) And (Tot_dias < 330) Then
                impto = 0
                calculoRetencion sum(1), impto, 1, regtro
            Else
                rgtro = ConNom1.TextMatrix(li, C_NUM)
                aoalta = Val(Mid(personal.fal, 7, 4))
                If (aoalta = empresa.ao) And (Tot_dias < 330) Then
                    impto = 0
                    calculoRetencion sum(1), impto, 1, regtro
                Else
                    impto = 0
                    calc_anual sum(1), impto, 1
                    If impto < 0 And ArAcum.DImpret > 0 Then GoTo BRTodo
                End If
            End If
        End If

BRTodo:
        If impto <> 0 And Option3 = True Then
            ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)
        Else
            ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)
            If Tot_dias > 330 And cal_anual = 1 Then
            GoTo saltaAnual:
        End If

            If impto <> 0 And Option4 = True Then
                mesNomina1 = Trim(Left(Combo1.Text, 3))
                AnoNomina1 = Right(Trim(Label7.Caption), 4)
                Close 39
                Open CStr(Trim(Form1.Dir1)) + "\" + mesNomina1 + "1" + AnoNomina1 + ".NOM" For Random As 39 Len = Len(nomina)
                Get 39, regtro, nomina
                ImpuestoAnteriorRetenido = nomina.ispt
                ConNom1.TextMatrix(li, C_ISPT) = Format((impto - ImpuestoAnteriorRetenido), z1$)
                Close 39
            End If
saltaAnual:
        End If
        If impto = 0 Then ConNom1.TextMatrix(li, C_ISPT) = ""

    End If

    If si_imss1 = 0 Then
        checanum C_DIAS, dato_sal
        If IsNumeric(ConNom1.TextMatrix(li, C_DIAS)) Then
            diaseg = ConNom1.TextMatrix(li, C_DIAS)
        End If
        seguro = 0
        If IsNumeric(ConNom1.TextMatrix(li, cPV)) Then
            p_vacacional = ConNom1.TextMatrix(li, cPV)
        Else
            p_vacacional = 0
        End If

        inte_grar

        personal.integrado = integrado
        ' calcula el imss
        imss personal.integrado, seguro, diaseg

        ConNom1.TextMatrix(li, C_IMSS) = Format(seguro, z1$)
        If seguro = 0 Then ConNom1.TextMatrix(li, C_IMSS) = ""
    End If

    ' CALCULO ESPECIAL ****************************************************
    ' *********************************************************************
    If ConNom1.TextMatrix(li, C_FLAG_ESP) = "1" Then
        normal_1 = (sum(1) / 365 * 30.4)
        If (normal_1 > 0) And (ConNom1.TextMatrix(li, C_FLAG_ISR) = "0") Then
            rgtro = ConNom1.TextMatrix(li, C_NUM)
            Get #2, regtro, personal
            Normal = personal.ingr * 30
            DifImptos.CalcDoble = 0

            ' Funcion calcula mensual
            calculo_compl2 Normal, impto, 1

            DifImptos.ImpTIni = nom_com.ImpTot
            DifImptos.SubAIni = nom_com.subapl
            DifImptos.SubNIni = nom_com.subNap
            impto_1 = impto
            normal_2 = normal_1 + Normal
            impto = 0
            calculo_compl2 normal_2, impto, 1
            impto_2 = impto

            If (impto_1 = 0) And (impto_2 > 0) Then
                DifImptos.CalcDoble = 1
                ' Funcion que hace el calculo anual dos
                calculo_compl2 Normal, impto, 1
                DifImptos.CalcDoble = 0
                impto_1 = impto + creere
                DifImptos.ImpTIni = nom_com.ImpTot
                DifImptos.SubAIni = nom_com.subapl
                DifImptos.SubNIni = nom_com.subNap
            End If

            DifImptos.ImpTFin = nom_com.ImpTot
            DifImptos.SubAFin = nom_com.subapl
            DifImptos.SubNFin = nom_com.subNap
            DifImptos.ImpTIni = 0
            DifImptos.SubAIni = 0
            DifImptos.SubNIni = 0
            impto_3 = impto_2 - impto_1

            porc_apli = CCur((impto_3 / normal_1))
            impto = sum(1) * porc_apli

            If impto < 0 Then impto = 0
            nom_com.ImpTot = impto
            nom_com.subapl = 0
            nom_com.subNap = 0
            Put 14, regtro, nom_com
            ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)
        Else

            If impto < 0 Then impto = 0
            nom_com.ImpTot = impto
            nom_com.subapl = 0
            nom_com.subNap = 0
            Put 14, regtro, nom_com

        End If
        ConNom1.TextMatrix(li, C_SUBEMP) = ""
        ConNom1.TextMatrix(li, C_IMSS) = ""
    End If

    ' La exenta se suma solo para el total de ingresos (no entra a la base de ISR)
    ConNom1.Col = C_EXENTA
    If ConNom1.TextMatrix(li, C_EXENTA) <> "" Then sum(1) = sum(1) + ConNom1.TextMatrix(li, C_EXENTA)
    If sum(1) Then
        ConNom1.TextMatrix(li, C_TOTING) = Format(sum(1), z1$)
    Else
        ConNom1.TextMatrix(li, C_TOTING) = ""
    End If
    For late = C_DED_INI To C_DED_FIN
        If ConNom1.TextMatrix(li, late) <> "" Then
            sum(2) = sum(2) + ConNom1.TextMatrix(li, late)
        End If
    Next late
    If sum(2) <> 0 Then
        ConNom1.TextMatrix(li, C_TOTDED) = Format(sum(2), z1$)
    Else
        ConNom1.TextMatrix(li, C_TOTDED) = ""
    End If
    sum(3) = sum(1) - sum(2)
    If sum(3) <> 0 Then
        ConNom1.TextMatrix(li, C_NETO) = Format(sum(3), z1$)
    Else
        ConNom1.TextMatrix(li, C_NETO) = ""
    End If
    ConNom1.Col = colanti
End Sub

Sub inte_grar()

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then Exit Sub
     aoingr = Val(Mid$(personal.fal, 7, 4))
     If aoingr < 1900 Then
        Select Case aoingr
           Case Is > 50
           aoingr = aoingr + 1900
           Case Else
           aoingr = aoingr + 2000
        End Select
        antig = CalcularAntiguedad(personal.fal)
        Else
        antig = CalcularAntiguedad(personal.fal)
     End If
     
     facto = 0
     factor antig, facto
     
      exe_nto = 0
     If IsNumeric(ConNom1.TextMatrix(li, C_EXENTA)) Then
            exe_nto = ConNom1.TextMatrix(li, C_EXENTA) - (empresa.sm * diaseg * 0.4)
        If exe_nto < 0 Then exe_nto = 0
                    
     End If
     If diaseg > 0 Then
        Dim viat_per As Currency
        Dim otras_per As Currency
        
        viat_per = 0
        otras_per = 0
        If IsNumeric(ConNom1.TextMatrix(li, C_CONC7)) Then viat_per = ConNom1.TextMatrix(li, C_CONC7)
        If IsNumeric(ConNom1.TextMatrix(li, C_OTRAS)) Then otras_per = ConNom1.TextMatrix(li, C_OTRAS)
        
        Dim base_integ As Currency
        base_integ = sum(1) - viat_per - otras_per - p_vacacional + exe_nto
        
        integrado = (base_integ / diaseg * facto) + (viat_per / diaseg) + (otras_per / diaseg)
        If personal.integrado > (empresa.sm * 25) Then personal.integrado = (empresa.sm * 25)
        personal.integrado = integrado
        Put 2, regtro, personal
     End If
     
End Sub

Private Sub PonCel(ByVal f As Long, ByVal c As Long, ByVal v As Currency, Optional ByVal fmt As String = "")
    If fmt = "" Then fmt = z1$
    If v <> 0 Then
        ConNom1.TextMatrix(f, c) = Format(v, fmt)
    Else
        ConNom1.TextMatrix(f, c) = ""
    End If
End Sub

Sub carganom()
    li = ConNom1.Row
    ingresos = 0
    deducciones = 0
    neto = 0

    ' Nombre completo del empleado
    ConNom1.TextMatrix(li, C_NOMBRE) = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)

    PonCel li, C_DIAS, nomina.dias, "##0.00"
    PonCel li, C_SUELDO, nomina.sueldo

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        ' ---- Liquidacion / Finiquito: un concepto por columna ----
        PonCel li, C_LQ_SDOVAC, nomina.viaticos
        PonCel li, C_LQ_PRIMAVAC, nomina.pvac
        PonCel li, C_LQ_AGUI, nomina.aguin
        PonCel li, C_LQ_ANTIG, nomina.hs_dbl
        PonCel li, C_LQ_COMPENS, nomina.hs_nor
        PonCel li, C_LQ_INDEMN, nomina.hs_tri
        PonCel li, C_LQ_BONO, nomina.hsnor
        ingresos = nomina.sueldo + nomina.viaticos + nomina.pvac + nomina.aguin + _
                   nomina.hs_dbl + nomina.hs_nor + nomina.hs_tri + nomina.hsnor
    Else
        ' ---- Normal y especiales PTU / Aguinaldo / Bono: igual que siempre ----
        PonCel li, C_CONC4, nomina.hs_nor
        If N_ormal = 1 Then
            PonCel li, C_CONC5, nomina.aguin
        Else
            PonCel li, C_CONC5, nomina.hs_dbl
        End If
        If N_ormal = 1 Then
            PonCel li, C_CONC6, nomina.ptu
        Else
            PonCel li, C_CONC6, nomina.hs_tri
        End If
        PonCel li, C_CONC7, nomina.viaticos
        PonCel li, C_CONC8, nomina.pvac
        ingresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + _
                   nomina.aguin + nomina.ptu + nomina.viaticos + nomina.pvac
    End If

    PonCel li, C_OTRAS, nomina.otras
    PonCel li, C_EXENTA, nomina.exentos
    ingresos = ingresos + nomina.otras + nomina.exentos
    PonCel li, C_TOTING, ingresos

    PonCel li, C_ISPT, nomina.ispt
    PonCel li, C_SUBEMP, nomina.crdsal
    PonCel li, C_IMSS, nomina.imss
    PonCel li, C_PRESTAMOS, nomina.prestamos
    PonCel li, C_FONACOT, nomina.fonacot
    PonCel li, C_PENSION, nomina.telefono
    PonCel li, C_INFONAVIT, nomina.otraded

    deducciones = nomina.crdsal + nomina.ispt + nomina.imss + nomina.prestamos + _
                  nomina.fonacot + nomina.telefono + nomina.otraded
    PonCel li, C_TOTDED, deducciones

    neto = ingresos - deducciones
    PonCel li, C_NETO, neto

    ' ---- No visibles ----
    ConNom1.TextMatrix(li, C_FLAG_ISR) = 0
    ConNom1.TextMatrix(li, C_FLAG_IMSS) = 0
    ConNom1.TextMatrix(li, C_FLAG_ESP) = 0
    If N_ormal = 1 Then ConNom1.TextMatrix(li, C_FLAG_ESP) = "1"

    locobra

    Get 12, regtro, Clbnx

    ConNom1.TextMatrix(li, C_BANAMEX) = (" " + Clbnx.Q1)
End Sub


Sub define()
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
        Encab C_LQ_SDOVAC, "Sdo.Vac.", 1200
        Encab C_LQ_PRIMAVAC, "P.Vacac.", 1200
        Encab C_LQ_AGUI, "Aguinaldo", 1200
        Encab C_LQ_ANTIG, "P.Antig.", 1200
        Encab C_LQ_COMPENS, "Compens.", 1200
        Encab C_LQ_INDEMN, "Indemn.", 1200
        Encab C_LQ_BONO, "BONO", 1200
    Else
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

    ConNom1.Col = C_FLAG_ISR
    ConNom1.ColWidth(C_FLAG_ISR) = 0
    ConNom1.Col = C_FLAG_IMSS
    ConNom1.ColWidth(C_FLAG_IMSS) = 0
    ConNom1.Col = C_FLAG_ESP
    ConNom1.ColWidth(C_FLAG_ESP) = 0

    Encab C_BANAMEX, "Banamex", 2400
End Sub

Sub genenom(gg)
    Dim cPV As Long
    Dim esLiq As Boolean

    esLiq = (g_TipoNominaActiva = tnLiquidacionFiniquito)
    If esLiq Then
        cPV = C_LQ_PRIMAVAC
    Else
        cPV = C_CONC8
    End If

    sum(1) = 0
    sum(2) = 0
    ConNom1.Font.Size = 8
    ConNom1.Font.Bold = True
    li = ConNom1.Row
    ii = 0

    ConNom1.TextMatrix(li, C_NUM) = Format(gg, "###0")
    ConNom1.TextMatrix(li, C_NOMBRE) = RTrim$(personal.ape1) + " " + RTrim$(personal.ape2) + " " + RTrim$(personal.nom)
    ConNom1.TextMatrix(li, C_DIAS) = Format(diat, "###0.00")
    ConNom1.TextMatrix(li, C_SUELDO) = Format((personal.ingr * diat), z1$): sum(1) = sum(1) + (personal.ingr * diat)

    If Not esLiq Then
        If personal.viat > 0 Then
            ConNom1.TextMatrix(li, C_CONC7) = Format((personal.viat * diat), z1$)
            sum(1) = sum(1) + (personal.viat * diat)
        End If

        If personal.otras > 0 Then
            ConNom1.TextMatrix(li, C_OTRAS) = Format((personal.otras * diat), z1$)
            sum(1) = sum(1) + (personal.otras * diat)
        End If
    End If

    ConNom1.TextMatrix(li, C_TOTING) = Format(sum(1), z1$)

    Acum_Doble

    If cal_anual = 0 Then
        impto = 0
        calculoRetencion sum(1), impto, 1, regtro
    Else
        rgtro = ConNom1.TextMatrix(li, C_NUM)
        aoalta = Val(Mid(personal.fal, 7, 4))
        If (aoalta = empresa.ao) And (Tot_dias < 330) Then
            impto = 0
            calculoRetencion sum(1), impto, 1, regtro
        Else
            impto = 0: calc_anual sum(1), impto, 1
        End If
    End If

    If ArAcum.DImpret > 0 And impto < 0 Then
        crd_sal = 0
    Else
        If impto < 0 Then
            crd_sal = impto: impto = 0
        Else
            crd_sal = 0
        End If
    End If

    diaseg = diat: integrado = personal.integrado
    seguro = 0

    If ConNom1.TextMatrix(li, cPV) > "" Then
        p_vacacional = ConNom1.TextMatrix(li, cPV)
    Else
        p_vacacional = 0
    End If

    imss integrado, seguro, diaseg

    If impto <> 0 Then
        ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)
        sum(2) = sum(2) + impto
    Else
        ConNom1.TextMatrix(li, C_ISPT) = ""
    End If

    If crd_sal <> 0 Then
        ConNom1.TextMatrix(li, C_SUBEMP) = Format(crd_sal, z1$)
        sum(2) = sum(2) + crd_sal
    Else
        ConNom1.TextMatrix(li, C_SUBEMP) = ""
    End If

    If seguro > 0 Then
        ConNom1.TextMatrix(li, C_IMSS) = Format(seguro, z1$)
        sum(2) = sum(2) + seguro
    Else
        ConNom1.TextMatrix(li, C_IMSS) = ""
    End If

    sum(3) = sum(1) - sum(2)
    ConNom1.TextMatrix(li, C_TOTDED) = Format(sum(2), z1$)
    ConNom1.TextMatrix(li, C_NETO) = Format(sum(3), z1$)
    ConNom1.TextMatrix(li, C_FLAG_ISR) = 0
    ConNom1.TextMatrix(li, C_FLAG_IMSS) = 0

    If diat = 0 Then
        ConNom1.TextMatrix(li, C_FLAG_ESP) = 1
    Else
        ConNom1.TextMatrix(li, C_FLAG_ESP) = 0
    End If

    Get 12, regtro, Clbnx

    ConNom1.TextMatrix(li, C_BANAMEX) = (" " + Clbnx.Q1)
    locobra

End Sub

Private Sub LimpiaNomina()
    With nomina
        .dias = 0: .hsnor = 0: .hs_no = 0: .hsdbl = 0: .hs_db = 0: .hstri = 0
        .hs_tr = 0: .ispt = 0: .crdsal = 0: .imss = 0: .sueldo = 0: .hs_nor = 0
        .hs_dbl = 0: .hs_tri = 0: .viaticos = 0: .pvac = 0: .otras = 0: .aguin = 0
        .ptu = 0: .exentos = 0: .prestamos = 0: .fonacot = 0: .telefono = 0
        .otraded = 0
    End With
End Sub

' Valor numerico de una celda (0 si esta vacia o no es numerica)
Private Function CelCur(ByVal f As Long, ByVal c As Long) As Currency
    Dim s As String
    s = Trim$(ConNom1.TextMatrix(f, c))
    If Len(s) > 0 Then
        If IsNumeric(s) Then CelCur = CCur(s)
    End If
End Function

' Pasa el renglon f de la grilla al registro global 'nomina'
' (el Get del registro ya se hizo antes de llamar).
Private Sub GridANomina(ByVal f As Long)
    nomina.dias = CelCur(f, C_DIAS)
    nomina.sueldo = CelCur(f, C_SUELDO)

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        nomina.viaticos = CelCur(f, C_LQ_SDOVAC)
        nomina.pvac = CelCur(f, C_LQ_PRIMAVAC)
        nomina.aguin = CelCur(f, C_LQ_AGUI)
        nomina.hs_dbl = CelCur(f, C_LQ_ANTIG)
        nomina.hs_nor = CelCur(f, C_LQ_COMPENS)
        nomina.hs_tri = CelCur(f, C_LQ_INDEMN)
        nomina.hsnor = CelCur(f, C_LQ_BONO)
    Else
        nomina.hs_nor = CelCur(f, C_CONC4)
        If N_ormal = 1 Then
            nomina.aguin = CelCur(f, C_CONC5)
        Else
            nomina.hs_dbl = CelCur(f, C_CONC5)
        End If
        If N_ormal = 1 Then
            nomina.ptu = CelCur(f, C_CONC6)
        Else
            nomina.hs_tri = CelCur(f, C_CONC6)
        End If
        nomina.viaticos = CelCur(f, C_CONC7)
        nomina.pvac = CelCur(f, C_CONC8)
    End If

    nomina.otras = CelCur(f, C_OTRAS)
    nomina.exentos = CelCur(f, C_EXENTA)

    nomina.ispt = CelCur(f, C_ISPT)
    nomina.crdsal = CelCur(f, C_SUBEMP)
    nomina.imss = CelCur(f, C_IMSS)
    nomina.prestamos = CelCur(f, C_PRESTAMOS)
    nomina.fonacot = CelCur(f, C_FONACOT)
    nomina.telefono = CelCur(f, C_PENSION)
    nomina.otraded = CelCur(f, C_INFONAVIT)
End Sub

Private Sub archnom_Click(Index As Integer)
    Dim f As Long
    Dim colanti As Long, renati As Long

    Text4.Text = ""
    Close 6:: Dm = LOF(2) / Len(personal)
    Open Arch$ For Random As 6 Len = Len(nomina)
    nm = LOF(6) / Len(nomina)

    colanti = ConNom1.Col
    renati = ConNom1.Row

    ' 1) Dejar en ceros todos los registros del archivo
    LimpiaNomina
    For f = 1 To Dm
        Put 6, f, nomina
    Next f

    ' 2) Guardar cada renglon de la grilla en su registro
    For f = 1 To limite
        regtro = Val(ConNom1.TextMatrix(f, C_NUM))
        If regtro > 0 Then
            Get #6, regtro, nomina
            GridANomina f
            Put 6, regtro, nomina
        End If
    Next f

    archiva_o = 0
    ConNom1.Col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus

    MsgBox ("Archivada con exito!")
End Sub

Private Sub cap_act_Click(Index As Integer)
    sumavert
End Sub

Private Sub capacimp_Click(Index As Integer)
    FuncionImpuesto
End Sub

Private Sub FuncionImpuesto()
    rrenat = ConNom1.Row
    rcolant = ConNom1.Col
     
    For li = 1 To ConNom1.Rows - 3
      ProgressBar1.Min = 0
      ProgressBar1.Max = ConNom1.Rows
      ProgressBar1.Value = 0
      ProgressBar1.Visible = True
      ProgressBar1.Value = li
          
        If IsNumeric(ConNom1.TextMatrix(li, 0)) Then
            regtro = ConNom1.TextMatrix(li, 0)
            Get #2, regtro, personal
        End If
    
        If IsNumeric(ConNom1.TextMatrix(li, C_ISPT)) Then
            impto1 = ConNom1.TextMatrix(li, C_ISPT)
        Else
            impto1 = 0
        End If
      
        ' Se fuerza la actualizacion del IMSS ignorando el bloqueo
        imss1 = 0
         
        If regtro > 0 Then
            sumah impto1, imss1
        End If
    Next li
     
    ProgressBar1.Visible = False
    ConNom1.Row = rrenat
    ConNom1.Col = rcolant
    archiva_o = 1

End Sub

Private Sub capimpno_Click(Index As Integer)
    For i = 1 To 9: sum(i) = 0: Next i
    hoja = 0: conta = 0: conta3 = 0: conta1 = 0: conta2 = 0
    GoTo Sale
    nomb_e$ = Printer.FontName
    tama_o = Printer.FontSize
    Printer.FontName = "courier new"
    Printer.FontSize = 7
    Printer.FontBold = True
    Printer.Orientation = 1
    If ConNom1.TextMatrix(1, 0) <> "" Then rgtro = ConNom1.TextMatrix(1, 0) Else rgtro = 0
If rgtro > 0 Then
    impr_tit
    For r = 1 To limite
      li = r
      rgtro = ConNom1.TextMatrix(li, 0)
      Printer.Line (0 + Aumento1, Printer.CurrentY)-(12200 + Aumento1, Printer.CurrentY), , BF
      valor = rgtro: uso$ = "####0.": ancho2 = 0
      colocar ancho2, valor, uso$
      Printer.CurrentX = 0 + ancho2 + Aumento1
      Printer.Print rgtro;
      lin_nom
    Next r
    Printer.Print: Printer.Print
    tot_sub 1
    Printer.EndDoc
   Else
     MsgBox "Necesita Capturar Nomina para Imprimirla"
  End If
  Printer.FontName = nomb_e$
  Printer.FontSize = tama_o
Sale:
End Sub

Private Sub capnoalf_Click(Index As Integer)
    colanti = ConNom1.Col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.Col = 1
    ConNom1.RowSel = limite
    ConNom1.Sort = 1
    ConNom1.Col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus
    End Sub

Private Sub capnomnum_Click(Index As Integer)
    colanti = ConNom1.Col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.Col = 0
    ConNom1.RowSel = limite
    ConNom1.Sort = 3
    ConNom1.Col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus
End Sub

Private Sub capnosal_Click(Index As Integer)
    Clipboard.Clear
    
On Error GoTo salidaverdad

    If archiva_o <> 0 Then
        responde = MsgBox("Desea salir sin Archivar ", vbYesNo + vbCritical + vbDefaultButton2)
        If responde = vbYes Then
            archiva_o = 0
            ConNom1.Clear
            Close 2: Close 3: Close 4: Close 5
            Form8.Hide
            Form1.Show
        Else
            If ConNom1.TextMatrix(1, 0) <> "" Then archnom_Click 1
            ConNom1.Clear
            Close 2: Close 3: Close 4: Close 5
            Form8.Hide
            Form1.Show
        End If
    End If
    Exit Sub

salidaverdad:
    MsgBox "Ups! ocurio un error" & Err.Description, vbOKCancel, Err.Number

End Sub

Private Sub caprecind_Click(Index As Integer)
    colanti = ConNom1.Col
    renati = ConNom1.Row
    
    If ConNom1.Text <> "" Then
        rgtro = ConNom1.Text
    Else
        rgtro = 0
    End If
    
    If rgtro > 0 Then
        regtro = ConNom1.TextMatrix(ConNom1.Row, 0)
        Get 14, rgtro, nom_com
        lin_rbo ConNom1.Row, ConNom1.RowSel
        ConNom1.SelectionMode = flexSelectionFree
        ConNom1.Col = colanti
        ConNom1.Row = renati
        ConNom1.SetFocus
    Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If
End Sub

Private Sub caprectd_Click(Index As Integer)
   colanti = ConNom1.Col
   renati = ConNom1.Row
   ConNom1.Row = 1
    ConNom1.Col = 0
   If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
   If rgtro > 0 Then
            regtro = ConNom1.TextMatrix(ConNom1.Row, 0)
            Get 14, rgtro, nom_com
            lin_rbo 1, limite
            ConNom1.Col = colanti
            ConNom1.Row = renati
            ConNom1.SetFocus
     Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If
End Sub

Private Sub Combo1_Click()
   Combo1_Change
End Sub

Private Sub Combo1_Change()

    meselegido = Combo1.ListIndex + 1
    
    If Option3 = True Then
        Label7.Caption = "Nomina de la 1a.quincena de " + Combo1.Text + " de " + Str$(empresa.ao)
        dia_pago = 15 - 1
        
    ElseIf Option4 = True Then
        Label7.Caption = "Nomina de la 2a.quincena de " + Combo1.Text + " de " + Str$(empresa.ao)
        dia_pago = dd(meselegido) - 1
      
        If meselegido = 12 Then
            respuesta = MsgBox("Desea hacer Calculo Anual", vbCritical + vbYesNo, "Captura Ultima Nomina")
            If respuesta = vbYes Then
                cal_anual = 1
            Else
                cal_anual = 0
            End If
        End If
    End If

End Sub

Private Sub Combo1_dblClick()
   Combo1_Change
   ConNom1.SetFocus
   ConNom1.Row = 1
   ConNom1.Col = 2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Combo1_Change
       ConNom1.SetFocus
       ConNom1.Row = 1
       ConNom1.Col = 2
    End If
End Sub

Private Sub ConNom1_EnterCell()
    If ConNom1.Col > 1 And ConNom1.Row > 0 Then
        ConNom1.CellBackColor = &H80FF80
    End If
End Sub

Private Sub ConNom1_KeyDown(KeyCode As Integer, Shift As Integer)
       Select Case KeyCode
            Case vbKeyDelete
                ConNom1.Text = ""
                checar
            Case vbKeyF2
                Text2.Text = ConNom1.Text
                Text2.SetFocus
       End Select
End Sub

Private Sub ConNom1_KeyPress(KeyAscii As Integer)
         If IsNumeric(ConNom1.Text) Then
            Nu_mero = ConNom1.Text
            valcelant = Nu_mero
            Else
            valcelant = ConNom1.Text
         End If
         If KeyAscii <> 13 Then
            Text2.Text = Chr$(KeyAscii)
            Text2.SetFocus
         End If
End Sub

Private Sub ConNom1_LeaveCell()
   If ConNom1.Col > 1 And ConNom1.Row > 0 Then
         ConNom1.CellBackColor = vbWhite
   End If
End Sub

Private Sub ConNom1_RowColChange()
        Text2.Text = ConNom1.Text
End Sub


Public Sub main()
    ' Funcion que viene desde el form1
    generarNominas
End Sub

Private Sub generarNominas()
    ProgressBar1.Visible = True
    QUIN = 0
  
  ' Si es una Nomina normal
    If Option1 = True Then
        ' Si es Primera quincena
        If Option3 = True Then
            QUIN = 1
            Arch$ = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.Mes = (Combo1.ListIndex + 1)
            ReferOper.dia = 15
        Else
        ' si es Segunda Quincena
            QUIN = 2
            Arch$ = UCase(Mid$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(Mid$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ArchAnterior = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1Anterior = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.Mes = (Combo1.ListIndex + 1)
            ReferOper.dia = dd(ReferOper.Mes)
            
            Close 95
    
            Open ArchAnterior For Random As 95 Len = Len(nominaAnterior)
            largoNomina = LOF(95) / Len(nominaAnterior)
    
       End If
    Else ' Si es una nomina especial
        Arch$ = Trim(Text1.Text) + ".NOM"
        Arch1 = Trim(Text1.Text) + ".cmp"
        diat = 0
    End If

    Rem *******     NOMBRE DEL ARCHIVO   ********
    ' Modifica el mes para que nunca sea cero
    If meselegido < 0 Then meselegido = 1
    ' Agrega cero a los numeros inferiores a 10
    If meselegido < 10 Then mif$ = LTrim$(Str$(dia_pago)) + "/0" + LTrim$(Str$(meselegido)) + "/" + LTrim$(Str$(empresa.ao))
    ' Emite el nombre de las nomnas antes del mes de octubre
    If meselegido > 9 Then mif$ = LTrim$(Str$(dia_pago)) + "/" + LTrim$(Str$(meselegido)) + "/" + LTrim$(Str$(empresa.ao))
    
    MiFecha = mif$
    checa_fecha MiFecha, 0, 0
    
    ' Abre archivos necesarios para Nomina
    Close 6
    Close 12
    Close 14

    Open Arch$ For Random As 6 Len = Len(nomina)
    nm = LOF(6) / Len(nomina)
    
    Open Arch1 For Random As 14 Len = Len(nom_com)
    Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
    
    ProgressBar1.Min = 0
    ProgressBar1.Max = Dm
    ProgressBar1.Value = 0
            
    If nm > 0 Then
        If nm < Dm Then
            fi_nm = (Dm)
        Else
            fi_nm = (nm)
        End If
        
        ConNom1.Clear
        
        define

        ConNom1.Rows = fi_nm + 2
        limite = 0
        renglon = 0
        
        For r = 1 To fi_nm
            Get #6, r, nomina
            Get #2, r, personal
            Get #8, r, maestro
            Get #14, r, nom_com
    
            regtro = r
            yavas = 0
            
            verifica yavas
            
            aobaja = Val(Mid$(personal.fab, 7, 4))
            mesbaja = Val(Mid$(personal.fab, 4, 2))
            diabaja = Val(Mid$(personal.fab, 1, 2))
            
            verifica yavas
                
            If yavas > 0 Then ' Aplica para primera quincena y nomina especial
                renglon = renglon + 1
                ConNom1.Row = renglon
                ConNom1.Col = 0
                ConNom1.Text = Format(r, "#####")
                limite = limite + 1
                ' Solo imprime primera quincena
                carganom
            Else
                ' Que si la nomina es especial y no es primera quincena
            End If
                
            If Option2 = True Then ' Especial
                If ((aobaja > 0) And (aobaja < (empresa.ao) - 1) And (yavas = 0)) Then
                    GoTo sigueLE
                End If
            Else
                If ((aobaja > 0) And (aobaja < empresa.ao) And (yavas = 0)) Then
                    GoTo sigueLE
                End If
            End If
                 
            If Option2 = True Then ' Aplica para nomina especial y bajas
                'If ((mesbaja > 0) And (mesbaja <= meselegido)) And (aobaja = (empresa.ao)) Then
                 '   renglon = renglon + 1
                 '   ConNom1.Row = renglon
                 '   ConNom1.Col = 0
                 '   ConNom1.Text = Format(r, "#####")
                 '   limite = limite + 1
                 '   ConNom1.CellBackColor = vbRed
                 '   ConNom1.CellForeColor = vbWhite
                    ' ___________________________________2
                    'carganom
                'End If
            Else
                If ((mesbaja > 0) And (mesbaja < meselegido)) And (aobaja = (empresa.ao)) Then
                    GoTo sigueLE
                End If
            End If
sigueLE:
            
        Next r
          
    Else
            ConNom1.Clear
            define
            ConNom1.Rows = Dm + 2
            limite = 0
            Close 13
            Open "Empcomp.dno" For Random As 13 Len = Len(Dat_ide)
            Get 13, 1, Dat_ide
            ceroscompl
            
            For r = 1 To Dm: Get #2, r, personal
                rgtro = r
                regtro = r
                
                    
                Get #8, r, maestro
                aobaja = Val(Mid$(personal.fab, 7, 4))
                mesbaja = Val(Mid$(personal.fab, 4, 2))
                diabaja = Val(Mid$(personal.fab, 1, 2))
                diat = 0
                If Option2 = True Then ' Especial
                    If ((aobaja > 0) And (aobaja < (empresa.ao) - 1) And (yavas = 0)) Then
                        GoTo SIGUE
                    End If
                Else
                    If ((aobaja > 0) And (aobaja < empresa.ao) - 1) Then
                        GoTo SIGUE
                    End If
                End If
                
                If Option2 = False Then ' Segunda quincena
                    If ((mesbaja > 0) And (mesbaja < meselegido)) And (aobaja = (empresa.ao)) Then GoTo sigueLE
                End If
                
                If Mid$(personal.nom, 1, 3) <= "   " Then GoTo SIGUE
                
                If (Option3 = True) Then diat = 15 ' Primera
                
                If (Option4 = True) And (Dat_ide.dias = 1) Then ' Segunda
                    diat = (dd(meselegido) - 15)
                Else
                    diat = 15
                End If
                
                If Option2 = True Then diat = 0 ' Segunda quincena
                
                renglon = renglon + 1: ConNom1.Row = renglon
                limite = limite + 1
                genenom r
SIGUE:
                ProgressBar1.Value = r
            Next r
        End If
        
    sumavert
        
    Close 10
    
    ConNom1.SetFocus
    ConNom1.Row = 1
    ConNom1.Col = 2
    ConNom1.Rows = limite + 3
    nomordeli_Click (Index)
    ProgressBar1.Visible = False

End Sub
Private Sub generarNominaGeneral()
    ProgressBar1.Visible = True
    QUIN = 0
  ' Si es una Nomina normal
    If Option1 = True Then
        ' Si es Primera quincena
        If Option3 = True Then
            QUIN = 1
            Arch$ = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.Mes = (Combo1.ListIndex + 1)
            ReferOper.dia = 15
        Else
        ' si es Segunda Quincena
            QUIN = 2
            Arch$ = UCase(Mid$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(Mid$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ArchAnterior = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1Anterior = UCase(Mid$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.Mes = (Combo1.ListIndex + 1)
            ReferOper.dia = dd(ReferOper.Mes)
       End If
    Else
           
        ' Instrucciones validas para SACMAG
        Arch$ = Trim(Text1.Text) + ".NOM"
        Arch1 = Trim(Text1.Text) + ".cmp"
        
        diat = 0
    End If

  Rem *******     NOMBRE DEL ARCHIVO   ********
    ' Modifica el mes para que nunca sea cero
    If meselegido = 0 Then meselegido = 1
    
    ' Agrega cero a los numeros inferiores a 10
    If meselegido < 10 Then mif$ = LTrim$(Str$(dia_pago)) + "/0" + LTrim$(Str$(meselegido)) + "/" + LTrim$(Str$(empresa.ao))
    
    ' Le da el formato a la fecha en que se captura la nomina
    If meselegido > 9 Then mif$ = LTrim$(Str$(dia_pago)) + "/" + LTrim$(Str$(meselegido)) + "/" + LTrim$(Str$(empresa.ao))
    
    MiFecha = mif$
    checa_fecha MiFecha, 0, 0
    
    ' Abre archivos necesarios para Nomina
    Close 6: Open Arch$ For Random As 6 Len = Len(nomina): nm = LOF(6) / Len(nomina)
    Close 14: Open Arch1 For Random As 14 Len = Len(nom_com)
    Close 12: Open "bnxcla.dno" For Random As 12 Len = Len(Clbnx)
            
    ProgressBar1.Min = 0
    ProgressBar1.Max = Dm
    ProgressBar1.Value = 0
            
    If nm > 0 Then
        If nm < Dm Then
            fi_nm = (Dm)
        Else
            fi_nm = (nm)
        End If
        
        ConNom1.Clear
        define ' funcion que define los encabezados
        ConNom1.Rows = fi_nm + 2
        limite = 0
        renglon = 0
        
        For r = 1 To fi_nm
            Get #6, r, nomina
            Get #2, r, personal
            Get #8, r, maestro
            Get #14, r, nom_com
    
            regtro = r
            yavas = 0
            
            verifica yavas
            
            aobaja = Val(Mid$(personal.fab, 7, 4))
            mesbaja = Val(Mid$(personal.fab, 4, 2))
            diabaja = Val(Mid$(personal.fab, 1, 2))
            
            verifica yavas
                
            renglon = renglon + 1
            ConNom1.Row = renglon
            ConNom1.Col = 0
            ConNom1.Text = Format(r, "#####")
            limite = limite + 1
            carganom
                
            If Option2 = True Then
                If ((aobaja > 0) And (aobaja < (empresa.ao) - 1) And (yavas = 0)) Then
                    GoTo sigueLE
                End If
            Else
                If ((aobaja > 0) And (aobaja < empresa.ao) And (yavas = 0)) Then
                    GoTo sigueLE
                End If
            End If
                 
            If Option2 = True Then
                If ((mesbaja > 0) And (mesbaja <= meselegido)) And (aobaja = (empresa.ao)) Then
                    renglon = renglon + 1
                    ConNom1.Row = renglon
                    ConNom1.Col = 0
                    ConNom1.Text = Format(r, "#####")
                    limite = limite + 1
                    carganom
                End If
            Else
                If ((mesbaja > 0) And (mesbaja < meselegido)) And (aobaja = (empresa.ao)) Then
                    GoTo sigueLE
                End If
            End If
sigueLE:
            ProgressBar1.Value = r
        Next r
          
    Else
            ConNom1.Clear
            define
            ConNom1.Rows = Dm + 2
            limite = 0
            Close 13
            Open "Empcomp.dno" For Random As 13 Len = Len(Dat_ide)
            Get 13, 1, Dat_ide
            ceroscompl
            
            For r = 1 To Dm: Get #2, r, personal
                rgtro = r
                regtro = r
                Get #8, r, maestro
                aobaja = Val(Mid$(personal.fab, 7, 4))
                mesbaja = Val(Mid$(personal.fab, 4, 2))
                diabaja = Val(Mid$(personal.fab, 1, 2))
                diat = 0
                If Option2 = True Then
                    If ((aobaja > 0) And (aobaja < (empresa.ao) - 1) And (yavas = 0)) Then
                        GoTo SIGUE
                    End If
                End If
                
                If Option2 = False Then
                    If ((mesbaja > 0) And (mesbaja < meselegido)) And (aobaja = (empresa.ao)) Then GoTo sigueLE
                End If
                
                renglon = renglon + 1: ConNom1.Row = renglon
                limite = limite + 1
                genenom r
SIGUE:
                ProgressBar1.Value = r
            Next r
        End If
        Close 10
        ConNom1.SetFocus
        ConNom1.Row = 1
        ConNom1.Col = 2
        ConNom1.Rows = limite + 3
        ProgressBar1.Visible = False

End Sub
 
Sub ceroscompl()
    For g = 1 To fi_nm: Get 14, g, nom_com
        nom_com.ArchImp = " ": nom_com.CredNe = 0: nom_com.CreTot = 0
        nom_com.ImpTot = 0: nom_com.PSubDi = 0: nom_com.subapl = 0
        nom_com.subdio = 0: nom_com.subNap = 0
        Put 14, g, nom_com
    Next g
End Sub

Private Sub nomordeli_Click(Index As Integer)
    If ConNom1.Rows < 5 Then
        Exit Sub
    End If
        
    eliminacion
    If (ConNom1.Rows <> 4) Then
        colanti = ConNom1.Col
        renati = ConNom1.Row
        ConNom1.Row = 1
        ConNom1.Col = 1
        ConNom1.RowSel = limite
        
        ConNom1.Sort = 1
        ConNom1.Col = colanti
        ConNom1.Row = renati
        ConNom1.SetFocus
    End If
    
    If (ConNom1.TextMatrix(1, 0) = "") Then
        ConNom1.RemoveItem 1
        ConNom1.Rows = ConNom1.Rows + 1
        
    End If
End Sub


Private Sub Option1_GotFocus()
    If Option1 = True Then
        N_ormal = 0
        Option3.SetFocus
    End If
End Sub

Private Sub Option1_Click()

    Combo1.locked = False
    Combo1.Visible = True

    MostrarControlesEspecial False

    txtPre.Text = ""
    cboEspecial.ListIndex = -1

End Sub

Private Sub Option2_Click()

On Error GoTo ErrorManejador

    If Option2 = True Then

        Option3 = False
        Option4 = False

        N_ormal = 1
        dia_pago = 15 - 1

        Combo1.locked = True
        Combo1.Visible = False

        ' Limpiar nombre anterior.
        txtPre.Text = ""
        cboEspecial.ListIndex = -1
        Text1.Text = ""

        Label7.Caption = ""

        MostrarControlesEspecial True

        cboEspecial.SetFocus

    End If

Exit Sub

ErrorManejador:
    MsgBox Err.Number & " " & Err.Description

End Sub

Public Sub presionarOption2()
    Option2_Click
End Sub

Private Sub Option3_Click()
MostrarControlesEspecial False

    If meselegido = 0 Then meselegido = 1

    If Option3 = True Then
        Label7.Caption = "Nomina de la 1a.quincena " + _
                         Combo1.Text + Str$(empresa.ao)

        dia_pago = 15 - 1

        Combo1.Visible = True
        Combo1.locked = False
        Combo1.SetFocus
    End If
End Sub

Private Sub Option4_Click()
 MostrarControlesEspecial False

    If meselegido = 0 Then meselegido = 1

    If Option4 = True Then
        Label7.Caption = "Nomina de la 2a.quincena " + _
                         Combo1.Text + Str$(empresa.ao)

        dia_pago = dd(meselegido) - 1

        If meselegido = 12 Then
            respuesta = MsgBox("Desea hacer Calculo Anual", _
                               vbCritical + vbYesNo, _
                               "Captura Ultima Nomina")

            If respuesta = vbYes Then
                cal_anual = 1
            Else
                cal_anual = 0
            End If
        End If

        Combo1.Visible = True
        Combo1.SetFocus
    End If
End Sub

Public Sub presionarEnterForm8()
    Text1_KeyPress 13
End Sub
Private Sub Text2_Change()
    ConNom1.Text = Text2.Text
End Sub
Private Sub Text2_GotFocus()

    On Error Resume Next

    antiguo = ConNom1.Col

    ConNom1.Col = 1
    ConNom1.CellBackColor = &H80FF80

    ConNom1.Col = antiguo

    ' Colocar el cursor al final del contenido.
    Text2.SelStart = Len(Text2.Text)

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
      Select Case KeyCode
         Case vbKeyEscape
           Text2.Text = ""
           Text2.Text = valcelant
           ConNom1.Text = Format(valcelant, z1$)
           ConNom1.SetFocus
      End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   Dim aguinaldo As Currency
   Dim exentoLocal As Currency
   Dim montoTemporal As Currency

    If KeyAscii = 13 Then
         ConNom1.Text = Text2.Text
         checar
         archiva_o = 1
         ConNom1.SetFocus
    End If
    
End Sub

Private Sub Text2_LostFocus()
   antiguo = ConNom1.Col
   ConNom1.Col = 1: ConNom1.CellBackColor = Color_gris
   ConNom1.Col = antiguo
   
   ' Buscar la siguiente fila visible resultante del filtro
   Dim nextRow As Long
   nextRow = ConNom1.Row + 1
   Do While nextRow < ConNom1.Rows
       If ConNom1.RowHeight(nextRow) > 0 Then
           ConNom1.Row = nextRow
           Exit Do
       End If
       nextRow = nextRow + 1
   Loop
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      exento = Val(Text3.Text)
      Text3.FontBold = True
      Text3.Text = Format(exento, z1$)
      ConNom1.SetFocus
   End If
   
End Sub

Public Sub calculoRetencion(ingresoBruto, impuesto, porcentajeSubsidio, idEmpleado As Integer)
    Dim subsidioDiario As Currency
    Dim salarioMinimoDiario As Currency
    Dim limiteInferior As Currency
    Dim limiteSuperior As Currency
    Dim diasTrabajadosBrutos  As Integer
    
    ' Hace el calculo de la retencion quincenal
    
    If Form8.Option3.Value = True Then
        ' Abre los archivos de tarifas
        Close 3: Open (Trim(Dir_imptos) + "Tab08Kin.ISR") For Random As #3 Len = Len(articulo): largoArticulo = LOF(3) / Len(articulo)
        Close 4: Open (Trim(Dir_imptos) + "Tab08Kin.SUB") For Random As #4 Len = Len(subsidio): largoSubsidio = LOF(4) / Len(subsidio)
                For i = 1 To largoArticulo
            Get 4, 1, subsidio
    
            Get 3, i, articulo
            If ingresoBruto > (articulo.liminf) And ingresoBruto < (articulo.limsup) Then
                excedente = ingresoBruto - articulo.liminf
                impuestoSobreExcedente = excedente * (articulo.porcsl / 100)
                impto = articulo.cuotaf + impuestoSobreExcedente
                nom_com.ImpTot = impto
                Exit For
            End If
        Next i
            diasTrabajadosBrutos = CInt(ConNom1.TextMatrix(li, 2))
        
        ' Verifica que los dias sean mayores que cero
        If diasTrabajadosBrutos <> 0 Then
            ingresoDiarioTotal = ingresoBruto / diasTrabajadosBrutos
            subsidioDiario = subsidio.cuotafs / 15
            salarioMinimoDiario = empresa.psub
            limiteSuperior = subsidio.limsups / 15
            Rem limiteInferior = 0.01
        End If
        
            If ingresoDiarioTotal <= limiteSuperior Then
                subdio = (-subsidioDiario) * diasTrabajadosBrutos
                ElseIf ingresoDiarioTotal >= limiteSuperior Then
                subdio = 0
            Rem Elseif ingresoDiarioTotal <= limiteInferior then
            Rem subdio = -196.06
            End If
            
        impto = impto - subdio
        ConNom1.TextMatrix(li, C_SUBEMP) = subdio
        
        If impto < 0 Then
            impto = 0
        End If
        
        nom_com.CredNe = 0
        
        Put 14, regtro, nom_com
    End If
     
' Hace el calculo de la retencion mensual
    If Form8.Option4.Value = True Then
        mesNomina = Trim(Left(Combo1.Text, 3))
        AnoNomina = Right(Trim(Label7.Caption), 4)
        
        Close 35: Open CStr(Trim(Form1.Dir1)) + "\" + mesNomina + "1" + AnoNomina + ".NOM" For Random As 35 Len = Len(nomina)
        Close 36: Open CStr(Trim(Form1.Dir1)) + "\" + mesNomina + "1" + AnoNomina + ".cmp" For Random As 36 Len = Len(nom_com)
        
        Get 35, regtro, nomina: Get 36, regtro, nom_com
        
        IngresosAnteriores = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.ptu + nomina.viaticos + nomina.pvac + nomina.otras
        ImpuestoAnteriorRetenido = nomina.ispt
        ImpuestoAnteriorTotal = nom_com.ImpTot
        subsidioCausado = nom_com.subapl
        diasTrabajados = ConNom1.TextMatrix(ConNom1.Row, 2)
        
        ' Tarifa mensual
        Close 3: Open (Trim(Dir_imptos) + "Tab08Mes.ISR") For Random As #3 Len = Len(articulo): largoArticulo = LOF(3) / Len(articulo)
        Close 4: Open (Trim(Dir_imptos) + "Tab08Mes.SUB") For Random As #4 Len = Len(subsidio): largoSubsidio = LOF(4) / Len(subsidio)
         
        ingresoMensual = ingresoBruto + IngresosAnteriores
        
        For i = 1 To largoArticulo
            Get 3, i, articulo
            If ingresoMensual > (articulo.liminf) And ingresoMensual < (articulo.limsup) Then
                excedente = ingresoMensual - articulo.liminf
                impuestoSobreExcedente = excedente * (articulo.porcsl / 100)
                impto = articulo.cuotaf + impuestoSobreExcedente
                nom_com.ImpTot = impto - ImpuestoAnteriorTotal
              Exit For
                            End If
                        Next

    End If
End Sub

Private Sub Text4_Change()
    Dim txt As String
    Dim i As Long
    Dim rowHt As Long
    Static origHt As Long
    
    ' Inicializar la altura original de fila
    If origHt = 0 Then
        If ConNom1.Rows > 1 Then
            If ConNom1.RowHeight(1) > 0 Then
                origHt = ConNom1.RowHeight(1)
            Else
                origHt = 240
            End If
        Else
            origHt = 240
        End If
    End If
    rowHt = origHt
    
    ' Obtener texto de bÃºsqueda sin espacios
    txt = Trim(Text4.Text)
    
    ' Mostrar u ocultar botÃ³n de restaurar
    If txt <> "" Then
        CmdRestoreFilter.Visible = True
    Else
        CmdRestoreFilter.Visible = False
    End If
    
    ' Ocultar/mostrar filas coincidentes
    ConNom1.Redraw = False
    For i = 1 To limite
        If txt = "" Then
            ConNom1.RowHeight(i) = rowHt
        Else
            If InStr(1, ConNom1.TextMatrix(i, 1), txt, vbTextCompare) > 0 Then
                ConNom1.RowHeight(i) = rowHt
            Else
                ConNom1.RowHeight(i) = 0
            End If
        End If
    Next i
    ConNom1.Redraw = True
End Sub

Private Sub CmdRestoreFilter_Click()
    Text4.Text = """"
    Text4.SetFocus
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim i As Long
    If KeyAscii = 13 Then
        ' Al presionar Enter, enfocar el primer empleado visible
        For i = 1 To limite
            If ConNom1.RowHeight(i) > 0 Then
                ConNom1.Row = i
                ConNom1.Col = 1
                ConNom1.SetFocus
                Exit For
            End If
        Next i
        KeyAscii = 0
    End If
End Sub

Private Sub Encab(ByVal c As Long, ByVal txt As String, ByVal ancho As Long)
    ConNom1.Col = c
    ConNom1.CellAlignment = 4
    ConNom1.ColWidth(c) = ancho
    ConNom1.Text = txt
End Sub

Private Sub MostrarControlesEspecial(ByVal v As Boolean)

    cboEspecial.Visible = v
    txtPre.Visible = v
    lblNombre.Visible = v

End Sub

Private Function SoloDigitos(ByVal texto As String) As String
    Dim i As Long
    Dim caracter As String
    Dim resultado As String

    resultado = ""

    For i = 1 To Len(texto)
        caracter = Mid$(texto, i, 1)

        If caracter >= "0" And caracter <= "9" Then
            resultado = resultado & caracter
        End If
    Next i

    SoloDigitos = resultado
End Function

Private Sub ActualizarNombreEspecial()

    If ActualizandoNombre Then Exit Sub

    On Error GoTo Salida

    ActualizandoNombre = True

    Dim numero As String
    Dim concepto As String
    Dim nombreFinal As String

    numero = Trim$(txtPre.Text)
    concepto = Trim$(cboEspecial.Text)

    ' Construir el nombre de la nómina especial.
    nombreFinal = numero & concepto

    ' Guardar únicamente el nombre.
    Text1.Text = nombreFinal

    ' Mostrarlo al usuario.
    Label7.Caption = nombreFinal

Salida:
    ActualizandoNombre = False

End Sub

Private Sub txtPre_Change()
    ActualizarNombreEspecial
End Sub

Private Sub txtPre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0

        ActualizarNombreEspecial

        If cboEspecial.ListIndex < 0 Then
            MsgBox "Seleccione un tipo de nómina especial.", vbExclamation, "Nómina"
            cboEspecial.SetFocus
            Exit Sub
        End If

        Text3.SetFocus
        Exit Sub
    End If

    ' Permitir Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Permitir únicamente números
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If

End Sub

Private Sub CargarPartesNombreEspecial()

    Dim nombre As String
    Dim concepto As String
    Dim i As Long
    Dim posicion As Long

    nombre = Trim$(Text1.Text)

    If nombre = "" Then
        ActualizarNombreEspecial
        Exit Sub
    End If

    ActualizandoNombre = True

    ' Buscar el concepto dentro del nombre existente.
    posicion = 0
    concepto = ""

    For i = 0 To cboEspecial.ListCount - 1

        If Trim$(cboEspecial.List(i)) <> "" Then

            If InStr(1, nombre, Trim$(cboEspecial.List(i)), vbTextCompare) > 0 Then
                concepto = Trim$(cboEspecial.List(i))
                posicion = InStr(1, nombre, concepto, vbTextCompare)
                Exit For
            End If

        End If

    Next i

    If posicion > 0 Then

        txtPre.Text = Left$(nombre, posicion - 1)

        cboEspecial.ListIndex = -1

        For i = 0 To cboEspecial.ListCount - 1
            If StrComp(Trim$(cboEspecial.List(i)), concepto, vbTextCompare) = 0 Then
                cboEspecial.ListIndex = i
                Exit For
            End If
        Next i

        txtSuf.Text = Mid$(nombre, posicion + Len(concepto))

    Else

        ' No se identificó ningún concepto conocido.
        txtPre.Text = ""
        txtSuf.Text = ""

    End If

    ActualizandoNombre = False

    ' Volver a construir el nombre con los controles.
    ActualizarNombreEspecial

End Sub
