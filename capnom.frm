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
   Begin VB.ComboBox cboMesNominaEspecial 
      Height          =   360
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtPre 
      Height          =   375
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox cboTipoNominaEspecial 
      Height          =   360
      Left            =   5520
      Style           =   2  'Dropdown List
      TabIndex        =   3
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
      Left            =   9720
      TabIndex        =   19
      Top             =   1200
      Width           =   2295
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   18
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
      Left            =   9720
      TabIndex        =   16
      Top             =   720
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
      TabIndex        =   14
      Top             =   2400
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid ConNom1 
      Height          =   4005
      Left            =   240
      TabIndex        =   6
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
      Left            =   5520
      TabIndex        =   11
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
      Left            =   5520
      List            =   "capnom.frx":0444
      TabIndex        =   2
      Top             =   1440
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
      TabIndex        =   1
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      Left            =   12240
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblMesNomina 
      Alignment       =   2  'Center
      Caption         =   "Mes de la nómina a capturar:"
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
      TabIndex        =   23
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblNumeroNominaEspecial 
      Caption         =   "Número de Nómina:"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   840
      Width           =   1455
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
      Left            =   8880
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "   Usando tarifas de impuestos : "
      Height          =   255
      Left            =   3000
      TabIndex        =   17
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
      Left            =   8160
      TabIndex        =   15
      Top             =   840
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
      Left            =   0
      TabIndex        =   13
      Top             =   1800
      Width           =   9375
   End
   Begin VB.Label lblTipoNominaEspecial 
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
      TabIndex        =   12
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
Dim requisito As Integer, comp_ti As String, clave As Integer
Dim si_hay As Integer, operacion, lim_conta As Integer
Dim fi_nm As Integer, anchopapel As Long, largopapel As Long
Dim Entrada As String * 32, abre, caso, Salida, numerillo As Integer, lincam As Currency
Dim Slinimpte As Currency, Slincam As Currency, Sabono As Integer
Dim porc_apli As Currency, p_vacacional As Currency, Detener As Integer, Detener1 As Integer
Dim PrincIngr As Long, PrincDesc As Long, Color_gris As Long, Nu_mero As Double, M_aximo As Long
Public subsidioParaElEmpleo As Currency
Private ActualizandoNombre As Boolean
Private FocoInicialSelector As Boolean

Private RutaNominaSesion As String
Private RutaCmpSesion As String

Private NominaNuevaSesion As Boolean
Private CmpNuevoSesion As Boolean

Private NominaArchivadaSesion As Boolean
Private GuardadoFallidoSesion As Boolean

Private CerrandoForm8 As Boolean
Private RegresoModalDesdeVisor As Boolean
Private GuardandoDuranteCierre As Boolean

Private SesionTemporalRegistrada As Boolean

Private CapturaActivaSesion As Boolean

Private EditandoNominaGuardada As Boolean

Private mSoloConsulta As Boolean
Private mDetalleConsulta As String

Private CopiaCmpEdicion() As Byte
Private HayCopiaCmpEdicion As Boolean

Private FalloRestauracionNom As Boolean
Private BloqueoPorPlazoEdicion As Boolean
Private RutaEdicionPendienteVisor As String

Private PreparandoEdicionDesdeVisor As Boolean

Private Sub Borrarblancos_Click(Index As Integer)
    Dim i As Integer
    Dim vacia As Boolean

    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If

    If ConNom1.Rows <= 1 Then
        MsgBox "No hay filas suficientes para procesar.", vbExclamation, "Error"
        Exit Sub
    End If

    For i = ConNom1.Rows - 1 To 1 Step -1
        vacia = True
        
        If Trim(ConNom1.TextMatrix(i, 5)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 6)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 7)) <> "" Then vacia = False
        If Trim(ConNom1.TextMatrix(i, 8)) <> "" Then vacia = False
        If vacia Then
            ConNom1.RemoveItem i
        End If
    Next i
End Sub

Private Sub cboTipoNominaEspecial_Click()
    RefrescarModoNomina
    ActualizarNombreEspecial
End Sub

Private Sub cboTipoNominaEspecial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If cboTipoNominaEspecial.ListIndex < 0 Then
            MsgBox "Seleccione el tipo de nomina especial.", vbExclamation, "Nomina"
            Exit Sub
        End If
        txtPre.SetFocus
        txtPre.SelStart = Len(txtPre.Text)
    End If
End Sub

Private Sub Form_Load()
    g_TipoNominaActiva = tnOrdinaria
    N_ormal = 0
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

    cboTipoNominaEspecial.Clear
    cboTipoNominaEspecial.AddItem "AGUINALDO"
    cboTipoNominaEspecial.AddItem "PTU"
    cboTipoNominaEspecial.AddItem "BONO"
    cboTipoNominaEspecial.AddItem "FINIQUITO"
    cboTipoNominaEspecial.AddItem "LIQUIDACION"
    
    cboMesNominaEspecial.Clear
    For i = 1 To 12
        cboMesNominaEspecial.AddItem RTrim$(mm(i))
    Next i
    cboMesNominaEspecial.ListIndex = Month(Date) - 1

    Text1.Visible = False
    RefrescarModoNomina
    txtPre.Text = ""
    FocoInicialSelector = True

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
    Dim solicitarEdicion As Boolean

    If CapturaActivaSesion Then
        MsgBox "Ya existe una nomina en captura." & vbCrLf & vbCrLf & "Para iniciar otra nomina, primero debe " & "salir de la captura actual." & vbCrLf & vbCrLf & "Puede guardar, salir sin guardar " & "o cancelar la salida.", vbExclamation, "Captura activa"
        Exit Sub
    End If
    
    If Option1.Value = True Then
        If Option3.Value = True Then
            Arch = UCase$(mID$(Combo1.Text, 1, 3)) & "1" & LTrim$(Str$(empresa.ao)) & ".NOM"
        ElseIf Option4.Value = True Then
            Arch = UCase$(mID$(Combo1.Text, 1, 3)) & "2" & LTrim$(Str$(empresa.ao)) & ".NOM"
        Else
            MsgBox "Seleccione Primera o Segunda quincena.", vbExclamation, "Nomina"
            Exit Sub
        End If
    ElseIf Option2.Value = True Then
        If cboTipoNominaEspecial.ListIndex < 0 Then
            MsgBox "Seleccione el tipo de nomina especial.", vbExclamation, "Nomina"
            Exit Sub
        End If

        ActualizarNombreEspecial

        If Len(Trim$(Text1.Text)) = 0 Then
            MsgBox "Indique el nombre de la nomina especial.", vbExclamation, "Nomina"
            Exit Sub
        End If

        Arch = Trim$(Text1.Text) & ".NOM"
    Else
        MsgBox "Seleccione el tipo de nomina.", vbExclamation, "Nomina"
        Exit Sub
    End If
    
    rutaNominas = Trim$(CStr(Form1.Dir1))

    If Len(rutaNominas) = 0 Then
        MsgBox "No se ha seleccionado un directorio de nominas.", vbExclamation, "Nomina"
        Exit Sub
    End If

    If Right$(rutaNominas, 1) <> "\" Then
        rutaNominas = rutaNominas & "\"
    End If

    fullArchPath = rutaNominas & Arch
    
    If Len(Dir$(fullArchPath)) > 0 Then
        If FileLen(fullArchPath) = 0 Then
            If MsgBox("Existe un archivo vacio con este nombre:" & vbCrLf & Arch & vbCrLf & vbCrLf & "Normalmente es el resto de una captura que se cerro sin guardar." & vbCrLf & "No contiene datos. Desea eliminarlo e iniciar la captura?", vbYesNo + vbQuestion + vbDefaultButton2, "Archivo vacio") = vbNo Then
                Exit Sub
            End If
            Kill fullArchPath
        End If
    End If
    
    If Len(Dir$(fullArchPath)) > 0 Then

        tamArchivo = FileLen(fullArchPath)
        If tamArchivo = 0 Then
            MsgBox "Ya existe un archivo con ese nombre, " & "pero esta vacio." & vbCrLf & vbCrLf & Arch & vbCrLf & vbCrLf & "No se iniciara otra captura para evitar " & "sobrescribirlo." & vbCrLf & "Revise el archivo antes de continuar.", vbExclamation, "Archivo vacio"
            Exit Sub
        End If
        
        If tamArchivo Mod Len(nomina) <> 0 Then
            MsgBox "El archivo de nomina existe, pero su " & "tamano no corresponde a registros completos." & vbCrLf & vbCrLf & Arch & vbCrLf & "No se modificara el archivo.", vbCritical, "Nomina"
            Exit Sub
        End If

        registrosArchivo = tamArchivo \ Len(nomina)
        AbrirNominaGuardadaParaEditar fullArchPath

        Exit Sub
    End If
    
    If StrComp(CurDir$, Left$(rutaNominas, Len(rutaNominas) - 1), vbTextCompare) <> 0 Then
        MsgBox "El directorio de trabajo no coincide " & "con el directorio seleccionado." & vbCrLf & vbCrLf & "Directorio activo: " & CurDir$ & vbCrLf & "Directorio seleccionado: " & rutaNominas & vbCrLf & vbCrLf & "No se iniciara la captura para evitar " & "crear archivos en una ubicacion incorrecta.", vbExclamation, "Verificar directorio"
        Exit Sub
    End If
    
    If Option1.Value = True And Option4.Value = True Then
        Dim primeraQna As String
        Dim faltaPrimera As Boolean
        primeraQna = rutaNominas & UCase$(mID$(Combo1.Text, 1, 3)) & "1" & LTrim$(Str$(empresa.ao)) & ".NOM"
        faltaPrimera = (Len(Dir$(primeraQna)) = 0)
        If Not faltaPrimera Then faltaPrimera = (FileLen(primeraQna) = 0)
        If faltaPrimera Then
            If MsgBox("No existe (o esta vacia) la nomina de la 1a quincena:" & vbCrLf & UCase$(mID$(Combo1.Text, 1, 3)) & "1" & LTrim$(Str$(empresa.ao)) & ".NOM" & vbCrLf & vbCrLf & "La retencion de la 2a quincena depende de ella. Desea continuar?", vbYesNo + vbExclamation + vbDefaultButton2, "Falta 1a quincena") = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    AplicarTipoNomina Arch

    Text3.Locked = (g_TipoNominaActiva = tnPTU)
    If g_TipoNominaActiva = tnPTU Then
        Text3.Text = ""
        exento = 0
    End If
    
    RutaNominaSesion = fullArchPath
    RutaCmpSesion = Left$(fullArchPath, Len(fullArchPath) - 4) & ".cmp"
    
    NominaNuevaSesion = (Len(Dir$(RutaNominaSesion)) = 0)
    CmpNuevoSesion = (Len(Dir$(RutaCmpSesion)) = 0)
    
    SesionTemporalRegistrada = True
    CapturaActivaSesion = True
    
    NominaArchivadaSesion = False
    GuardadoFallidoSesion = False
    
generarNominas
FuncionImpuesto

BloquearSeleccionNomina True
    Exit Sub
ErrorManejador:
    MsgBox "No fue posible iniciar o consultar la nomina." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbCritical, "Nomina"
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

    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    With FlexGrid
        For fila = 1 To .Rows - 1
            For columna = 0 To 1
                o_Hoja.Cells(fila, columna + 1).Value = .TextMatrix(fila, columna)
            Next
        Next
    End With
    
    o_Libro.Close True, sOutputPath
    o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function

Error_Handler:
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
        If (mID(Entrada, t, 1) >= Chr(48) And mID(Entrada, t, 1) <= Chr(57)) And (t < 4) Then
        Else
            caso = caso + mID(Entrada, t, 1)
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
        slinca$ = mID(Str(Slincam), 3, 2)
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
                linca$ = mID(Str(lincam), 3, 2)
            End If
        If Len(linca$) < 2 Then linca$ = linca$ + "0"
            Impte_Tr = linimpte
            linimpte = Int(linimpte)
            linimp$ = String(16 - Len(LTrim(Str(linimpte))), "0") + LTrim(Str(linimpte))
            plastico = ""
        For i = 1 To Len(Clbnx.Q1)
            If mID(Clbnx.Q1, i, 1) = "-" Then
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
        claveban = mID(Clbnx.Q1, (punto + 2))
        claveban = String(7 - Len(Trim(claveban)), "0") + claveban
    End If
    
    LINARCH$ = "30001" + linimp$ + linca$ + plastico + claveban + String(30, " ") + referencia$
    nomban$ = Trim(personal.nom) & Chr(9) & Trim(personal.ape1) & Chr(9) & Trim(personal.ape2)
    nomban$ = mID$(nomban$, 1, 55)
    
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
            mID(Label7.Caption, 1, 7) & Chr(9) & _
            "Mismo dia" & Chr(9) & _
            Format(Siguiente, "####0")
nuevo:
    Next r1
    
    slinca$ = "":
    LINARCH$ = "4001"
    Slincam = Int((Slinimpte - Int(Slinimpte)) * 100)
    slinca$ = Trim(Str(Slincam))
    
    If Slincam <= 0 Then slinca$ = "00"
    
    sabon$ = mID(Str(Sabono), 2)
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
    CFDI2_Click
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
    Printer.Print "-  " & Printer.Page & "  -"
End Sub
Private Sub CapDis_Click()
    final = limite
    Poliza.Show
End Sub

Private Sub CapInrt_Click()
    CapArch_Click 2
End Sub

Private Sub capturaGeneral_Click()
    generarNominaGeneral
End Sub

Private Sub CFDI2_Click()

    Dim nombreNomina As String
    On Error GoTo ErrorCFDI
        
    If Not AutorizarCambioNominaForm8() Then
        MsgBox "No se permite iniciar el CFDI desde una nomina con el plazo vencido o abierta en solo consulta.", vbExclamation, "CFDI"
        Exit Sub

    End If

    OrigenCFDI = 0
    g_QuincenaCFDI2026 = 0
    g_FechaPagoCFDI2026 = ""

    If Len(Trim$(RutaNominaSesion)) = 0 Then
    Err.Raise vbObjectError + 5221, "Form8.CFDI2_Click", "No hay una nomina activa para generar CFDI."
    End If
    
    g_RutaNominaCFDI2026 = RutaNominaSesion
    
    nombreNomina = NombreBaseNomina(RutaNominaSesion)

    If EsNominaNormal(RutaNominaSesion) Then
        g_QuincenaCFDI2026 = CInt(mID$(nombreNomina, 4, 1))
        g_FechaPagoCFDI2026 = FechaCFDI_PagoOrdinariaLegacy(RutaNominaSesion)
    Else
        g_QuincenaCFDI2026 = 0
        g_FechaPagoCFDI2026 = FechaCFDI_FormatoDMY(FechaCFDI_InterpretarDMY(CStr(MiFecha)))
    End If

    NOMCF2.Show
    Exit Sub

ErrorCFDI:
    MsgBox "No fue posible preparar el CFDI." & vbCrLf & "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "CFDI"
End Sub

Private Sub cheajte_Click(Index As Integer)
   Load AJTECH
   AJTECH.Show 1
End Sub

Private Sub cheind_Click(Index As Integer)
    colanti = ConNom1.col
    renati = ConNom1.Row
    If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
    If rgtro > 0 Then
        lin_che ConNom1.Row, ConNom1.RowSel
        ConNom1.SelectionMode = flexSelectionFree
        ConNom1.col = colanti
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
     Direc_torio = mID(CommonDialog1.FileName, 1, Len(CommonDialog1.FileName) - Len(CommonDialog1.FileTitle))
     Open "C:\GconTA\perma.dno" For Random As #7 Len = Len(basico)
     fin_basico = LOF(7) / Len(basico)
     basico.datoarch = Direc_torio
     Put #7, 2, basico
     Close #7
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub chetot_Click(Index As Integer)
   colanti = ConNom1.col
   renati = ConNom1.Row
   ConNom1.Row = 1
    ConNom1.col = 0
   If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
   If rgtro > 0 Then
            
            lin_che 1, limite
            ConNom1.col = colanti
            ConNom1.Row = renati
            ConNom1.SetFocus
     Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If
End Sub

Private Sub edicionPersonalCaptura_Click()
    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If
    Form4.Show
End Sub

Private Sub EdiNomSel_Click()
    Clipboard.Clear
    ConNom1.Row = 0
    ConNom1.col = 0
    ConNom1.RowSel = ConNom1.Rows - 1
    ConNom1.ColSel = ConNom1.Cols - 1
End Sub

Private Sub EdNomCopiar_Click()
    Dim textoSeleccionado As String
    Dim i As Integer, j As Integer
    
    Clipboard.Clear
    
    textoSeleccionado = ""
    For i = ConNom1.Row To ConNom1.RowSel
        For j = ConNom1.col To ConNom1.ColSel
            textoSeleccionado = textoSeleccionado & ConNom1.TextMatrix(i, j)
            If j < ConNom1.ColSel Then
                textoSeleccionado = textoSeleccionado & vbTab
            End If
        Next j
        If i < ConNom1.RowSel Then
            textoSeleccionado = textoSeleccionado & vbCrLf
        End If
    Next i
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
    
    Set objXL = CreateObject("Excel.Application")

    objXL.Visible = True

    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet

    If WorkSheetName <> "" Then
        wsXL.name = WorkSheetName
    End If

    For intRow = 1 To TheFlexgrid.Rows - 1
        j = 1
        For intCol = 0 To 1
            wsXL.Cells(intRow, j).Value = TheFlexgrid.TextMatrix(intRow, intCol)
            j = j + 1
        Next intCol
    Next intRow
    Exit Sub

ErrHandler:
    MsgBox "Necesitas tener instalado Microsoft Excel en tu equipo." & vbCrLf & "Detalle: " & Err.Description, vbExclamation, "Exportar a Excel"
    
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error GoTo ErrorCierre
    If CerrandoForm8 Then Exit Sub
    If Not ConfirmarSalidaNomina() Then
        Cancel = 1
    End If
    Exit Sub
ErrorCierre:
    Cancel = 1
    GuardandoDuranteCierre = False
    MsgBox "No fue posible completar la salida." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Nomina"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim mostrarForm1 As Boolean
    
    mostrarForm1 = Not RegresoModalDesdeVisor
    CerrandoForm8 = True
    FinalizarSalidaNomina
    CerrandoForm8 = False
    If mostrarForm1 Then
        Form1.Show
    End If
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

    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If
    
    Temporal1 = ""
    temporal = Clipboard.GetText()

    RetornoCarro = ConNom1.col
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
            Select Case mID(temporal, i, 1)
                Case Chr(13)
                        If EsColumnaNumericaCaptura(ConNom1.col) And ImporteCapturaValido(Trim$(mID$(temporal, DeAqui, i - DeAqui))) Then
                        Temporal1 = Temporal1 + ConNom1.TextMatrix(ConNom1.Row, ConNom1.col) & Chr(13)
                        ConNom1.Text = mID(temporal, DeAqui, (i - DeAqui))
                        Text2.Text = mID(temporal, DeAqui, (i - DeAqui))
                        Text2_KeyPress (13)
                        ConNom1.Row = ConNom1.Row + 1
                        
                        If ConNom1.Row >= ConNom1.Rows - 3 Then
                            ConNom1.Row = ConNom1.Rows - 3
                        End If
                        ConNom1.TopRow = 1
                    Else
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
        ConNom1.col = RetornoCarro
    End If

    Clipboard.Clear
    ConNom1.SetFocus

End Sub

Private Sub Importar_Click()
Dim oRS As New ADODB.Recordset
Dim oConn As New ADODB.Connection
    
    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If

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

Private Sub Llenar_FlexGrid(cn As Connection, rs As Recordset)
  
On Error GoTo ErrSub
    Screen.MousePointer = vbHourglass
    ConNom1.Redraw = False
    rs.MoveFirst
    
    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If
    ConNom1.Rows = rs.RecordCount + 1
    ConNom1.Cols = rs.Fields.Count
      
    Dim c As Integer
    ConNom1.Row = 1
    ConNom1.col = 0
      
    ConNom1.RowSel = ConNom1.Rows - 1
    ConNom1.ColSel = ConNom1.Cols - 1
    ConNom1.Clip = rs.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
    ConNom1.Row = 1
    ConNom1.Redraw = True
    
    rs.Close
    Set rs = Nothing
    cn.Close
    Set cn = Nothing
    Screen.MousePointer = vbDefault
Exit Sub
ErrSub:
MsgBox Err.Description, vbCritical
Screen.MousePointer = vbDefault
End Sub

Private Sub minToMax_Click()
    colanti = ConNom1.col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.col = 3
    ConNom1.RowSel = limite
    ConNom1.Sort = 3
    ConNom1.col = colanti
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
                       
            requisito = Val(mID(Mnunomobr.Item(Index).Caption, 1, 5))
            Entrada = mID(Mnunomobr.Item(Index).Caption, 7)
            If RTrim(Entrada) = "Capturar" Then
                    
                    comp_ti = InputBox("Teclee el nombre de la obra " + Str(requisito), "Impresion nomina obra")
                    Else
                    comp_ti = (Mnunomobr.Item(Index).Caption)
            End If
            impr_nomina requisito
     End Select
End Sub
Sub impr_nomina(clave As Integer)
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
                
                Select Case clave
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
                            If maestro.O_1 = clave Then
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
            If maestro.O_1 = Val(mID$(Mnunomobr(i).Caption, 1, 5)) Then ya = 1
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
            ConNom1.col = 1
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
                    Printer.Print LTrim$(mID$(Label7.Caption, 1, 20));
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
    Dim Borrar As Boolean

    antecol = ConNom1.col
    anteren = ConNom1.Row
    re = 1

    Do While re <= limite
        If EsSeparacionNueva() Then
            Borrar = Not IsNumeric(Trim$(ConNom1.TextMatrix(re, C_NUM)))
            If Not Borrar Then
                Borrar = (Val(ConNom1.TextMatrix(re, C_NUM)) <= 0)
            End If
        Else
            Borrar = (Trim$(ConNom1.TextMatrix(re, C_BANAMEX)) = "")
        End If
        If Borrar Then
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
        ConNom1.col = antecol
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
    Printer.Print RTrim$(mID$(ConNom1.TextMatrix(li, 1), 1, 27));
        
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
        ConNom1.col = i
        If ConNom1.Text <> "" Then
            sum(i) = ConNom1.Text
        Else
            sum(i) = 0
        End If
    Next i
    
    ConNom1.col = 0
    sum(0) = ConNom1.Text

End Sub
Sub ida_1()
    For i = 2 To 22: ConNom1.col = i
        If ConNom1.Text <> "" Then
            sumv(i) = ConNom1.Text
        Else
            sumv(i) = 0
        End If
    Next i
    
ConNom1.col = 0
sumv(0) = ConNom1.Text
End Sub
Sub regre()
         ConNom1.col = 0
         ConNom1.Text = Format(sumv(0), "###0")
         ConNom1.col = 1
         ConNom1.Text = compa$
         For i = 2 To 20: ConNom1.col = i
            If sumv(i) <> 0 Then
              ConNom1.Text = Format(sumv(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i
         For i = 21 To 22: ConNom1.col = i
            If sumv(i) <> 0 Then
              ConNom1.Text = sumv(i)
              Else
              ConNom1.Text = ""
            End If
         Next i
End Sub
Sub regre_1()
    ConNom1.col = 0
    ConNom1.Text = Format(sum(0), "###0")
         For i = 2 To 20: ConNom1.col = i
              If sum(i) <> 0 Then
              ConNom1.Text = Format(sum(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i
         For i = 21 To 22: ConNom1.col = i
              If sum(i) <> 0 Then
              ConNom1.Text = Format(sum(i), z1$)
              Else
              ConNom1.Text = ""
              End If
         Next i

End Sub
Sub datoper()
On Error GoTo handler
      ruta = ConNom1.col

    If ConNom1.Text = "" Then
        dato_ent = 0
    Else
        dato_ent = ConNom1.Text
    End If

      ConNom1.col = C_FLAG_ISR: If ConNom1.Text <> "" Then si_impto = ConNom1.Text Else si_impto = 0
      ConNom1.col = C_FLAG_IMSS: If ConNom1.Text <> "" Then si_imss = ConNom1.Text Else si_imss = 0
      dato_sal = 0: rgtro = 0
      ConNom1.col = C_NUM

      If ConNom1.Text <> "" Then
          regtro = ConNom1.Text
          Get 2, regtro, personal
          Get 8, regtro, maestro
      End If
      ConNom1.col = ruta

Exit Sub

handler:
Close 2: Close 8
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Open "maestro.dno" For Random As 8 Len = Len(maestro)
    ddm = LOF(8) / Len(maestro)
End Sub


Sub checanum(camcol, datoret As Currency)
   
   antecol = ConNom1.col
   ConNom1.col = camcol
   If ConNom1.Text = "" Then
      datoret = 0
   Else: datoret = ConNom1.Text
   End If
   
   ConNom1.col = antecol
End Sub

Private Function ExentoAutomaticoPTU(ByVal montoTotal As Currency) As Currency
    Dim tope As Currency
    If montoTotal <= 0 Then Exit Function
    If empresa.sm <= 0 Then Exit Function
    tope = CCur(empresa.sm) * 15
    If montoTotal < tope Then
        ExentoAutomaticoPTU = montoTotal
    Else
        ExentoAutomaticoPTU = tope
    End If
End Function

Sub checar()
    Dim esLiq As Boolean
    Dim utilidad As Currency
    
    If EsSeparacionNueva() Then
        checarSeparacionNueva
        Exit Sub
    End If

    esLiq = (g_TipoNominaActiva = tnLiquidacionFiniquito)

    datoper
    If regtro > 0 Then
      Select Case ConNom1.col
      
         Case C_DIAS
              diat = dato_ent
              If diat > 0 Then
                  ConNom1.Text = Format(diat, "##0.00")
                  ConNom1.col = C_SUELDO
                  ConNom1.Text = Format(diat * personal.ingr, z1$)
              Else
                  ConNom1.Text = ""
                  ConNom1.col = C_SUELDO
                  ConNom1.Text = ""
                  If g_TipoNominaActiva = tnOrdinaria And N_ormal = 0 Then
                      ConNom1.col = C_CONC7
                      ConNom1.Text = ""
                      ConNom1.col = C_CONC8
                      ConNom1.Text = ""
                  End If
              End If
              If g_TipoNominaActiva = tnOrdinaria And N_ormal = 0 Then
                  If personal.viat > 0 Then
                      ConNom1.col = C_CONC7
                      ConNom1.Text = Format(diat * personal.viat, z1$)
                  End If
                  If personal.otras > 0 Then
                      ConNom1.col = C_OTRAS
                      ConNom1.Text = Format(diat * personal.otras, z1$)
                  End If
              End If
              ConNom1.col = ruta
         Case C_SUELDO
              MsgBox "El sueldo se modifica editando su archivo de personal", vbCritical, "Captura de Nomina"
              Text2.Text = valcelant
         Case C_CONC4 To C_CONC6
              hs_ext = personal.ingr / 8
              If g_TipoNominaActiva = tnOrdinaria And N_ormal = 0 Then
                  dato_sal = dato_ent * hs_ext
                  ConNom1.col = ruta
                  ConNom1.Text = Format$(dato_sal, z1$)
              Else
                  Select Case ruta
                     Case C_CONC4
                        ConNom1.col = ruta
                        ConNom1.Text = Format$(dato_ent, z1$)
                     
            Case C_CONC5
                If g_TipoNominaActiva = tnAguinaldo Then
                    If dato_ent < 0 Or dato_ent < CelCur(ConNom1.Row, C_EXENTA) Then
                        MsgBox "El aguinaldo no puede ser negativo ni menor que su percepcion exenta.", vbExclamation, "Captura de Aguinaldo"
                        If IsNumeric(valcelant) Then
                            ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = Format$(CCur(valcelant), z1$)
                        Else
                            ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = ""
                        End If
                    ElseIf dato_ent = 0 Then
                        ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = ""
                    Else
                        ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = Format$(dato_ent, z1$)
                    End If
                Else
                    ConNom1.TextMatrix(ConNom1.Row, C_CONC5) = Format$(dato_ent, z1$)
                End If
                           
                Case C_CONC6
                    If g_TipoNominaActiva = tnPTU Then
                        If dato_ent < 0 Or dato_ent < CelCur(ConNom1.Row, C_EXENTA) Then
                            MsgBox "La PTU no puede ser negativa ni menor que su percepcion exenta.", vbExclamation, "Captura de PTU"
                            If IsNumeric(valcelant) Then
                                ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = Format$(CCur(valcelant), z1$)
                            Else
                                ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = ""
                            End If
                        ElseIf dato_ent = 0 Then
                            ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = ""
                        Else
                            ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = Format$(dato_ent, z1$)
                        End If
                    Else
                        ConNom1.TextMatrix(ConNom1.Row, C_CONC6) = Format$(dato_ent, z1$)
                    End If
            End Select
        End If
        
         Case C_CONC7, C_OTRAS
              If g_TipoNominaActiva = tnBonoPremio And ruta = C_CONC7 Then
                  If dato_ent < 0 Or dato_ent < CelCur(ConNom1.Row, C_EXENTA) Then
                      MsgBox "El bono no puede ser negativo ni menor que su percepcion exenta.", vbExclamation, "Captura de Bono"
                      If IsNumeric(valcelant) Then
                          ConNom1.TextMatrix(ConNom1.Row, C_CONC7) = Format$(CCur(valcelant), z1$)
                      Else
                          ConNom1.TextMatrix(ConNom1.Row, C_CONC7) = ""
                      End If
                  ElseIf dato_ent = 0 Then
                      ConNom1.TextMatrix(ConNom1.Row, C_CONC7) = ""
                  Else
                      ConNom1.TextMatrix(ConNom1.Row, C_CONC7) = Format$(dato_ent, z1$)
                  End If
              ElseIf g_TipoNominaActiva <> tnOrdinaria Or N_ormal <> 0 Then
                  ConNom1.col = ruta

                  If dato_ent <> 0 Then
                      ConNom1.Text = Format$(dato_ent, z1$)
                  Else
                      ConNom1.Text = ""
                  End If
              Else
                  aoingr = Val(mID$(personal.fal, 7, 4))
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

                  personal.integrado = (personal.ingr * facto) + _
                                       personal.viat + personal.otras
                  If personal.integrado > (empresa.sm * 25) Then
                      personal.integrado = (empresa.sm * 25)
                  End If
                  MsgBox "El salario integrado en el archivo de personal fue cambiado", vbExclamation
                  Put #2, regtro, personal
                  ConNom1.col = ruta
                  ConNom1.Text = Format$(dato_ent, z1$)
              End If

         Case C_ISPT
              respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
              If respuesta = vbYes Then
                  si_impto = 1
                  ConNom1.Text = Format(dato_ent, z1$)
                  ConNom1.col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.col = C_SUBEMP: ConNom1.Text = ""
                  ConNom1.col = ruta
              Else
                  si_impto = 0
                  ConNom1.col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.col = ruta
                  ConNom1.Text = Format(valcelant, z1$)
              End If
              
         Case C_SUBEMP
              respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
              If respuesta = vbYes Then
                  si_impto = 1
                  ConNom1.Text = Format(dato_ent, z1$)
                  ConNom1.col = C_FLAG_ISR: ConNom1.Text = 1
                  ConNom1.col = C_ISPT: ConNom1.Text = ""
                  ConNom1.col = ruta
              Else
                  si_impto = 0
                  ConNom1.col = C_FLAG_ISR: ConNom1.Text = 0
                  ConNom1.col = ruta
                  ConNom1.Text = Format(valcelant, z1$)
              End If
              
          Case C_IMSS
               If g_TipoNominaActiva = tnLiquidacionFiniquito Then
                   si_imss = 1
                   ConNom1.col = C_FLAG_IMSS
                   ConNom1.Text = "1"
                   ConNom1.col = C_IMSS
                   If dato_ent <> 0 Then
                       ConNom1.Text = Format$(dato_ent, z1$)
                   Else
                       ConNom1.Text = ""
                   End If
               Else
                   respuesta = MsgBox("Si modifica esta celda ya no se efectuara el calculo de la retencion", vbYesNo + vbExclamation + vbDefaultButton2, "Mensaje de Nomina")
                   If respuesta = vbYes Then
                       si_imss = 1
                       ConNom1.Text = Format$(dato_ent, z1$)
                       ConNom1.col = C_FLAG_IMSS
                       ConNom1.Text = "1"
                       ConNom1.col = ruta
                   Else
                       si_imss = 0
                       ConNom1.col = C_FLAG_IMSS
                       ConNom1.Text = "0"
                       ConNom1.col = ruta
                       ConNom1.Text = Format$(valcelant, z1$)
                   End If
               End If
               
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
            
            
            Case C_EXENTA
                If g_TipoNominaActiva = tnAguinaldo Then
                    If dato_ent < 0 Or dato_ent > CelCur(ConNom1.Row, C_CONC5) Then
                        MsgBox "La percepcion exenta no puede ser negativa ni superar el aguinaldo total.", vbExclamation, "Captura de Aguinaldo"
                        If IsNumeric(valcelant) Then
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(CCur(valcelant), z1$)
                        Else
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                        End If
                    ElseIf dato_ent = 0 Then
                        ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                    Else
                        ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(dato_ent, z1$)
                    End If
                                    
                ElseIf g_TipoNominaActiva = tnPTU Then
                        If dato_ent < 0 Or dato_ent > CelCur(ConNom1.Row, C_CONC6) Then
                            MsgBox "La percepcion exenta no puede ser negativa ni superar el importe total de PTU.", vbExclamation, "Captura de PTU"
                            If IsNumeric(valcelant) Then
                                ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(CCur(valcelant), z1$)
                            Else
                                ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                            End If
                        ElseIf dato_ent = 0 Then
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                        Else
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(dato_ent, z1$)
                        End If
                ElseIf g_TipoNominaActiva = tnBonoPremio Then
                    If dato_ent < 0 Or dato_ent > CelCur(ConNom1.Row, C_CONC7) Then
                        MsgBox "La percepcion exenta no puede ser negativa ni superar el bono total.", vbExclamation, "Captura de Bono"
                        If IsNumeric(valcelant) Then
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(CCur(valcelant), z1$)
                        Else
                            ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                        End If
                    ElseIf dato_ent = 0 Then
                        ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = ""
                    Else
                        ConNom1.TextMatrix(ConNom1.Row, C_EXENTA) = Format$(dato_ent, z1$)
                    End If

                Else
                    ConNom1.col = ruta
                    If dato_ent <> 0 Then
                        ConNom1.Text = Format$(dato_ent, z1$)
                    Else
                        ConNom1.Text = ""
                    End If
                End If

         Case Else
              ConNom1.col = ruta
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
    Dim Vw As Long
    Dim columna As Long
    Dim valorCelda As String
    Dim colantiLocal As Long
    Dim renatiLocal As Long

    ' REVISION 2026: sumar solo renglones con numero de empleado valido;
    ' no tomar las filas vacias ni las filas de total como empleados.
    colantiLocal = ConNom1.col
    renatiLocal = ConNom1.Row
    ProgressBar1.Min = 0
    ProgressBar1.Max = ConNom1.Rows
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True

    For columna = C_SUELDO To C_NETO
        sumv(columna) = 0
    Next columna

    Vw = 0
    For li = 1 To limite
        ProgressBar1.Value = li
        If IsNumeric(ConNom1.TextMatrix(li, C_NUM)) Then
            If Val(ConNom1.TextMatrix(li, C_NUM)) > 0 Then
                Vw = Vw + 1
                For columna = C_SUELDO To C_NETO
                    valorCelda = Trim$(ConNom1.TextMatrix(li, columna))
                    If Len(valorCelda) > 0 And IsNumeric(valorCelda) Then
                        sumv(columna) = sumv(columna) + CCur(valorCelda)
                    End If
                Next columna
            End If
        End If
    Next li

    li = limite + 1
    If ConNom1.Rows <= li Then ConNom1.Rows = li + 1
    ConNom1.TextMatrix(li, C_NOMBRE) = _
        "Empleados... " & CStr(Vw) & " S u m a s ....."

    For columna = C_SUELDO To C_NETO
        If sumv(columna) <> 0 Then
            ConNom1.TextMatrix(li, columna) = Format$(sumv(columna), z1$)
        Else
            ConNom1.TextMatrix(li, columna) = ""
        End If
    Next columna

    ConNom1.col = colantiLocal
    ConNom1.Row = renatiLocal
    ProgressBar1.Visible = False
End Sub

Sub sumah(si_impto1 As Integer, si_imss1 As Integer)

    Dim cPV As Long
    Dim baseISRAguinaldo As Currency

    If EsSeparacionNueva() Then
        sumahSeparacionNueva si_impto1, si_imss1
        Exit Sub
    End If

    sum(1) = 0
    sum(2) = 0
    sum(3) = 0
    colanti = ConNom1.col

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        cPV = C_LQ_PRIMAVAC
    Else
        cPV = C_CONC8
    End If
    
    If g_TipoNominaActiva = tnOrdinaria And N_ormal = 0 Then
        For late = 3 To 9
            If ConNom1.TextMatrix(li, late) <> "" Then
                sum(1) = sum(1) + CCur(ConNom1.TextMatrix(li, late))
            End If
        Next late
    Else
        For late = C_PERC_INI To C_PERC_FIN
            If ConNom1.TextMatrix(li, late) <> "" Then
                sum(1) = sum(1) + CCur(ConNom1.TextMatrix(li, late))
            End If
        Next late
    End If

    rgtro = ConNom1.TextMatrix(li, C_NUM)
    Acum_Doble
    
    '=========================================================
    ' BASE PARA EL CALCULO DE ISR
    '
    ' AGUINALDO, PTU y BONO:
    '
    ' Su columna principal contiene el importe TOTAL.
    ' La percepcion exenta forma parte de ese importe.
    '
    ' No modificar sum(1).
    '=========================================================

    baseISRAguinaldo = sum(1)

    If g_TipoNominaActiva = tnAguinaldo Or g_TipoNominaActiva = tnPTU Or g_TipoNominaActiva = tnBonoPremio Then
        baseISRAguinaldo = sum(1) - CelCur(li, C_EXENTA)

        If baseISRAguinaldo < 0 Then
            Err.Raise vbObjectError + 5901, "sumah", "La percepcion exenta supera el total de ingresos."
        End If
    End If

    
    If si_impto1 = 0 Then
        impto = 0
        cr_sal = 0

        If cal_anual = 0 Then
            If baseISRAguinaldo > 0 Then
                impto = 0
                calculoRetencion baseISRAguinaldo, impto, 1, regtro

            End If

        Else
            aoalta = Val(mID(personal.fal, 7, 4))

            If (aoalta = empresa.ao) And (Tot_dias < 330) Then
                impto = 0
                calculoRetencion baseISRAguinaldo, impto, 1, regtro

            Else
                rgtro = ConNom1.TextMatrix(li, C_NUM)
                aoalta = Val(mID(personal.fal, 7, 4))

                If (aoalta = empresa.ao) And (Tot_dias < 330) Then
                    impto = 0
                    calculoRetencion baseISRAguinaldo, impto, 1, regtro

                Else
                    impto = 0
                    calc_anual baseISRAguinaldo, impto, 1
                    If impto < 0 And ArAcum.DImpret > 0 Then
                        GoTo BRTodo
                    End If
                End If
            End If
        End If
        
BRTodo:
        If impto <> 0 And Option3 = True Then
            ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)

        Else
            ConNom1.TextMatrix(li, C_ISPT) = Format(impto, z1$)

            If Tot_dias > 330 And cal_anual = 1 Then
                GoTo saltaAnual
            End If

            If impto <> 0 And Option4 = True Then
            
                mesNomina1 = Trim(Left(Combo1.Text, 3))
                AnoNomina1 = Right(Trim(Label7.Caption), 4)

                Close 39

                Open CStr(Trim(Form1.Dir1)) & "\" & mesNomina1 & "1" & AnoNomina1 & ".NOM" For Random As 39 Len = Len(nomina)

                Get 39, regtro, nomina

                ImpuestoAnteriorRetenido = nomina.ispt

                ConNom1.TextMatrix(li, C_ISPT) = Format((impto - ImpuestoAnteriorRetenido), z1$)
                
                Close 39

            End If

saltaAnual:

        End If

        If impto = 0 Then
            ConNom1.TextMatrix(li, C_ISPT) = ""
        End If

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
        
        If g_TipoNominaActiva = tnOrdinaria And N_ormal = 0 Then
            inte_grar
            personal.integrado = integrado
        Else
            integrado = personal.integrado
        End If

        imss integrado, seguro, diaseg
        ConNom1.TextMatrix(li, C_IMSS) = Format(seguro, z1$)

        If seguro = 0 Then
            ConNom1.TextMatrix(li, C_IMSS) = ""
        End If
    End If

    If ConNom1.TextMatrix(li, C_FLAG_ESP) = "1" Then

        normal_1 = (baseISRAguinaldo / 365 * 30.4)
        
        If (normal_1 > 0) And (ConNom1.TextMatrix(li, C_FLAG_ISR) = "0") Then

            rgtro = ConNom1.TextMatrix(li, C_NUM)

            Get #2, regtro, personal

            Normal = personal.ingr * 30

            DifImptos.CalcDoble = 0

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
            
            impto = baseISRAguinaldo * porc_apli

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

        If Not (g_TipoNominaActiva = tnLiquidacionFiniquito And Val(ConNom1.TextMatrix(li, C_FLAG_IMSS)) = 1) Then
            ConNom1.TextMatrix(li, C_IMSS) = ""
        End If
    End If

    '=========================================================
    ' PERCEPCION EXENTA
    '
    ' En AGUINALDO, PTU y BONO:
    ' La exenta ya esta incluida en el importe bruto.
    '
    ' En las otras nominas:
    ' Conservar la logica anterior.
    '=========================================================

    ConNom1.col = C_EXENTA

    If g_TipoNominaActiva <> tnAguinaldo And g_TipoNominaActiva <> tnPTU And g_TipoNominaActiva <> tnBonoPremio Then
        If ConNom1.TextMatrix(li, C_EXENTA) <> "" Then
            sum(1) = sum(1) + CCur(ConNom1.TextMatrix(li, C_EXENTA))
        End If
    End If
    
    If sum(1) <> 0 Then
        ConNom1.TextMatrix(li, C_TOTING) = Format$(sum(1), z1$)
    Else
        ConNom1.TextMatrix(li, C_TOTING) = ""
    End If

    If sum(1) <> 0 Then
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

    ConNom1.col = colanti

End Sub


Sub inte_grar()

    ' REVISION 2026: ninguna nomina especial actualiza personal.dno.
    If g_TipoNominaActiva <> tnOrdinaria Or N_ormal <> 0 Then Exit Sub
     aoingr = Val(mID$(personal.fal, 7, 4))
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

    ConNom1.TextMatrix(li, C_NOMBRE) = RTrim$(personal.ape1) & " " & RTrim$(personal.ape2) & " " & RTrim$(personal.nom)
    PonCel li, C_DIAS, nomina.dias, "##0.00"
    ' Recuperar sueldo tanto en nuevas especiales como en normales/legacy.
    PonCel li, C_SUELDO, nomina.sueldo

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        If EsSeparacionNueva() Then
            If C_LQ_SDOVAC >= 0 Then PonCel li, C_LQ_SDOVAC, nomina.viaticos
            If C_LQ_PRIMAVAC >= 0 Then PonCel li, C_LQ_PRIMAVAC, nomina.pvac
            If C_LQ_AGUI >= 0 Then PonCel li, C_LQ_AGUI, nomina.aguin
            If C_LQ_ANTIG >= 0 Then PonCel li, C_LQ_ANTIG, nomina.hs_dbl
            If C_LQ_COMPENS >= 0 Then PonCel li, C_LQ_COMPENS, nomina.hs_nor
            If C_LQ_INDEMN >= 0 Then PonCel li, C_LQ_INDEMN, nomina.hs_tri
            If C_LQ_BONO >= 0 Then PonCel li, C_LQ_BONO, nomina.hsnor
            ingresos = nomina.sueldo + nomina.viaticos + nomina.pvac + nomina.aguin + nomina.hs_dbl + nomina.hs_nor + nomina.hs_tri + nomina.hsnor
        Else
            PonCel li, C_LQ_SDOVAC, nomina.viaticos
            PonCel li, C_LQ_PRIMAVAC, nomina.pvac
            PonCel li, C_LQ_AGUI, nomina.aguin
            PonCel li, C_LQ_ANTIG, nomina.hs_dbl
            PonCel li, C_LQ_COMPENS, nomina.hs_nor
            PonCel li, C_LQ_INDEMN, nomina.hs_tri
            PonCel li, C_LQ_BONO, nomina.hsnor
            ingresos = nomina.sueldo + nomina.viaticos + nomina.pvac + nomina.aguin + nomina.hs_dbl + nomina.hs_nor + nomina.hs_tri + nomina.hsnor
        End If
    Else

        '=====================================================
        ' CARGA DE PERCEPCIONES
        '
        ' Nomina ordinaria y especiales no-separacion.
        '=====================================================

        PonCel li, C_CONC4, nomina.hs_nor

        If N_ormal = 1 Then

            '-------------------------------------------------
            ' AGUINALDO ESPECIAL
            '
            ' El archivo NOM conserva:
            '
            ' nomina.aguin  = parte gravada
            ' nomina.exentos = parte exenta
            '
            ' En captura se muestra el TOTAL.
            '-------------------------------------------------

            If g_TipoNominaActiva = tnAguinaldo Then
                PonCel li, C_CONC5, nomina.aguin + nomina.exentos
            Else
                'Resto de especiales: comportamiento original.
                PonCel li, C_CONC5, nomina.aguin
            End If
        Else
            'Nomina ordinaria.
            PonCel li, C_CONC5, nomina.hs_dbl

        End If

        '=========================================================
        ' RECUPERAR PTU EN CAPTURA
        '
        ' Mostrar el importe TOTAL.
        ' No cambiar el registro NOM almacenado.
        '=========================================================
    
        If N_ormal = 1 Then
            If g_TipoNominaActiva = tnPTU Then
                PonCel li, C_CONC6, nomina.ptu + nomina.exentos
            Else
                PonCel li, C_CONC6, nomina.ptu
            End If
        Else
            PonCel li, C_CONC6, nomina.hs_tri
        End If

         If g_TipoNominaActiva = tnBonoPremio Then
            PonCel li, C_CONC7, nomina.viaticos + nomina.exentos
        Else
            PonCel li, C_CONC7, nomina.viaticos
        End If

        'Prima vacacional: comportamiento sin cambios.
        PonCel li, C_CONC8, nomina.pvac
        '=====================================================
        ' TOTAL DE INGRESOS ALMACENADOS
        '
        ' Aqui se utilizan los registros NOM.
        ' nomina.aguin sigue siendo la parte gravada.
        '
        ' El exento se agregara una sola vez mas abajo
        ' para reconstruir el importe bruto.
        '=====================================================

        ingresos = nomina.sueldo + nomina.hs_nor + nomina.hs_dbl + nomina.hs_tri + nomina.aguin + nomina.ptu + nomina.viaticos + nomina.pvac
    End If

    PonCel li, C_OTRAS, nomina.otras
    PonCel li, C_EXENTA, nomina.exentos
    ingresos = ingresos + nomina.otras
    ' LEGACY y normales: comportamiento original. Exenta es adicional.
    If Not EsSeparacionNueva() Then ingresos = ingresos + nomina.exentos
    PonCel li, C_TOTING, ingresos

    PonCel li, C_ISPT, nomina.ispt
    PonCel li, C_SUBEMP, nomina.crdsal
    PonCel li, C_IMSS, nomina.imss
    PonCel li, C_PRESTAMOS, nomina.prestamos
    PonCel li, C_FONACOT, nomina.fonacot
    PonCel li, C_PENSION, nomina.telefono
    PonCel li, C_INFONAVIT, nomina.otraded

    deducciones = nomina.crdsal + nomina.ispt + nomina.imss + nomina.prestamos + nomina.fonacot + nomina.telefono + nomina.otraded
    PonCel li, C_TOTDED, deducciones
    neto = ingresos - deducciones
    PonCel li, C_NETO, neto

    ConNom1.TextMatrix(li, C_FLAG_ISR) = 0
    
    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        'Conservar el IMSS almacenado en la nomina.
        ConNom1.TextMatrix(li, C_FLAG_IMSS) = "1"
    Else
        ConNom1.TextMatrix(li, C_FLAG_IMSS) = "0"
    End If

    
    '=========================================================
    ' BANDERA DE NOMINA ESPECIAL
    '
    ' Restablecer el comportamiento original:
    ' N_ormal = 1 identifica estas nominas especiales.
    '=========================================================

    ConNom1.TextMatrix(li, C_FLAG_ESP) = "0"
    
    If N_ormal = 1 Then
        ConNom1.TextMatrix(li, C_FLAG_ESP) = "1"
    End If

    '=========================================================
    ' DATOS BANCARIOS
    '
    ' Conservar el procedimiento original.
    '=========================================================

    locobra
    Get #12, regtro, Clbnx
    ConNom1.TextMatrix(li, C_BANAMEX) = " " & Clbnx.Q1

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

    ' Sueldo se muestra tambien en Liquidacion y Finiquito nuevos.
    Encab C_SUELDO, "Sueldo", 1200

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        If C_LQ_SDOVAC >= 0 Then Encab C_LQ_SDOVAC, "Sdo.Vacaciones", 1450
        If C_LQ_PRIMAVAC >= 0 Then Encab C_LQ_PRIMAVAC, "Prima Vacacional", 1450
        If C_LQ_AGUI >= 0 Then Encab C_LQ_AGUI, "Aguinaldo", 1200
        If C_LQ_ANTIG >= 0 Then Encab C_LQ_ANTIG, "Prima Antiguedad", 1450
        If C_LQ_COMPENS >= 0 Then Encab C_LQ_COMPENS, "Compensaciones", 1450
        If C_LQ_INDEMN >= 0 Then Encab C_LQ_INDEMN, "Indemnizacion", 1450
        If C_LQ_BONO >= 0 Then Encab C_LQ_BONO, "Bono", 1200
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
            Encab C_CONC7, "Bono", 1200
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

    
    '=========================================================
    ' COLUMNAS INTERNAS DEL SISTEMA
    '
    ' Mantener ocultas las banderas originales.
    '=========================================================

    ConNom1.ColWidth(C_FLAG_ISR) = 0
    ConNom1.ColWidth(C_FLAG_IMSS) = 0
    ConNom1.ColWidth(C_FLAG_ESP) = 0

    '=========================================================
    ' ULTIMA COLUMNA DE LA CAPTURA
    '=========================================================

    Encab C_BANAMEX, "Banamex", 2400

    '=========================================================
    ' VISIBILIDAD DE COLUMNAS POR TIPO DE NOMINA
    '
    ' Solo modifica el ancho visual de las columnas.
    '
    ' No elimina columnas.
    ' No modifica indices.
    ' No borra importes.
    ' No altera el guardado del archivo NOM.
    '=========================================================

    ConfigurarColumnasVisiblesEspeciales

End Sub

Private Sub ConfigurarColumnasVisiblesEspeciales()

    '=========================================================
    ' CONFIGURACION VISUAL DE NOMINAS ESPECIALES
    '
    ' IMPORTANTE:
    '
    ' El grid conserva todas sus columnas originales.
    '
    ' ColWidth = 0:
    '   Oculta visualmente una columna.
    '
    ' Los indices permanecen sin modificaciones.
    ' Los valores almacenados permanecen intactos.
    '
    ' No afecta el calculo ni el guardado de nomina.
    '=========================================================

    Select Case g_TipoNominaActiva

        '=====================================================
        ' AGUINALDO, PTU Y BONO
        '
        ' COLUMNAS COMUNES QUE NO SE UTILIZAN:
        '
        ' C_CONC4  = hs.Norm.
        ' C_CONC8  = P.Vacac.
        ' C_OTRAS  = Otras
        ' C_SUBEMP = Sub.P/Empl.
        '=====================================================

        Case tnAguinaldo, tnPTU, tnBonoPremio
            ConNom1.ColWidth(C_CONC4) = 0
            ConNom1.ColWidth(C_CONC8) = 0
            ConNom1.ColWidth(C_OTRAS) = 0
            ConNom1.ColWidth(C_SUBEMP) = 0

            '=================================================
            ' COLUMNAS ESPECIFICAS SEGUN EL TIPO DE NOMINA
            '
            ' Solamente debe quedar visible la columna
            ' correspondiente al concepto seleccionado.
            '
            ' C_EXENTA permanece visible en todos los casos.
            '=================================================

            Select Case g_TipoNominaActiva

                '---------------------------------------------
                ' AGUINALDO
                '
                ' MOSTRAR:
                '   Aguinaldo (C_CONC5)
                '   Perc.exenta. (C_EXENTA)
                '
                ' OCULTAR:
                '   PTU
                '   BONO
                '---------------------------------------------

                Case tnAguinaldo
                    ConNom1.ColWidth(C_CONC6) = 0
                    ConNom1.ColWidth(C_CONC7) = 0

                '---------------------------------------------
                ' PTU
                '
                ' MOSTRAR:
                '   PTU (C_CONC6)
                '   Perc.exenta. (C_EXENTA)
                '
                ' OCULTAR:
                '   Aguinaldo
                '   BONO
                '---------------------------------------------

                Case tnPTU
                    ConNom1.ColWidth(C_CONC5) = 0
                    ConNom1.ColWidth(C_CONC7) = 0

                '---------------------------------------------
                ' BONO
                '
                ' MOSTRAR:
                '   BONO (C_CONC7)
                '   Perc.exenta. (C_EXENTA)
                '
                ' OCULTAR:
                '   Aguinaldo
                '   PTU
                '---------------------------------------------

                Case tnBonoPremio
                    ConNom1.ColWidth(C_CONC5) = 0
                    ConNom1.ColWidth(C_CONC6) = 0

            End Select

        '=====================================================
        ' FINIQUITO Y LIQUIDACION
        '
        ' Conservar visibles todos los conceptos propios
        ' del formato seleccionado.
        '
        ' OCULTAR UNICAMENTE:
        '   Sub.P/Empl.
        '
        ' Funciona tanto con los indices del formato
        ' nuevo como con los del formato anterior.
        '=====================================================

        Case tnLiquidacionFiniquito
            ConNom1.ColWidth(C_SUBEMP) = 0
        
        '=====================================================
        ' NOMINA ORDINARIA / NORMAL
        '
        ' COLUMNAS QUE SE OCULTAN VISUALMENTE:
        '
        ' C_CONC4  = hs.Norm.
        ' C_CONC5  = hs.Dobles
        ' C_CONC6  = hs.Triples
        ' C_SUBEMP = Sub.P/Empl.
        '
        ' IMPORTANTE:
        ' Solamente se modifica el ancho de las columnas.
        '
        ' No eliminar columnas del grid.
        ' No modificar los indices.
        ' No borrar valores capturados.
        ' No alterar calculos ni archivos NOM.
        '=====================================================

        Case tnOrdinaria

            'Horas normales.
            ConNom1.ColWidth(C_CONC4) = 0

            'Horas dobles.
            ConNom1.ColWidth(C_CONC5) = 0

            'Horas triples.
            ConNom1.ColWidth(C_CONC6) = 0

            'Subsidio para el empleo.
            ConNom1.ColWidth(C_SUBEMP) = 0

        '=====================================================
        ' OTRAS MODALIDADES NO CONTEMPLADAS
        '
        ' Conservar su configuracion visual original.
        '=====================================================

        Case Else

            'Sin cambios adicionales.

    End Select
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
        aoalta = Val(mID(personal.fal, 7, 4))
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
        .dias = 0: .hsnor = 0: .hs_no = 0: .hsdbl = 0
        .hs_db = 0: .hstri = 0: .hs_tr = 0
        .ispt = 0: .crdsal = 0: .imss = 0: .sueldo = 0
        .hs_nor = 0: .hs_dbl = 0: .hs_tri = 0
        .viaticos = 0: .pvac = 0: .otras = 0: .aguin = 0
        .ptu = 0: .exentos = 0: .prestamos = 0
        .fonacot = 0: .telefono = 0: .otraded = 0
        If EsSeparacionNueva() Then .hstri = CCur(MARCA_SEPARACION_2026)
    End With
End Sub

Private Sub GridANomina(ByVal f As Long)
    nomina.dias = CelCur(f, C_DIAS)

    If EsSeparacionNueva() Then
        ' Conservar el sueldo calculado: dias T. x personal.ingr.
        nomina.sueldo = CelCur(f, C_SUELDO)
        nomina.viaticos = 0
        nomina.pvac = 0
        nomina.aguin = 0
        nomina.hs_dbl = 0
        nomina.hs_nor = 0
        nomina.hs_tri = 0
        nomina.hsnor = 0

        If C_LQ_SDOVAC >= 0 Then nomina.viaticos = CelCur(f, C_LQ_SDOVAC)
        If C_LQ_PRIMAVAC >= 0 Then nomina.pvac = CelCur(f, C_LQ_PRIMAVAC)
        If C_LQ_AGUI >= 0 Then nomina.aguin = CelCur(f, C_LQ_AGUI)
        If C_LQ_ANTIG >= 0 Then nomina.hs_dbl = CelCur(f, C_LQ_ANTIG)
        If C_LQ_COMPENS >= 0 Then nomina.hs_nor = CelCur(f, C_LQ_COMPENS)
        If C_LQ_INDEMN >= 0 Then nomina.hs_tri = CelCur(f, C_LQ_INDEMN)
        If C_LQ_BONO >= 0 Then nomina.hsnor = CelCur(f, C_LQ_BONO)

        ' Es una marca DE FORMATO, no importe ni horas triples.
        nomina.hstri = CCur(MARCA_SEPARACION_2026)
    Else
        ' Codigo previo: nominas ordinarias y TODAS las legacy.
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

            If g_TipoNominaActiva = tnAguinaldo Then
                nomina.aguin = CelCur(f, C_CONC5) - CelCur(f, C_EXENTA)
            Else
                nomina.aguin = CelCur(f, C_CONC5)
            End If
            Else
                nomina.hs_dbl = CelCur(f, C_CONC5)
            End If
            
            If N_ormal = 1 Then
                If g_TipoNominaActiva = tnPTU Then
                    nomina.ptu = CelCur(f, C_CONC6) - CelCur(f, C_EXENTA)
                Else
                    nomina.ptu = CelCur(f, C_CONC6)
                End If
            Else
                nomina.hs_tri = CelCur(f, C_CONC6)
            End If

            If g_TipoNominaActiva = tnBonoPremio Then
                nomina.viaticos = CelCur(f, C_CONC7) - CelCur(f, C_EXENTA)
            Else
                'Otras modalidades: conservar el guardado.
                nomina.viaticos = CelCur(f, C_CONC7)
            End If
            nomina.pvac = CelCur(f, C_CONC8)
        End If
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

    On Error GoTo ErrorGuardar

    '=====================================================
    ' PROTECCION DEL GUARDADO
    '
    ' CONSULTA:
    ' Nunca permitir archivar.
    '
    ' EDICION:
    ' Verificar nuevamente el plazo antes de
    ' continuar con el guardado original.
    '
    ' CAPTURA NUEVA:
    ' Conservar el comportamiento existente.
    '=====================================================

    If Not AutorizarCambioNominaForm8() Then
        GuardadoFallidoSesion = True
        MsgBox "Esta nomina no tiene autorizacion para guardar modificaciones.", vbExclamation, "Guardar nomina"
        Exit Sub
    End If

    GuardadoFallidoSesion = False

    Dim f As Long
    Dim colanti As Long
    Dim renati As Long
    Dim nuevaSeparacion As Boolean
    Dim primerRegistro As nom
    Dim hayRespaldo As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim detalleEdicion As String
    Dim copiaNom() As Byte
    Dim copiaCmpActual() As Byte
    Dim restauracionNomOK As Boolean

    Text4.Text = ""

    If Len(RutaNominaSesion) = 0 Then
        Err.Raise vbObjectError + 2027, "archnom_Click", "No hay una captura activa que archivar."
    End If
                    
    
    'Validaciones de seguridad para nominas existentes.
    If EditandoNominaGuardada Then
        If FalloRestauracionNom Then
            GuardadoFallidoSesion = True
            MsgBox "Existe un error de recuperacion pendiente." & vbCrLf & "No se permite otro guardado hasta revisar los archivos originales.", vbCritical, "Proteccion de nomina"
            Exit Sub
        End If
    
        If Not HayCopiaCmpEdicion Then
            GuardadoFallidoSesion = True
            MsgBox "No existe respaldo del CMP en memoria." & vbCrLf & "No se modificaran los archivos.", vbCritical, "Proteccion de nomina"
            Exit Sub
        End If
    End If
                
                    
    '====================================================
    ' VALIDAR PLAZO ANTES DE ESCRIBIR NOM O CMP
    '====================================================
    
    BloqueoPorPlazoEdicion = False
    
    If EditandoNominaGuardada Then
        If Not PuedeModificarNomina5Dias(RutaNominaSesion, detalleEdicion) Then
    
            BloqueoPorPlazoEdicion = True
            GuardadoFallidoSesion = True
    
            MsgBox detalleEdicion, vbExclamation, "Edicion vencida"
            
            Exit Sub
        End If
    End If
        
    '=========================================================
    ' VALIDAR AGUINALDO ANTES DE MODIFICAR EL ARCHIVO NOM
    '
    ' No altera las validaciones de otras modalidades.
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Then
        If Not ValidarCapturaAguinaldo() Then
            GuardadoFallidoSesion = True

            'No cerrar la captura ni archivar datos que no correspondan a esta modalidad.
            Exit Sub
        End If
    End If
    
    
    '=========================================================
    ' VALIDAR PTU ANTES DE GUARDAR EL ARCHIVO NOM
    '
    ' No modificar los datos capturados.
    ' No intervenir en otras modalidades.
    '=========================================================

    If g_TipoNominaActiva = tnPTU Then
        If Not ValidarCapturaPTU() Then
            GuardadoFallidoSesion = True
            Exit Sub
        End If
    End If
    
    
    '=========================================================
    ' VALIDAR BONO ANTES DE GUARDAR
    '
    ' No modifica percepciones ni deducciones.
    ' Solo impide archivar una captura inconsistente.
    '=========================================================

    If g_TipoNominaActiva = tnBonoPremio Then
        If Not ValidarCapturaBono() Then
            GuardadoFallidoSesion = True

            'No modificar el archivo NOM.
            Exit Sub
        End If
    End If
    
    
    '=========================================================
    ' VALIDAR FINIQUITO NUEVO ANTES DE ESCRIBIR EL NOM
    '
    ' No modifica los valores capturados.
    ' Si existe una inconsistencia, no guarda.
    '=========================================================

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        If g_EsFiniquito And EsSeparacionNueva() Then
            If Not ValidarCapturaFiniquitoNuevo() Then
                GuardadoFallidoSesion = True
                Exit Sub
            End If
        End If
    End If
    
    
    
    '=========================================================
    ' VALIDAR LIQUIDACION NUEVA ANTES DE GUARDAR
    '
    ' No modificar importes de la captura.
    ' No alterar otros tipos de nomina.
    '=========================================================

    If g_TipoNominaActiva = tnLiquidacionFiniquito Then
        If Not g_EsFiniquito And EsSeparacionNueva() Then
            If Not ValidarCapturaLiquidacionNueva() Then
                GuardadoFallidoSesion = True
                Exit Sub
            End If
        End If
    End If

    
    
    '=========================================================
    ' IDENTIFICAR FORMATO DE NOMINA
    '=========================================================

    nuevaSeparacion = EsSeparacionNueva()

    '=========================================================
    ' VALIDAR CMP DE NOMINA NORMAL NUEVA
    '
    ' Aplica a:
    '
    ' PRIMERA QUINCENA
    ' SEGUNDA QUINCENA
    '
    ' No modifica otras modalidades.
    '
    ' IMPORTANTE:
    ' Esta validacion solamente verifica registros.
    ' NO debe envolver el resto del guardado.
    '=========================================================

    If g_TipoNominaActiva = tnOrdinaria And NominaNuevaSesion And Not EditandoNominaGuardada Then

        '-----------------------------------------------------
        ' 1. VALIDAR EMPLEADOS
        '-----------------------------------------------------

        If Dm <= 0 Then
            Err.Raise vbObjectError + 7302, "archnom_Click", "No existen registros de personal para archivar esta nomina."
        End If

        '-----------------------------------------------------
        ' 2. CONSULTAR EL CMP ABIERTO
        '
        ' El archivo se encuentra abierto como #14.
        ' Utilizar LOF(14).
        '-----------------------------------------------------

        If LOF(14) <= 0 Then
            Err.Raise vbObjectError + 7304, "archnom_Click", "El complemento CMP se encuentra vacio." & vbCrLf & "No se archivara una nomina incompleta."
        End If

        '-----------------------------------------------------
        ' 3. VALIDAR LONGITUD DE REGISTROS
        '-----------------------------------------------------

        If LOF(14) Mod Len(nom_com) <> 0 Then
            Err.Raise vbObjectError + 7303, "archnom_Click", "El complemento CMP contiene registros incompletos."
        End If

        '-----------------------------------------------------
        ' 4. VALIDAR CANTIDAD DE EMPLEADOS
        '-----------------------------------------------------

        If (LOF(14) \ Len(nom_com)) < Dm Then
            Err.Raise vbObjectError + 7304, "archnom_Click", "El complemento CMP no contiene los registros necesarios."
        End If
    End If

    '=========================================================
    ' FIN DE LA VALIDACION
    '
    ' IMPORTANTE:
    ' A PARTIR DE AQUI EL GUARDADO DEBE CONTINUAR
    ' PARA AMBAS QUINCENAS Y TODAS LAS MODALIDADES.
    '=========================================================

    'Conservar una copia del NOM en memoria.
    Close #6

    hayRespaldo = False
    
    If Len(Dir$(RutaNominaSesion)) > 0 Then
        If FileLen(RutaNominaSesion) > 0 Then
            LeerBytesArchivo RutaNominaSesion, copiaNom
            hayRespaldo = True
        End If
    End If
    
    Dm = LOF(2) / Len(personal)
    Open RutaNominaSesion For Random As #6 Len = Len(nomina)
    nm = LOF(6) / Len(nomina)

    colanti = ConNom1.col
    renati = ConNom1.Row
    
    '====================================================
    ' INICIALIZAR REGISTROS SOLO EN NOMINAS NUEVAS
    '====================================================
    If Not EditandoNominaGuardada Then
        LimpiaNomina
    
        If nuevaSeparacion Then
            nomina.hstri = CCur(MARCA_SEPARACION_2026)
        End If
        For f = 1 To Dm
            Put #6, f, nomina
        Next f
    End If

    ' 2) Guardar cada empleado presente en la cuadricula.
    For f = 1 To limite
        regtro = Val(ConNom1.TextMatrix(f, C_NUM))
        If regtro > 0 Then
            Get #6, regtro, nomina
            GridANomina f

            If nuevaSeparacion Then
                nomina.hstri = CCur(MARCA_SEPARACION_2026)
            End If
            
            If EditandoNominaGuardada Then
                If regtro > nm Then
                    Err.Raise vbObjectError + 3020, , "La edicion no puede ampliar el archivo NOM."
                End If
            End If

            Put #6, regtro, nomina
        End If
    Next f

    ' 3) Verificacion basica del primer registro.
    If nuevaSeparacion And Dm > 0 Then
        Get #6, 1, primerRegistro
        If primerRegistro.hstri <> CCur(MARCA_SEPARACION_2026) Then
            Err.Raise vbObjectError + 2026, "archnom_Click", "No se pudo conservar la marca de formato nuevo."
        End If
    End If

    Close #6

    archiva_o = 0

    ConNom1.col = colanti
    ConNom1.Row = renati
    
    'No cambiar el foco cuando se esta cerrando Form8.
    If Not GuardandoDuranteCierre Then
        If Me.Visible And ConNom1.Visible Then
            ConNom1.SetFocus
        End If
    End If
    'Registrar fecha fiscal solamente en nominas nuevas. En una edicion conservar la fecha fiscal existente.
    If nuevaSeparacion And Not EditandoNominaGuardada Then
        FechaCFDI_GuardarFicha RutaNominaSesion, CStr(MiFecha)
    End If
    
    
    '=========================================================
    ' RESPALDAR EL ULTIMO CMP GUARDADO CORRECTAMENTE
    '
    ' EDICION:
    ' Conservar el comportamiento de respaldo actual.
    '
    ' NOMINA ORDINARIA NUEVA:
    ' Guardar tambien una copia en memoria despues
    ' de archivar por primera vez.
    '
    ' Esto permite descartar cambios posteriores
    ' sin sobrescribir el CMP archivado.
    '=========================================================

    If EditandoNominaGuardada Or (g_TipoNominaActiva = tnOrdinaria And NominaNuevaSesion) Then

        '-----------------------------------------------------
        ' VERIFICAR RESPALDO PREVIO EN EDICION
        '-----------------------------------------------------

        If EditandoNominaGuardada Then
            If Not HayCopiaCmpEdicion Then
                Err.Raise vbObjectError + 3021, "archnom_Click", "No existe respaldo CMP en memoria."
            End If
        End If

        '-----------------------------------------------------
        ' LEER LA VERSION QUE ACABA DE GUARDARSE
        '-----------------------------------------------------

        Close #14
        LeerBytesArchivo RutaCmpSesion, copiaCmpActual

        '-----------------------------------------------------
        ' RESTABLECER EL ACCESO NORMAL AL CMP
        '-----------------------------------------------------

        Open RutaCmpSesion For Random As #14 Len = Len(nom_com)

        '-----------------------------------------------------
        ' ACTUALIZAR RESPALDO EN MEMORIA
        '
        ' La copia representa el ultimo guardado
        ' exitoso, no los cambios posteriores.
        '-----------------------------------------------------

        CopiarBytesMemoria copiaCmpActual, CopiaCmpEdicion
        HayCopiaCmpEdicion = True
    End If

    '=========================================================
    ' REGISTRAR NUEVA GENERACION DE NOMINA
    '
    ' Solamente cuando:
    '
    ' 1. Es una captura nueva.
    ' 2. No estamos editando una nomina existente.
    ' 3. La captura aun no tiene un guardado exitoso.
    '
    ' Evitar incrementar el historial al guardar
    ' varias veces durante la misma sesion.
    '=========================================================

    If Not EditandoNominaGuardada And NominaNuevaSesion And Not NominaArchivadaSesion Then
        RegistrarPrimerGuardadoNomina RutaNominaSesion
    End If

    '=====================================================
    ' CONFIRMAR GUARDADO ORIGINAL
    '=====================================================

    '=====================================================
    ' GUARDADO COMPLETADO CORRECTAMENTE
    '
    ' Este es el unico punto donde se confirma
    ' que ya no existen cambios pendientes.
    '=====================================================

    NominaArchivadaSesion = True
    GuardadoFallidoSesion = False

    archiva_o = 0

    MsgBox "Archivada con exito!", vbInformation, "Nomina"
    Exit Sub



ErrorGuardar:

    errNum = Err.Number
    errDesc = Err.Description

    GuardadoFallidoSesion = True

    On Error Resume Next

    'Cerrar el archivo NOM antes de cualquier recuperacion.
    Close #6

    restauracionNomOK = True

    'Recuperar NOM desde la memoria si se habia guardado copia.
    If hayRespaldo Then

        Err.Clear

        RestaurarBytesArchivo RutaNominaSesion, copiaNom

        restauracionNomOK = (Err.Number = 0)

        If Not restauracionNomOK Then
            FalloRestauracionNom = True
        End If

    ElseIf EditandoNominaGuardada Then

        'Una nomina existente siempre debio tener respaldo.
        FalloRestauracionNom = True

    End If

    'Recuperar el acceso normal al CMP.
    If EditandoNominaGuardada Then

        Close #14

        If Len(RutaCmpSesion) > 0 Then
            If Len(Dir$(RutaCmpSesion, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then

                Err.Clear

                Open RutaCmpSesion For Random As #14 Len = Len(nom_com)
                
                If Err.Number <> 0 Then
                    Debug.Print "No fue posible reabrir CMP: " & Err.Description
                End If
            End If
        End If
    End If

    On Error GoTo 0

    If FalloRestauracionNom Then
        MsgBox "Ocurrio un error durante el guardado." & vbCrLf & "No fue posible asegurar la recuperacion del archivo NOM." & vbCrLf & vbCrLf & "No continue modificando esta nomina." & vbCrLf & "Error " & CStr(errNum) & ": " & errDesc, vbCritical, "Error grave al guardar"
    Else
        MsgBox "No fue posible completar el guardado." & vbCrLf & "El formulario permanecera abierto." & vbCrLf & vbCrLf & "Error " & CStr(errNum) & ": " & errDesc, vbExclamation, "Guardar nomina"
    End If
End Sub

Private Sub cap_act_Click(Index As Integer)
    sumavert
End Sub

Private Sub capacimp_Click(Index As Integer)
    FuncionImpuesto
End Sub

Private Sub FuncionImpuesto()
    '=====================================================
    ' NO RECALCULAR NOMINAS EN CONSULTA
    '
    ' FuncionImpuesto utiliza sumah y puede
    ' actualizar registros del complemento CMP.
    '=====================================================

    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If
    rrenat = ConNom1.Row
    rcolant = ConNom1.col

    ProgressBar1.Min = 0
    ProgressBar1.Max = ConNom1.Rows
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True

    For li = 1 To ConNom1.Rows - 3
        ProgressBar1.Value = li
        regtro = 0

        If IsNumeric(ConNom1.TextMatrix(li, C_NUM)) Then
            regtro = Val(ConNom1.TextMatrix(li, C_NUM))
            If regtro > 0 Then
                Get #2, regtro, personal
            End If
        End If

        ' La bandera (0 = calcular, 1 = ISR capturado a mano) vive en C_FLAG_ISR para TODOS los tipos de nomina. El importe de C_ISPT nunca es bandera.
        impto1 = Val(ConNom1.TextMatrix(li, C_FLAG_ISR))
        imss1 = 0
        
        'En FINIQUITO y LIQUIDACION respetar la captura manual del IMSS de cada empleado.
        
        If g_TipoNominaActiva = tnLiquidacionFiniquito Then
            imss1 = Val(ConNom1.TextMatrix(li, C_FLAG_IMSS))
        End If
        
        If regtro > 0 Then
            sumah impto1, imss1
        End If

    Next li

    ProgressBar1.Visible = False
    ConNom1.Row = rrenat
    ConNom1.col = rcolant
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
    colanti = ConNom1.col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.col = 1
    ConNom1.RowSel = limite
    ConNom1.Sort = 1
    ConNom1.col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus
    End Sub

Private Sub capnomnum_Click(Index As Integer)
    colanti = ConNom1.col
    renati = ConNom1.Row
    ConNom1.Row = 1
    ConNom1.col = 0
    ConNom1.RowSel = limite
    ConNom1.Sort = 3
    ConNom1.col = colanti
    ConNom1.Row = renati
    ConNom1.SetFocus
End Sub

Private Sub capnosal_Click(Index As Integer)

    '=====================================================
    ' SALIDA DE FORM8
    '
    ' Centralizar el cierre en Form_QueryUnload.
    '
    ' Form_QueryUnload:
    '   Pregunta si desea guardar.
    '   Permite cancelar el cierre.
    '
    ' Form_Unload:
    '   Finaliza y limpia el estado de la captura.
    '
    ' No limpiar la sesion antes de confirmar.
    '=====================================================

    Unload Me
End Sub



Private Sub caprecind_Click(Index As Integer)
    colanti = ConNom1.col
    renati = ConNom1.Row
    archnom_Click (1)
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
        ConNom1.col = colanti
        ConNom1.Row = renati
        ConNom1.SetFocus
    Else
        MsgBox "Necesita Capturar Nomina para Imprimirla"
    End If
End Sub

Private Sub caprectd_Click(Index As Integer)
   colanti = ConNom1.col
   renati = ConNom1.Row
   ConNom1.Row = 1
    ConNom1.col = 0
   If ConNom1.Text <> "" Then rgtro = ConNom1.Text Else rgtro = 0
   If rgtro > 0 Then
            regtro = ConNom1.TextMatrix(ConNom1.Row, 0)
            Get 14, rgtro, nom_com
            lin_rbo 1, limite
            ConNom1.col = colanti
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
   ConNom1.col = 2
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not Option1.Value Then Exit Sub
        If Not Option3.Value And Not Option4.Value Then
            MsgBox "Seleccione Primera o Segunda quincena.", vbExclamation, "Nomina"
            Exit Sub
        End If

        If Combo1.ListIndex < 0 Then
            MsgBox "Seleccione un mes valido.", vbExclamation, "Nomina"
            Exit Sub
        End If

        'Mantener la actualizacion original del mes.
        Combo1_Change
        
        'Ejecutar Captura -> Iniciar captura.
        IniciarCapturaDesdeTeclado
    End If
End Sub

'=========================================================
' IDENTIFICAR FILA DE SUMAS
'
' Esta fila contiene los totales calculados
' automaticamente mediante sumavert().
'
' Se permite:
' - Seleccionar la fila.
' - Navegar por sus columnas.
' - Consultar los importes.
'
' Se impide:
' - Modificar importes.
' - Borrar cantidades.
' - Abrir el editor de celdas.
'
' No modifica sumavert ni los calculos.
'=========================================================

Private Function EsFilaSumasCaptura() As Boolean
    EsFilaSumasCaptura = False

    'Validar que exista una fila seleccionada.
    If ConNom1.Row <= 0 Then Exit Function
    If ConNom1.Row >= ConNom1.Rows Then
        Exit Function
    End If

    'sumavert coloca SUMAS despues del ultimo empleado.
    If ConNom1.Row <> limite + 1 Then
        Exit Function
    End If

    'Confirmar que sea realmente la fila de totales.
    If InStr(1, ConNom1.TextMatrix(ConNom1.Row, C_NOMBRE), "S u m a s", vbTextCompare) = 0 Then
        Exit Function
    End If
    EsFilaSumasCaptura = True
End Function

Private Sub ConNom1_EnterCell()
    If ConNom1.col > 1 And ConNom1.Row > 0 Then
        ConNom1.CellBackColor = &H80FF80
    End If
End Sub

Private Sub ConNom1_KeyDown(KeyCode As Integer, Shift As Integer)

    '=====================================================
    ' PROTEGER FILA SUMAS
    '
    ' F2: No abrir el editor.
    ' SUPR: No borrar importes.
    '
    ' Las flechas y Tab siguen funcionando.
    '=====================================================

    If EsFilaSumasCaptura() Then
        Select Case KeyCode
            Case vbKeyDelete, vbKeyF2
                KeyCode = 0
                Exit Sub
        End Select
    End If

    '=====================================================
    ' CONSERVAR PROTECCION DE SOLO CONSULTA
    '=====================================================

    If KeyCode = vbKeyDelete Or KeyCode = vbKeyF2 Then
        If Not AutorizarCambioNominaForm8() Then
            KeyCode = 0
            Exit Sub
        End If
    End If

    '=====================================================
    ' COMPORTAMIENTO ORIGINAL DEL GRID
    '=====================================================

    Select Case KeyCode
        Case vbKeyDelete
            If Not EsColumnaNumericaCaptura(ConNom1.col) Then
                KeyCode = 0
                Exit Sub
            End If
            
            valcelant = ConNom1.Text
            ConNom1.Text = ""
            checar
            KeyCode = 0

        Case vbKeyF2
            If Not EsColumnaNumericaCaptura(ConNom1.col) Then
                If ConNom1.col <> C_BANAMEX Then
                    KeyCode = 0
                    Exit Sub
                End If
            End If

            valcelant = ConNom1.Text
            
            If EsColumnaNumericaCaptura(ConNom1.col) And IsNumeric(ConNom1.Text) Then
                'Retirar separadores de miles al editar.
                Text2.Text = LTrim$(Str$(CCur(ConNom1.Text)))
            Else
                Text2.Text = ConNom1.Text
            End If
            Text2.SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub ConNom1_KeyPress(KeyAscii As Integer)

    '=====================================================
    ' IMPEDIR ESCRITURA DIRECTA EN LA FILA SUMAS
    '
    ' Si el usuario selecciona una celda de totales
    ' y empieza a escribir un numero, letra o simbolo,
    ' no se abrira Text2.
    '
    ' Enter se conserva para la navegacion.
    '=====================================================

    If EsFilaSumasCaptura() Then
        If KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    '=====================================================
    ' IMPEDIR APERTURA DEL EDITOR EN SOLO CONSULTA

    '
    ' Conservar Enter para navegacion.
    '=====================================================

    If KeyAscii <> vbKeyReturn Then
        If Not AutorizarCambioNominaForm8() Then
            KeyAscii = 0
            Exit Sub
        End If
    End If

    '=========================================================
    ' AGUINALDO - PROTEGER COLUMNAS NO CAPTURABLES
    '
    ' Impedir que escribir directamente sobre el grid
    ' abra el editor de una percepcion no permitida.
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Or g_TipoNominaActiva = tnPTU Or g_TipoNominaActiva = tnBonoPremio Then
        If Not EsColumnaNumericaCaptura(ConNom1.col) And ConNom1.col <> C_BANAMEX Then

            'Conservar Enter para navegacion.
            If KeyAscii <> 13 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If

    '=========================================================
    ' COMPORTAMIENTO ORIGINAL
    '=========================================================

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
   If ConNom1.col > 1 And ConNom1.Row > 0 Then
         ConNom1.CellBackColor = vbWhite
   End If
End Sub

Private Sub ConNom1_RowColChange()
        Text2.Text = ConNom1.Text
End Sub


Public Sub main()
    ' REVISION 2026: no abrir archivos sin registrar la sesion.
    nomcap_ini_Click 1
End Sub

Private Sub IniciarCapturaDesdeTeclado()

    '=====================================================
    ' INICIAR CAPTURA MEDIANTE ENTER
    '
    ' Utilizar EXACTAMENTE la misma rutina del menu:
    '
    ' Captura -> Iniciar captura
    '
    ' No duplicar la logica de generacion ni guardado.
    '=====================================================

    If CapturaActivaSesion Then Exit Sub
    nomcap_ini_Click 1

    '=====================================================
    ' SI SE INICIO CORRECTAMENTE LA CAPTURA,
    ' POSICIONAR EL FOCO EN LA CUADRICULA.
    '
    ' Si el archivo ya existia o hubo un error,
    ' no forzar el foco sobre el grid.
    '=====================================================

    If CapturaActivaSesion Then
        If limite > 0 Then
            ConNom1.Row = 1
            Select Case g_TipoNominaActiva
                Case tnAguinaldo
                    ConNom1.col = C_CONC5
                Case tnPTU
                    ConNom1.col = C_CONC6
                Case tnBonoPremio
                    ConNom1.col = C_CONC7
                    
                Case tnLiquidacionFiniquito
                    If g_EsFiniquito Then
                        If C_LQ_SDOVAC >= 0 Then
                            ConNom1.col = C_LQ_SDOVAC
                        End If
                    Else
                        If C_LQ_COMPENS >= 0 Then
                            ConNom1.col = C_LQ_COMPENS
                        End If
                    End If
                Case Else
                    ConNom1.col = C_DIAS
            End Select
            ConNom1.SetFocus
        End If
    End If
End Sub

Private Sub generarNominas()
    Dim complementoNuevo As nomco
    Dim indiceCMP As Long

    ProgressBar1.Visible = True
    QUIN = 0
  
  ' Si es una Nomina normal
    If Option1 = True Then
        ' Si es Primera quincena
        If Option3 = True Then
            QUIN = 1
            Arch$ = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.mes = (Combo1.ListIndex + 1)
            ReferOper.dia = 15
        Else
        ' si es Segunda Quincena
            QUIN = 2
            Arch$ = UCase(mID$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(mID$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ArchAnterior = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1Anterior = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.mes = (Combo1.ListIndex + 1)
            ReferOper.dia = dd(ReferOper.mes)
            
            Close 95
    
            If Len(Dir$(ArchAnterior)) > 0 Then
                Open ArchAnterior For Random As 95 Len = Len(nominaAnterior)
                largoNomina = LOF(95) / Len(nominaAnterior)
            End If
    
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

       ' No crear el .NOM al iniciar: se crea hasta que se archiva.
    If Len(Dir$(Arch$)) > 0 Then
        Open Arch$ For Random As 6 Len = Len(nomina)
        nm = LOF(6) / Len(nomina)
    Else
        nm = 0
    End If
    
  
    '=========================================================
    ' INICIALIZAR COMPLEMENTO DE NOMINA NORMAL
    '
    ' Aplica para:
    ' - Primera quincena nueva.
    ' - Segunda quincena nueva.
    '
    ' No modifica nominas especiales.
    ' No sobrescribe complementos existentes.
    '=========================================================

    Open Arch1 For Random As #14 Len = Len(nom_com)

    If Option1.Value = True And nm = 0 Then

        '-----------------------------------------------------
        ' 1. COMPROBAR NUMERO DE EMPLEADOS
        '-----------------------------------------------------

        If Dm <= 0 Then
            Err.Raise vbObjectError + 7401, "generarNominas", "No existen empleados para generar el complemento de nomina."
        End If

        '-----------------------------------------------------
        ' 2. ESTABLECER CANTIDAD DE REGISTROS
        '
        ' ceroscompl utiliza fi_nm.
        '
        ' Debe recibir el numero real de empleados
        ' tanto en primera como en segunda quincena.
        '-----------------------------------------------------

        fi_nm = Dm

        '-----------------------------------------------------
        ' 3. VALIDAR ESTRUCTURA DEL CMP EXISTENTE
        '
        ' No alterar archivos que ya contienen
        ' registros incompletos.
        '-----------------------------------------------------

        If LOF(14) Mod Len(nom_com) <> 0 Then
            Err.Raise vbObjectError + 7402, "generarNominas", "El complemento CMP contiene registros incompletos."
        End If

        '-----------------------------------------------------
        ' 4. INICIALIZAR COMPLEMENTO NUEVO
        '
        ' Si el archivo tiene cero bytes, escribir
        ' un registro por cada empleado.
        '
        ' Se conserva el tipo nomco original.
        '-----------------------------------------------------

        If LOF(14) = 0 Then
            For indiceCMP = 1 To Dm
                Put #14, indiceCMP, complementoNuevo
            Next indiceCMP
        Else

            '-------------------------------------------------
            ' CMP EXISTENTE
            '
            ' No completar silenciosamente un archivo
            ' que podria proceder de una captura fallida.
            '-------------------------------------------------

            If (LOF(14) \ Len(nom_com)) < Dm Then
                Err.Raise vbObjectError + 7403, "generarNominas", "El complemento CMP existente o contiene todos los registros." & vbCrLf & "Debe revisarse antes de continuar."
            End If
        End If
    End If

    '=========================================================
    ' CONTINUAR CON LA APERTURA ORIGINAL
    '=========================================================

    Open "bnxcla.dno" For Random As #12 Len = Len(Clbnx)


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
            
            aobaja = Val(mID$(personal.fab, 7, 4))
            mesbaja = Val(mID$(personal.fab, 4, 2))
            diabaja = Val(mID$(personal.fab, 1, 2))
            
            verifica yavas
                
            If yavas > 0 Then ' Aplica para primera quincena y nomina especial
                renglon = renglon + 1
                ConNom1.Row = renglon
                ConNom1.col = 0
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
                aobaja = Val(mID$(personal.fab, 7, 4))
                mesbaja = Val(mID$(personal.fab, 4, 2))
                diabaja = Val(mID$(personal.fab, 1, 2))
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
                
                If mID$(personal.nom, 1, 3) <= "   " Then GoTo SIGUE
                
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
    ConNom1.col = 2
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
            Arch$ = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.mes = (Combo1.ListIndex + 1)
            ReferOper.dia = 15
        Else
        ' si es Segunda Quincena
            QUIN = 2
            Arch$ = UCase(mID$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1 = UCase(mID$(Combo1.Text, 1, 3)) + "2" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ArchAnterior = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".NOM"
            Arch1Anterior = UCase(mID$(Combo1.Text, 1, 3)) + "1" + LTrim$(Str$(empresa.ao)) + ".cmp"
            ReferOper.mes = (Combo1.ListIndex + 1)
            ReferOper.dia = dd(ReferOper.mes)
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
            
            aobaja = Val(mID$(personal.fab, 7, 4))
            mesbaja = Val(mID$(personal.fab, 4, 2))
            diabaja = Val(mID$(personal.fab, 1, 2))
            
            verifica yavas
                
            renglon = renglon + 1
            ConNom1.Row = renglon
            ConNom1.col = 0
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
                    ConNom1.col = 0
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
                aobaja = Val(mID$(personal.fab, 7, 4))
                mesbaja = Val(mID$(personal.fab, 4, 2))
                diabaja = Val(mID$(personal.fab, 1, 2))
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
        ConNom1.col = 2
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
    'Solo permitir la operacion durante captura nueva o edicion autorizada.

    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If

    
    If ConNom1.Rows < 5 Then
        Exit Sub
    End If
        
    eliminacion
    If (ConNom1.Rows <> 4) Then
        colanti = ConNom1.col
        renati = ConNom1.Row
        ConNom1.Row = 1
        ConNom1.col = 1
        ConNom1.RowSel = limite
        
        ConNom1.Sort = 1
        ConNom1.col = colanti
        ConNom1.Row = renati
        ConNom1.SetFocus
    End If
    
    If (ConNom1.TextMatrix(1, 0) = "") Then
        ConNom1.RemoveItem 1
        ConNom1.Rows = ConNom1.Rows + 1
        
    End If
End Sub

Private Sub Option1_GotFocus()

    'No avanzar automaticamente al recibir el foco.
    'Permitir seleccionar Normal o Especial con flechas.

    If Option1.Value = True Then
        N_ormal = 0
    End If

End Sub


Private Sub Option1_Click()
    N_ormal = 0
    Text3.Locked = False

    txtPre.Text = ""
    cboTipoNominaEspecial.ListIndex = -1

    RefrescarModoNomina
End Sub

Private Sub Option2_Click()
On Error GoTo ErrorManejador

    If Option2 = True Then
        Option3 = False
        Option4 = False

        N_ormal = 1
        Text3.Locked = False
        dia_pago = 15 - 1

        ' Reinicio de captura especial
        txtPre.Text = ""
        cboTipoNominaEspecial.ListIndex = -1
        Text1.Text = ""
        Label7.Caption = ""
        cboMesNominaEspecial.ListIndex = Month(Date) - 1

        RefrescarModoNomina
    End If
Exit Sub
ErrorManejador:
    MsgBox Err.Number & " " & Err.Description
End Sub

Public Sub presionarOption2()
    Option2_Click
End Sub

Private Sub Option3_Click()
    If meselegido = 0 Then meselegido = 1

    If Option3 = True Then
        Label7.Caption = "Nomina de la 1a.quincena " + Combo1.Text + Str$(empresa.ao)

        dia_pago = 15 - 1

        Combo1.Visible = True
        Combo1.Locked = False
    End If
End Sub

Private Sub Option4_Click()
    If meselegido = 0 Then meselegido = 1

    If Option4 = True Then
        Label7.Caption = "Nomina de la 2a.quincena " + Combo1.Text + Str$(empresa.ao)
        dia_pago = dd(meselegido) - 1
        If meselegido = 12 Then
            respuesta = MsgBox("Desea hacer Calculo Anual", vbCritical + vbYesNo, "Captura Ultima Nomina")

            If respuesta = vbYes Then
                cal_anual = 1
            Else
                cal_anual = 0
            End If
        End If

        Combo1.Visible = True
    End If
End Sub

'=========================================================
' NAVEGACION DEL TIPO DE NOMINA
'=========================================================

Private Sub Option1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        KeyCode = 0
        If Not Option1.Value Then
            Option1.Value = True
        End If

        'Normal -> Seleccion de quincena.
        Option3.SetFocus
    End If
End Sub

Private Sub Option2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        KeyCode = 0
        If Not Option2.Value Then
            Option2.Value = True
        End If

        'Especial -> Seleccionar concepto.
        If cboTipoNominaEspecial.Visible And cboTipoNominaEspecial.Enabled Then
            cboTipoNominaEspecial.SetFocus
        End If
    End If
End Sub

'=========================================================
' NAVEGACION DE QUINCENA
'=========================================================

Private Sub Option3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        KeyCode = 0
        If Not Option3.Value Then
            Option3.Value = True
        End If

        'Primera -> Mes.
        Combo1.SetFocus
    End If
End Sub

Private Sub Option4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
        KeyCode = 0
        If Not Option4.Value Then
            Option4.Value = True
        End If

        'Segunda -> Mes.
        Combo1.SetFocus
    End If
End Sub

Public Sub presionarEnterForm8()
Rem    Text1_KeyPress 13
End Sub

Private Function EsColumnaNumericaCaptura(ByVal columna As Long) As Boolean
    EsColumnaNumericaCaptura = False
    
      '=========================================================
    ' NOMINA ESPECIAL BONO
    '
    ' PERCEPCIONES PERMITIDAS:
    '   C_CONC7  = Bono total
    '   C_EXENTA = Parte exenta del bono
    '
    ' Se conservan las deducciones.
    '
    ' No permitir editar las otras percepciones.
    '=========================================================

    If g_TipoNominaActiva = tnBonoPremio Then
        Select Case columna
            Case C_CONC7, C_EXENTA, C_ISPT, C_SUBEMP, C_IMSS, C_PRESTAMOS, C_FONACOT, C_PENSION, C_INFONAVIT
                EsColumnaNumericaCaptura = True
        End Select
        Exit Function
    End If
    
    '=========================================================
    ' NOMINA ESPECIAL PTU
    '
    ' Percepciones editables:
    '   C_CONC6  = PTU total
    '   C_EXENTA = PTU exenta
    '
    ' Conservar las deducciones capturables.
    ' No permitir otras percepciones.
    '=========================================================

    If g_TipoNominaActiva = tnPTU Then
        Select Case columna
            Case C_CONC6, C_EXENTA, C_ISPT, C_SUBEMP, C_IMSS, C_PRESTAMOS, C_FONACOT, C_PENSION, C_INFONAVIT
                EsColumnaNumericaCaptura = True

        End Select
        Exit Function
    End If
    
    '=========================================================
    ' NOMINA ESPECIAL AGUINALDO
    '
    ' Percepciones permitidas:
    '   C_CONC5  = Aguinaldo
    '   C_EXENTA = Percepcion exenta de Aguinaldo
    '
    ' Se conservan las deducciones capturables.
    '
    ' No permitir editar:
    ' - Dias
    ' - Sueldo
    ' - Horas
    ' - PTU
    ' - Bono
    ' - Prima vacacional
    ' - Otras percepciones
    ' - Totales
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Then
        Select Case columna
            Case C_CONC5, C_EXENTA, C_ISPT, C_SUBEMP, C_IMSS, C_PRESTAMOS, C_FONACOT, C_PENSION, C_INFONAVIT
                EsColumnaNumericaCaptura = True
        End Select

        Exit Function

    End If

    '=========================================================
    ' RESTO DE NOMINAS
    '
    ' Conservar exactamente las reglas anteriores.
    '=========================================================

    Select Case columna
        Case C_NUM, C_NOMBRE, C_SUELDO, C_TOTING, C_TOTDED, C_NETO, C_BANAMEX, C_FLAG_ISR, C_FLAG_IMSS, C_FLAG_ESP
            Exit Function
    End Select

    EsColumnaNumericaCaptura = (columna = C_DIAS Or columna = C_EXENTA Or (columna >= C_PERC_INI And columna <= C_PERC_FIN) Or (columna >= C_DED_INI And columna <= C_DED_FIN))

End Function


Private Function ImporteCapturaValido(ByVal texto As String) As Boolean

    Dim i As Long
    Dim c As String
    Dim posicionPunto As Long
    Dim numero As Currency

    ImporteCapturaValido = False
    texto = Trim$(texto)

    'Permitir borrar una cantidad.
    If Len(texto) = 0 Then
        ImporteCapturaValido = True
        Exit Function
    End If

    If texto = "." Then Exit Function

    For i = 1 To Len(texto)

        c = mID$(texto, i, 1)

        If c >= "0" And c <= "9" Then
            'Digito permitido.

        ElseIf c = "." Then

            If posicionPunto > 0 Then Exit Function
            posicionPunto = i

        Else
            'Rechazar letras, signos y simbolos.
            Exit Function
        End If

    Next i

    If posicionPunto > 0 Then
        If Len(texto) - posicionPunto > 2 Then
            Exit Function
        End If
    End If

    On Error GoTo NumeroIncorrecto

    numero = CCur(texto)

    ImporteCapturaValido = True
    Exit Function

NumeroIncorrecto: ImporteCapturaValido = False

End Function

Private Sub Text2_GotFocus()

    On Error Resume Next

    antiguo = ConNom1.col

    ConNom1.col = 1
    ConNom1.CellBackColor = &H80FF80

    ConNom1.col = antiguo

    ' Colocar el cursor al final del contenido.
    Text2.SelStart = Len(Text2.Text)

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '=====================================================
    ' BLOQUEAR OPERACIONES DEL EDITOR SOBRE SUMAS
    '
    ' Incluye Escape, que en el procedimiento
    ' original puede volver a escribir una celda.
    '=====================================================

    If EsFilaSumasCaptura() Then
        KeyCode = 0
        Exit Sub
    End If

    If Not AutorizarCambioNominaForm8() Then
        KeyCode = 0
        Exit Sub
    End If
      Select Case KeyCode
         Case vbKeyEscape
           Text2.Text = ""
           Text2.Text = valcelant
           ConNom1.Text = Format(valcelant, z1$)
           ConNom1.SetFocus
      End Select
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    Dim columna As Long
    Dim textoSeleccionado As String
    
    '=====================================================
    ' BLOQUEAR CONFIRMACION DE CAMBIOS EN SUMAS
    '=====================================================

    If EsFilaSumasCaptura() Then

        KeyAscii = 0
        Exit Sub

    End If
    
    '=====================================================
    ' IMPEDIR CONFIRMAR CAMBIOS EN SOLO CONSULTA
    '=====================================================

    If Not AutorizarCambioNominaForm8() Then

        KeyAscii = 0

        Exit Sub

    End If

    columna = ConNom1.col

    '=========================================================
    ' AGUINALDO - EVITAR EDICIONES NO PERMITIDAS
    '
    ' Esta validacion es necesaria porque Text2 puede
    ' recibir el foco independientemente del grid.
    '
    ' BANAMEX conserva su comportamiento anterior.
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Or g_TipoNominaActiva = tnPTU Or g_TipoNominaActiva = tnBonoPremio Then
        If Not EsColumnaNumericaCaptura(columna) And columna <> C_BANAMEX Then
            If KeyAscii = vbKeyReturn Then
                KeyAscii = 0

                MsgBox "Esta columna no se puede capturar en la modalidad especial seleccionada.", vbExclamation, "Captura de Nomina"
                Text2.Text = ConNom1.Text
                ConNom1.SetFocus
            Else
                KeyAscii = 0
            End If
            Exit Sub
        End If
    End If

    '=========================================================
    ' CONFIRMAR VALOR CON ENTER
    '=========================================================

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If EsColumnaNumericaCaptura(ConNom1.col) Then
            If Not ImporteCapturaValido(Text2.Text) Then
                MsgBox "Ingrese una cantidad numerica valida.", vbExclamation, "Captura"
                Exit Sub
            End If
        End If

        ConNom1.Text = Text2.Text

        'Aplicar los calculos originales.
        checar

        archiva_o = 1
        ConNom1.SetFocus

        Exit Sub
    End If

    '=========================================================
    ' COMPORTAMIENTO ORIGINAL DEL EDITOR NUMERICO
    '=========================================================

    If Not EsColumnaNumericaCaptura(columna) Then
        Exit Sub
    End If

    If KeyAscii = vbKeyBack Then Exit Sub

    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Exit Sub
    End If

    If KeyAscii = 46 Then
        textoSeleccionado = mID$(Text2.Text, Text2.SelStart + 1, Text2.SelLength)

        If InStr(1, Text2.Text, ".") = 0 Or InStr(1, textoSeleccionado, ".") > 0 Then
            Exit Sub
        End If
    End If
    KeyAscii = 0
End Sub



Private Sub Text2_LostFocus()
   antiguo = ConNom1.col
   ConNom1.col = 1: ConNom1.CellBackColor = Color_gris
   ConNom1.col = antiguo
   
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
    If Not AutorizarCambioNominaForm8() Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
        If g_TipoNominaActiva = tnPTU Then
            ' Exento calculado al capturar PTU en C_CONC6.
            ConNom1.SetFocus
            Exit Sub
        End If
        exento = Val(Text3.Text)
        Text3.FontBold = True
        Text3.Text = Format$(exento, z1$)
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

Private Sub cboMesNominaEspecial_Click()
    If Option2.Value = True Then
        ' Que MiFecha y meselegido sigan el mes elegido.
        If cboMesNominaEspecial.ListIndex < Combo1.ListCount Then
            Combo1.ListIndex = cboMesNominaEspecial.ListIndex
        End If
        ActualizarNombreEspecial
    End If
End Sub

Private Sub cboMesNominaEspecial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not Option2.Value Then Exit Sub
        If cboMesNominaEspecial.ListIndex < 0 Then
            MsgBox "Seleccione el mes de la nomina especial.", vbExclamation, "Nomina"
            Exit Sub
        End If

        'Aplicar el nombre con el mes seleccionado.
        ActualizarNombreEspecial

        'Ejecutar Captura -> Iniciar captura.
        IniciarCapturaDesdeTeclado
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    Dim i As Long
    If KeyAscii = 13 Then
        ' Al presionar Enter, enfocar el primer empleado visible
        For i = 1 To limite
            If ConNom1.RowHeight(i) > 0 Then
                ConNom1.Row = i
                ConNom1.col = 1
                ConNom1.SetFocus
                Exit For
            End If
        Next i
        KeyAscii = 0
    End If
End Sub

Private Sub Encab(ByVal c As Long, ByVal txt As String, ByVal ancho As Long)
    ConNom1.col = c
    ConNom1.CellAlignment = 4
    ConNom1.ColWidth(c) = ancho
    ConNom1.Text = txt
End Sub

' REVISION 2026: unico lugar que decide que controles se ven.
Private Sub ConfigurarModoNomina(ByVal especial As Boolean)
    Dim concepto As String
    Dim llevaMes As Boolean
    Dim llevaExento As Boolean

    concepto = UCase$(Trim$(cboTipoNominaEspecial.Text))
    llevaMes = especial And ConceptoLlevaMes(concepto)
    
    'AGUINALDO captura su exencion directamente en la columna Perc.exenta. del grid. PTU y las otras modalidades no se modifican aqui.

    llevaExento = especial And (concepto = "AGUINALDO" Or concepto = "PTU")
    If especial And (concepto = "AGUINALDO" Or concepto = "PTU") Then
        llevaExento = False
    End If

    ' --- NORMAL ---
    Frame2.Visible = Not especial
    Combo1.Visible = Not especial
    Combo1.Locked = especial

    ' --- ESPECIAL (siempre) ---
    lblTipoNominaEspecial.Visible = especial
    cboTipoNominaEspecial.Visible = especial
    lblNumeroNominaEspecial.Visible = especial
    txtPre.Visible = especial

    ' --- Mes: Combo1 en normal; cboMesNominaEspecial en FINIQUITO/LIQUIDACION ---
    lblMesNomina.Visible = (Not especial) Or llevaMes
    cboMesNominaEspecial.Visible = llevaMes

    ' --- Exento: solo AGUINALDO / PTU ---
    Label8.Visible = llevaExento
    Text3.Visible = llevaExento
End Sub

' Reaplica la visibilidad segun la opcion Normal/Especial seleccionada.
Private Sub RefrescarModoNomina()
    ConfigurarModoNomina (Option2.Value = True)
End Sub

' REVISION 2026: sufijo de mes (3 letras) solo para FINIQUITO y LIQUIDACION.
' PTU, AGUINALDO y BONO conservan el nombre original (numero + concepto).
Private Function SufijoMesEspecial(ByVal concepto As String) As String
    Dim mes As Integer

    If Not ConceptoLlevaMes(concepto) Then Exit Function

    mes = cboMesNominaEspecial.ListIndex + 1
    If mes < 1 Or mes > 12 Then mes = Month(Date)

    SufijoMesEspecial = Choose(mes, "ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")
End Function

Private Function SoloDigitos(ByVal texto As String) As String
    Dim i As Long
    Dim caracter As String
    Dim resultado As String

    resultado = ""

    For i = 1 To Len(texto)
        caracter = mID$(texto, i, 1)

        If caracter >= "0" And caracter <= "9" Then
            resultado = resultado & caracter
        End If
    Next i

    SoloDigitos = resultado
End Function

Private Function ConceptoLlevaMes(ByVal concepto As String) As Boolean
    Select Case UCase$(Trim$(concepto))
        Case "BONO", "FINIQUITO", "LIQUIDACION"
            ConceptoLlevaMes = True
    End Select
End Function

Private Sub ActualizarNombreEspecial()

    If ActualizandoNombre Then Exit Sub

    On Error GoTo Salida

    ActualizandoNombre = True

    Dim numero As String
    Dim concepto As String
    Dim nombreFinal As String

    numero = Trim$(txtPre.Text)
    concepto = Trim$(cboTipoNominaEspecial.Text)

    ' Construir el nombre de la nómina especial.
    nombreFinal = numero & concepto & SufijoMesEspecial(concepto)

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
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        '=================================================
        ' VALIDAR TIPO DE NOMINA ESPECIAL
        '=================================================

        If cboTipoNominaEspecial.ListIndex < 0 Then
            MsgBox "Seleccione un tipo de nomina especial.", vbExclamation, "Nomina"
            cboTipoNominaEspecial.SetFocus
            Exit Sub
        End If

        '=================================================
        ' VALIDAR NUMERO DE NOMINA
        '=================================================

        If Len(Trim$(txtPre.Text)) = 0 Then
            MsgBox "Ingrese el numero de nomina.", vbExclamation, "Nomina"
            txtPre.SetFocus
            Exit Sub
        End If

        'Actualizar el nombre antes de iniciar.
        ActualizarNombreEspecial

        '=================================================
        ' NOMINAS ESPECIALES QUE REQUIEREN MES
        '
        ' BONO
        ' FINIQUITO
        ' LIQUIDACION
        '
        ' AGUINALDO y PTU no requieren seleccionar mes.
        '=================================================

        If cboMesNominaEspecial.Visible Then
            cboMesNominaEspecial.SetFocus
        Else

            'Iniciar directamente desde el teclado.
            IniciarCapturaDesdeTeclado
            
        End If
        Exit Sub
    End If

    '=====================================================
    ' VALIDACION ORIGINAL: SOLAMENTE NUMEROS
    '=====================================================

    If KeyAscii = vbKeyBack Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Function EliminarTemporalesSesion() As Boolean
    On Error GoTo ErrorEliminar
    EliminarTemporalesSesion = False
    
    If Not SesionTemporalRegistrada Then
    MsgBox "No se puede confirmar el origen de los archivos." & vbCrLf & "No se eliminara ningun archivo.", vbExclamation, "Proteccion de nomina"
        Exit Function
    End If

    ' Nunca eliminar una nomina archivada
    If NominaArchivadaSesion Then
        EliminarTemporalesSesion = True
        Exit Function
    End If

    ' Nunca eliminar despues de un guardado fallido
    If GuardadoFallidoSesion Then
        MsgBox "Hubo un error durante el guardado." & vbCrLf & "Los archivos no se eliminaran.", vbExclamation, "Nomina"
        Exit Function
    End If

    ' Cerrar archivos de nomina y complemento
    Close #6
    Close #14

    ' Eliminar solo archivos nuevos registrados
    If NominaNuevaSesion Then
        If Len(RutaNominaSesion) > 0 Then
            If Len(Dir$(RutaNominaSesion)) > 0 Then
                Kill RutaNominaSesion
            End If
        End If
    End If

    If CmpNuevoSesion Then
        If Len(RutaCmpSesion) > 0 Then
            If Len(Dir$(RutaCmpSesion)) > 0 Then
                Kill RutaCmpSesion
            End If
        End If
    End If

    EliminarTemporalesSesion = True
    Exit Function

ErrorEliminar:

    MsgBox "No fue posible completar la limpieza de archivos." & vbCrLf & "Error " & Err.Number & ": " & Err.Description & vbCrLf & "Revise los archivos antes de continuar.", vbCritical, "Nomina"

End Function

Private Function ConfirmarSalidaNomina() As Boolean

    Dim numeroErrorCMP As Long
    Dim descripcionErrorCMP As String
    Dim responde As VbMsgBoxResult

    ConfirmarSalidaNomina = False

    '=========================================================
    ' 1. EVITAR CIERRES SIMULTANEOS
    '=========================================================

    If CerrandoForm8 Then Exit Function

    '=========================================================
    ' 2. NO EXISTE CAPTURA ACTIVA
    '
    ' No preguntar.
    '=========================================================

    If Not CapturaActivaSesion Then
        ConfirmarSalidaNomina = True
        Exit Function
    End If

    '=========================================================
    ' 3. NOMINA ABIERTA EN SOLO CONSULTA
    '
    ' No se permite modificar ni archivar.
    ' Cerrar directamente.
    '=========================================================

    If mSoloConsulta Then
        ConfirmarSalidaNomina = True
        Exit Function
    End If

    '=========================================================
    ' 4. NOMINA YA GUARDADA SIN CAMBIOS PENDIENTES
    '
    ' Se utiliza el estado confirmado por
    ' archnom_Click.
    '
    ' No volver a preguntar.
    ' No ejecutar nuevamente archnom_Click.
    ' No reescribir archivos NOM ni CMP.
    '=========================================================

    If NominaArchivadaSesion And archiva_o = 0 And Not GuardadoFallidoSesion And Not FalloRestauracionNom Then
        ConfirmarSalidaNomina = True
        Exit Function
    End If

    '=========================================================
    ' 5. EXISTEN CAMBIOS SIN GUARDAR O ERROR PREVIO
    '
    ' Preguntar:
    '
    ' SI:
    '   Intentar guardar y cerrar.
    '
    ' NO:
    '   Descartar cambios y cerrar.
    '
    ' CANCELAR:
    '   Permanecer en Form8.
    '=========================================================

    responde = MsgBox("Desea archivar la nomina antes de salir?" & vbCrLf & vbCrLf & "SI: Guardar y salir." & vbCrLf & "NO: Salir sin guardar cambios." & vbCrLf & "CANCELAR: Continuar capturando.", vbYesNoCancel + vbQuestion + vbDefaultButton3, "Salir de captura")

    Select Case responde

        '=====================================================
        ' SI - GUARDAR Y SALIR
        '=====================================================

        Case vbYes

            GuardandoDuranteCierre = True

            'Utilizar el guardado original.
            archnom_Click 1

            GuardandoDuranteCierre = False

            'No cerrar si el guardado fallo.
            If GuardadoFallidoSesion Then
                ConfirmarSalidaNomina = False
                Exit Function
            End If

            'Comprobar que realmente se archivo.
            If Not NominaArchivadaSesion Then
                MsgBox "No se confirmo el guardado de la nomina.", vbExclamation, "Nomina"
                Exit Function
            End If

            'Guardado confirmado.
            ConfirmarSalidaNomina = True

            Exit Function

        '=====================================================
        ' NO - SALIR SIN GUARDAR NUEVOS CAMBIOS
        '=====================================================

        Case vbNo

            '-------------------------------------------------
            ' A. ERROR GRAVE DE RESTAURACION DEL NOM
            '
            ' No intentar nuevas escrituras.
            ' Conservar archivos para revision.
            '
            ' Cerrar no equivale a confirmar que
            ' los archivos sean recuperables.
            '-------------------------------------------------

            If FalloRestauracionNom Then
                MsgBox "No fue posible verificar la recuperacion completa de la nomina." & vbCrLf & vbCrLf & "Se conservaran todos los archivos para su revision." & vbCrLf & "No se intentara otro guardado.", vbExclamation, "Archivos pendientes de revision"
                ConfirmarSalidaNomina = True
                Exit Function
            End If

            '-------------------------------------------------
            ' B. NOMINA GUARDADA ANTERIORMENTE
            '
            ' Aplica a:
            '
            ' - Nomina existente abierta en edicion.
            ' - Nomina ordinaria nueva ya archivada.
            '
            ' No reescribir el NOM.
            '
            ' Si el CMP tiene cambios posteriores,
            ' restaurar el ultimo respaldo confirmado.
            '-------------------------------------------------

            If EditandoNominaGuardada Or NominaArchivadaSesion Then
                If HayCopiaCmpEdicion Then
                    On Error GoTo ErrorRestaurarCmp

                    'Liberar el archivo antes de comparar
                    'y, en su caso, restaurar sus bytes.

                    Close #14

                    If Not ArchivoCoincideConMemoria(RutaCmpSesion, CopiaCmpEdicion) Then
                        RestaurarBytesArchivo RutaCmpSesion, CopiaCmpEdicion
                    End If
                    On Error GoTo 0

                ElseIf EditandoNominaGuardada Then

                    'No existe respaldo recuperable.
                    'No borrar ni sobrescribir archivos.

                    MsgBox "No existe una copia del complemento que permita descartar los cambios." & vbCrLf & vbCrLf & "Los archivos se conservaran para revision.", vbExclamation, "Proteccion de nomina"
                End If

                'La nomina ya estaba guardada.
                'No eliminar NOM ni CMP.

                ConfirmarSalidaNomina = True
                Exit Function
            End If

            '-------------------------------------------------
            ' C. CAPTURA NUEVA CON GUARDADO FALLIDO
            '
            ' Es importante NO eliminar archivos
            ' generados durante un intento fallido.
            '
            ' Permitir cerrar Form8.
            '-------------------------------------------------

            If GuardadoFallidoSesion Then
                MsgBox "La captura se cerrara sin realizar otro guardado." & vbCrLf & vbCrLf & "Los archivos existentes se conservaran para su revision.", vbInformation, "Cerrar captura"

                ConfirmarSalidaNomina = True
                Exit Function
            End If

            '-------------------------------------------------
            ' D. CAPTURA NUEVA NUNCA GUARDADA
            '
            ' Solamente borrar archivos temporales
            ' cuando su origen esta registrado.
            '-------------------------------------------------

            If Not SesionTemporalRegistrada Then

                'No se conoce el origen de los archivos.
                'Cerrar sin eliminarlos.

                MsgBox "No se pudo confirmar el origen de los archivos temporales." & vbCrLf & "Se conservaran para revision.", vbExclamation, "Cerrar captura"
                ConfirmarSalidaNomina = True
                Exit Function
            End If

            '-------------------------------------------------
            ' E. ELIMINAR SOLO TEMPORALES IDENTIFICADOS
            '
            ' No modifica nominas archivadas.
            '-------------------------------------------------

            If Not EliminarTemporalesSesion() Then
                ConfirmarSalidaNomina = False
                Exit Function
            End If

            SesionTemporalRegistrada = False
            NominaNuevaSesion = False
            CmpNuevoSesion = False

            'Salir sin guardar.
            ConfirmarSalidaNomina = True

            Exit Function

        '=====================================================
        ' CANCELAR - PERMANECER EN FORM8
        '=====================================================

        Case vbCancel
            ConfirmarSalidaNomina = False
            Exit Function
    End Select

    Exit Function

'=========================================================
' ERROR AL DESCARTAR CAMBIOS DEL CMP
'
' No eliminar ni sobrescribir archivos adicionales.
' Mantener Form8 abierto para evitar un descarte
' incompleto y no confirmado.
'=========================================================

ErrorRestaurarCmp:

    numeroErrorCMP = Err.Number
    descripcionErrorCMP = Err.Description

    On Error Resume Next

    Close #14

    'Intentar recuperar el acceso al complemento
    'solamente si el archivo original existe.

    If Len(RutaCmpSesion) > 0 Then
        If Len(Dir$(RutaCmpSesion, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) > 0 Then
            Open RutaCmpSesion For Random As #14 Len = Len(nom_com)
        End If
    End If

    On Error GoTo 0

    MsgBox "No fue posible descartar completamente los cambios de esta nomina." & vbCrLf & vbCrLf & "El formulario permanecera abierto." & vbCrLf & "No se eliminara ningun archivo." & vbCrLf & vbCrLf & _
        "Error " & CStr(numeroErrorCMP) & ": " & descripcionErrorCMP, vbCritical, "Restauracion de nomina"
    ConfirmarSalidaNomina = False
End Function

Private Sub BloquearSeleccionNomina(ByVal bloquear As Boolean)

    Option1.Enabled = Not bloquear
    Option2.Enabled = Not bloquear
    Option3.Enabled = Not bloquear
    Option4.Enabled = Not bloquear
    
    cboMesNominaEspecial.Enabled = Not bloquear

    Combo1.Enabled = Not bloquear
    cboTipoNominaEspecial.Enabled = Not bloquear
    txtPre.Enabled = Not bloquear

End Sub

Private Sub ReiniciarEstadoSesion()

    CapturaActivaSesion = False
    SesionTemporalRegistrada = False

    NominaNuevaSesion = False
    CmpNuevoSesion = False

    NominaArchivadaSesion = False
    GuardadoFallidoSesion = False

    RutaNominaSesion = ""
    RutaCmpSesion = ""

    archiva_o = 0

    BloquearSeleccionNomina False
    
    'Liberar los respaldos mantenidos en memoria.
    Erase CopiaCmpEdicion
    
    HayCopiaCmpEdicion = False
    EditandoNominaGuardada = False
    BloqueoPorPlazoEdicion = False
    FalloRestauracionNom = False
    
    GuardandoDuranteCierre = False

    '=====================================================
    ' RESTABLECER ESTADO DE CONSULTA
    '
    ' Una nueva instancia de Form8 no debe heredar
    ' permisos de la nomina anterior.
    '=====================================================

    mSoloConsulta = False
    mDetalleConsulta = ""

End Sub

Private Sub FinalizarSalidaNomina()

    On Error Resume Next

    Close #6
    Close #14
    Close #95

    Close #2
    Close #3
    Close #4
    Close #5
    Close #8
    Close #12
    
    ConNom1.Clear

    ReiniciarEstadoSesion

    On Error GoTo 0

End Sub

Private Sub checarSeparacionNueva()
    Dim fila As Long
    Dim valor As Currency
    Dim maximoRelacionable As Currency
    Dim respuestaLocal As VbMsgBoxResult

    ' La cuenta bancaria es TEXTO; no debe pasar por datoper/CCur.
    If ConNom1.col = C_BANAMEX Then
        fila = ConNom1.Row
        regtro = Val(ConNom1.TextMatrix(fila, C_NUM))
        If regtro > 0 Then
            Clbnx.Q1 = Left$(Trim$(ConNom1.Text), 16)
            Put #12, regtro, Clbnx
        End If
        Exit Sub
    End If

    datoper
    If regtro <= 0 Then Exit Sub

    fila = ConNom1.Row
    valor = CCur(dato_ent)

    Select Case ruta
        
        '=====================================================
        ' DIAS DEL PERIODO
        '
        ' FINIQUITO NUEVO:
        ' Los dias son informativos para el periodo.
        ' No generan automaticamente una percepcion
        ' en C_SUELDO, que esta reservada.
        '
        ' RESTO DE MODALIDADES:
        ' Conservar la operacion original.
        '=====================================================

        
        '=====================================================
        ' DIAS DEL PERIODO - SEPARACION NUEVA
        '
        ' LIQUIDACION:
        ' Los dias no deben generar automaticamente sueldo.
        ' C_SUELDO es una columna auxiliar reservada.
        '
        ' FINIQUITO:
        ' Conservar la correccion previamente aplicada.
        '=====================================================

        Case C_DIAS
            If Not g_EsFiniquito And EsSeparacionNueva() Then

                '---------------------------------------------
                ' LIQUIDACION NUEVA
                '---------------------------------------------

                If valor < 0 Then
                    MsgBox "Los dias no pueden ser negativos.", vbExclamation, "Liquidacion"
                    ConNom1.TextMatrix(fila, C_DIAS) = ""
                ElseIf valor = 0 Then
                    ConNom1.TextMatrix(fila, C_DIAS) = ""
                Else
                    ConNom1.TextMatrix(fila, C_DIAS) = Format$(valor, "##0.00")
                End If

                'No crear una percepcion salarial adicional.
                ConNom1.TextMatrix(fila, C_SUELDO) = ""

            ElseIf g_EsFiniquito And EsSeparacionNueva() Then

                '---------------------------------------------
                ' FINIQUITO NUEVO
                ' Mantener el comportamiento ya implementado.
                '---------------------------------------------

                If valor < 0 Then
                    MsgBox "Los dias no pueden ser negativos.", vbExclamation, "Finiquito"
                    ConNom1.TextMatrix(fila, C_DIAS) = ""
                ElseIf valor = 0 Then
                    ConNom1.TextMatrix(fila, C_DIAS) = ""
                Else
                    ConNom1.TextMatrix(fila, C_DIAS) = Format$(valor, "##0.00")
                End If
                ConNom1.TextMatrix(fila, C_SUELDO) = ""
            Else

                '---------------------------------------------
                ' COMPORTAMIENTO ORIGINAL
                '---------------------------------------------

                If valor < 0 Then
                    MsgBox "Los dias no pueden ser negativos.", vbExclamation
                    ConNom1.TextMatrix(fila, C_DIAS) = ""
                    ConNom1.TextMatrix(fila, C_SUELDO) = ""
                Else
                    If valor = 0 Then
                        ConNom1.TextMatrix(fila, C_DIAS) = ""
                        ConNom1.TextMatrix(fila, C_SUELDO) = ""
                    Else
                        ConNom1.TextMatrix(fila, C_DIAS) = Format$(valor, "##0.00")
                        PonCel fila, C_SUELDO, CCur(valor * personal.ingr)
                    End If
                End If
            End If
            
        
        '=====================================================
        ' COLUMNA SUELDO
        '
        ' LIQUIDACION NUEVA:
        ' C_SUELDO se encuentra reservada.
        '
        ' OTRAS RUTAS:
        ' Conservar el funcionamiento original.
        '=====================================================

        Case C_SUELDO
            If Not g_EsFiniquito And EsSeparacionNueva() Then
                MsgBox "La columna Sueldo no se captura en una Liquidacion nueva." & vbCrLf & "Capture los importes en los conceptos correspondientes.", vbInformation, "Liquidacion"
                ConNom1.TextMatrix(fila, C_SUELDO) = ""
            Else
                MsgBox "El sueldo se modifica editando su archivo de personal", vbInformation, "Captura de Nomina"
                PonCel fila, C_SUELDO, CCur(CelCur(fila, C_DIAS) * personal.ingr)
            End If

        Case C_ISPT, C_SUBEMP
            respuestaLocal = MsgBox("Si modifica esta celda no se recalculara " & _
                          "automaticamente la retencion de este empleado. " & _
                          "Desea hacerlo?", vbYesNo + vbExclamation, "ISR manual")
            If respuestaLocal = vbYes Then
                ConNom1.TextMatrix(fila, C_FLAG_ISR) = "1"
                If ruta = C_ISPT Then
                    ConNom1.TextMatrix(fila, C_SUBEMP) = ""
                Else
                    ConNom1.TextMatrix(fila, C_ISPT) = ""
                End If
                If valor = 0 Then
                    ConNom1.TextMatrix(fila, ruta) = ""
                Else
                    ConNom1.TextMatrix(fila, ruta) = Format$(valor, z1$)
                End If
            Else
                ConNom1.TextMatrix(fila, ruta) = Format$(valcelant, z1$)
            End If

        
        Case C_IMSS
            'FINIQUITO y LIQUIDACION: IMSS es una deduccion de captura manual. No solicitar confirmacion.
            ConNom1.TextMatrix(fila, C_FLAG_IMSS) = "1"
            If valor = 0 Then
                ConNom1.TextMatrix(fila, C_IMSS) = ""
            Else
                ConNom1.TextMatrix(fila, C_IMSS) = Format$(valor, z1$)
            End If

        
        '=====================================================
        ' PERCEPCION EXENTA
        '=====================================================

        Case C_EXENTA
            If g_EsFiniquito And EsSeparacionNueva() Then

                '---------------------------------------------
                ' FINIQUITO NUEVO
                '
                ' La exenta es parte de los conceptos
                ' capturados. No es ingreso adicional.
                '
                ' Ante un dato invalido, restaurar el valor
                ' anterior en lugar de borrarlo.
                '---------------------------------------------

                maximoRelacionable = MaximoExentoFiniquitoNuevo(fila)

                If valor < 0 Or valor > maximoRelacionable Then
                    MsgBox "La percepcion exenta no puede ser negativa ni superar la suma de los conceptos potencialmente exentos del finiquito." & vbCrLf & "Revise Aguinaldo, Prima Vacacional y Prima de Antiguedad.", vbExclamation, "Finiquito"

                    If IsNumeric(valcelant) Then
                        ConNom1.TextMatrix(fila, C_EXENTA) = Format$(CCur(valcelant), z1$)

                    Else
                        ConNom1.TextMatrix(fila, C_EXENTA) = ""

                    End If
                ElseIf valor = 0 Then
                    ConNom1.TextMatrix(fila, C_EXENTA) = ""

                Else
                    ConNom1.TextMatrix(fila, C_EXENTA) = Format$(valor, z1$)

                End If
                
            
            ElseIf Not g_EsFiniquito And EsSeparacionNueva() Then

                '=============================================
                ' LIQUIDACION NUEVA - PERCEPCION EXENTA
                '
                ' C_EXENTA representa una parte de los
                ' conceptos brutos de la liquidacion.
                '
                ' NO es un ingreso adicional.
                ' NO se distribuye automaticamente.
                '=============================================

                maximoRelacionable = MaximoExentoLiquidacionNueva(fila)

                '---------------------------------------------
                ' 1. VALIDAR IMPORTE
                '---------------------------------------------

                If valor < 0 Or _
                   valor > maximoRelacionable Then
                    MsgBox "La percepcion exenta no puede ser negativa ni superar los conceptos potencialmente exentos de la Liquidacion." & vbCrLf & "Revise Aguinaldo, Prima Vacacional, Prima de Antiguedad e Indemnizacion.", vbExclamation, "Liquidacion"

                    '-----------------------------------------
                    ' 2. RECUPERAR EL IMPORTE ANTERIOR
                    '
                    ' No borrar automaticamente una
                    ' percepcion exenta ya capturada.
                    '-----------------------------------------

                    If IsNumeric(valcelant) Then
                        ConNom1.TextMatrix(fila, C_EXENTA) = Format$(CCur(valcelant), z1$)
                    Else
                        ConNom1.TextMatrix(fila, C_EXENTA) = ""
                    End If
                ElseIf valor = 0 Then
                    ConNom1.TextMatrix(fila, C_EXENTA) = ""
                Else

                    '-----------------------------------------
                    ' 3. CONSERVAR LA EXENCION CAPTURADA
                    '-----------------------------------------

                    ConNom1.TextMatrix(fila, C_EXENTA) = Format$(valor, z1$)
                End If
            Else

                '---------------------------------------------
                ' LIQUIDACION Y OTRAS RUTAS
                '
                ' Comportamiento original sin cambios.
                '---------------------------------------------

                maximoRelacionable = 0

                If C_LQ_AGUI >= 0 Then
                    maximoRelacionable = maximoRelacionable + CelCur(fila, C_LQ_AGUI)
                End If

                If C_LQ_PRIMAVAC >= 0 Then
                    maximoRelacionable = maximoRelacionable + CelCur(fila, C_LQ_PRIMAVAC)
                End If

                If C_LQ_ANTIG >= 0 Then
                    maximoRelacionable = maximoRelacionable + CelCur(fila, C_LQ_ANTIG)
                End If

                If C_LQ_INDEMN >= 0 Then
                    maximoRelacionable = maximoRelacionable + CelCur(fila, C_LQ_INDEMN)
                End If

                If valor < 0 Or valor > maximoRelacionable Then
                    MsgBox "La percepcion exenta no puede ser negativa ni mayor que la suma de conceptos potencialmente exentos." & vbCrLf & "Revise el desglose fiscal antes de guardar.", vbExclamation, "Percepcion exenta"
                    ConNom1.TextMatrix(fila, C_EXENTA) = ""
                ElseIf valor = 0 Then
                    ConNom1.TextMatrix(fila, C_EXENTA) = ""
                Else
                    ConNom1.TextMatrix(fila, C_EXENTA) = Format$(valor, z1$)
                End If
            End If

        Case C_TOTING, C_TOTDED, C_NETO, C_FLAG_ISR, C_FLAG_IMSS, C_FLAG_ESP
            MsgBox "Esta columna es calculada o reservada.", vbInformation
            ConNom1.TextMatrix(fila, ruta) = ""
        
        Case Else
            '=================================================
            ' CONCEPTOS Y DEDUCCIONES
            '
            ' Mantener la validacion original.
            '=================================================

            If valor < 0 Then
                MsgBox "No se permiten cantidades negativas en este concepto.", vbExclamation, "Captura"
                ConNom1.TextMatrix(fila, ruta) = ""
            ElseIf valor = 0 Then
                ConNom1.TextMatrix(fila, ruta) = ""
            Else
                ConNom1.TextMatrix(fila, ruta) = Format$(valor, z1$)
            End If

            '=================================================
            ' PROTECCION EXCLUSIVA DE FINIQUITO NUEVO
            '
            ' Si se reduce un concepto por debajo de la
            ' exencion previamente capturada, recuperar
            ' el importe anterior de ese concepto.
            '
            ' Nunca borrar automaticamente C_EXENTA.
            '=================================================

            If g_EsFiniquito And EsSeparacionNueva() Then
                If ruta = C_LQ_AGUI Or ruta = C_LQ_PRIMAVAC Or ruta = C_LQ_ANTIG Then
                    maximoRelacionable = MaximoExentoFiniquitoNuevo(fila)
                    If CelCur(fila, C_EXENTA) > maximoRelacionable Then
                        MsgBox "No es posible reducir este concepto porque la percepcion exenta capturada supera los conceptos potencialmente exentos." & vbCrLf & "Modifique primero Perc.exenta.", vbExclamation, "Finiquito"
                        If IsNumeric(valcelant) Then
                            ConNom1.TextMatrix(fila, ruta) = Format$(CCur(valcelant), z1$)
                        Else
                            ConNom1.TextMatrix(fila, ruta) = ""
                        End If
                    End If
                End If
            End If
            
            
            '=================================================
            ' LIQUIDACION NUEVA
            '
            ' Si se reduce una percepcion que respalda
            ' el exento, comprobar que la nueva cantidad
            ' siga siendo suficiente.
            '
            ' No modificar la percepcion exenta.
            ' No intervenir en Finiquito.
            '=================================================

            If Not g_EsFiniquito And EsSeparacionNueva() Then
                Select Case ruta
                    Case C_LQ_AGUI, C_LQ_PRIMAVAC, C_LQ_ANTIG, C_LQ_INDEMN
                        maximoRelacionable = MaximoExentoLiquidacionNueva(fila)

                        If CelCur(fila, C_EXENTA) > maximoRelacionable Then
                            MsgBox "No es posible reducir este concepto porque la percepcion exenta quedaria por encima de los conceptos que la respaldan." & vbCrLf & "Modifique primero Perc.exenta.", vbExclamation, "Liquidacion"

                            '---------------------------------
                            ' RESTAURAR LA PERCEPCION ANTERIOR
                            '
                            ' No borrar C_EXENTA.
                            '---------------------------------

                            If IsNumeric(valcelant) Then
                                ConNom1.TextMatrix(fila, ruta) = Format$(CCur(valcelant), z1$)
                            Else
                                ConNom1.TextMatrix(fila, ruta) = ""
                            End If
                        End If
                End Select
            End If
    End Select

    li = fila
    si_impto = Val(ConNom1.TextMatrix(fila, C_FLAG_ISR))
    si_imss = Val(ConNom1.TextMatrix(fila, C_FLAG_IMSS))
    ' En especiales conservar la politica de IMSS existente.
    If ConNom1.TextMatrix(fila, C_FLAG_ESP) = "1" Then si_imss = 1
    sumah si_impto, si_imss
    ConNom1.col = ruta
End Sub

Private Sub sumahSeparacionNueva(ByVal bloquearISR As Integer, ByVal bloquearIMSS As Integer)
    Dim colAnterior As Long
    Dim columna As Long
    Dim bruto As Currency
    Dim exentoFila As Currency
    Dim gravado As Currency
    Dim deduccionFila As Currency
    Dim normalBruto As Currency

    colAnterior = ConNom1.col
    bruto = 0
    deduccionFila = 0
    rgtro = Val(ConNom1.TextMatrix(li, C_NUM))
    If rgtro <= 0 Then Exit Sub

    For columna = C_PERC_INI To C_PERC_FIN
        bruto = bruto + CelCur(li, columna)
    Next columna
    
    '=========================================================
    ' VALIDACION DE PERCEPCION EXENTA
    '
    ' FINIQUITO NUEVO:
    ' No borrar el importe exento al recalcular.
    '
    ' LIQUIDACION Y OTRAS RUTAS:
    ' Conservar el comportamiento anterior.
    '=========================================================

    
    '=========================================================
    ' VALIDACION DEL EXENTO
    '
    ' LIQUIDACION NUEVA:
    ' La exenta es parte de las percepciones brutas.
    ' Nunca se suma como ingreso adicional.
    '
    ' Si el desglose es invalido, no borrar C_EXENTA.
    '
    ' FINIQUITO NUEVO:
    ' Conservar la proteccion que ya implementamos.
    '=========================================================

    exentoFila = CelCur(li, C_EXENTA)

    If exentoFila < 0 Or exentoFila > bruto Then
        If g_EsFiniquito And EsSeparacionNueva() Then
        
            '---------------------------------------------
            ' FINIQUITO - COMPORTAMIENTO YA CORREGIDO
            '---------------------------------------------

            MsgBox "La percepcion exenta del finiquito no coincide con las percepciones capturadas." & vbCrLf & "Revise los importes antes de continuar.", vbExclamation, "Finiquito"

            ConNom1.col = colAnterior
            Exit Sub

        ElseIf Not g_EsFiniquito And EsSeparacionNueva() Then

            '---------------------------------------------
            ' LIQUIDACION NUEVA
            '---------------------------------------------

            MsgBox "La percepcion exenta de la Liquidacion no coincide con las percepciones capturadas." & vbCrLf & "El importe exento no sera eliminado.", vbExclamation, "Liquidacion"

            ConNom1.col = colAnterior
            Exit Sub
        Else

            '---------------------------------------------
            ' COMPORTAMIENTO ORIGINAL
            '---------------------------------------------

            MsgBox "Perc.exenta no puede exceder las percepciones.", vbExclamation

            exentoFila = 0
            
            ConNom1.TextMatrix(li, C_EXENTA) = ""
        End If
    End If

    '=========================================================
    ' COMPROBAR EL LIMITE ESTRUCTURAL DE LIQUIDACION
    '
    ' No se esta calculando ninguna exencion fiscal.
    '=========================================================

    If Not g_EsFiniquito And EsSeparacionNueva() Then
        If exentoFila > MaximoExentoLiquidacionNueva(li) Then
            MsgBox "El exento de la Liquidacion supera los conceptos potencialmente exentos." & vbCrLf & "Revise la captura antes de continuar.", vbExclamation, "Liquidacion"
            ConNom1.col = colAnterior
            Exit Sub
        End If
    End If

    '=========================================================
    ' DETERMINAR LA PARTE GRAVADA
    '
    ' NO modificar el total bruto.
    '=========================================================

    gravado = bruto - exentoFila

    ' Las percepciones ya contienen la parte exenta: NO agregarla.
    PonCel li, C_TOTING, bruto

    Acum_Doble
    If bloquearISR = 0 Then
        impto = 0
        If gravado > 0 Then
            If cal_anual = 0 Then
                calculoRetencion gravado, impto, 1, regtro
            Else
                aoalta = Val(mID$(personal.fal, 7, 4))
                If aoalta = empresa.ao And Tot_dias < 330 Then
                    calculoRetencion gravado, impto, 1, regtro
                Else
                    calc_anual gravado, impto, 1
                End If
            End If
        End If
        If impto < 0 Then impto = 0
        PonCel li, C_ISPT, impto
    End If

    If bloquearIMSS = 0 Then
        diaseg = CelCur(li, C_DIAS)
        p_vacacional = CelCur(li, C_LQ_PRIMAVAC)
        integrado = personal.integrado
        seguro = 0
        imss integrado, seguro, diaseg
        PonCel li, C_IMSS, seguro
    End If

    If ConNom1.TextMatrix(li, C_FLAG_ESP) = "1" Then
        ' Conservar logica original de retencion de nomina especial,
        ' usando SOLO ingreso gravado, no el bruto.
        normalBruto = (gravado / 365 * 30.4)
        If normalBruto > 0 And _
           ConNom1.TextMatrix(li, C_FLAG_ISR) = "0" Then
            Get #2, regtro, personal
            Normal = personal.ingr * 30
            DifImptos.CalcDoble = 0
            calculo_compl2 Normal, impto, 1
            impto_1 = impto
            normal_2 = normalBruto + Normal
            impto = 0
            calculo_compl2 normal_2, impto, 1
            impto_2 = impto
            If impto_1 = 0 And impto_2 > 0 Then
                DifImptos.CalcDoble = 1
                calculo_compl2 Normal, impto, 1
                DifImptos.CalcDoble = 0
                impto_1 = impto + creere
            End If
            impto_3 = impto_2 - impto_1
            porc_apli = CCur(impto_3 / normalBruto)
            impto = gravado * porc_apli
            If impto < 0 Then impto = 0
            nom_com.ImpTot = impto
            nom_com.subapl = 0
            nom_com.subNap = 0
            Put #14, regtro, nom_com
            PonCel li, C_ISPT, impto
        End If
            ConNom1.TextMatrix(li, C_SUBEMP) = ""
            
            'No borrar el IMSS que fue capturado manualmente.
            If Val(ConNom1.TextMatrix(li, C_FLAG_IMSS)) = 0 Then
                ConNom1.TextMatrix(li, C_IMSS) = ""
            End If
    End If

    For columna = C_DED_INI To C_DED_FIN
        deduccionFila = deduccionFila + CelCur(li, columna)
    Next columna
    PonCel li, C_TOTDED, deduccionFila
    PonCel li, C_NETO, bruto - deduccionFila
    sum(1) = bruto
    sum(2) = deduccionFila
    sum(3) = bruto - deduccionFila
    ConNom1.col = colAnterior
End Sub

Private Function CelCur(ByVal f As Long, ByVal c As Long) As Currency
    Dim s As String
    s = Trim$(ConNom1.TextMatrix(f, c))
    If Len(s) > 0 Then
        If IsNumeric(s) Then
            CelCur = CCur(s)
        End If
    End If
End Function

Private Sub Text2_Change()

    '=====================================================
    ' NO REFLEJAR CAMBIOS SOBRE LA FILA SUMAS
    '
    ' La fila solamente puede actualizarse desde
    ' los calculos originales de sumavert().
    '=====================================================

    If EsFilaSumasCaptura() Then
        Exit Sub
    End If

    '=====================================================
    ' IMPEDIR MODIFICACIONES DE CELDAS EN CONSULTA
    '=====================================================

    If Not AutorizarCambioNominaForm8() Then
        Exit Sub
    End If

    '=========================================================
    ' VALIDACION DE CAMBIO EN TIEMPO REAL
    '=========================================================

    If ConNom1.Row <= 0 Then Exit Sub

    '=========================================================
    ' AGUINALDO - NO MODIFICAR COLUMNAS PROTEGIDAS
    '
    ' Se permite mantener la edicion bancaria existente.
    '=========================================================

    If g_TipoNominaActiva = tnAguinaldo Or g_TipoNominaActiva = tnPTU Or g_TipoNominaActiva = tnBonoPremio Then
        If Not EsColumnaNumericaCaptura(ConNom1.col) And ConNom1.col <> C_BANAMEX Then
            Exit Sub
        End If
    End If

    '=========================================================
    ' VALIDACION NUMERICA ORIGINAL
    '=========================================================

    If EsColumnaNumericaCaptura(ConNom1.col) Then
        'Permitir escribir el punto decimal antes
        'de completar la cantidad.
        If Text2.Text <> "." Then
            If Not ImporteCapturaValido(Text2.Text) Then
                Exit Sub
            End If
        End If
    End If
    ConNom1.Text = Text2.Text
End Sub

Public Sub PrepararCapturaNuevaDesdeForm1()
    RegresoModalDesdeVisor = False
    RutaEdicionPendienteVisor = ""
    PreparandoEdicionDesdeVisor = False
End Sub

Public Function ProgramarEdicionDesdeVisor(ByVal rutaNom As String) As Boolean

    ProgramarEdicionDesdeVisor = False

    If CapturaActivaSesion Then
        MsgBox "Ya existe una nomina en captura." & vbCrLf & "Finalice esa captura antes de editar otra.", vbExclamation, "Modificar nomina"
        Exit Function
    End If

    RegresoModalDesdeVisor = True
    RutaEdicionPendienteVisor = rutaNom
    ProgramarEdicionDesdeVisor = True

End Function

Public Sub AbrirNominaGuardadaParaEditar(ByVal rutaNom As String)

    Dim rutaBase As String
    Dim rutaCmp As String
    Dim detalle As String

    Dim r As Long
    Dim yavas As Integer

    Dim errorNum As Long
    Dim errorTexto As String

    Dim rutaActiva As String

    Dim totalMaestro As Long
    Dim totalBnx As Long
    Dim totalCmp As Long

    Dim soloConsultaApertura As Boolean

    On Error GoTo ErrorAbrir

    If CapturaActivaSesion Then
        Err.Raise vbObjectError + 2901, "AbrirNominaGuardadaParaEditar", "Ya existe una nomina abierta en Form8."
    End If

    soloConsultaApertura = Not PuedeModificarNomina5Dias(rutaNom, detalle)
    
    If InStrRev(rutaNom, "\") = 0 Then
        Err.Raise vbObjectError + 2903, "AbrirNominaGuardadaParaEditar", "La ruta de nomina no es absoluta."
    End If
    If UCase$(Right$(rutaNom, 4)) <> ".NOM" Then
        Err.Raise vbObjectError + 2906, "AbrirNominaGuardadaParaEditar", "El archivo debe tener extension NOM."
    End If

    rutaBase = Left$(rutaNom, InStrRev(rutaNom, "\"))
    rutaCmp = Left$(rutaNom, Len(rutaNom) - 4) & ".cmp"
    
    rutaActiva = CurDir$

    If Right$(rutaActiva, 1) <> "\" Then
        rutaActiva = rutaActiva & "\"
    End If

    If StrComp(rutaActiva, rutaBase, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 2904, "AbrirNominaGuardadaParaEditar", "El directorio activo no coincide con el de la nomina seleccionada."
    End If
    If Len(Dir$(rutaNom)) = 0 Then
        Err.Raise vbObjectError + 2909, "AbrirNominaGuardadaParaEditar", "No existe el archivo NOM."
    End If
    If Len(Dir$(rutaCmp)) = 0 Then
        Err.Raise vbObjectError + 2905, "AbrirNominaGuardadaParaEditar", "Falta el complemento CMP."
    End If

    If FileLen(rutaNom) <= 0 Then
        Err.Raise vbObjectError + 2910, "AbrirNominaGuardadaParaEditar", "El archivo NOM esta vacio."
    End If

    If FileLen(rutaNom) Mod Len(nomina) <> 0 Then
        Err.Raise vbObjectError + 2911, "AbrirNominaGuardadaParaEditar", "El archivo NOM contiene registros incompletos."
    End If

    HayCopiaCmpEdicion = False
    Erase CopiaCmpEdicion
    FalloRestauracionNom = False

    If Not soloConsultaApertura Then
        LeerBytesArchivo rutaCmp, CopiaCmpEdicion
        HayCopiaCmpEdicion = True
    End If
    
    Close #2
    Close #8
    Close #6
    Close #12
    Close #14

    If soloConsultaApertura Then
        Open rutaBase & "personal.dno" For Random Access Read As #2 Len = Len(personal)
        Open rutaBase & "maestro.dno" For Random Access Read As #8 Len = Len(maestro)
        Open rutaBase & "bnxcla.dno" For Random Access Read As #12 Len = Len(Clbnx)
        Open rutaNom For Random Access Read As #6 Len = Len(nomina)
        Open rutaCmp For Random Access Read As #14 Len = Len(nom_com)

    Else
        Open rutaBase & "personal.dno" For Random As #2 Len = Len(personal)
        Open rutaBase & "maestro.dno" For Random As #8 Len = Len(maestro)
        Open rutaBase & "bnxcla.dno" For Random As #12 Len = Len(Clbnx)
        Open rutaNom For Random As #6 Len = Len(nomina)
        Open rutaCmp For Random As #14 Len = Len(nom_com)
    End If

    Dm = LOF(2) \ Len(personal)
    nm = LOF(6) \ Len(nomina)

    If nm <= 0 Or nm > Dm Then
        Err.Raise vbObjectError + 2907, "AbrirNominaGuardadaParaEditar", "Los registros NOM no corresponden con personal.dno."
    End If

    totalMaestro = LOF(8) \ Len(maestro)
    totalBnx = LOF(12) \ Len(Clbnx)
    totalCmp = LOF(14) \ Len(nom_com)

    AplicarTipoNomina rutaNom

    For r = 1 To nm
        Get #6, r, nomina
        yavas = 0
        verifica yavas

        If yavas > 0 Then
            If r > totalMaestro Or r > totalBnx Or r > totalCmp Then
                Err.Raise vbObjectError + 2908, "AbrirNominaGuardadaParaEditar", "Faltan registros auxiliares para el empleado ID " & CStr(r) & "." & vbCrLf & "MAESTRO: " & CStr(totalMaestro) & " / BNXCLA: " & CStr(totalBnx) & " / CMP: " & CStr(totalCmp)
            End If
        End If
    Next r

    '=========================================================
    ' 11. RECUPERAR EL CONTEXTO DE LA NOMINA
    '
    ' No generar nomina nueva.
    ' No ejecutar FuncionImpuesto.
    ' No ejecutar sumah.
    '=========================================================

    QUIN = 0

    If Option1.Value = True Then
        ReferOper.mes = Combo1.ListIndex + 1

        If Option3.Value = True Then
            QUIN = 1
            ReferOper.dia = 15
        ElseIf Option4.Value = True Then
            QUIN = 2
            ReferOper.dia = dd(ReferOper.mes)
        End If
    End If

    If Option2.Value = True Then
        N_ormal = 1
    Else
        N_ormal = 0
    End If

    '=========================================================
    ' 12. RECONSTRUIR LA CUADRICULA
    '
    ' Se utiliza el carganom original.
    ' No modificar sus formulas ni campos.
    '=========================================================

    ConNom1.Clear
    define
    ConNom1.Rows = nm + 2
    limite = 0
    
    For r = 1 To nm
        Get #6, r, nomina
        yavas = 0
        verifica yavas

        If yavas > 0 Then
            Get #2, r, personal
            Get #8, r, maestro
            Get #14, r, nom_com
            
            limite = limite + 1

            ConNom1.Row = limite

            ConNom1.col = C_NUM

            regtro = r
            rgtro = r

            ConNom1.Text = Format$(r, "#####")

            carganom
        End If
    Next r

    '=========================================================
    ' 13. MOSTRAR TOTALES ALMACENADOS
    '
    ' sumavert:
    ' Mantener la suma vertical de las celdas.
    '
    ' No recalcular impuestos ni percepciones.
    '=========================================================

    ConNom1.Rows = limite + 2

    sumavert

    If limite > 0 Then
        ConNom1.Row = 1
        ConNom1.col = C_DIAS
    End If

    '=========================================================
    ' 14. IDENTIFICAR LA SESION ABIERTA
    '
    ' La nomina ya existe en ambos modos.
    '
    ' Nunca debe considerarse temporal o nueva.
    '=========================================================
    
    RutaNominaSesion = rutaNom
    RutaCmpSesion = rutaCmp
    NominaNuevaSesion = False
    CmpNuevoSesion = False
    NominaArchivadaSesion = True
    GuardadoFallidoSesion = False
    SesionTemporalRegistrada = False
    CapturaActivaSesion = True

    '=========================================================
    ' 15. CONFIGURAR EL PERMISO DE MODIFICACION
    '
    ' CONSULTA:
    ' EditandoNominaGuardada = False.
    '
    ' EDICION:
    ' EditandoNominaGuardada = True.
    '=========================================================

    mSoloConsulta = soloConsultaApertura
    mDetalleConsulta = detalle
    EditandoNominaGuardada = Not soloConsultaApertura
    BloqueoPorPlazoEdicion = soloConsultaApertura
    archiva_o = 0

    '=========================================================
    ' 16. BLOQUEAR SELECCION DE OTRA NOMINA
    '=========================================================

    BloquearSeleccionNomina True

    '=========================================================
    ' 17. APLICAR PERMISOS DE CONSULTA
    '=========================================================

    If mSoloConsulta Then
        AplicarModoConsultaForm8
    Else
        Me.Caption = "Nomina - EDICION"
    End If

    '=========================================================
    ' 18. MOSTRAR LA CUADRICULA
    '=========================================================

    If Me.Visible And ConNom1.Visible Then
        ConNom1.SetFocus
    End If

    '=========================================================
    ' 19. INFORMAR CUANDO NO SE PERMITE EDITAR
    '=========================================================

    
    '=========================================================
    ' MENSAJE DE APERTURA DE NOMINA
    '
    ' Informar solamente el permiso concedido.
    '
    ' No mostrar:
    ' - Rutas de archivos.
    ' - Registros internos.
    ' - Controles utilizados para validar fechas.
    '=========================================================

    If mSoloConsulta Then
        MsgBox "Esta nomina se encuentra fuera del periodo permitido para modificaciones." & vbCrLf & vbCrLf & "Se ha abierto en modo SOLO CONSULTA." & vbCrLf & "Puede visualizar su informacion, pero no modificar sus importes.", _
            vbInformation, "Nomina - Solo consulta"
    Else
        MsgBox "Esta nomina se encuentra dentro del plazo de cinco dias habiles para realizar modificaciones." & vbCrLf & vbCrLf & "La EDICION esta permitida.", vbInformation, "Nomina - Edicion permitida"
    End If
    Exit Sub

ErrorAbrir:

    errorNum = Err.Number

    errorTexto = Err.Description

    On Error Resume Next

    Close #6
    Close #14
    Close #12
    Close #8
    Close #2

    mSoloConsulta = False
    mDetalleConsulta = ""
    EditandoNominaGuardada = False
    BloqueoPorPlazoEdicion = False
    ReiniciarEstadoSesion
    
    Open "personal.dno" For Random As #2 Len = Len(personal)
    Open "maestro.dno" For Random As #8 Len = Len(maestro)
    On Error GoTo 0

    MsgBox "No fue posible abrir la nomina." & vbCrLf & "Error " & CStr(errorNum) & ": " & errorTexto, vbExclamation, "Abrir nomina"

End Sub

Private Sub Form_Activate()

    Dim rutaNom As String
    Dim nombreArchivo As String
    Dim nombreBase As String
    Dim indiceTipo As Integer
    Dim mes As Integer
    Dim i As Integer
    Dim entradaMes As String
    Dim encontrado As Boolean

    
    
    '=====================================================
    ' FOCO INICIAL DE FORM8
    '
    ' Solamente al abrir una nueva pantalla de captura.
    ' No interferir con la edicion desde FormViewer.
    '=====================================================

    If Len(RutaEdicionPendienteVisor) = 0 Then
        If FocoInicialSelector And _
           Not CapturaActivaSesion Then
            FocoInicialSelector = False
            If Me.Visible And Option1.Enabled Then
                Option1.SetFocus
            End If
        End If
        Exit Sub
    End If
    
    'Activar proteccion de foco durante la preparacion de la nomina.
    PreparandoEdicionDesdeVisor = True
    
    rutaNom = RutaEdicionPendienteVisor
    RutaEdicionPendienteVisor = ""


    On Error GoTo ErrorEdicion

    nombreArchivo = mID$(rutaNom, InStrRev(rutaNom, "\") + 1)

    If Len(nombreArchivo) <= 4 Then
        Err.Raise vbObjectError + 2950, , "Nombre de nomina invalido."
    End If

    nombreBase = Left$(nombreArchivo, Len(nombreArchivo) - 4)

    If EsNominaNormal(nombreArchivo) Then

        'Nomina ordinaria.
        Option1.Value = True

        encontrado = False

        For i = 0 To Combo1.ListCount - 1

            If UCase$(Left$(Combo1.List(i), 3)) = UCase$(Left$(nombreArchivo, 3)) Then

                Combo1.ListIndex = i
                encontrado = True
                Exit For

            End If

        Next i

        If Not encontrado Then
            Err.Raise vbObjectError + 2951, , "No se pudo identificar el mes de la nomina."
        End If

        If mID$(nombreArchivo, 4, 1) = "1" Then
            Option3.Value = True
        Else
            Option4.Value = True
        End If

        N_ormal = 0

    Else

        'Nomina especial.
        Option2.Value = True
        indiceTipo = -1
        If InStr(1, nombreBase, "FINIQUITO", vbTextCompare) > 0 Then
            indiceTipo = 3
        ElseIf InStr(1, nombreBase, "LIQUIDACION", vbTextCompare) > 0 Then
            indiceTipo = 4
        ElseIf InStr(1, nombreBase, "AGUINALDO", vbTextCompare) > 0 Then
            indiceTipo = 0
        ElseIf InStr(1, nombreBase, "PTU", vbTextCompare) > 0 Then
            indiceTipo = 1
        ElseIf InStr(1, nombreBase, "BONO", vbTextCompare) > 0 Then
            indiceTipo = 2
        End If

        If indiceTipo < 0 Then
            Err.Raise vbObjectError + 2952, , "No se pudo identificar el tipo de nomina especial."
        End If

        cboTipoNominaEspecial.ListIndex = indiceTipo

        mes = 0

        If indiceTipo = 3 Or indiceTipo = 4 Then

            'Finiquito y liquidacion incluyen mes en el sufijo del nombre del archivo.
            For i = 0 To cboMesNominaEspecial.ListCount - 1
                If UCase$(Left$(cboMesNominaEspecial.List(i), 3)) = UCase$(Right$(nombreBase, 3)) Then

                    mes = i + 1
                    Exit For
                End If
            Next i

            If mes = 0 Then
                Err.Raise vbObjectError + 2953, , "No se pudo identificar el mes de la separacion."
            End If

        
    ElseIf indiceTipo = 2 Then

        '=====================================================
        ' BONO
        '
        ' NUEVO:
        ' El mes se obtiene del sufijo del archivo.
        '
        ' EJEMPLOS:
        '   1BONOENE.NOM
        '   1BONOJUN.NOM
        '   1BONONOV.NOM
        '
        ' ANTIGUO:
        ' Si el archivo no tiene sufijo de mes,
        ' conservar la solicitud manual del mes.
        '=====================================================

        For i = 0 To cboMesNominaEspecial.ListCount - 1
            If UCase$(Left$(cboMesNominaEspecial.List(i), 3)) = UCase$(Right$(nombreBase, 3)) Then
                mes = i + 1
                Exit For
            End If
        Next i

        '=====================================================
        ' COMPATIBILIDAD CON BONOS ANTERIORES
        '=====================================================

        If mes = 0 Then
            entradaMes = InputBox("Indique el mes fiscal original de esta nomina BONO (del 1 al 12).", "Periodo de nomina especial", CStr(Month(Date)))
            If Len(Trim$(entradaMes)) = 0 Then
                Me.Hide
                Exit Sub
            End If

            If Not IsNumeric(entradaMes) Then
                Err.Raise vbObjectError + 2954, , "El mes debe ser numerico."

            End If

            mes = CInt(entradaMes)

            If mes < 1 Or mes > 12 Then
                Err.Raise vbObjectError + 2955, , "El mes debe estar entre 1 y 12."
            End If
        End If
    Else
    
        'AGUINALDO y PTU: Conservar el comportamiento original.
            entradaMes = InputBox("Indique el mes fiscal original de esta nomina (del 1 al 12).", "Periodo de nomina especial", CStr(Month(Date)))

            If Len(Trim$(entradaMes)) = 0 Then
                Me.Hide
                Exit Sub
            End If

            If Not IsNumeric(entradaMes) Then
                Err.Raise vbObjectError + 2954, , "El mes debe ser numerico."
            End If

            mes = CInt(entradaMes)

            If mes < 1 Or mes > 12 Then
                Err.Raise vbObjectError + 2955, , "El mes debe estar entre 1 y 12."
            End If

        End If

        cboMesNominaEspecial.ListIndex = mes - 1

        'Mantener el nombre original archivado.
        Text1.Text = nombreBase
        Label7.Caption = nombreBase

        N_ormal = 1

    End If

    
    '=========================================================
    ' ABRIR NOMINA GUARDADA EN FORM8
    '
    ' El mismo procedimiento puede devolver:
    '
    ' 1. Edicion autorizada.
    ' 2. Consulta de solo lectura.
    '
    ' Ambos modos son aperturas validas.
    '=========================================================

    AbrirNominaGuardadaParaEditar rutaNom
    PreparandoEdicionDesdeVisor = False

    '=========================================================
    ' COMPROBAR SI LA CARGA SE REALIZO
    '
    ' No comprobar EditandoNominaGuardada:
    '
    ' En consulta esa variable sera False, aunque
    ' la nomina se haya cargado correctamente.
    '
    ' CapturaActivaSesion sera True para ambos modos.
    '=========================================================

    If Not CapturaActivaSesion Then
        'La apertura fallo.
        'Conservar el regreso modal existente.
        Me.Hide
        Exit Sub
    End If

    '=========================================================
    ' LA NOMINA ESTA CARGADA
    '
    ' No cerrar Form8.
    ' No redirigir al Viewer.
    '=========================================================

    Exit Sub
ErrorEdicion:
    PreparandoEdicionDesdeVisor = False
    MsgBox "No se pudo preparar la edicion." & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Modificar nomina"
    Me.Hide
End Sub

Private Sub LeerBytesArchivo(ByVal ruta As String, ByRef datos() As Byte)

    Dim h As Integer
    Dim tam As Long
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorLectura

    Erase datos

    h = FreeFile
    Open ruta For Binary Access Read As #h

    tam = LOF(h)

    If tam <= 0 Then
        Err.Raise vbObjectError + 3001, , "El archivo esta vacio: " & ruta
    End If

    ReDim datos(0 To tam - 1)
    Get #h, 1, datos

    Close #h
    Exit Sub

ErrorLectura:
    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next
    If h <> 0 Then Close #h
    On Error GoTo 0

    Err.Raise numeroError, "LeerBytesArchivo", descripcionError

End Sub

Private Sub RestaurarBytesArchivo(ByVal ruta As String, ByRef datos() As Byte)

    Dim h As Integer
    Dim tam As Long
    Dim numeroError As Long
    Dim descripcionError As String

    On Error GoTo ErrorRestauracion

    'Nunca crear un archivo durante la restauracion.
    If Len(Dir$(ruta, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then

        Err.Raise vbObjectError + 3002, , "El archivo original ya no existe."
    End If

    tam = UBound(datos) - LBound(datos) + 1

    h = FreeFile

    Open ruta For Binary Access Read Write As #h

    'No modificar un archivo cuya estructura cambio.
    If LOF(h) <> tam Then
        Err.Raise vbObjectError + 3003, , "El tamano del archivo cambio."
    End If

    Put #h, 1, datos

    Close #h
    Exit Sub

ErrorRestauracion:
    numeroError = Err.Number
    descripcionError = Err.Description

    On Error Resume Next
    If h <> 0 Then Close #h
    On Error GoTo 0

    Err.Raise numeroError, "RestaurarBytesArchivo", descripcionError

End Sub

Private Sub CopiarBytesMemoria(ByRef origen() As Byte, ByRef destino() As Byte)

    Dim i As Long
    ReDim destino(0 To UBound(origen))
    
    For i = 0 To UBound(origen)
        destino(i) = origen(i)
    Next i
End Sub

Private Function ArchivoCoincideConMemoria(ByVal ruta As String, ByRef copia() As Byte) As Boolean

    Dim datosActuales() As Byte
    Dim i As Long

    ArchivoCoincideConMemoria = False

    If Len(Dir$(ruta, vbNormal Or vbHidden Or vbReadOnly Or vbSystem)) = 0 Then
        Err.Raise vbObjectError + 3050, , "El archivo CMP ya no existe."
    End If

    LeerBytesArchivo ruta, datosActuales

    'Si los tamaños son diferentes, no son iguales.
    If UBound(datosActuales) <> UBound(copia) Then
        Exit Function
    End If

    For i = 0 To UBound(copia)
        If datosActuales(i) <> copia(i) Then
            Exit Function
        End If
    Next i

    ArchivoCoincideConMemoria = True

End Function

Private Function ValidarCapturaAguinaldo() As Boolean

    Dim fila As Long
    Dim columna As Long
    Dim empleado As String

    Dim importeAguinaldo As Currency
    Dim importeExento As Currency
    Dim totalIngresos As Currency
    Dim totalDeducciones As Currency
    Dim netoCapturado As Currency

    ValidarCapturaAguinaldo = False

    '=========================================================
    ' SOLO APLICA PARA AGUINALDO
    '=========================================================

    If g_TipoNominaActiva <> tnAguinaldo Then
        ValidarCapturaAguinaldo = True
        Exit Function

    End If

    '=========================================================
    ' VALIDAR CADA EMPLEADO DE LA CAPTURA
    '=========================================================

    For fila = 1 To limite
        empleado = Trim$(ConNom1.TextMatrix(fila, C_NUM))
        
        If Val(empleado) > 0 Then

            '-------------------------------------------------
            ' 1. COMPROBAR QUE NO EXISTAN OTRAS PERCEPCIONES
            '
            ' Aguinaldo solamente admite C_CONC5.
            ' C_EXENTA se revisa por separado.
            '-------------------------------------------------

            For columna = C_PERC_INI To C_PERC_FIN
                If columna <> C_CONC5 Then
                    If CelCur(fila, columna) <> 0 Then
                        MsgBox "Empleado: " & empleado & vbCrLf & vbCrLf & "Se encontro un importe en: " & ConNom1.TextMatrix(0, columna) & vbCrLf & "Esta nomina solamente permite Aguinaldo y su percepcion exenta.", vbExclamation, "Validacion Aguinaldo"
                        Exit Function
                    End If
                End If
            Next columna

            '-------------------------------------------------
            ' 2. OBTENER AGUINALDO Y EXENCION
            '-------------------------------------------------

            importeAguinaldo = CelCur(fila, C_CONC5)
            importeExento = CelCur(fila, C_EXENTA)
            
            If importeAguinaldo < 0 Or importeExento < 0 Then
                MsgBox "Empleado: " & empleado & vbCrLf & "Existen importes negativos en Aguinaldo o Perc.exenta.", vbExclamation, "Validacion Aguinaldo"
                Exit Function
            End If

            If importeExento > importeAguinaldo Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta supera el aguinaldo total.", vbExclamation, "Validacion Aguinaldo"
                Exit Function
            End If

            '-------------------------------------------------
            ' 3. COMPROBAR EL TOTAL DE INGRESOS
            '
            ' El exento ya forma parte del aguinaldo.
            '-------------------------------------------------

            totalIngresos = CelCur(fila, C_TOTING)
            
            If Abs(totalIngresos - importeAguinaldo) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El total de ingresos no coincide." & vbCrLf & "Aguinaldo: " & Format$(importeAguinaldo, "#,##0.00") & vbCrLf & "Tot.Ingr.: " & Format$(totalIngresos, "#,##0.00"), vbExclamation, "Validacion Aguinaldo"
                Exit Function
            End If

            '-------------------------------------------------
            ' 4. COMPROBAR EL NETO
            '-------------------------------------------------

            totalDeducciones = CelCur(fila, C_TOTDED)
            netoCapturado = CelCur(fila, C_NETO)
            
            If Abs(netoCapturado - (totalIngresos - totalDeducciones)) > CCur(0.01) Then

                MsgBox "Empleado: " & empleado & vbCrLf & "El neto no coincide con los ingresos menos las deducciones.", vbExclamation, "Validacion Aguinaldo"
                Exit Function
            End If
        End If
    Next fila
    ValidarCapturaAguinaldo = True
End Function

Private Function ValidarCapturaPTU() As Boolean

    Dim fila As Long
    Dim columna As Long
    Dim empleado As String

    Dim ptuTotal As Currency
    Dim ptuExenta As Currency
    Dim totalIngresos As Currency
    Dim totalDeducciones As Currency
    Dim netoCapturado As Currency

    ValidarCapturaPTU = False

    '=========================================================
    ' SOLO APLICA PARA PTU
    '=========================================================

    If g_TipoNominaActiva <> tnPTU Then
        ValidarCapturaPTU = True
        Exit Function
    End If

    '=========================================================
    ' VALIDAR CADA EMPLEADO
    '=========================================================

    For fila = 1 To limite
        empleado = Trim$(ConNom1.TextMatrix(fila, C_NUM))
        If Val(empleado) > 0 Then

            '-------------------------------------------------
            ' 1. NO PERMITIR OTRAS PERCEPCIONES
            '-------------------------------------------------

            For columna = C_PERC_INI To C_PERC_FIN
                If columna <> C_CONC6 Then
                    If CelCur(fila, columna) <> 0 Then
                        MsgBox "Empleado: " & empleado & vbCrLf & vbCrLf & "Existe un importe en la columna: " & ConNom1.TextMatrix(0, columna) & vbCrLf & "Esta nomina solamente permite PTU y su percepcion exenta.", vbExclamation, "Validacion PTU"
                        Exit Function
                    End If
                End If
            Next columna

            '-------------------------------------------------
            ' 2. LEER PTU TOTAL Y EXENTA
            '-------------------------------------------------

            ptuTotal = CelCur(fila, C_CONC6)
            ptuExenta = CelCur(fila, C_EXENTA)

            If ptuTotal < 0 Or ptuExenta < 0 Then
                MsgBox "Empleado: " & empleado & vbCrLf & "PTU y Perc.exenta no pueden contener importes negativos.", vbExclamation, "Validacion PTU"
                Exit Function
            End If

            If ptuExenta > ptuTotal Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La PTU exenta supera el importe total de PTU.", vbExclamation, "Validacion PTU"
                Exit Function
            End If

            '-------------------------------------------------
            ' 3. CONCILIAR TOTAL DE INGRESOS
            '
            ' PTU exenta ya forma parte de PTU total.
            '-------------------------------------------------

            totalIngresos = CelCur(fila, C_TOTING)

            If Abs(totalIngresos - ptuTotal) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "No coincide el total de ingresos." & vbCrLf & "PTU: " & Format$(ptuTotal, "#,##0.00") & vbCrLf & "Tot.Ingr.: " & Format$(totalIngresos, "#,##0.00"), vbExclamation, "Validacion PTU"
                Exit Function
            End If

            '-------------------------------------------------
            ' 4. CONCILIAR EL NETO
            '-------------------------------------------------

            totalDeducciones = CelCur(fila, C_TOTDED)
            netoCapturado = CelCur(fila, C_NETO)

            If Abs(netoCapturado - (totalIngresos - totalDeducciones)) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El neto no coincide con ingresos menos deducciones.", vbExclamation, "Validacion PTU"

                Exit Function

            End If

        End If

    Next fila

    ValidarCapturaPTU = True

End Function

Private Function ValidarCapturaBono() As Boolean

    Dim fila As Long
    Dim columna As Long
    Dim empleado As String

    Dim bonoTotal As Currency
    Dim bonoExento As Currency
    Dim totalIngresos As Currency
    Dim totalDeducciones As Currency
    Dim netoCapturado As Currency

    ValidarCapturaBono = False

    '=========================================================
    ' SOLO NOMINA ESPECIAL BONO
    '=========================================================

    If g_TipoNominaActiva <> tnBonoPremio Then
        ValidarCapturaBono = True
        Exit Function
    End If

    '=========================================================
    ' VALIDAR CADA EMPLEADO
    '=========================================================

    For fila = 1 To limite
        empleado = Trim$(ConNom1.TextMatrix(fila, C_NUM))
        If Val(empleado) > 0 Then

            '-------------------------------------------------
            ' 1. VERIFICAR QUE NO EXISTAN OTRAS PERCEPCIONES
            '-------------------------------------------------

            For columna = C_PERC_INI To C_PERC_FIN
                If columna <> C_CONC7 Then
                    If CelCur(fila, columna) <> 0 Then
                        MsgBox "Empleado: " & empleado & vbCrLf & vbCrLf & "Existe un importe en: " & ConNom1.TextMatrix(0, columna) & vbCrLf & "Esta nomina solamente permite BONO y su percepcion exenta.", vbExclamation, "Validacion BONO"
                        Exit Function
                    End If
                End If
            Next columna

            '-------------------------------------------------
            ' 2. OBTENER BONO TOTAL Y EXENTO
            '-------------------------------------------------

            bonoTotal = CelCur(fila, C_CONC7)
            bonoExento = CelCur(fila, C_EXENTA)

            If bonoTotal < 0 Or bonoExento < 0 Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El bono y su percepcion exenta no pueden ser negativos.", vbExclamation, "Validacion BONO"
                Exit Function
            End If

            If bonoExento > bonoTotal Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta supera el importe total del bono.", vbExclamation, "Validacion BONO"
                Exit Function
            End If

            '-------------------------------------------------
            ' 3. CONCILIAR TOTAL DE INGRESOS
            '
            ' La parte exenta ya esta dentro del bono.
            '-------------------------------------------------

            totalIngresos = CelCur(fila, C_TOTING)
            
            If Abs(totalIngresos - bonoTotal) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El total de ingresos no coincide." & vbCrLf & "BONO: " & Format$(bonoTotal, "#,##0.00") & vbCrLf & "Tot.Ingr.: " & Format$(totalIngresos, "#,##0.00"), vbExclamation, "Validacion BONO"
                Exit Function
            End If

            '-------------------------------------------------
            ' 4. CONCILIAR EL NETO
            '-------------------------------------------------

            totalDeducciones = CelCur(fila, C_TOTDED)
            netoCapturado = CelCur(fila, C_NETO)

            If Abs(netoCapturado - (totalIngresos - totalDeducciones)) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El neto no coincide con ingresos menos deducciones.", vbExclamation, "Validacion BONO"
                Exit Function
            End If
        End If
    Next fila
    ValidarCapturaBono = True
End Function

Private Function MaximoExentoFiniquitoNuevo(ByVal fila As Long) As Currency

    Dim total As Currency
    
    total = 0

    '=====================================================
    ' CONCEPTOS POTENCIALMENTE EXENTOS
    '
    ' Se conserva el mismo criterio que actualmente
    ' utiliza checarSeparacionNueva().
    '
    ' No calcula limites fiscales.
    ' No distribuye el importe exento por concepto.
    '=====================================================

    If C_LQ_AGUI >= 0 Then
        total = total + CelCur(fila, C_LQ_AGUI)
    End If

    If C_LQ_PRIMAVAC >= 0 Then
        total = total + CelCur(fila, C_LQ_PRIMAVAC)
    End If

    If C_LQ_ANTIG >= 0 Then
        total = total + CelCur(fila, C_LQ_ANTIG)
    End If

    If C_LQ_INDEMN >= 0 Then
        total = total + CelCur(fila, C_LQ_INDEMN)
    End If

    MaximoExentoFiniquitoNuevo = total

End Function

Private Function ValidarCapturaFiniquitoNuevo() As Boolean

    Dim fila As Long
    Dim columna As Long
    Dim empleado As String

    Dim totalPercepciones As Currency
    Dim totalExento As Currency
    Dim totalIngresos As Currency
    Dim totalDeducciones As Currency
    Dim totalNeto As Currency

    ValidarCapturaFiniquitoNuevo = False

    '=========================================================
    ' SOLO FINIQUITO DE FORMATO NUEVO
    '=========================================================

    If g_TipoNominaActiva <> tnLiquidacionFiniquito Then
        ValidarCapturaFiniquitoNuevo = True
        Exit Function

    End If

    If Not g_EsFiniquito Then
        ValidarCapturaFiniquitoNuevo = True
        Exit Function

    End If

    If Not EsSeparacionNueva() Then
        ValidarCapturaFiniquitoNuevo = True
        Exit Function

    End If

    '=========================================================
    ' RECORRER LOS EMPLEADOS
    '=========================================================

    For fila = 1 To limite
        empleado = Trim$(ConNom1.TextMatrix(fila, C_NUM))
        If Val(empleado) > 0 Then

            '-------------------------------------------------
            ' 1. COMPROBAR SUELDO AUXILIAR
            '
            ' En esta estructura, C_SUELDO esta reservado.
            '-------------------------------------------------

            If CelCur(fila, C_SUELDO) <> 0 Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La columna Sueldo contiene un importe en el finiquito nuevo." & vbCrLf & "Revise la captura de dias.", vbExclamation, "Validacion Finiquito"
                Exit Function
            End If

            '-------------------------------------------------
            ' 2. SUMAR PERCEPCIONES
            '
            ' C_EXENTA no se suma.
            ' Es una parte del total bruto.
            '-------------------------------------------------

            totalPercepciones = 0

            For columna = C_PERC_INI To C_PERC_FIN
                If CelCur(fila, columna) < 0 Then
                    MsgBox "Empleado: " & empleado & vbCrLf & "La columna " & ConNom1.TextMatrix(0, columna) & " contiene una percepcion negativa.", vbExclamation, "Validacion Finiquito"
                    Exit Function
                End If

                totalPercepciones = totalPercepciones + CelCur(fila, columna)
            Next columna

            '-------------------------------------------------
            ' 3. VALIDAR PERCEPCION EXENTA
            '-------------------------------------------------

            totalExento = CelCur(fila, C_EXENTA)

            If totalExento < 0 Or totalExento > totalPercepciones Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta supera el importe de las percepciones.", vbExclamation, "Validacion Finiquito"
                Exit Function
            End If

            If totalExento > MaximoExentoFiniquitoNuevo(fila) Then

                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta es mayor que los conceptos potencialmente exentos registrados.", vbExclamation, "Validacion Finiquito"
                Exit Function
            End If

            '-------------------------------------------------
            ' 4. VALIDAR TOTAL DE INGRESOS
            '-------------------------------------------------

            totalIngresos = CelCur(fila, C_TOTING)
            
            If Abs(totalIngresos - totalPercepciones) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El Total de Ingresos no coincide." & vbCrLf & "Conceptos: " & Format$(totalPercepciones, "#,##0.00") & vbCrLf & "Tot.Ingr.: " & Format$(totalIngresos, "#,##0.00"), vbExclamation, "Validacion Finiquito"
                Exit Function
            End If

            '-------------------------------------------------
            ' 5. VALIDAR DEDUCCIONES
            '-------------------------------------------------

            totalDeducciones = 0

            For columna = C_DED_INI To C_DED_FIN
                totalDeducciones = totalDeducciones + CelCur(fila, columna)
            Next columna

            If Abs(totalDeducciones - CelCur(fila, C_TOTDED)) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El total de deducciones no coincide.", vbExclamation, "Validacion Finiquito"
                Exit Function

            End If

            '-------------------------------------------------
            ' 6. VALIDAR NETO
            '-------------------------------------------------

            totalNeto = totalPercepciones - totalDeducciones

            If Abs(totalNeto - CelCur(fila, C_NETO)) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El neto calculado no coincide." & vbCrLf & "Calculado: " & Format$(totalNeto, "#,##0.00") & vbCrLf & "Capturado: " & Format$(CelCur(fila, C_NETO), "#,##0.00"), vbExclamation, "Validacion Finiquito"
                Exit Function
            End If
        End If
    Next fila

    ValidarCapturaFiniquitoNuevo = True

End Function

Private Function MaximoExentoLiquidacionNueva(ByVal fila As Long) As Currency

    Dim total As Currency

    total = 0

    '=====================================================
    ' LIQUIDACION NUEVA
    '
    ' Conceptos que actualmente utiliza el proyecto
    ' para limitar la percepcion exenta capturada.
    '
    ' No determina limites fiscales del SAT.
    ' No distribuye la exencion entre conceptos.
    '=====================================================

    If g_TipoNominaActiva <> tnLiquidacionFiniquito Then
        Exit Function
    End If

    If g_EsFiniquito Then
        Exit Function
    End If

    If Not EsSeparacionNueva() Then
        Exit Function
    End If

    '-----------------------------------------------------
    ' AGUINALDO
    '-----------------------------------------------------

    If C_LQ_AGUI >= 0 Then
        total = total + CelCur(fila, C_LQ_AGUI)
    End If

    '-----------------------------------------------------
    ' PRIMA VACACIONAL
    '-----------------------------------------------------

    If C_LQ_PRIMAVAC >= 0 Then
        total = total + CelCur(fila, C_LQ_PRIMAVAC)
    End If

    '-----------------------------------------------------
    ' PRIMA DE ANTIGUEDAD
    '-----------------------------------------------------

    If C_LQ_ANTIG >= 0 Then
        total = total + CelCur(fila, C_LQ_ANTIG)
    End If

    '-----------------------------------------------------
    ' INDEMNIZACION
    '-----------------------------------------------------

    If C_LQ_INDEMN >= 0 Then
        total = total + CelCur(fila, C_LQ_INDEMN)

    End If

    MaximoExentoLiquidacionNueva = total

End Function

Private Function ValidarCapturaLiquidacionNueva() As Boolean

    Dim fila As Long
    Dim columna As Long
    Dim empleado As String

    Dim totalPercepciones As Currency
    Dim totalExento As Currency
    Dim totalIngresos As Currency
    Dim totalDeducciones As Currency
    Dim totalNeto As Currency

    ValidarCapturaLiquidacionNueva = False

    '=========================================================
    ' 1. IDENTIFICAR LIQUIDACION NUEVA
    '
    ' No aplicar a:
    ' - Nomina ordinaria
    ' - Aguinaldo
    ' - PTU
    ' - Bono
    ' - Finiquito
    ' - Liquidaciones antiguas
    '=========================================================

    If g_TipoNominaActiva <> tnLiquidacionFiniquito Then
        ValidarCapturaLiquidacionNueva = True
        Exit Function
    End If

    If g_EsFiniquito Then
        ValidarCapturaLiquidacionNueva = True
        Exit Function
    End If

    If Not EsSeparacionNueva() Then
        ValidarCapturaLiquidacionNueva = True
        Exit Function
    End If

    '=========================================================
    ' 2. RECORRER EMPLEADOS
    '=========================================================

    For fila = 1 To limite
        empleado = Trim$(ConNom1.TextMatrix(fila, C_NUM))
        If Val(empleado) > 0 Then

            '-------------------------------------------------
            ' 3. VALIDAR SUELDO AUXILIAR
            '-------------------------------------------------

            If CelCur(fila, C_SUELDO) <> 0 Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La columna Sueldo contiene un importe que no corresponde a la Liquidacion nueva.", vbExclamation, "Validacion Liquidacion"

                Exit Function

            End If

            '-------------------------------------------------
            ' 4. SUMAR PERCEPCIONES
            '
            ' Columnas 4 a 9 contienen conceptos brutos.
            ' C_EXENTA no se suma nuevamente.
            '-------------------------------------------------

            totalPercepciones = 0

            For columna = C_PERC_INI To C_PERC_FIN
                If CelCur(fila, columna) < 0 Then
                    MsgBox "Empleado: " & empleado & vbCrLf & "Existe una percepcion negativa en la columna: " & ConNom1.TextMatrix(0, columna), vbExclamation, "Validacion Liquidacion"
                    Exit Function
                End If
                totalPercepciones = totalPercepciones + CelCur(fila, columna)

            Next columna

            '-------------------------------------------------
            ' 5. VALIDAR EXENTO
            '
            ' Es parte de las percepciones, no adicional.
            '-------------------------------------------------

            totalExento = CelCur(fila, C_EXENTA)

            If totalExento < 0 Or totalExento > totalPercepciones Then
                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta es invalida para el total de la Liquidacion.", vbExclamation, "Validacion Liquidacion"
                Exit Function
            End If

            If totalExento > _
               MaximoExentoLiquidacionNueva(fila) Then

                MsgBox "Empleado: " & empleado & vbCrLf & "La percepcion exenta supera los conceptos potencialmente exentos de la Liquidacion.", vbExclamation, "Validacion Liquidacion"
                Exit Function
            End If

            '-------------------------------------------------
            ' 6. VALIDAR TOTAL DE INGRESOS
            '-------------------------------------------------

            totalIngresos = CelCur(fila, C_TOTING)

            If Abs(totalIngresos - totalPercepciones) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El total de ingresos no coincide." & vbCrLf & "Conceptos: " & Format$(totalPercepciones, "#,##0.00") & vbCrLf & "Tot.Ingr.: " & Format$(totalIngresos, "#,##0.00"), vbExclamation, "Validacion Liquidacion"
                Exit Function
            End If

            '-------------------------------------------------
            ' 7. VALIDAR DEDUCCIONES
            '
            ' Respetar los signos utilizados por el grid.
            '-------------------------------------------------

            totalDeducciones = 0

            For columna = C_DED_INI To C_DED_FIN
                totalDeducciones = totalDeducciones + CelCur(fila, columna)

            Next columna

            If Abs(totalDeducciones - CelCur(fila, C_TOTDED)) > CCur(0.01) Then

                MsgBox "Empleado: " & empleado & vbCrLf & "El total de deducciones no coincide.", vbExclamation, "Validacion Liquidacion"
                Exit Function
            End If

            '-------------------------------------------------
            ' 8. VALIDAR NETO
            '-------------------------------------------------

            totalNeto = totalIngresos - totalDeducciones

            If Abs(totalNeto - CelCur(fila, C_NETO)) > CCur(0.01) Then
                MsgBox "Empleado: " & empleado & vbCrLf & "El neto calculado no coincide." & vbCrLf & "Calculado: " & Format$(totalNeto, "#,##0.00") & vbCrLf & "Captura: " & Format$(CelCur(fila, C_NETO), "#,##0.00"), vbExclamation, "Validacion Liquidacion"

                Exit Function
            End If
        End If
    Next fila
    ValidarCapturaLiquidacionNueva = True
End Function

Private Sub AplicarModoConsultaForm8()

    '=========================================================
    ' FORM8 - MODO SOLO CONSULTA
    '
    ' IMPORTANTE:
    '
    ' Este procedimiento NO modifica importes.
    ' NO recalcula percepciones.
    ' NO archiva NOM ni CMP.
    '
    ' Solamente deshabilita las operaciones que
    ' podrian modificar informacion.
    '
    ' Las opciones de impresion y exportacion
    ' permanecen disponibles.
    '=========================================================

        'Aplicar tanto a consulta como a una
    'edicion cuyo plazo acaba de vencer.

    If Not mSoloConsulta And Not BloqueoPorPlazoEdicion Then
        Exit Sub
    End If

    If mSoloConsulta Then
        Me.Caption = "Nomina - SOLO CONSULTA"
    Else
        Me.Caption = "Nomina - EDICION VENCIDA"
    End If
    '=========================================================
    ' 2. IMPEDIR GUARDADO Y NUEVAS CAPTURAS
    '=========================================================

    archnom(2).Enabled = False
    nomcap_ini(1).Enabled = False
    capturaGeneral.Enabled = False

    '=========================================================
    ' 3. IMPEDIR MODIFICACIONES DESDE LOS MENUS
    '=========================================================

    funcionPegar.Enabled = False
    Importar.Enabled = False
    Borrar.Enabled = False

    'No abrir la edicion del archivo de personal
    'desde una consulta de nomina.

    edicionPersonalCaptura.Enabled = False

    '=========================================================
    ' 4. DESHABILITAR RECALCULOS MANUALES
    '
    ' En este proyecto estas operaciones pueden
    ' actualizar ISR, IMSS y el complemento CMP.
    '
    ' Los importes se consultaran tal como fueron
    ' almacenados.
    '=========================================================

    capact(0).Enabled = False

    '=========================================================
    ' 5. DESHABILITAR ORDENAMIENTOS Y ELIMINACIONES
    '
    ' Evitar cambios accidentales sobre las filas
    ' de la cuadricula cargada.
    '=========================================================

    capnomord(0).Enabled = False

    '=========================================================
    ' 6. BLOQUEAR LOS EDITORES DE CELDAS
    '
    ' El grid permanecera habilitado para:
    ' - Consultar importes.
    ' - Desplazarse por filas y columnas.
    ' - Seleccionar y copiar informacion.
    '=========================================================

    Text2.Enabled = False
    Text3.Enabled = False
    ConNom1.Enabled = True

    '=========================================================
    ' 7. CFDI
    '
    ' Una nomina fuera del plazo interno no debe
    ' iniciar timbrado desde esta ruta.
    '
    ' La politica fiscal definitiva de timbrado
    ' se verificara por separado.
    '=========================================================

    CFDI2.Enabled = False

    '=========================================================
    ' 8. CONSERVAR OPERACIONES DE CONSULTA
    '
    ' No modificar:
    '
    ' capnoimp     = Impresion
    ' exportar     = Exportacion
    ' CapArch      = Archivo Banamex
    ' EdNomCopiar  = Copiar
    ' capnosal     = Salir
    '
    ' IMPORTANTE:
    ' No se habilitan opciones que originalmente
    ' estuvieran deshabilitadas por otras razones.
    '=========================================================

End Sub

Private Function AutorizarCambioNominaForm8() As Boolean

    Dim detalle As String
    AutorizarCambioNominaForm8 = False

    '=========================================================
    ' 1. NOMINA ABIERTA EN SOLO CONSULTA
    '
    ' No autorizar cambios de importes.
    '=========================================================

    If mSoloConsulta Then Exit Function

    '=========================================================
    ' 2. EDICION QUE YA FUE BLOQUEADA
    '
    ' La sesion debe cerrarse descartando los
    ' cambios pendientes.
    '=========================================================

    If BloqueoPorPlazoEdicion Then Exit Function

    '=========================================================
    ' 3. NOMINA GUARDADA EN MODO EDICION
    '
    ' Comprobar el plazo nuevamente.
    '
    ' Esto evita que una ventana abierta antes
    ' del vencimiento permita continuar editando
    ' despues de la fecha limite.
    '=========================================================

    If EditandoNominaGuardada Then
        If Len(RutaNominaSesion) = 0 Then
            Exit Function
        End If

        If Not PuedeModificarNomina5Dias(RutaNominaSesion, detalle) Then
            BloqueoPorPlazoEdicion = True
            mDetalleConsulta = detalle

            'Bloquear nuevas operaciones modificadoras.
            AplicarModoConsultaForm8

            MsgBox "El plazo de edicion de esta nomina ha finalizado." & vbCrLf & vbCrLf & detalle & vbCrLf & vbCrLf & "Los cambios pendientes no podran guardarse." & vbCrLf & "Cierre la ventana y seleccione NO " & _
                   "cuando se pregunte si desea guardar, para restaurar el CMP original.", vbExclamation, "Edicion vencida"
            Exit Function
        End If
    End If

    '=========================================================
    ' 4. AUTORIZAR
    '
    ' Captura nueva:
    '   Funcionamiento original.
    '
    ' Edicion vigente:
    '   Funcionamiento original.
    '=========================================================

    AutorizarCambioNominaForm8 = True
End Function


