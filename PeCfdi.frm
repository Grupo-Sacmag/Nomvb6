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
Dim temporal, Temporal1
Sub Col_Def()
PCpCfdi.Clear: PCpCfdi.Rows = 200: PCpCfdi.Cols = 9: PCpCfdi.FixedCols = 0
   PCpCfdi.FixedRows = 1
   PCpCfdi.Row = 0
   PCpCfdi.Col = 0: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(0) = 800: PCpCfdi.Text = "Num"
   PCpCfdi.Col = 1: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(1) = 3600: PCpCfdi.Text = "Nombre"
   PCpCfdi.Col = 2: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(2) = 2400: PCpCfdi.Text = "Direccion"
   PCpCfdi.Col = 3: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(3) = 2400: PCpCfdi.Text = "Colonia"
   PCpCfdi.Col = 4: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(4) = 2400: PCpCfdi.Text = "Ciudad"
   PCpCfdi.Col = 5: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(5) = 2400: PCpCfdi.Text = "Estado"
   PCpCfdi.Col = 6: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(6) = 2400: PCpCfdi.Text = "Delegacion"
   PCpCfdi.Col = 7: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(7) = 800: PCpCfdi.Text = "Codigo"
   PCpCfdi.Col = 8: PCpCfdi.CellAlignment = 4: PCpCfdi.ColWidth(8) = 2400: PCpCfdi.Text = "Correo "
   'pcpcfdi.Col = 9: pcpcfdi.CellAlignment = 4: pcpcfdi.ColWidth(9) = 600: pcpcfdi.Text = "Cons."
   
End Sub

Private Sub ArCfAbr_Click()
   Dim EmpNum As Long, nombre As String, Cadena1
   PCpCfdi.Rows = 1
    Close 1, 2
    
    Open "Personal.dno" For Random As 2 Len = Len(personal)
    cm = LOF(2) / Len(personal)
    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
    
    For i = 1 To cm: Get 2, i, personal: Get 1, i, Empleado_1
    
    If IsNumeric(Empleado_1.Cpostal) Then
           nombre = (Trim(personal.nom) + " " + Trim(personal.ape1) + " " + Trim(personal.ape2))
           Cadena1 = Format(i, "####0") & Chr(9) & nombre & Chr(9) & Empleado_1.Direccion _
                     & Chr(9) & Empleado_1.Colonia & Chr(9) & Empleado_1.Ciudad _
                     & Chr(9) & Empleado_1.Estado & Chr(9) & Empleado_1.Delegacion _
                     & Chr(9) & Empleado_1.Cpostal & Chr(9) & Empleado_1.correo
          PCpCfdi.AddItem Cadena1
    End If
   Next i
End Sub

Private Sub ArCfGuarda_Click()
Dim EmpNum As Long
    Close 1
    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
    cm = LOF(1) / Len(Empleado_1)
    For renglon = 1 To PCpCfdi.Rows - 1
    
    If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
                EmpNum = PCpCfdi.TextMatrix(renglon, 0)
                Empleado_1.Direccion = PCpCfdi.TextMatrix(renglon, 2)
                Empleado_1.Colonia = PCpCfdi.TextMatrix(renglon, 3)
                Empleado_1.Ciudad = PCpCfdi.TextMatrix(renglon, 4)
                Empleado_1.Estado = PCpCfdi.TextMatrix(renglon, 5)
                Empleado_1.Delegacion = PCpCfdi.TextMatrix(renglon, 6)
                Empleado_1.Cpostal = PCpCfdi.TextMatrix(renglon, 7)
                Empleado_1.correo = PCpCfdi.TextMatrix(renglon, 8)
                Put 1, EmpNum, Empleado_1
                
    End If
    Next renglon
    Close 1
    
End Sub

Private Sub ArCfSale_Click()
   Close
   Unload PeCfdi
End Sub

Private Sub EdCfCopia_Click()
 Dim Temporal1
 Clipboard.Clear
   
   difer = PCpCfdi.RowSel - PCpCfdi.Row
   For i = PCpCfdi.Row To PCpCfdi.RowSel
      
      For f = PCpCfdi.Col To PCpCfdi.ColSel
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
  Dim DeAqui As Integer, RetornoCarro As Long, InicioCopia As Long
  Temporal1 = ""
  temporal = Clipboard.GetText(vbCFText)

  RetornoCarro = PCpCfdi.Col
  InicioCopia = PCpCfdi.Row: DeAqui = InicioCopia
  InIPGr = InicioCopia: InIPGc = RetornoCarro
If temporal <> "" Then
  Clipboard.Clear
  DeAqui = 1
For i = 1 To Len(temporal)
    Select Case Mid(temporal, i, 1)
          Case Chr(9)
          LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
          Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(9)
          PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
          PCpCfdi.Col = PCpCfdi.Col + 1
          DeAqui = i + 1
          Case Chr(13)
          Temporal1 = Temporal1 + PCpCfdi.TextMatrix(PCpCfdi.Row, PCpCfdi.Col) & Chr(13)
          LTXT.Text = Mid(temporal, DeAqui, (i - DeAqui))
          PCpCfdi.Text = Mid(temporal, DeAqui, (i - DeAqui))
          If (PCpCfdi.Rows - 1) <= PCpCfdi.Row Then PCpCfdi.Rows = PCpCfdi.Rows + 1:
          PCpCfdi.Row = PCpCfdi.Row + 1: PCpCfdi.TopRow = 1
          DeAqui = i + 1
          Case Chr(10)
          Temporal1 = Temporal1 & Chr(10)
          PCpCfdi.Col = RetornoCarro
          DeAqui = i + 1
          Case Else
          Rem nada
    End Select
 Next i
 Rem pcpcfdi.Row = InicioCopia:
 PCpCfdi.Col = RetornoCarro
End If
 
 Clipboard.Clear
 Clipboard.SetText Clipboard.GetText + Temporal1
End Sub

Private Sub EdCfSelT_Click()
    Dim limite As Long
    Clipboard.Clear
    PCpCfdi.Row = 1: PCpCfdi.Col = 0
   For limite = 1 To PCpCfdi.Rows - 1
       renglon = limite
    If IsNumeric(PCpCfdi.TextMatrix(renglon, 0)) Then
           PCpCfdi.RowSel = PCpCfdi.Row
    End If
   Next limite
    PCpCfdi.ColSel = PCpCfdi.Cols - 1
End Sub

Private Sub Form_Load()
    Dim oRS As New ADODB.Recordset
    Dim sSQL
    Dim abrEmpresa
    Col_Def
    
    ''Conectar a base datos
    On Error GoTo Error
    
    abrEmpresa = Left(Trim(emp), 4)
  
    If (UCase(abrEmpresa) = "SACM") Then
        abrEmpresa = "SACMAG"
    End If
    If (UCase(abrEmpresa) = "COOR") Then
        abrEmpresa = "CORDINA"
    End If
    If (UCase(abrEmpresa) = "EPES") Then
        abrEmpresa = "EPESA"
    End If
    If (UCase(abrEmpresa) = "SUPE") Then
        abrEmpresa = "SUPERVISA"
    End If
    
    sSQL = "SELECT idNomina, nombre, apellidoP, apellidoM, nombreDeVialidad, nombreColonia, nombreLocalidad, entidadFederativa, nombreMunicipio, cp , correoElectronico " & _
             "FROM datosSat where empresa = '" & abrEmpresa & "'"
  
  ' Create and Open the Recordset object.
    
    Set oRS = New ADODB.Recordset
    oRS.CursorLocation = adUseClient
    oRS.Open sSQL, con, adOpenStatic, adLockBatchOptimistic, adCmdText
                
    oRS.MoveFirst
      
    'Agrega las filas necesarias en el FlexGRid
    
    dat.Rows = oRS.RecordCount + 1
    
    'Agrega las columnas necesarias
    
    dat.Cols = oRS.Fields.Count
    dat.Row = 0: dat.Col = 0
    dat.Col = 0: dat.CellAlignment = 4: dat.ColWidth(0) = 2800: dat.Text = "ID NÓMINA"
    dat.Col = 1: dat.CellAlignment = 4: dat.ColWidth(1) = 2800: dat.Text = "NOMBRE"
    dat.Col = 2: dat.CellAlignment = 4: dat.ColWidth(2) = 2800: dat.Text = "APELLIDO P"
    dat.Col = 3: dat.CellAlignment = 4: dat.ColWidth(3) = 2800: dat.Text = "APELLIDO M"
    dat.Col = 4: dat.CellAlignment = 4: dat.ColWidth(4) = 2800: dat.Text = "DIRECCIÓN"
    dat.Col = 5: dat.CellAlignment = 4: dat.ColWidth(5) = 2800: dat.Text = "COLONIA"
    dat.Col = 6: dat.CellAlignment = 4: dat.ColWidth(6) = 2800: dat.Text = "CIUDAD"
    dat.Col = 7: dat.CellAlignment = 4: dat.ColWidth(7) = 2800: dat.Text = "ESTADO"
    dat.Col = 8: dat.CellAlignment = 4: dat.ColWidth(8) = 2800: dat.Text = "DELEGACIÓN"
    dat.Col = 9: dat.CellAlignment = 4: dat.ColWidth(9) = 2800: dat.Text = "CÓDIGO POSTAL"
    dat.Col = 10: dat.CellAlignment = 4: dat.ColWidth(10) = 2800: dat.Text = "CORREO"
    vardatarows = oRS.GetRows()
     
     For i = 1 To dat.Rows - 1
        For h = 0 To dat.Cols - 1
            If (IsNull(vardatarows(h, i - 1))) Then
                dat.TextMatrix(i, h) = "N/A"
            Else
                dat.TextMatrix(i, h) = Trim(vardatarows(h, i - 1))
            End If
        Next h
     Next i
    oRS.MarshalOptions = adMarshalModifiedOnly
    
    ' Disconnect the Recordset.
    Set oRS.ActiveConnection = Nothing
    oRS.Close
    Set oRS = Nothing
    Close 1, 2
    Open "personal.dno" For Random As 2 Len = Len(personal)
    Dm = LOF(2) / Len(personal)
    Open "Perscfdi.dno" For Random As 1 Len = Len(Empleado_1)
   
    For i = 1 To dat.Rows - 1
        idNomina = dat.TextMatrix(i, 0)
        Get 1, CInt(idNomina), Empleado_1
        Empleado_1.Direccion = UCase(Trim(dat.TextMatrix(i, 4)))
        Empleado_1.Colonia = UCase(Trim(dat.TextMatrix(i, 5)))
        Empleado_1.Ciudad = UCase(Trim(dat.TextMatrix(i, 6)))
        Empleado_1.Estado = UCase(Trim(dat.TextMatrix(i, 7)))
        Empleado_1.Delegacion = UCase(Trim(dat.TextMatrix(i, 8)))
        Empleado_1.Cpostal = UCase(Trim(dat.TextMatrix(i, 9)))
        Empleado_1.correo = UCase(Trim(dat.TextMatrix(i, 10)))
        Put 1, idNomina, Empleado_1
    
    Next i
    
    PeCfdi.Caption = "Datos complementarios de Personal" & " - " & "Estas conectado"
    
    Close 1, 2
    
    Exit Sub
Error:
MsgBox ("Ocurrió un error:" + Err.Description)
   
    
End Sub


Private Sub Form_Resize()

    PCpCfdi.Height = PeCfdi.Height * 0.85
    PCpCfdi.Width = PeCfdi.Width * 0.95
    
End Sub

Private Sub LTXT_Change()
    PCpCfdi.Text = LTXT.Text
End Sub

Private Sub pcpcfdi_EnterCell()
    PCpCfdi.CellBackColor = vbYellow
    valcelant = PCpCfdi.Text
    LTXT.Text = valcelant
    Rem If pcpcfdi.Row > 25 Then pcpcfdi.TopRow = pcpcfdi.TopRow + 1
    
End Sub

Private Sub pcpcfdi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyDelete
               
               For Q = PCpCfdi.Row To PCpCfdi.RowSel
                 
                 For W = PCpCfdi.Col To PCpCfdi.ColSel
                    
                    PCpCfdi.TextMatrix(Q, W) = ""
                 Next W
               Next Q
               
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

