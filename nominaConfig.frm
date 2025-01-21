VERSION 5.00
Begin VB.Form nominaConfiguracion 
   Caption         =   "Configuración"
   ClientHeight    =   1050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4035
   LinkTopic       =   "Form10"
   ScaleHeight     =   1050
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "nominaConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim propiedad As String

Combo1.AddItem "Clasico"
Combo1.AddItem "Oscuro"
Combo1.AddItem "SACMAG"
Combo1.AddItem "Rosa"
Combo1.AddItem "Matrix"
Combo1.AddItem "Eclipse"

End Sub

Private Sub Combo1_Click()
    seleccion Combo1.Text
End Sub

Private Sub seleccion(opcion As String)

Select Case opcion
    Case "Clasico"
        Form1.Label1.BackColor = &H80FFFF
        Form1.Label2.BackColor = &H80FFFF
        Form1.Label3.BackColor = &H80FFFF
        Form1.Label4.BackColor = &H80FFFF
        Form1.Label5.BackColor = &H80FFFF
        Form1.Label6.BackColor = &H80FFFF
        
        Form8.Label7.BackColor = &H80FFFF
        Form8.Label7.ForeColor = &H80000012
        
        Form1.Label1.ForeColor = &H80000012
        Form1.Label2.ForeColor = &H80000012
        Form1.Label3.ForeColor = &H80000012
        Form1.Label4.ForeColor = &H80000012
        Form1.Label5.ForeColor = &H80000012
        Form1.Label6.ForeColor = &H80000012
        
        Form1.BackColor = &H8000000F
        
    
    Case "Oscuro"
        Form1.Label1.BackColor = &H404040
        Form1.Label2.BackColor = &H404040
        Form1.Label3.BackColor = &H404040
        Form1.Label4.BackColor = &H404040
        Form1.Label5.BackColor = &H404040
        Form1.Label6.BackColor = &H404040
                
        Form8.Label7.BackColor = &H404040
        Form8.Label7.ForeColor = &HFFFFFF
        
        Form1.Label1.ForeColor = &HFFFFFF
        Form1.Label2.ForeColor = &HFFFFFF
        Form1.Label3.ForeColor = &HFFFFFF
        Form1.Label4.ForeColor = &HFFFFFF
        Form1.Label5.ForeColor = &HFFFFFF
        Form1.Label6.ForeColor = &HFFFFFF
        
        Form1.BackColor = &HC0C0C0
    Case "SACMAG"
    
        Form1.Label1.BackColor = &H8000000D
        Form1.Label2.BackColor = &H8000000D
        Form1.Label3.BackColor = &H8000000D
        Form1.Label4.BackColor = &H8000000D
        Form1.Label5.BackColor = &H8000000D
        Form1.Label6.BackColor = &H8000000D
                
        Form8.Label7.BackColor = &H8000000D
        Form8.Label7.ForeColor = &HFFFFFF
        
        Form1.Label1.ForeColor = &HFFFFFF
        Form1.Label2.ForeColor = &HFFFFFF
        Form1.Label3.ForeColor = &HFFFFFF
        Form1.Label4.ForeColor = &HFFFFFF
        Form1.Label5.ForeColor = &HFFFFFF
        Form1.Label6.ForeColor = &HFFFFFF
        
        Form1.BackColor = &HC0C0C0
        
    Case "Rosa"
        ' Rosa claro
        Form1.Label1.BackColor = &HFFC0FF
        Form1.Label2.BackColor = &HFFC0FF
        Form1.Label3.BackColor = &HFFC0FF
        Form1.Label4.BackColor = &HFFC0FF
        Form1.Label5.BackColor = &HFFC0FF
        Form1.Label6.BackColor = &HFFC0FF
        Form8.Label5.BackColor = &HFFC0FF
        Form8.Label8.BackColor = &HFFC0FF
        Form8.Label9.BackColor = &HFFC0FF
        Form8.Frame1.BackColor = &HFFC0FF
        Form8.Frame2.BackColor = &HFFC0FF
         Form8.Text1.BackColor = &HFFC0FF
         Form8.Text2.BackColor = &HFFC0FF
         Form8.Text3.BackColor = &HFFC0FF
        
        
        ' Color de fuente claro
        Form8.Label7.ForeColor = &H80000012
        Form1.Label1.ForeColor = &H80000012
        Form1.Label2.ForeColor = &H80000012
        Form1.Label3.ForeColor = &H80000012
        Form1.Label4.ForeColor = &H80000012
        Form1.Label5.ForeColor = &H80000012
        Form1.Label6.ForeColor = &H80000012
         Form8.Text1.ForeColor = &H80000012
         Form8.Text2.ForeColor = &H80000012
         Form8.Text3.ForeColor = &H80000012
        
        '
              Form1.BackColor = &H400040
              Form8.BackColor = &H400040
        
        '
        Form8.Option1.BackColor = &HFFC0FF
        Form8.Option2.BackColor = &HFFC0FF
        Form8.Option3.BackColor = &HFFC0FF
        Form8.Option4.BackColor = &HFFC0FF
        Form8.Combo1.BackColor = &HFFC0FF
        
    Case "Matrix"
        Form1.Label1.BackColor = &H80FF80
        Form1.Label2.BackColor = &H80FF80
        Form1.Label3.BackColor = &H80FF80
        Form1.Label4.BackColor = &H80FF80
        Form1.Label5.BackColor = &H80FF80
        Form1.Label6.BackColor = &H80FF80
        
        Form8.Label5.BackColor = &H80FF80
        Form8.Label7.BackColor = &H80FF80
        Form8.Label8.BackColor = &H80FF80
        Form8.Label9.BackColor = &H80FF80
        
        Form8.Frame1.BackColor = &H80FF80
        Form8.Frame2.BackColor = &H80FF80
        
        Form8.Label7.ForeColor = &H0&
        
        Form1.Label1.ForeColor = &H0&
        Form1.Label2.ForeColor = &H0&
        Form1.Label3.ForeColor = &H0&
        Form1.Label4.ForeColor = &H0&
        Form1.Label5.ForeColor = &H0&
        Form1.Label6.ForeColor = &H0&
        
        Form1.BackColor = &H404040
        Form8.BackColor = &H404040
    Case "Fantasía"
        Form1.Label1.BackColor = &HC000&
        Form1.Label2.BackColor = &HC000&
        Form1.Label3.BackColor = &HC000&
        Form1.Label4.BackColor = &HC000&
        Form1.Label5.BackColor = &HC000&
        Form1.Label6.BackColor = &HC000&
        
        Form8.Label5.BackColor = &H80FF&
        Form8.Label7.BackColor = &H80FF&
        Form8.Label8.BackColor = &H80FF&
        Form8.Label9.BackColor = &H80FF&
        
        Form8.Frame1.BackColor = &H80FF&
        Form8.Frame2.BackColor = &H80FF&
        
        Form8.Label7.ForeColor = &HFFFFFF
        
        Form1.Label1.ForeColor = &HFFFFFF
        Form1.Label2.ForeColor = &HFFFFFF
        Form1.Label3.ForeColor = &HFFFFFF
        Form1.Label4.ForeColor = &HFFFFFF
        Form1.Label5.ForeColor = &HFFFFFF
        Form1.Label6.ForeColor = &HFFFFFF
        
        Form1.BackColor = &H800080
        Form8.BackColor = &H800080
    Case Else
        MsgBox "Opción no reconocida"
End Select

End Sub
