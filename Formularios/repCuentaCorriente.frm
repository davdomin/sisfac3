VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form42 
   Caption         =   "Reporte de Operaciones"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form42"
   ScaleHeight     =   3900
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   3720
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   8
      Top             =   0
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   390
      ItemData        =   "repCuentaCorriente.frx":0000
      Left            =   840
      List            =   "repCuentaCorriente.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2040
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   112984065
      CurrentDate     =   40796
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   112984065
      CurrentDate     =   40796
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Operación"
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function filtro()
    Dim f As String
    Dim f1 As String
    Dim f2 As String
    
    f1 = "#" & Format(DTPicker1.value, "mm/dd/yyyy") & "#"
    f2 = "#" & Format(DTPicker2.value, "mm/dd/yyyy") & "#"
    
    f = "{movCuenta.fecha}>=" & f1 & " and {movCuenta.fecha}<=" & f2
    If Combo1.ListIndex <> 0 Then
        f = f & "and {movCuenta.tipMovimiento}='" & Combo1.Text & "'"
    End If
    filtro = f
    
End Function
Private Sub Command1_Click()
    Dim f As String
    Dim Archivo As String
    Dim Formulas() As String
    Dim cont As Integer
    cont = 0
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "desde='" & Format(DTPicker1.value, "dd/mm/yyyy") & "'"
    cont = cont + 1
    
    ReDim Preserve Formulas(cont)
    Formulas(cont) = "hasta='" & Format(DTPicker2.value, "dd/mm/yyyy") & "'"
    cont = cont + 1
    
    ReDim Preserve Formulas(cont)

    If Combo1.ListIndex <> 0 Then
        Formulas(cont) = "operaciones='Todas'"
    Else
        Formulas(cont) = "operaciones='" & Combo1.Text & "'"
    End If
    cont = cont + 1

    
    
    f = filtro
    Archivo = App.Path & "\reportes\rptCuenta.rpt"
    
    
    Datos.CargarReporte filtro, Archivo, Formulas

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub
