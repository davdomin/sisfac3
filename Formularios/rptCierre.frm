VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form45 
   Caption         =   "Reporte de cierre de caja"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form45"
   ScaleHeight     =   2670
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox CrystalReport1 
      Height          =   480
      Left            =   1680
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   6
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   112984065
      CurrentDate     =   40796
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      Format          =   112984065
      CurrentDate     =   40796
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form45"
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
    
    f = "{cierre.fecha}>=" & f1 & " and {cierre.fecha}<=" & f2
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

    cont = cont + 1

    
    
    
    Archivo = App.Path & "\reportes\rptCierre.rpt"
    f = filtro
    
    Datos.CargarReporte f, Archivo, Formulas


End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub
