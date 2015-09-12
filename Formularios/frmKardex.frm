VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmKardex 
   Caption         =   "KARDEX - Reporte de Kardex por articulo"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   3570
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   3900
      TabIndex        =   7
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdReporte 
      Caption         =   "Ver Reporte"
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   3060
      Width           =   1275
   End
   Begin VB.ComboBox cmbArticulo 
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      Text            =   "cmbArticulo"
      Top             =   1980
      Width           =   4155
   End
   Begin MSComCtl2.DTPicker dtHasta 
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      Format          =   125435905
      CurrentDate     =   41929
   End
   Begin MSComCtl2.DTPicker dtDesde 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   125435905
      CurrentDate     =   41929
   End
   Begin VB.Label lblProducto 
      Caption         =   "Articulo"
      Height          =   255
      Left            =   1020
      TabIndex        =   4
      Top             =   1680
      Width           =   1035
   End
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDesde 
      Caption         =   "Desde"
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Formulas() As String
Function filtroCardex() As String
Dim f As String
Dim fechaDesde As String
Dim fechaHasta As String
    fechaDesde = "#" & Format(dtDesde.value, "mm/dd/yyyy") & "#"
    fechaHasta = "#" & Format(dtHasta.value, "mm/dd/yyyy") & "#"
    f = "{kardex.fecha}>=" & fechaDesde & " and {kardex.fecha}<=" & fechaHasta
    If (cmbArticulo.Text <> "") Then
        f = " AND {productos.descripcion} ='" & cmbArticulo.Text & "'"
    End If
    
    filtroCardex = f
End Function
Private Sub cmdReporte_Click()
Dim f As String
Dim Archivo As String
    f = filtroCardex
    Archivo = App.Path & "\reportes\kardex.rpt"
    ReDim Formulas(3)
    Formulas(0) = "desde=" & Format(dtDesde.value, "dd/mm/yyyy")
    Formulas(1) = "hasta=" & Format(dtHasta.value, "dd/mm/yyyy")
    Datos.CargarReporte f, Archivo, Formulas
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Datos.llenarCombo "SELECT descripcion FROM productos", cmbArticulo
    dtDesde.value = Now
    dtHasta.value = Now
End Sub
