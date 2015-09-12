VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form48 
   Caption         =   " "
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form48"
   ScaleHeight     =   3975
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Salir"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Desactivar"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activar"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   56033281
      CurrentDate     =   40853
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form48"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Conexion.Execute "delete from inventario"
rs.Open "select *from Inventario", Conexion
If rs.EOF Then
    Fecha = Format(DTPicker1.value, "dd/mm/yyyy")
    Conexion.Execute "Insert into inventario(cod_inventario,fechaInicio,Activo) values(0,'" & Fecha & "',true)"
    Conexion.Execute "update productos set verif=false"
    MsgBox "Activado"
End If
    
End Sub

Private Sub Command2_Click()
Conexion.Execute "delete from inventario"
 Conexion.Execute "update productos set verif=false"

MsgBox "Desactivado"
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Formularios.ColorLabels ColorLetras, Me
    
    DTPicker1.value = Date
    Dim rs As New ADODB.Recordset
    rs.Open "select *from Inventario", Conexion
    If Not rs.EOF Then
        Label2.Caption = "Inventario activo desde el " & rs("fechaInicio")
    Else
        Label2.Caption = ""
    End If
End Sub
