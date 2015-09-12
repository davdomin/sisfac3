VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos terminados"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2880
      TabIndex        =   71
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2940
      TabIndex        =   70
      Text            =   "9"
      Top             =   3840
      Width           =   3735
   End
   Begin VB.ComboBox Text13 
      Height          =   330
      Left            =   2880
      TabIndex        =   69
      Text            =   "Combo3"
      Top             =   4800
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Inventario"
      Height          =   1575
      Left            =   8880
      TabIndex        =   62
      Top             =   3480
      Width           =   3855
      Begin VB.CommandButton Command19 
         Caption         =   "Verificar"
         Height          =   255
         Left            =   840
         TabIndex        =   68
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         Height          =   315
         Left            =   2040
         TabIndex        =   66
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label33 
         Height          =   255
         Left            =   840
         TabIndex        =   67
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Existencia:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -240
         TabIndex        =   65
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2040
         TabIndex        =   64
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total de Ventas:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -240
         TabIndex        =   63
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   32
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   8295
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   8295
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   8295
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   29
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   8295
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   28
      ToolTipText     =   "Haga click Aqui para volver al menu prinicpal"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   27
      ToolTipText     =   "Haga Cliick Aqui para buscar un registro"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      ToolTipText     =   "Haga Click aqui para borra definitivamente este registrop"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Modificar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      ToolTipText     =   "Haga Clcik Aqui para cambiiar los valores de este registro"
      Top             =   7740
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   24
      ToolTipText     =   "Haga Click aqui para deshacer el registro actual"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      ToolTipText     =   "Haga Click Aqui para guardar los cambios en este registro"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   22
      ToolTipText     =   "Haga Click Aqui para Agregar un Nuevo Registro"
      Top             =   7740
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      TabIndex        =   21
      Top             =   -600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8280
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5640
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11160
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Catalogo1 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      ToolTipText     =   "Permite Buscar en el catalogo"
      Top             =   -1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Materiales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   -1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   17
      Top             =   -360
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command13"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   -360
      TabIndex        =   16
      Top             =   -240
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "productos.frx":0000
      Left            =   -2280
      List            =   "productos.frx":000A
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6420
      TabIndex        =   8
      Text            =   "8"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -1560
      Picture         =   "productos.frx":0016
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Calculadora de Precios"
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Producto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   -480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Servicios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   -720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7320
      Picture         =   "productos.frx":0C58
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Permite Buscar en el catalogo"
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2940
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   -1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2940
      TabIndex        =   6
      Text            =   "7"
      Top             =   2880
      Width           =   4335
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Codigo de Barras"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   360
      Picture         =   "productos.frx":0FFA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label25"
      Height          =   255
      Left            =   11880
      TabIndex        =   61
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Ventas Generales:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   60
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label25"
      Height          =   255
      Left            =   11880
      TabIndex        =   59
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Ventas este año:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   58
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label25"
      Height          =   255
      Left            =   11880
      TabIndex        =   57
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Ventas Mensuales:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9120
      TabIndex        =   56
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   55
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Producto(Familia):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   54
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Normal:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   53
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Maximo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6360
      TabIndex        =   52
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Minimo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   51
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Actual:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   50
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   49
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   48
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3720
      TabIndex        =   47
      Top             =   -960
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2160
      TabIndex        =   46
      Top             =   -600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Se Fabrica"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -3360
      TabIndex        =   45
      Top             =   5880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Marca o Procedencia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Medida:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   43
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Con IGV:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   42
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   41
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Especial:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   40
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio Compra:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   39
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo de Barras:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   38
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   36
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Con IGV:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4380
      TabIndex        =   35
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Presione F2 para rebajar el IGV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   34
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Presione F2 para rebajar el IGV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3000
      TabIndex        =   33
      Top             =   2520
      Width           =   4095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tabla As String
Public campoClave As String
Dim modificar As Boolean
Sub BuscarBarra(cod As String)
    Dim codProducto As Integer

    codProducto = Val(Datos.MostrarCampo("Productos", "codProducto", "barras='" & cod & "'"))
    If codProducto = 0 Then
        
        MsgBox "Producto no encontrado"
    Else
        Text1.Text = codProducto
        Formularios.Buscar Me
    End If
        
    
    
    
End Sub

Function getTipo() As String
    If Option1.value Then
        getTipo = "Producto"
    Else
        getTipo = "Servicio"
    End If
End Function
Function setTipo() As String
    If getTipo = "Producto" Then
        Option1.value = True
    Else
        Option2.value = True
    End If
End Function
Sub BloqueoServicios(st As Boolean)
    Text3.Enabled = st
    Text4.Enabled = st
    Text5.Enabled = st
    Label3.Enabled = st
    Label4.Enabled = st
    
    Label5.Enabled = st
    
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub

Sub CargarTablas()
    Tabla = "Productos"
    campoClave = "codProducto"
End Sub

Function SqlActualizacion()
    Dim iSql As String
    Dim rs As New ADODB.Recordset
    rs.Open "Select *from [" & Tabla & "] where [" & campoClave & "] =" & Val(Text1.Text), Conexion
    If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codProducto,Descripcion,Stock,StockMin," _
        & "StockMax," _
        & "Precio," _
        & "SeFabrica," _
        & "Talla," _
        & "Color," _
        & "PrecioPaga," _
        & "Tipo,Costo,barras,almacen," _
        & "CodTipoP)" _
        & " values(" _
        & "" & Text1.Text & "," _
        & "'" & Text2.Text & "'," _
        & "" & Val(Text3.Text) & "," _
        & "" & Val(Text4.Text) & "," _
        & "" & Val(Text5.Text) & "," _
        & "" & Val(Text6.Text) & "," _
        & "'" & Combo1.Text & "'," _
        & "'" & Text8.Text & "'," _
        & "'" & Text9.Text & "'," _
        & "" & Val(Text10.Text) & "," _
        & "'" & Me.getTipo & "'," & Val(Text11.Text) & "," _
        & "'" & Text12.Text & "'," _
        & "'" & Text13.Text & "'," _
        & "" & Val(Text7.Text) & ")"
       ' A = InputBox(sql, iSql, iSql)
    Else
        iSql = "UPDATE [" & Tabla & "] set " _
            & "Descripcion='" & Text2.Text & "'," _
            & "Stock=" & Text3.Text & "," _
            & "StockMin=" & Text4.Text & "," _
            & "StockMax=" & Text5.Text & "," _
            & "Precio=" & Text6.Text & "," _
            & "PrecioPaga=" & Val(Text10.Text) & "," _
            & "Talla='" & Text8.Text & "'," _
            & "Color ='" & Text9.Text & "'," _
            & "Costo ='" & Val(Text11.Text) & "'," _
            & "SeFabrica='" & Combo1.Text & "'," _
            & "codTipoP=" & Text7.Text & ", " _
            & "barras='" & Text12.Text & "', " _
            & "almacen='" & Text13.Text & "', " _
            & "Tipo='" & Me.getTipo & "' " _
            & "Where codProducto=" & Text1.Text & ""
    End If
    'MsgBox iSql
    SqlActualizacion = iSql
    
End Function

Sub bloquear(es As Boolean)
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Text9.Enabled = es
    Text12.Enabled = es
    Text13.Enabled = es
    Text10.Enabled = es
    Text11.Enabled = es
    Label1.Enabled = es
    Label2.Enabled = es
    Label3.Enabled = es
    Label4.Enabled = es
    Label5.Enabled = es
    Label6.Enabled = es
    Label7.Enabled = es
    Label9.Enabled = es
    Label10.Enabled = es
    Label10.Enabled = es
    Label11.Enabled = es
    Label12.Enabled = es
    Label13.Enabled = es
    Label14.Enabled = es
    Label15.Enabled = es
    Label16.Enabled = es
    Label17.Enabled = es
    Label18.Enabled = es
    Label19.Enabled = es
    Label20.Enabled = es
    Label21.Enabled = es
    Label22.Enabled = es
    
    Text10.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Catalogo1.Enabled = es
    Command12.Enabled = es
    Command15.Enabled = es
    Option1.value = True
    Option1.Enabled = es
    Option2.Enabled = es
    Command16.Enabled = es
    Command17.Enabled = es
    'Command19.Enabled = es
    If Option2.value Then
        BloqueoServicios False
    End If
    modificar = False
    'Frame1.Enabled = False

End Sub
Sub MostrarInventario(codProducto As Integer)
    Dim rs As New ADODB.Recordset
    
    rs.Open "select *from inventario", Conexion
    If rs.EOF Then
        Frame1.Visible = False
    Else
        Dim f1 As String
        Frame1.Visible = True
        f1 = "#" & Format(rs("fechaInicio"), "mm/dd/yyyy") & "#"
        Dim sql As String
        sql = "select count(*) from PedidoEnc,PedidoDet where PedidoEnc.codPedido = PedidoDet.codPedido and Fecha>=" & f1 & " and codProducto=" & codProducto
        Dim rs2 As New ADODB.Recordset
        rs2.Open sql, Conexion
        If Not rs2.EOF Then
            Label31.Caption = rs2(0)
        End If
    End If
        
End Sub
Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    
    Combo1.Text = "No"
    If Not Produccion Then Combo1.Text = "No"
    Option1.value = False
    Option2.value = True
    
    Label14.Caption = ""
    Label19.Caption = ""
    Label25.Caption = ""
    Label27.Caption = ""
    Label29.Caption = ""
    
End Sub
Sub Mostrar(rs As ADODB.Recordset)
On Error Resume Next
    Text1.Text = rs(0)
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = rs(3)
    Text5.Text = rs(4)
    Text6.Text = rs(5)
    Text7.Text = rs(6) 'Tipo de Producto
    Combo1.Text = rs(7)
    Text8.Text = rs(8)
    Text9.Text = rs(9)
    Text10.Text = rs(11)
    Text11.Text = rs(12)
    Text12.Text = rs("barras")
    Text13.Text = rs("almacen")
    
    mostrarEstadisticas Val(Text1.Text)
    If rs("Verif") Then
        Command19.Enabled = False
        Label33.Caption = "Verificado"
    Else
        Command19.Enabled = True
        Label33.Caption = "No verificado"
    End If

    MostrarInventario Val(Text1.Text)


End Sub


Private Sub Catalogo1_Click()
    MostrarCatalogo "select *from TipoP"
    Text7.Text = Catalogo.Resultado
End Sub

Private Sub Combo2_Click()
    Dim Descripcion As String
    Descripcion = Combo2.Text
    Text7.Text = Datos.MostrarCampo("TipoP", "codTipoP", "descripcion='" & Descripcion & "'")

End Sub

Private Sub Command1_Click()
    Formularios.Nuevo Me
End Sub


Private Sub Command12_Click()
    Form11.codProducto = Val(Text1.Text)
    Form11.Show vbModal
    
End Sub

Private Sub Command13_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Command14_Click()
    SendKeys "{TAB}"
End Sub

Private Sub Command15_Click()
    Form19.Show vbModal
End Sub

Private Sub Command16_Click()
    
    Form6.Show vbModal
    llenarCombo "select descripcion from TipoP order by descripcion", Combo2
End Sub

Private Sub Command17_Click()
    Form37.Show vbModal
    'Text11.Text = Form37.Resultado
End Sub

Private Sub Command18_Click()
Dim B As String
B = (InputBox("Int. Codigo de barras", "Buscar"))
Me.BuscarBarra B

End Sub

Private Sub Command19_Click()
    Conexion.Execute "update productos set stock=" & Val(Text14.Text) - Val(Label31.Caption) & ", verif=" & 1 & " where codProducto=" & Val(Text1.Text)
    MsgBox "Verificado"
    Command19.Enabled = False
    
End Sub

Private Sub Command2_Click()
    Formularios.Guardar Me
   ' Command1_Click
End Sub

Private Sub Command3_Click()
    Formularios.cancelar Me
End Sub

Private Sub Command4_Click()
    Formularios.modificar Me
    modificar = True
End Sub

Private Sub Command5_Click()
    Formularios.Eliminar Me
End Sub

Private Sub Command6_Click()
    MostrarCatalogo "select *from [ProductosOrdenados]"
    Text1.Text = Catalogo.Resultado
    Formularios.Buscar Me
End Sub

Private Sub Command7_Click()
    Unload Me
End Sub
Private Sub Command8_Click()
    Formularios.Primero Me
End Sub

Private Sub Command9_Click()
    Formularios.Anterior Me
End Sub
Private Sub Command10_Click()
    Formularios.Siguiente Me
End Sub

Private Sub Command11_Click()
    Formularios.Ultimo Me
End Sub
Sub mostrarTipoP()
    Dim codTipoP As Integer
    codTipoP = Val(Text7.Text)
    Combo2.Text = Datos.MostrarCampo("TipoP", "Descripcion", "codTipoP=" & codTipoP)
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    
    rs.Open "select *from inventario", Conexion
    If rs.EOF Then
        Frame1.Visible = False
    End If
    
    Formularios.ColorLabels ColorLetras, Me
    llenarCombo "select descripcion from TipoP order by descripcion", Combo2
    llenarCombo "select monedas from monedas", Text13
    Formularios.Botones Me, 1
    bloquear False
    Limpiar
    CargarTablas
End Sub



Private Sub Option1_Click()
'    BloqueoServicios True
End Sub

Private Sub Option2_Click()
    BloqueoServicios False
End Sub

Private Sub Text10_Change()
    Label19.Caption = Round((Val(Text10.Text) * Proyecto.PIVa) + Val(Text10.Text), 2)

End Sub

Private Sub Text10_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 113) Then
    Dim sinIGV As Double
        sinIGV = sinIva(Val(Text10.Text))
        Label19.Caption = Text10.Text
        Text10.Text = sinIGV
    End If

End Sub

Private Sub Text2_Change()
    Datos.AutoCompletar_TextBox Text2
End Sub

Private Sub Text2_GotFocus()
    Datos.CargarValores "select *from productos order by descripcion"
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text1.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    Dim codProducto As Integer
    codProducto = Val(Datos.MostrarCampo("productos", "codProducto", "descripcion='" & Text2.Text & "'"))
    If codProducto <> 0 And Not modificar Then
        Text1.Text = codProducto
        Formularios.Buscar Me
    End If
End Sub

Private Sub Text6_Change()
    Label14.Caption = Round((Val(Text6.Text) * Proyecto.PIVa) + Val(Text6.Text), 2)
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 113) Then
    Dim sinIGV As Double
        sinIGV = sinIva(Val(Text6.Text))
        Label14.Caption = Text6.Text
        Text6.Text = sinIGV
    End If
End Sub

Private Sub Text7_Change()
mostrarTipoP
End Sub
Sub mostrarEstadisticas(codProducto As Integer)
    Dim t As New ADODB.Recordset
    Dim sql As String
    Dim Fecha As String
    
    Fecha = Format(Date, "mm/yyyy")
    sql = "SELECT count(*)from pedidoEnc,PedidoDet where PedidoEnc.codPedido = pedidoDet.codPedido and status<>'Anulado' And format(fecha,'mm/yyyy') = '" & Fecha & "' and codProducto=" & codProducto
    t.Open sql, Conexion
    If IsNumeric(t(0)) Then
        Label25.Caption = t(0)
    Else
        Label25.Caption = 0
    End If
    t.Close
    
    Fecha = Format(Date, "yyyy")
    sql = "SELECT count(*)from pedidoEnc,PedidoDet where PedidoEnc.codPedido = pedidoDet.codPedido and status<>'Anulado' And format(fecha,'yyyy') = '" & Fecha & "' and codProducto=" & codProducto
    t.Open sql, Conexion
    If IsNumeric(t(0)) Then
        Label27.Caption = t(0)
    Else
        Label27.Caption = 0
    End If
    t.Close
    
    sql = "SELECT count(*)from pedidoEnc,PedidoDet where PedidoEnc.codPedido = pedidoDet.codPedido and status<>'Anulado' and codProducto=" & codProducto
    t.Open sql, Conexion
    If IsNumeric(t(0)) Then
        Label29.Caption = t(0)
    Else
        Label29.Caption = 0
    End If
    t.Close
    
    
    
    
End Sub

