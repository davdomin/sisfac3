VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00E4B8B8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos"
   ClientHeight    =   10740
   ClientLeft      =   -1905
   ClientTop       =   270
   ClientWidth     =   18915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   18915
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command25 
      Caption         =   "Cambiar a"
      Height          =   615
      Left            =   13320
      TabIndex        =   129
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox Text33 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13800
      TabIndex        =   128
      Text            =   "Text29"
      Top             =   4200
      Width           =   975
   End
   Begin VB.ComboBox Combo6 
      Height          =   330
      ItemData        =   "pedidos.frx":0000
      Left            =   9960
      List            =   "pedidos.frx":0002
      TabIndex        =   125
      Text            =   "Combo6"
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text32 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   7680
      TabIndex        =   122
      Top             =   260
      Width           =   1575
   End
   Begin VB.TextBox Text31 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   5760
      TabIndex        =   120
      Top             =   260
      Width           =   1575
   End
   Begin VB.TextBox Text30 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9480
      TabIndex        =   117
      Text            =   "Text29"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text29 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13800
      TabIndex        =   116
      Text            =   "Text29"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Datos de la guia de Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   99
      ToolTipText     =   "Haga doble click para expandir la pantalla"
      Top             =   7440
      Width           =   7335
      Begin VB.TextBox Text28 
         Height          =   315
         Left            =   3840
         TabIndex        =   115
         Top             =   1740
         Width           =   3315
      End
      Begin VB.TextBox Text27 
         Height          =   315
         Left            =   120
         TabIndex        =   113
         Top             =   1740
         Width           =   3135
      End
      Begin VB.TextBox Text26 
         Height          =   315
         Left            =   5100
         TabIndex        =   111
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text25 
         Height          =   315
         Left            =   960
         TabIndex        =   109
         Top             =   840
         Width           =   6195
      End
      Begin VB.TextBox Text24 
         Height          =   315
         Left            =   960
         TabIndex        =   107
         Top             =   1200
         Width           =   3315
      End
      Begin VB.TextBox Text23 
         Height          =   315
         Left            =   5460
         TabIndex        =   105
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtTranportista 
         Height          =   315
         Left            =   3060
         TabIndex        =   103
         Top             =   480
         Width           =   2235
      End
      Begin VB.TextBox Text21 
         Height          =   315
         Left            =   120
         TabIndex        =   101
         Top             =   480
         Width           =   2835
      End
      Begin VB.Label Label50 
         BackStyle       =   0  'Transparent
         Caption         =   "Llegada"
         Height          =   255
         Left            =   3840
         TabIndex        =   114
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "Partida"
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   "Placa:"
         Height          =   255
         Left            =   4500
         TabIndex        =   110
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio"
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "RUC"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Destino"
         Height          =   255
         Left            =   5460
         TabIndex        =   104
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "Tranportista"
         Height          =   255
         Left            =   3060
         TabIndex        =   102
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Señores"
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   96
      Top             =   2040
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8760
      Top             =   10080
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command24 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      Picture         =   "pedidos.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List8 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   17160
      TabIndex        =   94
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   10320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13080
      Picture         =   "pedidos.frx":0C16
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14400
      Picture         =   "pedidos.frx":1828
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   840
   End
   Begin VB.ListBox List5 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   11160
      TabIndex        =   82
      Top             =   4920
      Width           =   1935
   End
   Begin VB.ComboBox Text9 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7560
      TabIndex        =   4
      Text            =   "Text9"
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox Text19 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   -7680
      TabIndex        =   12
      Text            =   "Text19"
      Top             =   5040
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   72
      Tag             =   "Si no esta activado no se cobrara el IGV a este producto"
      Top             =   9000
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      Picture         =   "pedidos.frx":254E
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text18 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9720
      TabIndex        =   11
      Top             =   -960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7200
      TabIndex        =   10
      Top             =   -960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      TabIndex        =   9
      Top             =   -960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1920
      TabIndex        =   8
      Text            =   "Combo2"
      Top             =   -960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Anular"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   9720
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   8535
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   7575
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9000
      TabIndex        =   7
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   4560
      Width           =   6735
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6840
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   120
      TabIndex        =   49
      Top             =   4920
      Width           =   6735
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   6840
      TabIndex        =   48
      Top             =   4920
      Width           =   735
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   7560
      TabIndex        =   47
      Top             =   4920
      Width           =   1695
   End
   Begin VB.ListBox List4 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   9240
      TabIndex        =   46
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   8760
      Width           =   1575
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   8400
      Width           =   2055
   End
   Begin VB.TextBox Text15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   8760
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11160
      Picture         =   "pedidos.frx":28F0
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cargar Pedido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   9720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Guia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9720
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      Picture         =   "pedidos.frx":2C02
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Pagar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">|"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Haga Click Aqui para ir al ultimo Registro"
      Top             =   10320
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Haga Click Aqui para ir al siguiente Registro"
      Top             =   10320
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Haga Click Aqui para ir al anterior  Registro"
      Top             =   10320
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "|<"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Haga Click Aqui para ir al primer Registro"
      Top             =   10320
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Boleta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9720
      Width           =   975
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   -2520
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   75
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "88/88/8888"
      Top             =   75
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   12480
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   320
      Width           =   1815
   End
   Begin VB.PictureBox Crystal 
      Height          =   480
      Left            =   -4440
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   130
      Top             =   8280
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   12720
      TabIndex        =   45
      Top             =   7800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   76480513
      CurrentDate     =   40283
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   360
      TabIndex        =   76
      Text            =   "Combo1"
      Top             =   6360
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ListBox List6 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   7560
      TabIndex        =   92
      Top             =   4920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox Combo5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   5400
      TabIndex        =   5
      Text            =   "Combo5"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox List7 
      BackColor       =   &H00FAF5D3&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   9240
      TabIndex        =   93
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F0F0F0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label56 
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   127
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label55 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12720
      TabIndex        =   126
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pago: "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   124
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label53 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Factura"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7680
      TabIndex        =   123
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Boleta:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5760
      TabIndex        =   121
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label51"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7620
      TabIndex        =   119
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9480
      TabIndex        =   118
      Top             =   0
      Width           =   660
   End
   Begin VB.Label Label42 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   97
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   89
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   15360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label40 
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12480
      TabIndex        =   86
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   85
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Existencia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   84
      Top             =   3120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label32"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   83
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   81
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label34 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label32"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   80
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Marca/Procedencia:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   79
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label32"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   78
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Medida:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -2880
      TabIndex        =   77
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -960
      TabIndex        =   75
      Tag             =   "C"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8280
      TabIndex        =   73
      Top             =   2160
      Width           =   6015
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Año :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   70
      Top             =   -960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   69
      Top             =   -960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   68
      Top             =   -960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Placa:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   67
      Top             =   -960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "xxxx"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   9000
      TabIndex        =   65
      Top             =   3120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Index           =   4
      X1              =   -120
      X2              =   15120
      Y1              =   9600
      Y2              =   9600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   15360
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DNI o RUC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11280
      TabIndex        =   64
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Razon Social:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   63
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   62
      Top             =   1335
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   61
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Barras :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -7680
      TabIndex        =   60
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   59
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Base Imp."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   58
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   57
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Entrega"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -3240
      TabIndex        =   56
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   55
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "IGV"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   54
      Top             =   8400
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   53
      Top             =   8760
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Abono"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   52
      Top             =   8400
      Width           =   1815
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Restan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   51
      Top             =   8760
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   15480
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Factura:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -3600
      TabIndex        =   24
      Top             =   315
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Pago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -9480
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Trabajo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -5760
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label19 
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   -7680
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label18 
      Caption         =   "Label18"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00B50941&
      Height          =   255
      Left            =   -3720
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   17
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   13
      Top             =   315
      Width           =   1215
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim montoProductos As Double
Public Tabla As String
Public campoClave As String
Dim esNueva As Boolean
Function obtNuevoNumeroFactura()
Dim ultimoNumero As Long
Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM factura ORDER BY id DESC", Conexion
    If Not rs.EOF Then
        ultimoNumero = rs("nfac")
        obtNuevoNumeroFactura = ultimoNumero + 1
    Else
        obtNuevoNumeroFactura = 1
    End If
End Function
Function obtNuevoNumeroBoleta()
Dim rs As New ADODB.Recordset
Dim ultimoNumero As Long
    rs.Open "SELECT * FROM boleta ORDER BY id DESC", Conexion
    If Not rs.EOF Then
        ultimoNumero = rs("nbol")
        obtNuevoNumeroBoleta = ultimoNumero + 1
    Else
        obtNuevoNumeroBoleta = 1
    End If
End Function
Function obtNuevoNumeroGuia()
Dim rs As New ADODB.Recordset
Dim ultimoNumero As Long
    rs.Open "SELECT * FROM guias ORDER BY codGuia DESC", Conexion
    If Not rs.EOF Then
        ultimoNumero = rs("numeroGuia")
        obtNuevoNumeroGuia = ultimoNumero + 1
    Else
        obtNuevoNumeroGuia = 1
    End If
End Function
Function ObtMoneda()
    If Text29.Text = "$" Then
        ObtMoneda = " DOLARES AMERICANOS"
    Else
        ObtMoneda = " NUEVOS SOLES"
    End If
    
End Function
Sub colocarTamaño()
On Error Resume Next
Dim variacion As Integer
Dim Tamaño As Double
    variacion = 91005
    Tamaño = 8000
    DataGrid1.Columns(1).Width = Tamaño
    DataGrid1.Columns(0).Visible = False
    'Text1.Top = Me.ScaleHeight - variacion
    'Text2.Top = Text1.Top + Text1.Height + 1
    'Command1.Top = Me.ScaleHeight - variacion
    DataGrid1.Width = Me.ScaleWidth - 500
    DataGrid1.Height = 2000
End Sub
Sub mostrarGrilla()
Dim sql As String
    If Proyecto.inventario Then
        sql = "SELECT * FROM productosOrdenados2"
    Else
        sql = "SELECT * FROM productosOrdenados2"
    End If
    If Combo1.Text <> "" Then
        sql = sql & " WHERE  [descripcion articulo]like'%" & Combo1.Text & "%'"
    End If
    Adodc1.ConnectionString = Conexion.ConnectionString
    Adodc1.RecordSource = sql
    Adodc1.Refresh
    Set DataGrid1.datasource = Adodc1
    colocarTamaño
End Sub
Sub Bloqueo(st As Byte)
    If st = 1 Then 'Inicial
        Command1.Enabled = True 'Nuevo
        Command2.Enabled = False 'Guardar
        Command3.Enabled = True 'Cancelar
        Command9.Enabled = False 'Pagar
        Command6.Enabled = True 'Cargar Pedido
        Command14.Enabled = False 'Cargar Pedido
        Command15.Enabled = False 'Factura
        Command16.Enabled = False 'Anular
        Command7.Enabled = False 'Imprimir
        Command22.Enabled = False 'Factura
        Command25.Enabled = True 'Factura
    End If
    If st = 2 Then 'Edicion
        Command1.Enabled = False 'Nuevo
        Command2.Enabled = True 'Guardar
        Command3.Enabled = True 'Cancelar
        Command9.Enabled = False 'Pagar
        Command6.Enabled = True 'Cargar Pedido
        Command14.Enabled = False 'Cargar Pedido
        Command15.Enabled = True 'Factura
        Command16.Enabled = False 'Anular
        Command7.Enabled = True 'Imprimir
        Command22.Enabled = True 'Factura
        Command25.Enabled = True 'Factura
    End If
    If st = 3 Then 'Mostrar
        Command1.Enabled = True 'Nuevo
        Command2.Enabled = True 'Guardar
        Command3.Enabled = True  'Cancelar
        Command9.Enabled = True 'Pagar
        Command6.Enabled = True 'Cargar Pedido
        Command14.Enabled = True 'Cargar Pedido
        Command15.Enabled = True 'Factura
        Command16.Enabled = True 'Anular
        Command7.Enabled = True 'Imprimir
        Command22.Enabled = True 'Factura
        Command25.Enabled = False 'Factura
        If Label18.Caption = "Anulado" Then
            Command1.Enabled = True 'Nuevo
            Command2.Enabled = False 'Guardar
            Command3.Enabled = True 'Cancelar
            Command9.Enabled = False 'Pagar
            Command6.Enabled = True 'Cargar Pedido
            Command14.Enabled = False 'Cargar Pedido
            Command22.Enabled = False
            Command15.Enabled = False 'Factura
            Command16.Enabled = False 'Anular
            Command7.Enabled = False 'Imprimir
            Command25.Enabled = False 'Factura
        End If
    End If
        

End Sub
Sub ActivarBotones()
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command9.Enabled = True
    Command6.Enabled = True
    Command14.Enabled = True
    Command7.Enabled = True
    Command5.Enabled = True
End Sub
Sub CamposModificables()
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    List1.Enabled = True
    List2.Enabled = True
    List3.Enabled = True
    List4.Enabled = True
    List5.Enabled = True
    Combo1.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Combo5.Enabled = True
    Command5.Enabled = True
End Sub
Sub CargarTablas()
    Tabla = "pedidoEnc"
End Sub
Sub cancelar()
    Limpiar
    bloquear False
End Sub
Sub Nuevo()
    Limpiar
    esNueva = True
    Text1.Text = Datos.generarCodigo("numerosFacturas", "numeroFactura")
    Label18.Caption = "Pendiente"
    Label19.Caption = "Abierta"
    bloquear True
    Text5.SetFocus
    Text29.Text = "$"
    Totales
    Text32.Text = obtNuevoNumeroFactura
    Text31.Text = obtNuevoNumeroBoleta
    If Proyecto.inicialGuia Then
        Text30.Text = obtNuevoNumeroGuia
    End If
    Text33.Text = Me.obtTipoCambioDefault
End Sub
Sub BorrarDetalle(CualDetalle)
Dim codProducto As Long
Dim Cantidad As Double
Dim Monto As Double
Dim total As Integer
Dim esServicio As Boolean
Dim codPedido As Long
    If CualDetalle <> -1 Then
        Monto = Val(List4.List(CualDetalle))
        codPedido = Val(Text1.Text)
        codProducto = Datos.MostrarCampo("Productos", "CodProducto", "Descripcion='" & List1.List(CualDetalle) & "'")
        esServicio = Datos.MostrarCampo("Productos", "Tipo", "Descripcion='" & Combo1.Text & "'") = "Servicio"
        
        Cantidad = Val(List2.List(CualDetalle))
        Conexion.Execute "delete from pedidoDet where codPedido=" & codPedido & " and codproducto =" & codProducto, total
        If total > 0 And Not esServicio Then
            'Conexion.Execute "update Productos set stock = stock +" & Cantidad & " where codProducto=" & codProducto
        End If
        List1.RemoveItem CualDetalle
        List2.RemoveItem CualDetalle
        List3.RemoveItem CualDetalle
        List4.RemoveItem CualDetalle
        List5.RemoveItem CualDetalle
        List6.RemoveItem CualDetalle
        List7.RemoveItem CualDetalle
        List8.RemoveItem CualDetalle
        montoProductos = ObtmontoProductos
        Totales
        CuantoResta
    End If
End Sub
Sub BuscarBarra()
Dim codBarra As String
Dim Descripcion As String
    codBarra = Text19.Text
    Descripcion = Datos.MostrarCampo("Productos", "descripcion", "barras='" & codBarra & "'")
    Combo1.Text = Descripcion
    Combo1_Click
End Sub
Sub Mostrar(rs As ADODB.Recordset)
Dim rd As New ADODB.Recordset
Dim rf As New ADODB.Recordset
    On Error Resume Next
    Limpiar
    Text1.Text = rs(0)
    rf.Open "select *from Factura where codPedido=" & rs(0), Conexion
    If rf.EOF Then
        Text16.Text = "-"
    Else
        Text16.Text = rf("nfac")
    End If
    If rs("tipo") = "b" Then Text31.Text = rs("ncontrol")
    If rs("tipo") = "f" Then Text32.Text = rs("ncontrol")
    Text2.Text = rs(1)
    Text3.Text = rs(2)
    Text4.Text = Datos.MostrarCampo("Clientes", "cedrif", "codCliente=" & rs(3))
    Me.BuscarCliente
    Text11.Text = rs(4): montoProductos = rs(4)
    Text12.Text = rs(5)
    Text13.Text = rs(6)
    Text14.Text = rs(7)
    Text15.Text = rs(8)
    Text29.Text = rs("moneda")
    Label18.Caption = rs(9)
    DTPicker1.value = rs(10)
    Combo6.Text = Datos.MostrarCampo("formaPago", "descripcion", "codFormaPago=" & rs("codFormaPago") & "")
    'Detalle del pedido
    Combo2.Text = Datos.MostrarCampo("Vehiculos", "placa", "codVehiculo=" & rs("codVehiculo"))
    buscarVehiculo
    rd.Open "select *from PedidoDet where CodPedido=" & rs(0) & "", Conexion
    While Not rd.EOF
        List1.AddItem Datos.MostrarCampo("Productos", "Descripcion", "CodProducto=" & rd(1))
        List2.AddItem rd(2)
        List3.AddItem rd(3)
        List4.AddItem rd(4)
        List5.AddItem Datos.MostrarCampo("Productos", "Almacen", "CodProducto=" & rd(1))
        List6.AddItem Round((rd(3) * PIVa) + rd(3), 2)
        List7.AddItem Round((rd(4) * PIVa) + rd(4), 2)
        List8.AddItem Datos.MostrarCampo("Productos", "stock", "CodProducto=" & rd(1))
        rd.MoveNext
    Wend
    bloquear False
    Bloqueo 3
End Sub
Sub SeleccionarListas(n)
    List1.Selected(n) = True
    List2.Selected(n) = True
    List3.Selected(n) = True
    List4.Selected(n) = True
    List5.Selected(n) = True
    List6.Selected(n) = True
    List7.Selected(n) = True
    List8.Selected(n) = True
End Sub
Sub rebajarStock(Cantidad As Double, codProducto As Integer, i As Integer)
Dim iSql As String
Dim esServicio
    If Not Proyecto.inventario Then Exit Sub
    If Datos.MostrarCampo("Productos", "Stock", "codProducto=" & codProducto) = -999 Then Exit Sub
    esServicio = Datos.MostrarCampo("Productos", "Tipo", "codProducto=" & codProducto & "") = "Servicio"
    If Not esServicio Then
        iSql = "update Productos set stock=stock-" & Val(List2.List(i)) & " where codProducto=" & codProducto
        Conexion.Execute iSql
    End If
End Sub
Sub GuardarDetalle(codPedido As Integer)
Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim contFab As Integer
    Dim iSql As String
    Dim codProducto As Integer
    contFab = 0
    For i = 0 To List1.ListCount - 1
        codProducto = Datos.MostrarCampo("Productos", "CodProducto", "Descripcion='" & List1.List(i) & "'")
        rs.Open "select *from pedidoDet where CodPedido=" & codPedido & " and codProducto=" & codProducto, Conexion
        If rs.EOF Then
            iSql = "insert into pedidoDet(CodPedido,CodProducto,Cantidad,Precio,Total,orden) values(" _
            & "" & codPedido & "," _
            & "" & codProducto & "," _
            & "" & List2.List(i) & "," _
            & "" & List3.List(i) & "," _
            & "" & List4.List(i) & "," _
            & "" & i + 1 & ")"
            If Datos.MostrarCampo("Productos", "seFabrica", "Descripcion='" & List1.List(i) & "'") = "No" Then
                rebajarStock Val(List2.List(i)), codProducto, i
            Else
                If (Datos.MostrarCampo("Productos", "Stock", "Descripcion='" & List1.List(i) & "'") >= Val(List2.List(i))) Then
                    rebajarStock Val(List2.List(i)), codProducto, i
                Else
                    contFab = contFab + 1
                End If
            End If
        Else
            iSql = "UPDATE pedidoDet SET " _
            & "Cantidad=" & Val(List2.List(i)) & "," _
            & "Precio=" & Val(List3.List(i)) & "," _
            & "Total=" & Val(List4.List(i)) & " " _
            & "WHERE codPedido=" & codPedido & " " _
            & "AND  codProducto=" & codProducto & " "
        End If
        rs.Close
        Conexion.Execute iSql
    Next
    If contFab = 0 Then
    '    Conexion.Execute "update pedidoEnc set status='Terminado' where codPedido=" & codPedido
     '   Label18.Caption = "Terminado"
    End If
End Sub
Sub guardarPago(abono As Double)
    If abono > 0 Then
        Conexion.Execute "insert into pagos (CodPago,CodPedido,Fecha,Hora,CodUsuario,Monto) values(" _
        & "" & Datos.generarCodigo("Pagos", "CodPago") & "," _
        & "" & Text1.Text & "," _
        & "'" & Date & "'," _
        & "'" & Time & "'," _
        & "" & codUsuario & "," _
        & "" & abono & ")"
    End If
End Sub
Sub guardarGuia()
Dim sql As String
If Not Proyecto.inicialGuia Then Exit Sub
    Conexion.Execute "DELETE * FROM guias WHERE codPedido = " & Val(Text1.Text)
    sql = "insert into guias (codGuia,señores,transporte,destino,ruc,domicilio,placa,partida,llegada,numeroGuia,codPedido) values(" _
    & "" & Datos.generarCodigo("guias", "codGuia") & "," _
    & "'" & Text21.Text & "'," _
    & "'" & txtTranportista.Text & "'," _
    & "'" & Text23.Text & "'," _
    & "'" & Text24.Text & "'," _
    & "'" & Text25.Text & "'," _
    & "'" & Text26.Text & "'," _
    & "'" & Text27.Text & "'," _
    & "'" & Text28.Text & "'," _
    & "'" & Val(Text30.Text) & "'," _
    & "" & Val(Text1.Text) & ")"
    Conexion.Execute sql
End Sub
Sub GuardarEncabezado()
Dim rs As New ADODB.Recordset
Dim AbonoAnterior As Double
Dim iSql As String
    GuardarCliente
    guardarVehiculo
    guardarRelacion
    CuantoResta
'    codVeiculo = Datos.MostrarCampo("Vehiculos", "codVehiculo", "placa='" & Combo2.Text & "'")
    rs.Open "select *from pedidoEnc where codPedido=" & Text1.Text, Conexion
    If rs.EOF Then
        iSql = "insert into pedidoEnc (codPedido,Fecha,Hora,codCliente,subtotal,iva,total,abono,restan" _
        & ",status,fecha_entrega,codUsuario,ivaCobrado,enLetras,Moneda,codFormaPago,tipoCambio,statuspago) values(" _
        & "" & Text1.Text & "," _
        & "'" & Date & "'," _
        & "'" & Time & "'," _
        & "" & Datos.MostrarCampo("Clientes", "codCliente", "cedrif='" & Text4.Text & "'") & "," _
        & "" & Val(Text11.Text) & "," _
        & "" & Val(Text12.Text) & "," _
        & "" & Val(Text13.Text) & "," _
        & "" & Val(Text14.Text) & "," _
        & "" & Val(Text15.Text) & "," _
        & "'" & Label18.Caption & "'," _
        & "'" & DTPicker1.value & "'," _
        & "" & codUsuario & "," _
        & "" & PIVa & "," _
        & "'" & Label51.Caption & "'," _
        & "'" & Text29.Text & "'," _
        & "'" & Datos.MostrarCampo("formaPago", "codFormaPago", "descripcion ='" & Combo6.Text & "'") & "'," _
        & "'" & Val(Text33.Text) & "'," _
        & "'" & Label19.Caption & "')"
        'MsgBox iSql
        Conexion.Execute iSql
        Conexion.Execute "insert into numerosFacturas (numeroFactura) values(" & Text1.Text & ")"
        GuardarDetalle Text1.Text
        guardarPago Val(Text14.Text)
    Else
        AbonoAnterior = rs("Abono")
        iSql = "Update PedidoEnc set " _
        & " codCliente=" & Datos.MostrarCampo("Clientes", "codCliente", "cedrif='" & Text4.Text & "'") & "," _
        & " subtotal=" & Val(Text11.Text) & "," _
        & " Iva=" & Val(Text12.Text) & "," _
        & " Total=" & Val(Text13.Text) & "," _
        & " Abono=" & Val(Text14.Text) & "," _
        & " codFormapago=" & Datos.MostrarCampo("formaPago", "codFormaPago", "descripcion ='" & Combo6.Text & "'") & ", " _
        & " Restan=" & Val(Text15.Text) & "," _
        & " tipoCambio=" & Val(Text33.Text) & " " _
        & " Where CodPedido=" & Val(Text1.Text) & ""
      
        Conexion.Execute iSql
        If AbonoAnterior <> Val(Text14.Text) Then
            guardarPago Val(Text14.Text)
        End If
        GuardarDetalle Text1.Text
    End If
End Sub
Sub guardarVehiculo()
Dim rs As New ADODB.Recordset
Dim codModelo As String
    rs.Open "select *from vehiculos where placa='" & Combo2.Text & "'", Conexion
    codModelo = Datos.MostrarCampo("Modelos", "codModelo", "Nombre='" & Combo3.Text & "'")
'    If rs.EOF Then
        'iSql = "insert into vehiculos (codVehiculo,placa,codModelo,color,año) values (" _
        '& "" & Datos.generarCodigo("Vehiculos", "codVehiculo") & "," _
        '& "'" & Combo2.Text & "'," _
        '& "" & codModelo & "," _
        '& "'" & Text17.Text & "')"
    'Conexion.Execute iSql
    
    'End If
End Sub
Sub guardarRelacion()
    Dim rs As New ADODB.Recordset
    Dim codVehiculo As Integer
    Dim codCliente As Integer
    codVehiculo = Val(Datos.MostrarCampo("Vehiculos", "codVehiculo", "placa='" & Combo2.Text & "'"))
    codCliente = Val(Datos.MostrarCampo("Clientes", "codCliente", "cedrif='" & Text4.Text & "'"))
    
    
    rs.Open "select *from Clientevehiculos where codVehiculo=" & codVehiculo & " and codCliente=" & codCliente, Conexion
    If rs.EOF Then
    '    Conexion.Execute "insert into Clientevehiculos(codVehiculo,codCliente) values(" _
        & "" & codVehiculo & "," _
        & "" & codCliente & ")"
    End If
End Sub
Sub GuardarCliente()
Dim rs As New ADODB.Recordset
Dim iSql As String
Dim Tabla As String
Dim campoClave As String
Dim codCliente As Integer
Dim codClient As String
Tabla = "Clientes"
    campoClave = "CodCliente"
    codClient = Text4.Text
    If Text4.Text = "" Then
        Dim Fe As String
        Dim Nom As String

        Fe = Format(Date, "ddmm")
        Nom = Mid(Text5.Text, 1, 3)
        codClient = Nom & Fe
    End If
    
    rs.Open "select *from clientes where cedrif = '" & codClient & "'", Conexion
        If rs.EOF Then
        iSql = "insert into [" & Tabla & "] (" _
        & "codCliente," _
        & "cedRif," _
        & "razonsocial," _
        & "direccion," _
        & "telefono)" _
        & " values(" _
        & "" & Datos.generarCodigo(Tabla, campoClave) & "," _
        & "'" & codClient & "'," _
        & "'" & Text5.Text & "'," _
        & "'" & Text6.Text & "'," _
        & "'" & Text7.Text & "')"
    Else
        codCliente = rs(campoClave)
        iSql = "UPDATE [" & Tabla & "] set " _
            & "cedrif      = '" & codClient & "'," _
            & "RazonSocial = '" & Text5.Text & "'," _
            & "direccion   = '" & Text6.Text & "'," _
            & "telefono    = '" & Text7.Text & "'" _
            & "WHERE codCliente=" & codCliente & ""
    End If
     Conexion.Execute iSql
    Text4.Text = codClient
End Sub
Function estaCerrada(codPedido) As Boolean
Dim rs As New ADODB.Recordset
Dim vestaCerrada As Boolean
    vestaCerrada = False
    rs.Open "select *from pedidoEnc where codPedido=" & Val(Text1.Text), Conexion
    If Not rs.EOF Then
       vestaCerrada = rs("statusPago") = "Cerrada"
    End If
    estaCerrada = vestaCerrada
End Function
Sub CuantoResta()
    Dim Resta As Double
    Dim abono As Double
    abono = Val(Text14.Text)
    Resta = Val(Text13.Text) - abono
    Resta = Round(Resta, 2)
    If estaCerrada(Val(Text1.Text)) Then
        Resta = 0
    End If
    Text15.Text = Resta
    Label19.Caption = ObtenerStatusPago(Resta, Val(Text13.Text))
End Sub
Function ObtmontoProductos()
Dim s As Double
Dim i As Integer
    s = 0
    For i = 0 To List4.ListCount - 1
        s = s + Val(List4.List(i))
    Next i
    
    ObtmontoProductos = s
End Function
Sub Totales()
    Dim Iva As Double
    Dim total As Double
    Dim SubTotal As Double
    
    If Check1.value Then
        Iva = montoProductos * PIVa
    Else
        Iva = 0
    End If
    If Not Proyecto.inicialIvaIncluido Then
        SubTotal = montoProductos
    Else
        SubTotal = Formularios.sinIva(montoProductos)
    End If
    Iva = Round(Iva, 2)
    total = SubTotal + Iva
        
    Label51.Caption = Moneda.NumLetra(total, 2) & ObtMoneda
    
    Text11.Text = SubTotal
    Text12.Text = Iva
    Text13.Text = total
    
End Sub
Function buscarLista(Descripcion As String) As Boolean
    Dim Encontro As Boolean
    Dim Pos As Integer
    Encontro = False
    Pos = 0
    While Not Encontro And Pos < List1.ListCount
        If List1.List(Pos) = Descripcion Then
            Encontro = True
        Else
            Pos = Pos + 1
        End If
    Wend
    buscarLista = Encontro
    
End Function
Function posLista(Descripcion As String) As Integer
    Dim Encontro As Boolean
    Dim Pos As Integer
    Encontro = False
    Pos = 0
    While Not Encontro And Pos < List1.ListCount
        If List1.List(Pos) = Descripcion Then
            Encontro = True
        Else
            Pos = Pos + 1
        End If
    Wend
    posLista = Pos
    
End Function

Function Agregar() As Boolean
Dim Agrego As Boolean
Dim Existencia As Integer
Dim Cantidad As Double
Dim seFabrica As String
Dim esServicio As Boolean
    Agrego = False
    If Combo1.Text = "" Then Exit Function
    If Val(Text10.Text) = 0 Then Exit Function
    
    If Val(Combo5.Text) = 0 Then Exit Function
    If Text10.Text <> "" And Text8.Text <> "" Then
        If Not buscarLista(Combo1.Text) Then
            Existencia = Val(Datos.MostrarCampo("Productos", "Stock", "Descripcion='" & Combo1.Text & "'"))
            seFabrica = Datos.MostrarCampo("Productos", "seFabrica", "Descripcion='" & Combo1.Text & "'")
            esServicio = Datos.MostrarCampo("Productos", "Tipo", "Descripcion='" & Combo1.Text & "'") = "Servicio"
            Cantidad = Val(Text8.Text)
            If Not (Existencia >= Cantidad Or seFabrica = "Si" Or esServicio Or Existencia < 0) Then
                MsgBox "No stock para satisfacer este pedido, solo tenemos (" & Existencia & ")"
            End If
                Dim tot As Double
                Dim prec As Double
                List1.AddItem Combo1.Text
                List2.AddItem Cantidad
                prec = Val(Text9.Text)
                List3.AddItem prec
                tot = Val(Text10.Text)
                List4.AddItem tot
                List5.AddItem Text29.Text
                List6.AddItem Round((prec * PIVa) + prec, 2)
                List7.AddItem Round((tot * PIVa) + tot, 2)
                List8.AddItem Label37.Caption

                montoProductos = ObtmontoProductos
                montoProductos = Round(montoProductos, 2)
                Combo4.Text = ""
                Combo1.Text = ""
                Combo4.Text = ""
                Combo1.SetFocus
                Text8.Text = ""
                Text9.ListIndex = 0
                Combo5.ListIndex = 0
                Text9.Clear
                Text10.Text = ""
                Agrego = True
                Label23.Caption = "Precio + IGV:"
                Label32.Caption = ""
                Label34.Caption = ""
                Label37.Caption = ""
                
                Totales
                
            End If
        Else
            Dim posi As Integer
            Dim Cant As Integer
            Dim s As Double
            Dim ns As Double
            posi = posLista(Combo1.Text)
            Cant = List2.List(posi) + Val(Text8.Text)
            'Text9.ListIndex = Combo5.ListIndex
            s = Val(Text8.Text) * Val(Text9.Text)
            ns = s + Val(List4.List(posi))
            montoProductos = ObtmontoProductos
            List2.List(posi) = Cant
            List4.List(posi) = ns
           montoProductos = Round(montoProductos, 2)
                Combo4.Text = ""
                Combo1.Text = ""
                Combo1.SetFocus
                Text8.Text = ""
                Text9.ListIndex = 0
                Combo5.ListIndex = 0
                Text9.Clear
                Text10.Text = ""
                Agrego = True
                Label23.Caption = "Precio + IGV:"
                Label32.Caption = ""
                Label34.Caption = ""
                Label37.Caption = ""
                Totales
            'MsgBox "No puede agregar el mismo Detalle en la factura"
        End If
    Agregar = Agrego
End Function
Sub CalcularMonto()
    Text10.Text = (Val(Text8.Text)) * Val(Text9.Text)
End Sub
Sub BuscarProducto()
Dim filtro As String
Dim codTipoP As Integer
Dim codCliente As Integer
Dim codProducto As Integer
Dim sql As String
    filtro = "Descripcion='" & Combo1.Text & "'"
    codProducto = Val(Datos.MostrarCampo("Productos", "codProducto", filtro))
    
        If Not Datos.Existe("Productos", filtro) Then
            Exit Sub
        End If
    Text9.Clear
    Combo5.Clear
    codTipoP = Val(Datos.MostrarCampo("Productos", "codTipoP", filtro))
    Combo4.Text = Datos.MostrarCampo("TipoP", "Descripcion", "codTipoP=" & codTipoP)
    
    'Dim codCliente As Integer
    filtro = "cedrif='" & Text4.Text & "'"
    codCliente = Val(Datos.MostrarCampo("Clientes", "codCliente", filtro))
    sql = "select *from CotizacionEnc,CotizacionDet where CotizacionEnc.codCotizacion=CotizacionDet.codCotizacion " _
    & " and CodCliente =" & codCliente & " " _
    & " and CodProducto =" & codProducto & ""
    Dim rs As New ADODB.Recordset
    rs.Open sql, Conexion
    filtro = "Descripcion='" & Combo1.Text & "'"
    If rs.EOF Then
    'aqui quede
        
        Text9.AddItem Datos.MostrarCampo("Productos", "Precio", filtro)
        Combo5.AddItem Round((Datos.MostrarCampo("Productos", "Precio", filtro) * PIVa) + Datos.MostrarCampo("Productos", "Precio", filtro), 2)
        
        Text9.AddItem Datos.MostrarCampo("Productos", "PrecioPaga", filtro)
        Combo5.AddItem Round((Datos.MostrarCampo("Productos", "PrecioPaga", filtro) * PIVa) + Datos.MostrarCampo("Productos", "PrecioPaga", filtro), 2)
    Else
        Text9.AddItem rs("Precio")
        Combo5.AddItem conIva(rs("precio"))
    End If
    Combo5.ListIndex = 0
    Text9.ListIndex = 0
    
    Text19.Text = Datos.MostrarCampo("Productos", "barras", filtro)
    Label32.Caption = Datos.MostrarCampo("Productos", "talla", filtro)
    Label34.Caption = Datos.MostrarCampo("Productos", "color", filtro)
    'Text29.Text = Datos.MostrarCampo("Productos", "almacen", Filtro)
    Label37.Caption = Datos.MostrarCampo("Productos", "Stock", filtro)
    
    
    

End Sub
Sub cargarVehiculosClientes()
    Dim filtro As String
    Dim codCliente As Integer
    filtro = "cedrif='" & Text4.Text & "'"
    codCliente = Val(Datos.MostrarCampo("Clientes", "codCliente", filtro))
    Datos.llenarCombo "select placa from ClienteVehiculos,Vehiculos where Vehiculos.codVehiculo=ClienteVehiculos.codVehiculo and codCliente=" & codCliente, Combo2
    If Combo2.ListCount > 0 Then
        Combo2.Text = Combo2.List(0)
        buscarVehiculo
    End If
End Sub
Sub buscarVehiculo()
    Dim filtro As String
    Dim codModelo As Integer
    filtro = "Placa='" & Combo2.Text & "'"
    codModelo = Val(Datos.MostrarCampo("Vehiculos", "codModelo", filtro))
    Combo3.Text = Datos.MostrarCampo("Modelos", "nombre", "codModelo=" & codModelo)
    Text17.Text = Datos.MostrarCampo("Vehiculos", "Color", filtro)
    Text18.Text = Datos.MostrarCampo("Vehiculos", "Año", filtro)

End Sub
Sub BuscarCliente()
    Dim filtro As String
    If Text5.Text <> "" Then
        filtro = "razonsocial='" & Text5.Text & "'"
        Text4.Text = Datos.MostrarCampo("Clientes", "cedrif", filtro)
    Else
        filtro = "cedrif='" & Text4.Text & "'"
        Text5.Text = Datos.MostrarCampo("Clientes", "razonsocial", filtro)
    End If
    Text6.Text = Datos.MostrarCampo("Clientes", "direccion", filtro)
    Text7.Text = Datos.MostrarCampo("Clientes", "telefono", filtro)
 '   If Text5.Text <> "" Then
'        cargarVehiculosClientes
  '  End If
End Sub
Sub CargarPedido(n As Long)
    Dim rs As New ADODB.Recordset
    Dim a  As VbMsgBoxResult
    rs.Open "select *from PedidoEnc where codPedido=" & n, Conexion
    If rs.EOF Then
        a = MsgBox("El Pedido " & n & " no existe", vbExclamation)
    Else
        Mostrar rs
    End If

End Sub
Sub BuscarPedido()
    Dim n As Long
    Form39.Show vbModal

    n = Val(Form39.numero)
    If n = -1 Then Exit Sub
    CargarPedido n

End Sub

Sub Limpiar()
    Text1.Text = ""
    Text2.Text = Date
    Text3.Text = Time
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
'    Text9.ListIndex = 0
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    Text15.Text = ""
    Text16.Text = ""
    Text17.Text = ""
    Text18.Text = ""
    Text19.Text = ""
    txtTranportista.Text = ""
    Text27.Text = "ZARATE"
    Text25.Text = "AV GRAN CHIMU 1195 ZARATE"
    Text31.Text = ""
    Text32.Text = ""
    Combo1.Text = ""
    Combo2.Text = ""
    Combo3.Text = ""
    Combo4.Text = ""
    Combo5.Text = ""
    Label51.Caption = ""
    Label18.Caption = ""
    Label19.Caption = ""
    Label23.Caption = ""
    Label32.Caption = ""
    Text33.Text = ""
    Label34.Caption = ""
    Text29.Text = ""
    Text30.Text = ""
    Label37.Caption = ""
    
    montoProductos = 0
    DTPicker1.value = Date
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    List7.Clear
    List8.Clear
    Combo5.Clear
    Datos.llenarCombo "SELECT descripcion FROM  formaPago", Combo6
    Combo6.Text = Combo6.List(0)
    mostrarGrilla

End Sub
Sub bloquear(es As Boolean)
    DataGrid1.Enabled = es
    DataGrid1.EditActive = False
    
    Text1.Enabled = es
    Text2.Enabled = es
    Text3.Enabled = es
    Text4.Enabled = es
    Text5.Enabled = es
    Text6.Enabled = es
    Text7.Enabled = es
    Text8.Enabled = es
    Text9.Enabled = es
    Combo4.Enabled = es
    Text10.Enabled = es
    Combo5.Enabled = es

    Text11.Enabled = es
    Text12.Enabled = es
    Text13.Enabled = es
    Text14.Enabled = es
    Text15.Enabled = es
    Text16.Enabled = es
    Text17.Enabled = es
    Text18.Enabled = es
    Text19.Enabled = es
    Text20.Enabled = es
    Text31.Enabled = es
    Text32.Enabled = es
    Text33.Enabled = es
    Combo1.Enabled = es
    Combo2.Enabled = es
    Combo3.Enabled = es
    Combo4.Enabled = es
    Combo6.Enabled = es
    DTPicker1.Enabled = es
    Text29.Enabled = es
    Command25.Enabled = es
    List1.Enabled = es
    List2.Enabled = es
    List3.Enabled = es
    List4.Enabled = es
    List5.Enabled = es
    List6.Enabled = es
    List7.Enabled = es
    List8.Enabled = es
    Command5.Enabled = es
    Command8.Enabled = es
    Command17.Enabled = es
    Command21.Enabled = es
    Timer1.Enabled = es
    
End Sub
    
Private Sub Check1_Click()
Totales
CuantoResta
End Sub

Private Sub Combo1_Click()
    BuscarProducto
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
mostrarGrilla

End Sub

Private Sub Combo1_LostFocus()
    Dim filtro As String
    Dim codPedido As Long
    codPedido = Val(Text1.Text)
    
    filtro = "descripcion='" & Combo1.Text & "'"

    If Not Datos.Existe("Productos", filtro) And Not Adodc1.Recordset.EOF Then
        Combo1.Text = Adodc1.Recordset(1)
     End If
    
    BuscarProducto

End Sub

Private Sub Combo2_Click()
    buscarVehiculo
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
    buscarVehiculo
End Sub


Private Sub Combo4_Click()
ProductosPorCategoria
End Sub

Private Sub Combo4_Validate(Cancel As Boolean)
ProductosPorCategoria
End Sub

Private Sub Combo5_Click()
Text9_Click
End Sub

Private Sub Command1_Click()
    If Text4.Enabled Then
        If MsgBox("Al Hacer esta accion perdera los datos de la factura, Desea continuar ", vbYesNo) = vbYes Then
            Nuevo
            Bloqueo 2
        End If
    Else
        Nuevo
        Bloqueo 2
    End If

End Sub

Private Sub Command10_Click()
Formularios.Siguiente Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command11_Click()
Formularios.Ultimo Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command12_Click()
Formularios.Anterior Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command13_Click()

Formularios.Primero Me
'ActivarBotones
Bloqueo 3

End Sub

Private Sub Command14_Click()
    
    If Text4.Enabled Then
        MsgBox "Debe Tener una factura cargada"
    Else
        If Label19.Caption = "Cerrada" Then
            MsgBox "No puede modificar una factura cerrada"
        Else
            CamposModificables
            Bloqueo 2
            esNueva = False
        End If
    End If
End Sub
Sub crearBoleta()
Dim iSql As String
Dim restan As Double
Dim rs As New ADODB.Recordset
Dim Msj
Dim codPedido As Long
Dim j As Byte
Dim nBol As Long
Dim bol As String

    restan = Val(Text15.Text)
    
    
    If restan <> 0 Then
        GoTo 2
        Msj = MsgBox("Atención, Debe primero estar la venta cancelada", vbInformation)
    Else
2:
        codPedido = Val(Text1.Text)
        
        iSql = "Select *from Factura where CodPedido=" & codPedido & ""
        rs.Open iSql, Conexion
        If Not rs.EOF Then
            MsgBox "No puede transformar un factura a una boleta"
            Exit Sub
        End If
        rs.Close
        If Text31.Text = "" Then Exit Sub
        iSql = "SELECT * FROM boleta WHERE nbol =" & Text31.Text
        rs.Open iSql, Conexion
            If Not rs.EOF Then
                If MsgBox("Este numero de boleta ya esta guardado, ¿desea guardarlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        rs.Close
        
        iSql = "Select *from Boleta where CodPedido=" & codPedido & ""
        rs.Open iSql, Conexion
        If rs.EOF Then
            nBol = Text31.Text
            bol = rellenarCeros(nBol)
            'Text16.Text = nFac
            iSql = "update pedidoEnc set NControl=" & nBol & ", tipo = 'b'  where CodPedido=" & codPedido & ""
            Conexion.Execute iSql

            iSql = "insert into boleta (nbol,CodPedido,b,fecha,hora) values(" _
            & "" & nBol & "," _
            & "" & codPedido & "," _
            & "'" & bol & "'," _
            & "'" & Date & "'," _
            & "'" & Time & "')"
            Conexion.Execute iSql
            For j = 1 To 5
                Conexion.Close
                Conectar
            Next
        End If
        Dim filtro As String
        Dim Formulas() As String
        Dim Archivo As String
        ReDim Formulas(0)
        filtro = "{PedidoEnc.CodPedido}=" & codPedido
        Archivo = App.Path & "\reportes\reciboB.rpt"
        Datos.CargarReporte filtro, Archivo, Formulas, True
    
    End If
    
End Sub

Sub crearFactura()
Dim iSql As String
Dim restan As Double
Dim rs As New ADODB.Recordset
Dim Msj
Dim codPedido As Long
Dim nFac  As Long
Dim j As Byte
Dim f As String
    restan = Val(Text15.Text)
    If restan <> 0 Then
        GoTo 2
        Msj = MsgBox("Atención, Debe primero estar la factura cancelada", vbInformation)
    Else
2:
        codPedido = Val(Text1.Text)
        
        iSql = "Select *from Boleta where CodPedido=" & codPedido & ""
        rs.Open iSql, Conexion
        If Not rs.EOF Then
            MsgBox "No puede transformar un boleta una factura"
            Exit Sub
        End If
        rs.Close
        
        iSql = "SELECT * FROM factura WHERE nfac =" & Text32.Text
        rs.Open iSql, Conexion
            If Not rs.EOF Then
                If MsgBox("Este numero de factura ya esta guardado, ¿desea guardarlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        rs.Close

        
        iSql = "Select *from Factura where CodPedido=" & codPedido & ""
        rs.Open iSql, Conexion
        If rs.EOF Then
            nFac = Text32.Text
            f = rellenarCeros(nFac)
            Text16.Text = nFac
            iSql = "update pedidoEnc set NControl=" & nFac & ", tipo = 'f'  where CodPedido=" & codPedido & ""
            Conexion.Execute iSql

            iSql = "insert into factura (nfac,CodPedido,f,fecha,hora) values(" _
            & "" & nFac & "," _
            & "" & codPedido & "," _
            & "'" & f & "'," _
            & "'" & Date & "'," _
            & "'" & Time & "')"
            Conexion.Execute iSql
            For j = 1 To 5
                Conexion.Close
                Conectar
            Next
        End If
        Dim filtro As String
        Dim Formulas() As String
        Dim Archivo As String
        ReDim Formulas(1) As String
        Formulas(0) = "g='" & Text30.Text & "'"

        filtro = "{PedidoEnc.CodPedido}=" & codPedido
        Archivo = App.Path & "\reportes\RECIBOF.rpt"
        Datos.CargarReporte filtro, Archivo, Formulas, True
    End If
End Sub

Function obtTipoCambioDefault()
    Dim t As New ADODB.Recordset
    t.Open "SELECT TOP 1 * FROM pedidoEnc ORDER BY codPedido DESC", Conexion
    If Not t.EOF Then
        If Not IsNull(t("tipoCambio")) Then
            obtTipoCambioDefault = t("tipoCambio")
        Else
            obtTipoCambioDefault = 1
        End If
    Else
        obtTipoCambioDefault = 1
    End If
End Function
Sub transformarASoles(Optional revalidar As Boolean)
Dim i As Integer
Dim tipoCambio As Double
    If revalidar Then
        tipoCambio = 0
    Else
        tipoCambio = Val(Text33.Text)
    End If
    
    While tipoCambio = 0
    
        tipoCambio = Val(InputBox("Int. Tipo de Cambio (-1 para cancelar)", "Cambio de Moneda Actual", Val(Text33.Text)))
        If tipoCambio = -1 Then Exit Sub
    Wend
    Text33.Text = Val(tipoCambio)
    If Text29.Text = "$" Then
    montoProductos = 0
        For i = 0 To List3.ListCount - 1
            If List5.List(i) = "$" Then
                List3.List(i) = Val(List3.List(i)) * tipoCambio
                montoProductos = montoProductos + (Val(List4.List(i)) * tipoCambio)
                List4.List(i) = Val(List4.List(i)) * tipoCambio
                List5.List(i) = "S/."
            End If
        Next
        Text29.Text = "S/."
        Me.Totales
    Else
        MsgBox "La Factura ya está en Soles"
    End If
End Sub
Sub transformarADolares(Optional revalidar As Boolean)
Dim i As Integer
Dim tipoCambio As Double
    If revalidar Then
        tipoCambio = 0
    Else
        tipoCambio = Val(Text33.Text)
    End If
    
    While tipoCambio = 0
        tipoCambio = Val(InputBox("Int. Tipo de Cambio", "Cambio de Moneda Actual", Val(Text33.Text)))
    Wend
    Text33.Text = Val(tipoCambio)
    montoProductos = 0
    If Text29.Text = "S/." Then
        For i = 0 To List3.ListCount - 1
            If List5.List(i) = "S/." Then
                montoProductos = montoProductos + Val(Val(List4.List(i)) / tipoCambio)
                List3.List(i) = Val(List3.List(i)) / tipoCambio
                List4.List(i) = Val(List4.List(i)) / tipoCambio
                List5.List(i) = "$"
            End If
        Next
        Text29.Text = "$"
    Else
        MsgBox "La Factura ya está en Dolares"
    End If
End Sub

Private Sub Command15_Click()


   If List1.ListCount = 0 Then Exit Sub
        If Text29.Text = "$" Then transformarASoles True
        montoProductos = ObtmontoProductos
        
         Me.CalcularMonto
        CuantoResta
    
    GuardarEncabezado
    bloquear False
    crearBoleta
    Bloqueo 3
    

End Sub
Sub Anular()
    Dim RESP As VbMsgBoxResult
    Dim rd As New ADODB.Recordset
    
    Dim codPedido As Long
    Dim codProducto As Long
    Dim esServicio As Boolean
    Dim estaAnulada As Boolean
    RESP = MsgBox("Esta seguro de que desea Anular esta factura", vbYesNo)
    estaAnulada = Datos.MostrarCampo("pedidoenc", "status", "codPedido=" & codPedido) = "Anulado"
    If estaAnulada Then
        MsgBox "Ya esta factura se encuentra anulada"
        Exit Sub
    End If
    If RESP = vbYes Then
        codPedido = Val(Text1.Text)
        Conexion.Execute "update pedidoenc set  status='Anulado',statusPago='Anulado',subtotal=0,iva=0,total=0,abono=0,restan=0 where codPedido=" & codPedido
        rd.Open "select *from pedidoDet where codPedido=" & codPedido, Conexion
        While Not rd.EOF
            codProducto = rd("codProducto")
            esServicio = Datos.MostrarCampo("Productos", "tipo", "codProducto=" & codProducto & "") = "Servicio"
            If Not esServicio Then
                If Proyecto.inventario Then
                    Conexion.Execute "update Productos set stock = stock +" & rd("cantidad") & " where codProducto=" & codProducto
                End If
            End If
            rd.MoveNext
        Wend

    End If
    
End Sub

Private Sub Command16_Click()
    Anular
    bloquear False
    Label18.Caption = "Anulado"
    Bloqueo 3
End Sub

Private Sub Command17_Click()
    If NivelEntro = "Vendedor" Then Exit Sub
    Form5.Show vbModal
    Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    Datos.llenarCombo "select nombre from Modelos order by nombre", Combo3

End Sub


Private Sub Command2_Click()
If List1.ListCount = 0 Then Exit Sub
    GuardarEncabezado
    bloquear False
    Bloqueo 3

End Sub

Private Sub Command20_Click()
If NivelEntro = "Vendedor" Then Exit Sub
If Combo1.Text = "" Then Exit Sub
    Dim barra As String
    barra = InputBox("Nuevo Codigo de barras", "Cambiar Codigo de Barras", Text19.Text)
    If barra <> "" Then
        Conexion.Execute "update productos set barras='" & barra & "' where descripcion='" & Combo1.Text & "'"
        MsgBox "Cambiado"
        Text19.Text = barra
    End If

End Sub

Private Sub Command21_Click()
Dim codProducto As Integer
Dim Descripcion As String
    Catalogo.sql = "SELECT *from productosOrdenados2"
    Catalogo.Show vbModal
    codProducto = Val(Catalogo.Resultado)
    Descripcion = Datos.MostrarCampo("productos", "descripcion", "codProducto=" & codProducto)
    Combo1.Text = Descripcion
    Me.BuscarProducto
    Text8.SetFocus
End Sub

Private Sub Command22_Click()
    If Text4.Text = "" Or Text5.Text = "" Then
        MsgBox "Una factura debe tener un numero RUC y un nombres y apellidos"
        Exit Sub
    End If
    If Proyecto.inicialGuia Then
        If Trim(Text30.Text) = "" Then
            If MsgBox("Atencion, esta olvidando colocar el numero de guia, desea continuar", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    If List1.ListCount = 0 Then Exit Sub
    GuardarEncabezado
    bloquear False
    crearFactura
    Bloqueo 3
End Sub

Private Sub Command23_Click()
Dim codProducto As Integer
Dim Cantidad As Double
Dim i As Integer
    If Label18.Caption <> "Pendiente" Then
        MsgBox "Ya está cambiada"
        Exit Sub
    End If
    For i = 1 To List1.ListCount - 1
        codProducto = Val(Datos.MostrarCampo("productos", "codProducto", "descripcion='" & List1.List(i) & "'"))
        Cantidad = Val(List2.List(i))
        Conexion.Execute "update productos set stock=stock-" & Cantidad & " where codProducto=" & codProducto
    Next
    Conexion.Execute "update pedidoEnc set status='Cambiada' where codPedido=" & Val(Text1.Text)
    CargarPedido Val(Text1.Text)
End Sub
Private Sub Command24_Click()
    MostrarCatalogo "select cedrif,razonsocial,direccion,telefono from clientes"
    Text4.Text = Catalogo.Resultado
    Text5.Text = Catalogo.Adodc1.Recordset(1)
    Me.BuscarCliente
    On Error Resume Next
    Combo1.SetFocus
End Sub

Private Sub Command25_Click()
    If Command25.Tag = "S" Then
        Me.transformarASoles False
    Else
        Me.transformarADolares False
    End If
    montoProductos = ObtmontoProductos
    Totales
End Sub

Private Sub Command3_Click()
    If Text4.Enabled Then
        If MsgBox("Al Hacer esta accion perdera los datos de la factura, Desea continuar ", vbYesNo) = vbYes Then
            cancelar
            Bloqueo 1
        End If
    Else
        cancelar
        Bloqueo 1

    End If
End Sub
Private Sub Command4_Click()
    Unload Me
End Sub
Private Sub Command5_Click()
    If Me.ActiveControl = Combo1 Or Me.ActiveControl = Text19 Or Me.ActiveControl = Text29 Then
        If Me.ActiveControl = Text19 Then
            BuscarBarra
        End If
        Text8.SetFocus
        Exit Sub
    End If
    If Not Agregar Then
        On Error Resume Next
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Command6_Click()
    BuscarPedido
End Sub

Function ObtenerStatusPago(Resta As Double, total As Double) As String
Dim Status As String
    If Resta = 0 Then
        Status = "Cerrada"
    Else
        If Resta = total Then
            Status = "Abierta"
        Else
            Status = "Debe"
        End If
    End If
    ObtenerStatusPago = Status
End Function

Private Sub Command7_Click()
Dim filtro As String
Dim Formulas() As String
Dim Archivo As String
Dim codPedido As Long
    If List1.ListCount = 0 Then Exit Sub
    GuardarEncabezado
    Me.guardarGuia
    MsgBox "Generando guia, haga click para imprimir'"
    bloquear False
    codPedido = Val(Text1.Text)
    ReDim Formulas(0)
    Formulas(0) = ""
    filtro = "{PedidoEnc.CodPedido}=" & codPedido
    Archivo = App.Path & "\reportes\recibo.rpt"
    Datos.CargarReporte filtro, Archivo, Formulas, True
    Bloqueo 3

End Sub

Private Sub Command8_Click()
    BorrarDetalle List1.ListIndex
End Sub

Private Sub Command9_Click()
Dim a As VbMsgBoxResult
Dim restan As Double
    If NivelEntro = "Vendedor" Then
        MsgBox "Ni Puede realizar esta operacion"
        Exit Sub
    End If
    restan = Val(Text15.Text)
    If Val(Text14.Text) = 0 Then
        Form38.Label2.Caption = Val(Text13.Text)
        Form38.Son.Caption = ""
        Form38.Show vbModal
        If Form38.Monto = -1 Then
            Exit Sub
        End If
        Text14.Text = Val(Form38.Monto)
        bloquear False
        Text14.Enabled = True
        Text14.SetFocus
        Bloqueo 2
    Else
        If restan = 0 Then
            a = MsgBox("El Pedido ya esta cancelado en su totalidad")
        Else
            a = MsgBox("Cancelara los " & restan, vbYesNo)
            If a = vbYes Then
                Text15.Text = 0
                GuardarEncabezado
                Conexion.Execute "update pedidoEnc set statusPago='Cerrada', restan=0 where codPedido=" & Text1.Text
                guardarPago Val(restan)
                bloquear False
                Bloqueo 3
            End If
        End If
    End If
End Sub

Sub ProductosPorCategoria()
    If Combo4.Text = "" Then
        Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    Else
        Datos.llenarCombo "select productos.descripcion from productos,tipop where productos.codtipop=tipop.codtipop and tipop.descripcion='" & Combo4.Text & "' order by productos.descripcion", Combo1
    End If
End Sub

Private Sub DataGrid1_Click()
    Combo1.Text = Adodc1.Recordset(1)
    BuscarProducto

End Sub

Private Sub DataGrid1_DblClick()
    Text8.SetFocus
End Sub
Private Sub configurar()
    If Not Proyecto.inicialGuia Then 'Si no tiene guia
        Frame1.Visible = False
        Command7.Visible = False
        Command4.Left = Command7.Left
        Label54.Left = Label54.Left - 6500
        Combo6.Left = Combo6.Left - 6500
        Label55.Left = Label55.Left - 9300
        DTPicker1.Left = DTPicker1.Left - 9300
        Label55.Top = Label55.Top + 500
        DTPicker1.Top = DTPicker1.Top + 500
        
    Else
        Frame1.Visible = True
        Command7.Visible = True
    End If
End Sub

Private Sub Form_Load()
    configurar
    Formularios.ColorLabels ColorLetras, Me
    Frame1.BackColor = Me.BackColor
    cancelar
    Datos.llenarCombo "select descripcion from productos order by descripcion", Combo1
    Datos.llenarCombo "select descripcion from tipop order by descripcion", Combo4
    Label40.Caption = Datos.MostrarCampo("usuarios", "Alias", "codUsuario=" & codUsuario)
    Datos.llenarCombo "select nombre from Modelos order by nombre", Combo3
    CargarTablas
    Bloqueo 1
    campoClave = "codpedido"
End Sub

Private Sub Frame1_DblClick()
    Form51.Text21.Text = Me.Text21.Text
    Form51.Text22.Text = Me.txtTranportista.Text
    Form51.Text23.Text = Me.Text23.Text
    Form51.Text24.Text = Me.Text24.Text
    Form51.Text25.Text = Me.Text25.Text
    Form51.Text26.Text = Me.Text26.Text
    Form51.Text27.Text = Me.Text27.Text
    Form51.Text28.Text = Me.Text28.Text
    Form51.Show vbModal
    
End Sub

Private Sub Label32_Click()
Dim Unidad As String
    If NivelEntro = "Vendedor" Then Exit Sub
    If Combo1.Text = "" Then Exit Sub
    Unidad = InputBox("Nueva Unidad de Medida", "Cambiar Unidad de Medida", Label32.Caption)
    If Unidad <> "" Then
        Conexion.Execute "update productos set color='" & Unidad & "' where descripcion='" & Combo1.Text & "'"
        MsgBox "Cambiado"
        Label32.Caption = Unidad
    End If
End Sub

Private Sub Label34_Click()
Dim marca As String
    If NivelEntro = "Vendedor" Then Exit Sub
    If Combo1.Text = "" Then Exit Sub
    marca = InputBox("Nueva Marca o Procedencia", "Cambiar Marca o procedencia", Label34.Caption)
    If marca <> "" Then
        Conexion.Execute "update productos set talla='" & marca & "' where descripcion='" & Combo1.Text & "'"
        MsgBox "Cambiado"
        Label34.Caption = marca
    End If
End Sub

Private Sub Label37_Click()
Dim stock As String
    If NivelEntro = "Vendedor" Then Exit Sub
    If Combo1.Text = "" Then Exit Sub
        stock = Val(InputBox("Nuevo Stock", "Cambiar Stock", Label37.Caption))
        If stock <> 0 Then
            Conexion.Execute "update productos set stock=" & stock & " where descripcion='" & Combo1.Text & "'"
            MsgBox "Cambiado"
            Label37.Caption = stock
        End If

End Sub

Private Sub List1_Click()
    SeleccionarListas List1.ListIndex
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        BorrarDetalle List1.ListIndex
    End If
End Sub
Private Sub List2_Click()
    SeleccionarListas List2.ListIndex
End Sub

Private Sub List3_Click()
    SeleccionarListas List3.ListIndex
End Sub

Private Sub List4_Click()
    SeleccionarListas List4.ListIndex
End Sub

Private Sub Text10_Change()
    Text20.Text = Round((Val(Text10.Text) * PIVa) + Val(Text10.Text), 2)
End Sub

Private Sub Text14_Change()
    CuantoResta
End Sub

Private Sub Text14_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Text14.Text = ""
    End If
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
Dim a As VbMsgBoxResult
    If Val(Text15.Text) < 0 Then
        Text15.Text = Val(Text15.Text)
        a = MsgBox("No Puede Abonar un monto mayor al Total del pedido", vbCritical)
        Cancel = True
        Text14.SelStart = 0
        Text14.SelLength = Len(Text14.Text)
    End If
End Sub
Private Sub Text19_Validate(Cancel As Boolean)
    BuscarBarra
End Sub
Private Sub Text21_Change()
    Text21.ToolTipText = Text21.Text
End Sub

Private Sub txtTranportista_Change()
    txtTranportista.ToolTipText = txtTranportista.Text
End Sub
Private Sub Text23_Change()
    Text23.ToolTipText = Text23.Text
End Sub
Private Sub Text24_Change()
    Text24.ToolTipText = Text24.Text
End Sub
Private Sub Text25_Change()
    Text25.ToolTipText = Text25.Text
End Sub
Private Sub Text29_Change()
    Datos.AutoCompletar_TextBox Text29
    Me.MostrarCambio
End Sub
Private Sub Text29_GotFocus()
    Datos.CargarValores "select *from Monedas"
End Sub
Private Sub Text29_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text29.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
End Sub
Sub MostrarCambio()
Dim cambioMoneda As String
    If Text29.Text = "$" Then
        cambioMoneda = "Soles"
        Command25.Tag = "S"
    Else
        cambioMoneda = "Dolares"
        Command25.Tag = "$"
    End If
    Command25.Caption = "Cambiar a " & cambioMoneda
End Sub

Private Sub Text29_Validate(Cancel As Boolean)
    If Text29.Text <> "$" And Text29.Text <> "S/." Then
        MsgBox "Tipo de Moneda invalida, por favor coloque uno correcto"
        Text29.Text = ""
        Cancel = True
    Else
        Me.MostrarCambio
    End If
End Sub

Private Sub Text4_Change()
    Text24.Text = Text4.Text
End Sub
Private Sub Text5_Change()
    Datos.AutoCompletar_TextBox Text5
    Text21.Text = Text5.Text
End Sub
Private Sub Text5_GotFocus()
    Datos.CargarValores "select CODCliente,RAZONSOCIAL from Clientes order by RAZONSOCIAL"
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyBack, vbKeyDelete
            Select Case Len(Text5.Text)
                Case Is <> 0
                    bKeyBack = True
            End Select
    End Select
    If KeyCode = 114 Then
        Command1_Click
    End If
End Sub
Private Sub Text5_Validate(Cancel As Boolean)
    Text5.Text = Replace(Text5.Text, "'", "`")
    BuscarCliente
End Sub
Private Sub Text6_Change()
    Text23.Text = Text6.Text
End Sub
Private Sub Text8_Change()
    CalcularMonto
End Sub
Private Sub Text8_GotFocus()
    Text8.Text = 1
    Text8.SelStart = 0
    Text8.SelLength = 1
End Sub
Private Sub Text9_Change()
    Text8_Change
End Sub
Private Sub Text9_Click()
Dim Monto As Double
'Text9.ListIndex = Combo5.ListIndex    ç
    Monto = (Val(Text9.Text) * PIVa) + Val(Text9.Text)
    Label23.Caption = Round(Monto, 2)
    If Text9.ListIndex = 0 Then
        Label28.Caption = ""
    Else
        Label28.Caption = "Esta colocando un precio especial"
    End If
    CalcularMonto
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 113) Then
    Dim sinIGV As Double
        sinIGV = sinIva(Val(Text9.Text))
        Text9.Text = sinIGV
        CalcularMonto
    End If
End Sub
