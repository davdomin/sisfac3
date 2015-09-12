VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Catalogo 
   Caption         =   "Catalogo"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   13950
   LinkTopic       =   "Form3"
   ScaleHeight     =   9690
   ScaleWidth      =   13950
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   7320
      Width           =   13935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12303
      _Version        =   393216
      BackColor       =   12648447
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   19
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
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
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   14040
      TabIndex        =   2
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   13935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   12240
      Top             =   5640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
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
      RecordSource    =   "select *from almacen"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Catalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tamaño As Integer
Public sql As String
Public Resultado As String
Dim Campo

Private Sub Command1_Click()
    Resultado = Adodc1.Recordset(0)
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    Resultado = Adodc1.Recordset(0)
    Unload Me
End Sub
Sub colocarTamaño()
On Error Resume Next
    Dim variacion As Integer
    variacion = 1000
    DataGrid1.Columns(1).Width = Tamaño
    DataGrid1.Columns(0).Visible = False
    Text1.Top = Me.ScaleHeight - variacion
    Text2.Top = Text1.Top + Text1.Height + 1
    Command1.Top = Me.ScaleHeight - variacion
    DataGrid1.Width = Me.ScaleWidth
    DataGrid1.Height = Me.ScaleHeight - variacion
End Sub

Private Sub Form_Load()
On Error Resume Next
'    A = InputBox(sql, sql, sql)
    Tamaño = 6000
    Adodc1.ConnectionString = Conexion.ConnectionString
    Adodc1.RecordSource = sql
    Adodc1.Refresh
    Adodc1.Refresh
    Set DataGrid1.DataSource = Adodc1
    
    Campo = Adodc1.Recordset.Fields(1).Name
    colocarTamaño
End Sub

Private Sub Form_Resize()
colocarTamaño
End Sub

Private Sub Text1_Change()
    Adodc1.RecordSource = sql & " where [" & Campo & "] like'%" & Text1.Text & "%'"
    Adodc1.Refresh
    colocarTamaño
End Sub

Private Sub Text2_Change()
    If Text1.Text <> "" Then
        Adodc1.RecordSource = sql & " where [" & Campo & "] like'%" & Text1.Text & "%'  and [" & Campo & "] like'%" & Text2.Text & "%'"
    Else
        Adodc1.RecordSource = sql & " where [" & Campo & "] like'%" & Text2.Text & "%'"
    End If
        Adodc1.Refresh
        colocarTamaño

End Sub
