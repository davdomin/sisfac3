VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form52 
   Caption         =   "Tipo cambio SUNAT"
   ClientHeight    =   10260
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   11400
   LinkTopic       =   "Form52"
   ScaleHeight     =   10260
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4320
      Top             =   7680
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   6
      Top             =   3600
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5775
      Left            =   3720
      TabIndex        =   5
      Top             =   4320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   10186
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
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
         Size            =   9
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
            Type            =   1
            Format          =   "d. MMMM yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   5040
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3720
      Width           =   2535
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2670
      Left            =   360
      TabIndex        =   1
      Top             =   5400
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   4710
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   112459777
      CurrentDate     =   41336
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      ExtentX         =   26696
      ExtentY         =   5953
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Venta:"
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
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Form52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim sql As String
    Dim id As Integer
    Dim rs As New ADODB.Recordset
    sql = "SELECT * FROM cambioSunat WHERE fecha = " & getFecha("#")
    rs.Open sql, Conexion
    If rs.EOF Then
        sql = "INSERT INTO cambioSunat(fecha,precio_venta) VALUES (" & getFecha("#") & "," & Text1.Text & ")"
    Else
        id = rs("id")
        sql = "UPDATE cambioSunat SET precio_venta = " & Text1.Text _
            & " WHERE id = " & id
    End If
    Conexion.Execute sql
    MsgBox ("Cambio Registrado")
    ActualizarGrilla
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Function getFecha(delimitador)
    getFecha = delimitador & Format(Me.MonthView1.value, "mm/dd/yyyy") & delimitador
End Function

Private Function getMes()
    getMes = "MONTH(#" & Format(Me.MonthView1.value, "mm/dd/yyyy") & "#)"
End Function
Private Function getAnio()
    getAnio = "YEAR(#" & Format(Me.MonthView1.value, "mm/dd/yyyy") & "#)"
End Function
Sub ActualizarGrilla()
Dim sql As String
    Adodc1.ConnectionString = Conexion.ConnectionString
    sql = "SELECT  fecha as [Fecha], precio_venta as [Precio Venta] FROM cambioSunat WHERE " _
        & "MONTH(Fecha) = " & getMes & " AND YEAR(fecha) = " & getAnio & " ORDER BY fecha"
        
    Adodc1.RecordSource = sql
    Set DataGrid1.datasource = Adodc1
    Adodc1.Refresh
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    WebBrowser1.Navigate Proyecto.urlSunat
    Formularios.ColorLabels ColorLetras, Me
    ActualizarGrilla
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    ActualizarGrilla
End Sub

