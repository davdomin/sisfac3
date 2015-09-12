VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmReporte 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   15120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   13680
      Top             =   2160
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4800
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "impresora"
      BeginProperty Font 
         Name            =   "Eras Bold ITC"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   13800
      TabIndex        =   0
      Top             =   80
      Width           =   1400
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      CausesValidation=   0   'False
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      lastProp        =   500
      _cx             =   1931
      _cy             =   1931
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FrmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Redimensionar()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    CRViewer91.Width = ScaleWidth
End Sub
Private Sub Form_Load()
    Redimensionar
    Me.Caption = "Dacli - Reporte"
End Sub
Private Sub Form_Resize()
    Redimensionar
End Sub
Private Sub Command1_Click()
Dim controlador, impresora, puerto As String
On Error Resume Next
    cd1.CancelError = True
    cd1.ShowPrinter
    If Err = 0 Then
        impresora = Printer.DeviceName
        controlador = Printer.DriverName
        puerto = Mid(Printer.Port, 1, Len(Printer.Port) - 1)
        'Report.SelectPrinter controlador, impresora, puerto
    End If
Err = 0
End Sub


