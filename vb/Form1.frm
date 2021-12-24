VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   Caption         =   "Form1"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9840
   DrawMode        =   8  'Xor Pen
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7740
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   5640
      Top             =   6960
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   5640
      Top             =   7320
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "Stop"
      Height          =   735
      Left            =   4920
      TabIndex        =   17
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "00"
      Top             =   2760
      Width           =   1095
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5415
      Left            =   8160
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":2356
      Height          =   4575
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "nama"
         Caption         =   "nama"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "dalam"
         Caption         =   "dalam"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   494,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   3000,189
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      DataField       =   "ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "dalam"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton hapus 
      Caption         =   "Hapus All"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton simpan 
      Caption         =   "simpan"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "nama"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6720
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1200
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\data d\skripsi\alat\Pkl Air\vb\monitoring air.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\data d\skripsi\alat\Pkl Air\vb\monitoring air.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "t_kedalaman"
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
   Begin VB.ComboBox baudred 
      Height          =   315
      ItemData        =   "Form1.frx":236B
      Left            =   1200
      List            =   "Form1.frx":2375
      TabIndex        =   6
      Text            =   "Pilih Baudret"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Timer TimerBaca 
      Interval        =   2000
      Left            =   5640
      Top             =   6600
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Mulai"
      Height          =   735
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox port 
      Height          =   315
      ItemData        =   "Form1.frx":2385
      Left            =   1200
      List            =   "Form1.frx":23AD
      TabIndex        =   1
      Text            =   "Pilih Port"
      Top             =   840
      Width           =   2415
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budret"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   510
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monitoring Kedalaman Sungai"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrData()
Dim totalBaca As Integer
Dim tinggi As Single
Dim keterangan As String
Const maxBaca = 100

Dim conn1 As ADODB.Connection
Dim Cmd1 As ADODB.Command
Dim Param1 As ADODB.Parameter
Dim Rs1 As ADODB.Recordset


Private Sub btnStart_Click()
    With Chart
        MSChart1.chartType = VtChChartType2dBar
        MSChart1.AllowSelections = False
        MSChart1.ColumnCount = 2
        MSChart1.RowCount = 1
        MSChart1.Row = 1
        MSChart1.RowLabel = "label2"
        MSChart1.Data = Val(Text2(0).Text)
 End With
End Sub

Private Sub btnStop_Click()
btnStart.Enabled = True
btnStop.Enabled = False
TimerBaca.Enabled = False
'Timer1.Enabled = False
'Timer2.Enabled = False
End Sub

Private Sub Command1_Click()
If MSComm1.PortOpen = False Then
    MSComm1.CommPort = port
    MSComm1.Settings = baudred + ",n,8,1"
    MSComm1.PortOpen = True
    MSComm1.InputLen = 0
    MSComm1.RThreshold = 1
    PortisOpen = True
    Shape1.FillColor = vbGreen
Else
    MSComm1.PortOpen = False
    Shape1.FillColor = vbYellow
    PortisOpen = False
End If



End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
Adodc1.Recordset.Update
DataGrid1.Refresh

End Sub

Private Sub Command3_Click()
If MsgBox("Anda Ingin Keluar ?", vbYesNo, "DATA") = vbNo Then
    Cancel = 1
    Text1.SetFocus
Else
End
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Dim I As Byte
For I = 1 To 16
    port.AddItem (I)
Next I
dbConnect 'koneksi dari module
End Sub

Private Sub hapus_Click()
    Dim hapus_all As Integer
    For hapus_all = 1 To Adodc1.Recordset.RecordCount
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Delete
    Adodc1.Recordset.Update
    Adodc1.Recordset.MoveNext
    Next hapus_all
End Sub

Private Sub simpan_Click()
    Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset(1) = Text1.Text
    Me.Adodc1.Recordset(2) = Text2.Text
    Adodc1.Recordset.Update
End Sub

