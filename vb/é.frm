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
   ClientWidth     =   10740
   DrawMode        =   8  'Xor Pen
   FillColor       =   &H000000C0&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7740
   ScaleWidth      =   10740
   StartUpPosition =   3  'Windows Default
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   5535
      Left            =   8280
      OleObjectBlob   =   "é.frx":0000
      TabIndex        =   15
      Top             =   1320
      Width           =   2295
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
      TabIndex        =   14
      Text            =   "00"
      Top             =   2520
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "é.frx":24B4
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   8070
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
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
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   3000,189
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton hapus 
      Caption         =   "Hapus All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton simpan 
      Caption         =   "simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4680
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
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
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "é.frx":24C9
      Left            =   1200
      List            =   "é.frx":24D3
      TabIndex        =   6
      Text            =   "Pilih Baudret"
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Timer TimerBaca 
      Interval        =   1000
      Left            =   720
      Top             =   7320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Mulai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox port 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "é.frx":24E3
      Left            =   1200
      List            =   "é.frx":250B
      TabIndex        =   1
      Text            =   "Pilih Port"
      Top             =   840
      Width           =   3615
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InputLen        =   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Tinggi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Left            =   5880
      TabIndex        =   13
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Budret"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
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
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   525
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
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   6810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnStart_Click()

With MSComm1
    If .PortOpen = False Then
        .CommPort = port
        .Settings = baudred + ",n,8,1"
        .InputLen = 1
        .RThreshold = 1
        .PortOpen = True
        Shape1.FillColor = vbGreen
        
    Else
        .PortOpen = False
        Shape1.FillColor = vbYellow
    End If
End With

End Sub



Private Sub Command1_Click()
MSComm1.Output = "a"
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
    MSComm1.PortOpen = False
    Shape1.FillColor = vbYellow
End
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""

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

Private Sub MSComm1_OnComm()
Dim d As String
Dim data As String
Dim Angka As Integer
d = MSComm1.Input
Label4.Caption = Label4.Caption + d
If CBool(InStr(Label4.Caption, ",")) Then
    data = Replace(Label4.Caption, ",", "", 1, 10, 1)
    If Not data Then
        Text3.Text = data
        With Chart
            MSChart1.chartType = VtChChartType2dBar
            MSChart1.AllowSelections = False
            MSChart1.ColumnCount = 1
            MSChart1.RowCount = 1
            MSChart1.Row = 1
            MSChart1.RowLabel = "Tinggi"
            MSChart1.data = Val(Text3.Text)
        End With
        Adodc1.Recordset.AddNew
        Me.Adodc1.Recordset(1) = "tinggi"
        Me.Adodc1.Recordset(2) = data
        Adodc1.Recordset.Update
        Label4.Caption = ""
    Else
        MsgBox "Data Dari Arduino Tidak Ada", vbInformation, "Peringatan"
        Cancel = 0
    End If
    
End If
End Sub

Private Sub simpan_Click()
    Adodc1.Recordset.AddNew
    Me.Adodc1.Recordset(1) = Text1.Text
    Me.Adodc1.Recordset(2) = Text2.Text
    Adodc1.Recordset.Update
End Sub

Private Sub Timer2_Timer()

End Sub

