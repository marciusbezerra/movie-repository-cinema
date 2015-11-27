VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTitu 
   Caption         =   "Cadastro de títulos, atores e diretores"
   ClientHeight    =   7170
   ClientLeft      =   195
   ClientTop       =   285
   ClientWidth     =   9675
   Icon            =   "frmTitu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   9675
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   225
      TabIndex        =   0
      Top             =   6270
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   1376
      ButtonWidth     =   1508
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "Imagens"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Adicionar"
            Key             =   "Adicionar"
            Object.ToolTipText     =   "Adiciona um novo título"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "A&lterar"
            Key             =   "Alterar"
            Object.ToolTipText     =   "Altera o título atual"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salva as modificações"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "Cancelar"
            Object.ToolTipText     =   "Cancela as modificações"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "E&xcluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir o título atual"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgDiretor 
      Left            =   5415
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitu.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitu.frx":0626
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTitu.frx":0942
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab sstPaginas 
      Height          =   6045
      Left            =   60
      TabIndex        =   1
      Top             =   135
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   10663
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      MouseIcon       =   "frmTitu.frx":0C5E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Informações sobre títulos"
      TabPicture(0)   =   "frmTitu.frx":0C7A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Imagens"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "datGeral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Fotos do filme"
      TabPicture(1)   =   "frmTitu.frx":0C96
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Fotos dos atores"
      TabPicture(2)   =   "frmTitu.frx":0CB2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Fotos dos diretores"
      TabPicture(3)   =   "frmTitu.frx":0CCE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Cotações"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2385
         Left            =   6405
         TabIndex        =   26
         Top             =   885
         Width           =   2805
         Begin VB.TextBox txtCotacaoC 
            DataField       =   "CotacaoC"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   765
            TabIndex        =   28
            Top             =   900
            Width           =   465
         End
         Begin VB.TextBox txtCotacaoM 
            DataField       =   "CotacaoM"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   780
            TabIndex        =   27
            Top             =   1815
            Width           =   465
         End
         Begin MSComCtl2.UpDown updCotC 
            Height          =   285
            Left            =   1231
            TabIndex        =   29
            Top             =   900
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            _Version        =   393216
            OrigLeft        =   1425
            OrigTop         =   870
            OrigRight       =   1620
            OrigBottom      =   1170
            Max             =   5
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown updCotM 
            Height          =   285
            Left            =   1246
            TabIndex        =   30
            Top             =   1815
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "dbcDistribuidora"
            BuddyDispid     =   196637
            OrigLeft        =   1485
            OrigTop         =   1770
            OrigRight       =   1680
            OrigBottom      =   2070
            Max             =   5
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Crítica:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   225
            TabIndex        =   32
            Top             =   945
            Width           =   510
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Pessoal:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   150
            TabIndex        =   31
            Top             =   1860
            Width           =   600
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Diretor(es)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         Left            =   435
         TabIndex        =   23
         Top             =   3840
         Width           =   4335
         Begin MSAdodcLib.Adodc datDiretor 
            Height          =   330
            Left            =   915
            Top             =   1410
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   1
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
            Connect         =   $"frmTitu.frx":0CEA
            OLEDBString     =   $"frmTitu.frx":0D82
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM Diretores"
            Caption         =   "Adodc3"
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
         Begin MSDataGridLib.DataGrid grdDiretores 
            Bindings        =   "frmTitu.frx":0E1A
            Height          =   1080
            Left            =   105
            TabIndex        =   24
            Top             =   255
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1905
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Nome"
               Caption         =   "Nome"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Sobrenome"
               Caption         =   "Sobrenome"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1725,165
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1830,047
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   2940
            TabIndex        =   25
            Top             =   1365
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   688
            ButtonWidth     =   714
            ButtonHeight    =   688
            Style           =   1
            ImageList       =   "imgDiretor"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Adicionar"
                  Object.ToolTipText     =   "Adiciona um novo diretor"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Alterar"
                  Object.ToolTipText     =   "Altera o diretor atual"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excluir"
                  Object.ToolTipText     =   "Excluir o diretor atual"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ator(es)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   4905
         TabIndex        =   20
         Top             =   3840
         Width           =   4320
         Begin MSAdodcLib.Adodc datAtor 
            Height          =   330
            Left            =   150
            Top             =   1395
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\Dados\PROG\VBASIC\Filmotec\Filmotec.mdb"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\Dados\PROG\VBASIC\Filmotec\Filmotec.mdb"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Atores"
            Caption         =   "Adodc4"
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
         Begin MSDataGridLib.DataGrid grdAtores 
            Bindings        =   "frmTitu.frx":0E33
            Height          =   1065
            Left            =   105
            TabIndex        =   21
            Top             =   270
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1879
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "Nome"
               Caption         =   "Nome"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "Sobrenome"
               Caption         =   "Sobrenome"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
                  ColumnWidth     =   1755,213
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1800
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   390
            Left            =   2970
            TabIndex        =   22
            Top             =   1380
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   688
            ButtonWidth     =   714
            ButtonHeight    =   688
            Style           =   1
            ImageList       =   "imgDiretor"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Adicionar"
                  Object.ToolTipText     =   "Adiciona um novo ator"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Alterar"
                  Object.ToolTipText     =   "Altera o ator atual"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excluir"
                  Object.ToolTipText     =   "Excluir o ator atual"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Informações gerais"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         Left            =   435
         TabIndex        =   2
         Top             =   870
         Width           =   5925
         Begin VB.TextBox txtCodigo 
            DataField       =   "Codigo"
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1185
            TabIndex        =   8
            Top             =   315
            Width           =   660
         End
         Begin VB.TextBox txtTitulo 
            DataField       =   "Titulo"
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1185
            TabIndex        =   7
            Top             =   690
            Width           =   4425
         End
         Begin VB.TextBox txtTitOriginal 
            DataField       =   "TitOriginal"
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1185
            TabIndex        =   6
            Top             =   1080
            Width           =   4425
         End
         Begin VB.TextBox txtAnoProducao 
            DataField       =   "AnoProducao"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4950
            TabIndex        =   5
            Top             =   1485
            Width           =   465
         End
         Begin VB.TextBox txtDuracao 
            DataField       =   "Duracao"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datGeral"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4965
            TabIndex        =   4
            Top             =   1905
            Width           =   660
         End
         Begin VB.CheckBox chkAssisti 
            DataField       =   "Assisti"
            DataMember      =   "Titulos"
            DataSource      =   "datGeral"
            Height          =   285
            Left            =   5400
            TabIndex        =   3
            Top             =   2385
            Width           =   330
         End
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   4050
            Top             =   1785
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            Connect         =   $"frmTitu.frx":0E49
            OLEDBString     =   $"frmTitu.frx":0ED1
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM Distribuidoras ORDER By Descricao"
            Caption         =   "Adodc2"
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   4095
            Top             =   1650
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   4
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
            Connect         =   $"frmTitu.frx":0F59
            OLEDBString     =   $"frmTitu.frx":0FE1
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "SELECT * FROM Generos Order By Descricao"
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
         Begin MSComCtl2.UpDown updProducao 
            Height          =   315
            Left            =   5415
            TabIndex        =   9
            Top             =   1485
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   556
            _Version        =   393216
            BuddyControl    =   "chkAssisti"
            BuddyDispid     =   196621
            OrigLeft        =   6315
            OrigTop         =   1425
            OrigRight       =   6510
            OrigBottom      =   1725
            Max             =   5
            SyncBuddy       =   -1  'True
            Wrap            =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSDataListLib.DataCombo dbcGenero 
            Bindings        =   "frmTitu.frx":1069
            DataField       =   "Genero"
            DataSource      =   "datGeral"
            Height          =   315
            Left            =   1185
            TabIndex        =   10
            Top             =   1455
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "Descricao"
            BoundColumn     =   "Codigo"
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dbcDistribuidora 
            Bindings        =   "frmTitu.frx":1095
            DataField       =   "Distribuidora"
            DataSource      =   "datGeral"
            Height          =   315
            Left            =   1185
            TabIndex        =   11
            Top             =   1845
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Descricao"
            BoundColumn     =   "Codigo"
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   600
            TabIndex        =   19
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Título:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   705
            TabIndex        =   18
            Top             =   750
            Width           =   450
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tit. Original:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   270
            TabIndex        =   17
            Top             =   1125
            Width           =   885
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Produção:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   4185
            TabIndex        =   16
            Top             =   1530
            Width           =   735
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Gênero:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   555
            TabIndex        =   15
            Top             =   1530
            Width           =   585
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distribuidora:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   165
            TabIndex        =   14
            Top             =   1875
            Width           =   960
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duração:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   4275
            TabIndex        =   13
            Top             =   1950
            Width           =   660
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Já assistí este filme "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   3945
            TabIndex        =   12
            Top             =   2430
            Width           =   1425
         End
      End
      Begin MSAdodcLib.Adodc datGeral 
         Height          =   390
         Left            =   6405
         Top             =   3375
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   688
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
         Connect         =   $"frmTitu.frx":10AB
         OLEDBString     =   $"frmTitu.frx":1133
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Titulos"
         Caption         =   "Navegador"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList Imagens 
         Left            =   4710
         Top             =   5880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTitu.frx":11BB
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTitu.frx":14D7
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTitu.frx":17F3
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTitu.frx":1B0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTitu.frx":1E2B
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frmTitu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Deletando As Boolean

Private Sub datGeral_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    AtualizaRelacao
End Sub

Private Sub grdDiretores_AfterUpdate()
    AtualizaRelacao
End Sub

Private Sub grdDiretores_BeforeDelete(Cancel As Integer)
If MsgBox("Deletar o diretor " & datDiretor.Recordset("Nome") & " " & datDiretor.Recordset("Sobrenome") & " ?", vbYesNo, Caption) = vbNo Then
    Cancel = True
    Exit Sub
End If
Cancel = False
Deletando = True
End Sub

Private Sub grdDiretores_BeforeUpdate(Cancel As Integer)
    If IsNull(Me.datGeral.Recordset("Codigo")) Then
        MsgBox "Não existe um título cadastrado", , Caption
        Cancel = True
    End If
    If Not Deletando Then
        Me.datDiretor.Recordset("CodFilme") = Me.datGeral.Recordset("Codigo")
        Deletando = False
    End If
End Sub

Private Sub grdDiretores_Error(ByVal DataError As Integer, Response As Integer)
    Response = vbDataErrContinue
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Adicionar"
            Me.datGeral.Recordset.AddNew
        Case "Salvar"
            Me.datGeral.Recordset.UpdateBatch adAffectAllChapters
    End Select
End Sub

Sub AtualizaRelacao()
    If Me.datGeral.Recordset.EOF Then Exit Sub
    If Me.datGeral.Recordset.BOF Then Exit Sub
    Me.datDiretor.RecordSource = "SELECT * FROM Diretores WHERE CodFilme = " & Me.datGeral.Recordset("Codigo")
    Me.datDiretor.Refresh
End Sub
