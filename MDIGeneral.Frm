VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H000000C0&
   Caption         =   "MEDINAFARMA"
   ClientHeight    =   12885
   ClientLeft      =   90
   ClientTop       =   735
   ClientWidth     =   19755
   Icon            =   "MDIGeneral.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIGeneral.frx":57E2
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   14895
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      Begin VB.CommandButton TXTCIA 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   6720
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton cmdbuscacia 
         Height          =   350
         Left            =   6720
         Picture         =   "MDIGeneral.frx":55FF7
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         Height          =   500
         Left            =   15
         TabIndex        =   2
         Top             =   -120
         Width           =   11880
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Compañia Activa :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   135
            Width           =   1530
         End
         Begin VB.Label gcodusu 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   300
            Left            =   9960
            TabIndex        =   8
            Top             =   150
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "User :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   9240
            TabIndex        =   7
            Top             =   135
            Width           =   615
         End
         Begin VB.Label gfecha 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   7920
            TabIndex        =   6
            Top             =   135
            Width           =   1215
         End
         Begin VB.Label glblFecha 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Fecha :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   7100
            TabIndex        =   5
            Top             =   135
            Width           =   750
         End
         Begin VB.Label lblciagen 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   300
            Left            =   1755
            TabIndex        =   3
            Top             =   135
            Width           =   4890
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   28560
      _ExtentX        =   50377
      _ExtentY        =   688
      ButtonWidth     =   661
      ButtonHeight    =   635
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   30
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "FORMGEN"
            Object.ToolTipText     =   "Operaciones del Sistema"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Tablas"
            Object.ToolTipText     =   "Mant. Tablas"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ARTI"
            Object.ToolTipText     =   "Maestro de Articulos"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PRECIOS"
            Object.ToolTipText     =   "Actualización de Precios"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CLIENTES"
            Object.ToolTipText     =   "Maestro de Clientes"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "BANCOS"
            Object.ToolTipText     =   "Registro de Bancos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CUENTAS"
            Object.ToolTipText     =   "Cta. Cte. de Clientes y Proveedores"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CHEQUES"
            Object.ToolTipText     =   "Ctas. Cte. de Bancos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "KARDEX"
            Object.ToolTipText     =   "Kardex de Articulos"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Password"
            Object.ToolTipText     =   "Cambio de Claves de Usuarios"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "consultas"
            Object.ToolTipText     =   "Consultas de Documentos"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ALTERNO"
            Object.ToolTipText     =   "Modo Alterno"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Excel"
            Object.ToolTipText     =   "Calculadora"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CAL"
            Object.ToolTipText     =   "Suma y Resta"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COTI"
            Object.ToolTipText     =   "Cotizaciones"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ORDEN"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MSGBOX"
            Object.ToolTipText     =   "Registro de Sucesos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tope"
            Object.ToolTipText     =   "Editar Tope de Pago Diario"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CONSUL"
            Object.ToolTipText     =   "Consulta de Precios Competencia."
            ImageIndex      =   8
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "INCENTIVO"
            Object.ToolTipText     =   "Lista de Incentivos..."
            ImageIndex      =   11
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "AVISO"
            Object.ToolTipText     =   "Parar Impresión..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FECHAP"
            Object.ToolTipText     =   "Cambio Fecha de Proceso...!!!"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "IA"
            Object.ToolTipText     =   "Lista de Precios <= Costo"
            ImageIndex      =   15
         EndProperty
      EndProperty
      Begin VB.TextBox txtderechos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8520
         TabIndex        =   12
         Text            =   "Identificando sus accesos al sistema..."
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
      End
      Begin MSComctlLib.ProgressBar ProBar 
         Height          =   180
         Left            =   8520
         TabIndex        =   11
         Top             =   180
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3555
      Top             =   1335
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":56439
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1425
      Top             =   1545
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5688D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":56CE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":57135
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":57589
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":579DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":57E31
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":58285
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":586D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":58B2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":58E49
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5929D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":596F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":59B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":59F99
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5A3ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5A841
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5AC95
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIGeneral.frx":5B0E9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   135
      Top             =   2070
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   135
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport TablasC 
      Left            =   615
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\Archivos de programa\PROYECTO WINDOWS NT\tablas.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu menuTit1 
      Caption         =   "&Maestros"
      Enabled         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Unidades y Precios de  Articulos"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Datos de Cl&ientes y Proveedores"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^I
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Datos de Ven&dedores"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^D
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Definición de &Bancos"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^B
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas del Sistema"
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "C&ompañia"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Empresa &General"
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Tablas"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Definición de Transacciones"
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Calendario de Sistema /T. Cambio"
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Paquetes"
         Index           =   11
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Definir Incentivos"
         Index           =   12
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Cam&biar Password"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "Cambiar &Usuario"
         Enabled         =   0   'False
         Index           =   15
         Shortcut        =   ^U
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Reiniciar Aplicación"
         Enabled         =   0   'False
         Index           =   16
         Shortcut        =   ^R
      End
      Begin VB.Menu SubmenuTit1 
         Caption         =   "&Salir del Sistema"
         Enabled         =   0   'False
         Index           =   17
         Shortcut        =   +^{F9}
      End
   End
   Begin VB.Menu menuTit2 
      Caption         =   "&Procesos"
      Enabled         =   0   'False
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Operaciones"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^G
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Generar Pedidos de Ventas"
         Index           =   1
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Modulos de Letras"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^H
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Transfer. Exportar Datos"
         Index           =   4
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Tipo de Cambio y Costeo "
         Index           =   6
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit2 
         Caption         =   "&Cierre de Operaciones del Día"
         Enabled         =   0   'False
         Index           =   8
      End
   End
   Begin VB.Menu menugrifo 
      Caption         =   "&Grifos"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenugrifo 
         Caption         =   "&Definición de Turnos"
         Index           =   0
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Islas y Responsables"
         Index           =   1
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Parte Diario de Grifero"
         Index           =   2
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Stock por Regla"
         Index           =   3
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu submenugrifo 
         Caption         =   "&Transacciones II"
         Index           =   5
      End
   End
   Begin VB.Menu menutit3 
      Caption         =   "&Contabilidad General"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenutit3 
         Caption         =   "Definición de &Estructura"
         Index           =   0
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Plan contable"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Gestion de Vouchers"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu submenutit3 
         Caption         =   "&Mayorizacion "
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Cierre &Del Periodo de Trabajo "
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu submenutit3 
         Caption         =   "Pase a Con&tabilidad"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu menutit4 
      Caption         =   "&Emitir Reportes"
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Ventas"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu menuVent 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Compras"
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
         Begin VB.Menu Menucomp 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Principales"
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu menuAlm 
            Caption         =   " "
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Inventarios"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
         Begin VB.Menu menuMoli 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "C&ontabilidad"
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
         Begin VB.Menu MenuContab 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 1 "
         Enabled         =   0   'False
         Index           =   10
         Visible         =   0   'False
         Begin VB.Menu menudis1 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 2"
         Enabled         =   0   'False
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu menudis2 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 3"
         Enabled         =   0   'False
         Index           =   14
         Visible         =   0   'False
         Begin VB.Menu menudis3 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 9"
         Enabled         =   0   'False
         Index           =   16
         Visible         =   0   'False
         Begin VB.Menu menudis9 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   17
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 10"
         Enabled         =   0   'False
         Index           =   18
         Visible         =   0   'False
         Begin VB.Menu menudis10 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   19
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 11"
         Enabled         =   0   'False
         Index           =   20
         Visible         =   0   'False
         Begin VB.Menu menudis11 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   21
         Visible         =   0   'False
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "&Disponible 12"
         Enabled         =   0   'False
         Index           =   22
         Visible         =   0   'False
         Begin VB.Menu menudis12 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu SubmenuTit4 
         Caption         =   "-"
         Index           =   23
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menutit5 
      Caption         =   "&Utilidades"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu submenutit5 
         Caption         =   "&Backup"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Mostrar ultimo Reporte"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Configuración de Reporte"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Grupo de Trabajso"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu submenutit5 
         Caption         =   "&Reporte de Transacciones"
         Enabled         =   0   'False
         Index           =   4
      End
   End
   Begin VB.Menu menutit6 
      Caption         =   "&Consultas"
      Enabled         =   0   'False
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Kardex de Articulos"
         Enabled         =   0   'False
         Index           =   0
         Shortcut        =   ^K
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Cta.Ctes. Clientes y Proveedores"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   ^E
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Cta.Ctes. de &Bancos"
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^M
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Operaciones"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^N
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Detalle de Kardex de Lotes"
         Enabled         =   0   'False
         Index           =   6
         Shortcut        =   ^L
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Plantilla Stock Maximos / Escalas por Lineas"
         Enabled         =   0   'False
         Index           =   8
         Shortcut        =   ^S
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Pedidos a Proveedores"
         Enabled         =   0   'False
         Index           =   10
         Shortcut        =   ^P
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "&Actualización Automatico de Stock"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Documentos Anulados"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   14
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Establecer Cuotas de Vendedores"
         Index           =   15
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   16
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Copiar Estructura de Presupuestos"
         Index           =   17
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "-"
         Index           =   18
      End
      Begin VB.Menu SubmenuTit6 
         Caption         =   "Generar Presupuesto de Pedidos"
         Index           =   19
      End
   End
   Begin VB.Menu menutit7 
      Caption         =   "&Opciones"
      Begin VB.Menu submenutit7 
         Caption         =   "&Actualizar Precios en Tiempo Real"
         Index           =   0
         Shortcut        =   ^T
      End
      Begin VB.Menu submenutit7 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu submenutit7 
         Caption         =   "&Cotizaciones de Ventas"
         Index           =   2
         Shortcut        =   ^Z
      End
      Begin VB.Menu submenutit7 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu submenutit7 
         Caption         =   "&Generar Ordenes de Compra"
         Index           =   4
         Shortcut        =   ^O
      End
      Begin VB.Menu submenutit7 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu submenutit7 
         Caption         =   "&Bloqueo de Periodos"
         Index           =   6
         Shortcut        =   ^Q
      End
      Begin VB.Menu submenutit7 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu submenutit7 
         Caption         =   "&Generar Pase a Contablidad"
         Index           =   8
      End
   End
   Begin VB.Menu menuConsultorio 
      Caption         =   "Con&sultorio"
      Begin VB.Menu submenuConsultorio 
         Caption         =   "&Reservar Cita Medica"
      End
   End
   Begin VB.Menu menuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu menuAyudaSistema 
         Caption         =   "&Quienes somos..."
      End
      Begin VB.Menu menuRepoAyuda 
         Caption         =   "&Acerca de..."
      End
   End
   Begin VB.Menu Menucompania 
      Caption         =   "compania"
      Visible         =   0   'False
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu menuCia 
         Caption         =   ""
         Index           =   40
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Wfile As String
Dim WFORM As String
Public Sub LLENA_ACCESO_OTROS(WSOTROS As String)
Dim W1 As String * 2
Dim I, wPosF, WPosV, cuenta As Integer
Dim sal As Boolean
Dim cade As String
Dim WNUM As Integer
Dim F As Integer
Dim a As Integer
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(WSOTROS)
cade = Trim(WSOTROS)
cuenta = 0
wPosF = 1
a = 0
For I = 1 To Len(cade)
 If Mid(cade, I, 1) = "." Then
   a = a + 1
 End If
Next I
lk_OTROS_Count = a
ReDim lk_OTROS(a)
Do Until cuenta = a
   cuenta = cuenta + 1
   wPosF = InStr(wPosF, cade, ".", 1) + 1
   WNUM = Mid(cade, wPosF, 2)
   If Right(WNUM, 1) = "." Then
     WNUM = Left(WNUM, 2)
     wPosF = wPosF - 1
   End If
   lk_OTROS(cuenta) = Format(CStr(WNUM), "00")
Loop
End Sub
Private Sub cmdbuscacia_Click()
  MDIForm1.PopupMenu Menucompania, , 1800, 6300
End Sub


Private Sub lblciagen_Click()
 MDIForm1.PopupMenu Menucompania, , 1800, 6300
End Sub
Private Sub MDIForm_Load()
Dim WVALORACT As Boolean
Dim ciades As String * 80
Dim WUSU_CIAS As String
Dim WUSU_OTROS As String
Dim h, a, b, C, d, E As Integer
Dim a1  As Integer
Dim a2 As Integer
Dim a3 As Integer
Dim a4 As Integer

Dim wBAN As Integer
Dim WNUM, wPosF, WPosV, cuenta, cade
Dim wflag As String * 1
Dim F, g, fg As Integer
Dim wserie As Currency
Dim wfecha
Dim WS
LogInfo "MDI: Inicio MDIForm_Load"
cmdbuscacia.Enabled = False
ciades = ""
'On Error GoTo SALE
wflag = ""
TXTCIA.Enabled = False
PS_GEN(0) = 0
LogInfo "MDI: Antes de GEN.Requery"
GEN.Requery
LogInfo "MDI: Despues de GEN.Requery"
Pub_Titulo = Trim(GEN!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
LogInfo "MDI: Ruta Reportes = " & PUB_RUTA_REPORTE
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))
LK_IGV = Nulo_Valor0(GEN!GEN_IGV)
LK_DIG_RUC = Nulo_Valor0(GEN!GEN_DIG_RUC)
LK_TIPO_CAMBIO = Nulo_Valor0(GEN!gen_tipo_cambio)
LK_ACTIVA = Nulo_Valors(GEN!GEN_ACTIVA)
LK_FLAG_SOS = Nulo_Valors(GEN!gen_bloqueo)
'LK_RUTA_DOC = "C:\ADMIN\"

'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
If LK_FLAG_GRIFO = "A" Then
  MDIForm1.menugrifo.Visible = True
  MDIForm1.menugrifo.Enabled = True
Else
  MDIForm1.menugrifo.Visible = False
  MDIForm1.menugrifo.Enabled = False
End If

'TXTCIA.Visible = False
DoEvents
LogInfo "MDI: Antes de Splash.Hide"
Splash.Hide
LogInfo "MDI: Despues de Splash.Hide"
DoEvents
MDIForm1.Show
DoEvents
NOTPermisos
LogInfo "MDI: Antes de PASSFORM.Show"
PASSFORM.Show 1
DoEvents
Screen.MousePointer = 11
probar.Value = 0
probar.Min = 0
probar.max = 15
txtderechos.Visible = True
DoEvents
probar.Visible = True
DoEvents
If LK_CODCIA = "" Or LK_CODUSU = "" Then
  MsgBox "Existe un Problema con su Acceso , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

'StatusBar1.Enabled = False
PS_PAR(0) = " "
par.Requery
If par.EOF Then
  GoTo PASAR
End If
'If LK_CODUSU <> "ADMIN" Then
' MDIForm1.submenutit5(2).Visible = False
 'MDIForm1.submenutit5(3).Visible = False
'Else
' MDIForm1.submenutit5(2).Visible = True
' MDIForm1.submenutit5(3).Visible = True
'End If
probar.Value = probar.Value + 1
PSUSU_LLAVE(0) = LK_CODUSU
usu_llave.Requery
USU_ACC_TIPDOC = ""
If Not usu_llave.EOF Then
   WUSU_CIAS = Trim(Nulo_Valors(usu_llave!USU_CIAS))
   WUSU_OTROS = Trim(Nulo_Valors(usu_llave!USU_OTROS))
   LK_PRECIO = Trim(Nulo_Valors(usu_llave!USU_PRECIO))
   USU_ACC_TIPDOC = Trim(Nulo_Valors(usu_llave!USU_ACCTIPDOC))
   lk_acc_repoall = Nulo_Valor0(usu_llave!USU_REPOALL)
   pub_codvenrel = ""
   If Val(usu_llave!USU_CODVENREL) <> 0 Then
    pub_codvenrel = pub_codvenrel & usu_llave!USU_CODVENREL & ""
   End If
   If Val(usu_llave!USU_CODVENREL2) <> 0 Then
    pub_codvenrel = pub_codvenrel & "," & usu_llave!USU_CODVENREL2 & ""
   End If
   If Val(usu_llave!USU_CODVENREL3) <> 0 Then
    pub_codvenrel = pub_codvenrel & "," & usu_llave!USU_CODVENREL3 & ""
   End If
   If Val(usu_llave!USU_CODVENREL4) <> 0 Then
    pub_codvenrel = pub_codvenrel & "," & usu_llave!USU_CODVENREL4 & ""
   End If
End If
If Len(WUSU_CIAS) = 0 Then
  MsgBox "Usted NO Tiene accesos a ninguna compañia , consulte al Administrador ", 48, Pub_Titulo
  GoTo fin_de_pro
End If

If WUSU_OTROS = "" Then
 lk_OTROS_Count = 0
 ReDim lk_OTROS(0)
Else
 LLENA_ACCESO_OTROS WUSU_OTROS
End If
'LK_FLAG_SOS = Trim(Nulo_Valors(usu_llave!USU_FLAG_CONT))
MDIForm1.Toolbar1.Buttons(20).Visible = False
MDIForm1.Toolbar1.Buttons(28).Visible = False
MDIForm1.Toolbar1.Buttons(30).Visible = False
MDIForm1.Toolbar1.Buttons(20).Visible = True
MDIForm1.Toolbar1.Buttons(30).Visible = True
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "SUPER" Then
   If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
     MDIForm1.Toolbar1.Buttons(28).Visible = True

   End If
End If

MDIForm1.Toolbar1.Buttons(3).Enabled = True
probar.Value = probar.Value + 1
WNUM = 0
wPosF = 0
WPosV = 0
cuenta = 0
WPosV = Len(WUSU_CIAS)
cade = Trim(WUSU_CIAS)
cuenta = 0
wPosF = 1
a = 0
probar.Value = probar.Value + 1
For fila = 1 To Len(cade)
    If Mid(cade, fila, 1) = "." Then
        a = a + 1
    End If
Next fila
probar.Value = probar.Value + 1
wBAN = -1
h = 0
Do Until cuenta = a
    h = h + 1
    cuenta = cuenta + 1
    DoEvents
    wPosF = InStr(wPosF, cade, ".", 1) + 1
    DoEvents
    WNUM = Mid(cade, wPosF, 2)
    If Right(WNUM, 1) = "." Then
        WNUM = Left(WNUM, 2)
        wPosF = wPosF - 1
    End If
    PUB_CODCIA = Format(WNUM, "00")
    par.MoveFirst
    Do While Not par.EOF
      If Trim(par!PAR_CODCIA) = PUB_CODCIA Then Exit Do
      par.MoveNext
    Loop
    If Not par.EOF Then
       ciades = PUB_CODCIA & "-" & Trim(par!par_nombre)
       MDIForm1.menuCia(cuenta - 1).Caption = ciades
       MDIForm1.menuCia(cuenta - 1).Visible = True
       If LK_CODCIA = par!par_codcia Then wBAN = h - 1
    Else
       If a = 1 Then ' Cuando la unica compañia de acceso no exite
         MsgBox "Su Compañia de Acceso NO Existe , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
         GoTo fin_de_pro
       End If
       MDIForm1.menuCia(cuenta - 1).Caption = "* " & PUB_CODCIA & " , NO Procede"
       MDIForm1.menuCia(cuenta - 1).Visible = True
    End If
Loop
probar.Value = probar.Value + 1
If wBAN = -1 Then
  For a = 0 To h - 1
    If Left(MDIForm1.menuCia(a).Caption, 1) <> "*" Then
        wBAN = a
        Exit For
    End If
  Next a
  If wBAN = -1 Then
    MsgBox "Su(s) Compañias de Acceso NO Existen , consulte al Administrador, Procede a Salir del Sistema", 48, Pub_Titulo
    GoTo fin_de_pro
  End If
End If

For a = h To 40
  MDIForm1.menuCia(a).Visible = False
Next a
MDIForm1.menuCia(wBAN).Checked = True


probar.Value = probar.Value + 1
PASAR:
PS_GEN(0) = 0
GEN.Requery
probar.Value = probar.Value + 1
Pub_Titulo = Trim(GEN!GEN_NOMBRE)
PUB_RUTA_REPORTE = Trim(Nulo_Valors(GEN!gen_ruta_reportes))
PUB_RUTA_OTRO = Trim(Nulo_Valors(GEN!gen_ruta_otros))
'MDIForm1.StatusBar1.Panels(1).Text = "Tipo Cambio : S/." & Format(LK_TIPO_CAMBIO, "0.0000")
MDIForm1.gfecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'MDIForm1.StatusBar1.Panels(3).Text = Trim(GEN!gen_nombre)
MDIForm1.gcodusu.Caption = LK_CODUSU
'MDIForm1.StatusBar1.Panels(5).Text = MDIForm1.menuCia(wBAN).Caption
probar.Value = probar.Value + 1
MDIForm1.TXTCIA.Caption = Trim(MDIForm1.menuCia(wBAN).Caption)
MDIForm1.lblciagen.Caption = Trim(MDIForm1.menuCia(wBAN).Caption)
CONST_SERVER = Trim(Nulo_Valors(GEN!GEN_SERVER))
CONST_UID = Trim(Nulo_Valors(GEN!GEN_UID))
CONST_PWD = Trim(Nulo_Valors(GEN!GEN_PWD))
probar.Value = probar.Value + 1
If Len(CONST_SERVER) = 0 Or Len(CONST_UID) = 0 Then
  MsgBox "ADVERTENCIA ..!! DEBES DEFINIR EL NOMBRE DEL SERVIDOR , UID , PWD", 48, Pub_Titulo
End If
PUB_ODBC = "ODBC;DSN=OSBusinessCR;UID=" & CONST_UID & ";PWD=" & CONST_PWD & ";APP=Data From Designer;WSID=" & CONST_SERVER & ";LANGUAGE=us_english;DATABASE=BDATOS"
'CONn$ = "dsn=" & wdsn & ";uid=sa;pwd=" & wAcceso & ";database=bdatos2;"
' * Permisos para los Menus *
'Debug.Print PUB_ODBC

'asegura la compañia por defecto
ACTUALIZA_CIA Left(MDIForm1.menuCia(wBAN).Caption, 2)

If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
  Screen.MousePointer = 0
  Load frmclave2
  frmclave2.Show 1
End If
If LK_CODUSU = "SUPER" Then
   LK_ACCESO_REPORT = "A"
   LK_ACCESO_EXTORNO = "A"
End If

Screen.MousePointer = 11
Permisos
probar.Value = probar.Value + 1
Screen.MousePointer = 11


' * Carga menu de Reporte *
SQ_OPER = 2
PUB_INICIO = 5999
LEER_TRA_LLAVE
probar.Value = probar.Value + 1
a = 0
b = 0
C = 0
d = 0
E = 0
F = 0
g = 0
fg = 0
a1 = 0
a2 = 0
a3 = 0
a4 = 0
Do Until tra_menu.EOF
 If tra_menu!TRA_FLAG_EXT <> "A" Then
   GoTo OTRO_MENU
 End If
 If LK_FLAG_SOS = "A" And tra_menu!TRA_ACT14 = 1 Then GoTo OTRO_MENU
 pub_cadena = "select  * from negarepo WHERE ACC_CODUSU = '" & LK_CODUSU & "' AND ACC_CODTRA = " & tra_menu!TRA_KEY & " AND ACC_ACTIVO =1"
 Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
 If X.EOF Then
 WVALORACT = True
 Else
 WVALORACT = False
 End If
 If tra_menu(0) >= 8100 And tra_menu(0) < 8299 Then
     If a > 0 Then
        Load MDIForm1.menuVent(a)
     End If
     MDIForm1.menuVent(a).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
     MDIForm1.menuVent(a).Enabled = WVALORACT
     
     a = a + 1
 ElseIf tra_menu(0) >= 8300 And tra_menu(0) < 8499 Then
    If b > 0 Then
        Load MDIForm1.Menucomp(b)
    End If
    MDIForm1.Menucomp(b).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.Menucomp(b).Enabled = WVALORACT
    b = b + 1
 ElseIf tra_menu(0) >= 8500 And tra_menu(0) < 8699 Then
    If C > 0 Then
        Load MDIForm1.menuAlm(C)
    End If
    MDIForm1.menuAlm(C).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menuAlm(C).Enabled = WVALORACT
        
    C = C + 1
 ElseIf tra_menu(0) >= 8700 And tra_menu(0) < 8899 Then
    If d > 0 Then
        Load MDIForm1.menuMoli(d)
    End If
    MDIForm1.menuMoli(d).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menuMoli(d).Enabled = WVALORACT
    d = d + 1
 ElseIf tra_menu(0) >= 8900 And tra_menu(0) < 9099 Then
    If E > 0 Then
        Load MDIForm1.MenuContab(E)
    End If
    MDIForm1.MenuContab(E).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.MenuContab(E).Enabled = WVALORACT
    E = E + 1
 ElseIf tra_menu(0) >= 9100 And tra_menu(0) < 9299 Then
    If F > 0 Then
        Load MDIForm1.menudis1(F)
    End If
    MDIForm1.menudis1(F).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis1(F).Enabled = WVALORACT
    F = F + 1
 ElseIf tra_menu(0) >= 9300 And tra_menu(0) < 9499 Then
    If g > 0 Then
        Load MDIForm1.menudis2(g)
    End If
    MDIForm1.menudis2(g).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis2(g).Enabled = WVALORACT
    g = g + 1
 ElseIf tra_menu(0) >= 9500 And tra_menu(0) < 9699 Then
    If fg > 0 Then
        Load MDIForm1.menudis3(fg)
    End If
    MDIForm1.menudis3(fg).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis3(fg).Enabled = WVALORACT
    fg = fg + 1
  ElseIf tra_menu(0) >= 6000 And tra_menu(0) < 6199 Then
    If a1 > 0 Then
        Load MDIForm1.menudis9(a1)
    End If
    MDIForm1.menudis9(a1).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis9(a1).Enabled = WVALORACT
    a1 = a1 + 1
 ElseIf tra_menu(0) >= 6200 And tra_menu(0) < 6399 Then
    If a2 > 0 Then
        Load MDIForm1.menudis10(a2)
    End If
    MDIForm1.menudis10(a2).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis10(a2).Enabled = WVALORACT
    
    a2 = a2 + 1
 ElseIf tra_menu(0) >= 6400 And tra_menu(0) < 6599 Then
    If a3 > 0 Then
        Load MDIForm1.menudis11(a3)
    End If
    MDIForm1.menudis11(a3).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis11(a3).Enabled = WVALORACT
    a3 = a3 + 1
 ElseIf tra_menu(0) >= 6600 And tra_menu(0) < 6799 Then
    If a4 > 0 Then
        Load MDIForm1.menudis12(a4)
    End If
    MDIForm1.menudis12(a4).Caption = Trim(CStr(tra_menu(0))) & " - &" & Trim(tra_menu(1))
    MDIForm1.menudis12(a4).Enabled = WVALORACT
    a4 = a4 + 1
 End If


OTRO_MENU:
 tra_menu.MoveNext
Loop
SQ_OPER = 2
PUB_TIPREG = 99
PUB_CODCIA = "00"
LEER_TAB_LLAVE
If tab_mayor.EOF Then
 MsgBox " No existe TIPREG 99 en tablas crearlas con titulos del Menu correspondiente..", 48, Pub_Titulo
End If
fila = 0
SQ_OPER = 0
Do Until tab_mayor.EOF
 MDIForm1.SubmenuTit4(fila).Caption = "&" + Trim(tab_mayor!TAB_NOMLARGO)
 If Val(Nulo_Valor0(tab_mayor!tab_nomcorto)) = 1 Then
  MDIForm1.SubmenuTit4(fila).Visible = True
  MDIForm1.SubmenuTit4(fila + 1).Visible = True
  SQ_OPER = 9
 End If
 
 fila = fila + 2
 tab_mayor.MoveNext
Loop
If SQ_OPER = 9 Then
 MDIForm1.SubmenuTit4(23).Visible = False
End If

If LK_CODUSU <> "ADMIN" Then
 MDIForm1.SubmenuTit1(8).Caption = "&Usuarios"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. Usuarios"
Else
 MDIForm1.SubmenuTit1(8).Caption = "&Tablas"
 MDIForm1.Toolbar1.Buttons.Item(2).ToolTipText = "Mant. TABLAS"
End If
If LK_FLAG_SOS = "A" Then
 MDIForm1.SubmenuTit1(3).Visible = False
End If
If wflag = "A" Then
 MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
End If

probar.Value = probar.Value + 1
DoEvents
'StatusBar1.Enabled = True
'DoEvents

menutit7.Enabled = True
a3 = 0
flag_autoriza_pase = ""
For fila = 1 To lk_OTROS_Count
 If Val(lk_OTROS(fila)) = 22 Then ' ACCESO AL MENU DE OPCIONES
    a3 = 1
 End If
 If Val(lk_OTROS(fila)) = 71 Then flag_autoriza_pase = "A" ' ACCESO AL PASE CONTABLE
Next fila

If (LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER") Then

Else
If a3 = 0 Then menutit7.Enabled = False
End If
If iFormCount = 1 Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then FrmCal.lblacc.Visible = False


MDIForm1.Visible = True
DoEvents

MDIForm1.Show
DoEvents
probar.Value = probar.Value + 1
Screen.MousePointer = 11
PASSFORM.Hide
WF_REINICIAR = 0
probar.Min = 0
probar.Visible = False
txtderechos.Visible = False
'TXTCIA.Visible = True
Screen.MousePointer = 0
cmdbuscacia.Enabled = True
TXTCIA.Enabled = True
If TXTCIA.Enabled And TXTCIA.Visible Then TXTCIA.SetFocus


If LK_EMP = "HERuu" Then
  wfecha = #12/12/2001#
  If LK_FECHA_DIA >= wfecha Then
  If Val(GEN!GEN_CBS) <> -9 Then
sigue:
  '  WS = InputBox("Nro  Serial es : " & Format(GEN!GEN_CFP, "0000000000000000") & " , Llame al proveedor y pida su nueva serie." & Chr(13) & "{Sro  intenta colocar una serie que no es, por quinta vez... el Proveedor no responde por la informaciro n almacena.}", "Acceso Limitado ....")
  '  If WS = "" Then
      GoTo sigue
  '  End If
    wserie = Val(WS)
    If Val(PRO_DES(GEN!GEN_CFP)) = Val(wserie) Then
       GEN.Edit
       GEN!GEN_CBS = -9
       GEN.Update
    Else
      MsgBox "LLame a su Proveedor de OSBussines y pida su Nro  de Serie. Ingreso Incorrecto" & Chr(13) & "{No intentar si no tiene el nuevo Nro. de Serie.}", vbCritical, "Acceso limitado .... "
      End
    End If
  End If
  End If
End If


Exit Sub
sale:
   If Err.Number = 360 Then
     Resume Next
   Else
      MsgBox Err.Description, 48, Pub_Titulo
   End If
   Resume Next
Exit Sub
fin_de_pro:
    Screen.MousePointer = 11
    DoEvents
    MDIForm1.WindowState = 1
    MDIForm1.Caption = "Descargando Modulos..."
    EN.Close
    CN.Close
    Pub_ConnAdo.Close
    Screen.MousePointer = 0
    End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If LK_EMP = "PLA" Then
  If XLL Is Nothing Then
  Else
        XLL.Workbooks.Close
        XLL.Application.Quit
    End If
  Set XLL = Nothing
End If
If WF_REINICIAR <> 2 Then
    Screen.MousePointer = 11
    MDIForm1.Caption = "Descargando Modulos..."
    DoEvents
    MDIForm1.WindowState = 1
    EN.Close
    CN.Close
    Pub_ConnAdo.Close
    Unload MDIForm1
    Screen.MousePointer = 0
    End
End If
WF_REINICIAR = 0
End Sub
Private Sub menuAlm_Click(Index As Integer)
On Error GoTo sale
If Trim(menuAlm(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

pu_codtra = Val(Left(menuAlm(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Description
 End If
 Screen.MousePointer = 0
End Sub


Private Sub menuAyudaSistema_Click()
MsgBox "OSBusiness - Apisoft S.A.C. - . ", vbInformation, "Apisoft"
End Sub

Private Sub menuCia_Click(Index As Integer)
Dim iFormCount As Integer
On Error GoTo sigue
If Left(MDIForm1.menuCia(Index).Caption, 1) = "*" Then
  Exit Sub
End If
Screen.MousePointer = 11
If Forms.count - 1 > 0 Then
   For iFormCount = Forms.count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
For fila = 0 To 20
  MDIForm1.menuCia(fila).Checked = False
Next fila
MDIForm1.menuCia(Index).Checked = True
'MDIForm1.StatusBar1.Panels(7).Text = MDIForm1.menuCia(Index).Caption
MDIForm1.TXTCIA.Caption = Trim(MDIForm1.menuCia(Index).Caption)
MDIForm1.lblciagen.Caption = Trim(MDIForm1.menuCia(Index).Caption)
ACTUALIZA_CIA Left(MDIForm1.menuCia(Index).Caption, 2)
CONTA = 0
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until CONTA = 5
    CONTA = CONTA + 1
    If tab_mayor.EOF Then
       tab_precioss(CONTA) = "....."
    Else
       tab_precioss(CONTA) = tab_mayor!TAB_NOMLARGO
       tab_mayor.MoveNext
    End If
Loop
Screen.MousePointer = 0
Exit Sub
sigue:
Resume Next
End Sub
Private Sub menuComp_Click(Index As Integer)
On Error GoTo sale
If Trim(Menucomp(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

pu_codtra = Val(Left(Menucomp(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub

Private Sub MenuContab_Click(Index As Integer)
On Error GoTo sale
If Trim(MenuContab(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

pu_codtra = Val(Left(MenuContab(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0
End Sub

Private Sub menuctacte_Click()
Screen.MousePointer = 11
Load FrmCaa
FrmCaa.Show
Screen.MousePointer = 0
End Sub

Private Sub menuch_Click()
If LK_FLAG_SOS = "A" Then
  Exit Sub
End If
Screen.MousePointer = 11
Load frm_cheques
Screen.MousePointer = 0
frm_cheques.Show

End Sub

Private Sub menudis1_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis1(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis1(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis10_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis10(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis10(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0



End Sub

Private Sub menudis11_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis11(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis11(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0



End Sub

Private Sub menudis12_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis12(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis12(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0



End Sub

Private Sub menudis2_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis2(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis2(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menudis3_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis3(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis3(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0

End Sub

Private Sub menuKa_Click()
Screen.MousePointer = 11
Load KARDEX
KARDEX.Show
Screen.MousePointer = 0
End Sub

Private Sub menudis9_Click(Index As Integer)
On Error GoTo sale
If Trim(menudis9(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menudis9(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 Else
   MsgBox Err.Number
 End If
 Screen.MousePointer = 0


End Sub

Private Sub menuMoli_Click(Index As Integer)
If Trim(menuMoli(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

pu_codtra = Val(Left(menuMoli(Index).Caption, 4))
SQ_OPER = 3
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
 FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
End Sub

Private Sub menuRepoAyuda_Click()
Dim WL As Object
Dim var As String
'Dim ee As Words
MsgBox "En Construcciro n... " & Chr(13) & "Para mayor informaciro n llamar al Telf. Celular: (044) 9604779 Nextel 829*6446 - Trujillo - Perú", vbInformation, "Asesoria y Soporte"
Exit Sub
Dim ww As Form
Dim I
On Error GoTo OJO
var = Trim(MDIForm1.ActiveForm.Name)
If UCase(var) = "FORMGEN" Then
  var = Trim(MDIForm1.ActiveForm.Name) & "_" & LK_CODTRA
End If
On Error GoTo 0
If WL Is Nothing Then
    Set WL = CreateObject("word.Application")
End If
var = Left(PUB_RUTA_OTRO, 2) + "\admin\ayudas\" & var
On Error GoTo no_existe
WL.Application.WindowState = wdWindowStateMaximize
'wl.Application.ActiveDocument.Windows(1).WindowState = 2
'ee.Windows(0).WindowState = wdWindowStateMaximize
'wl.Windows(1).WindowState = wdWindowStateMaximize
'wl.Documents
WL.Documents.Open FileName:=var
WL.Application.Visible = True
Set WL = Nothing
On Error GoTo 0
Exit Sub
no_existe:
'Shell "C:\Archivos de programa\Microsoft Office\Office\winword.EXE", "c:\admin\ayudas\FORMGEN_2401.doc""
MsgBox Err.Description
OJO:
If Err.Number <> 91 Then
 MsgBox "Verificar " & Err.Description, 48, Pub_Titulo
ElseIf Err.Number = 91 Then
  var = "MDIForm1"
  On Error GoTo 0
  If WL Is Nothing Then
    Set WL = CreateObject("word.Application")
  End If
  var = Left(PUB_RUTA_OTRO, 2) + "\admin\ayudas\" & var
  On Error GoTo no_existe2
  WL.Application.WindowState = wdWindowStateMaximize
  WL.Documents.Open FileName:=var
  WL.Application.Visible = True
  On Error GoTo 0
End If
no_existe2:
End Sub

Private Sub menuVent_Click(Index As Integer)
On Error GoTo sale
If Trim(menuVent(Index).Caption) = "" Then
 Exit Sub
End If
Screen.MousePointer = 0

SQ_OPER = 3
pu_codtra = Val(Left(menuVent(Index).Caption, 4))
LEER_TRA_LLAVE
If Trim(retra_llave(7)) = "FRMIMP2" Then
 Load FrmImp2
 FrmImp2.Show
ElseIf Trim(retra_llave(7)) = "FRMIMP3" Then
 Load FrmImp3
FrmImp3.Show 1
ElseIf Trim(retra_llave(7)) = "FRMIMP" Then
 Load FRMIMP
 FRMIMP.Show 1
ElseIf Trim(retra_llave(7)) = "FRMREPO" Then
 Load FrmRepo
 FrmRepo.Show 1
ElseIf Trim(retra_llave(7)) = "RCRYSTAL" Then
 Load RCRYSTAL
 RCRYSTAL.Show
ElseIf Trim(retra_llave(7)) = "FRMIMPC1" Then
 Load FrmImpC1
 FrmImpC1.Show 1
End If
Screen.MousePointer = 0
Exit Sub
sale:
 If Err.Number = 424 Then
   MsgBox "Requiere el Formulario .", 48, Pub_Titulo
 ElseIf Err.Number = 40009 Then
   MsgBox "No se encontro el Registro del Reporte Reiniciar", 48, Pub_Titulo
 End If
 Screen.MousePointer = 0
End Sub

Private Sub submenugrifo_Click(Index As Integer)
Select Case Index
Case 0
    Screen.MousePointer = 11
    PUB_TIPREG = 2102
    PUB_CODCIA = LK_CODCIA
    Load FrmDatArti
    FrmDatArti.Caption = "TURNOS DE GRIFOS -  TAB_TIPREG = " & PUB_TIPREG
    Screen.MousePointer = 0
    FrmDatArti.Show 1
Case 1
    Screen.MousePointer = 11
    Load Frmislas
    Frmislas.Show
    Screen.MousePointer = 0
    Exit Sub
Case 2
    Screen.MousePointer = 11
    Load frmPARTE
    frmPARTE.Show
    Screen.MousePointer = 0
    Exit Sub
Case 3
    Screen.MousePointer = 11
    Load frmStockG
    Screen.MousePointer = 0
    frmStockG.Show 1
    Exit Sub
Case 5
    Screen.MousePointer = 11
    Load FORM_GRIFO
    FORM_GRIFO.Show
    Screen.MousePointer = 0
    Exit Sub
End Select


End Sub

Private Sub SubmenuTit1_Click(Index As Integer)
Dim h, wBAN As Integer
Dim iFormCount As Integer
On Error GoTo sale
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load frmARTI
        frmARTI.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load frmCLI
        frmCLI.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FrmVen
        FrmVen.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frmBancos
        Unload FORMGEN
        frmBancos.Show
        Screen.MousePointer = 0
    Case 5
        Screen.MousePointer = 11
        Load frmTablas
        frmTablas.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11

        If Forms.count - 1 > 0 Then
           For iFormCount = Forms.count - 1 To 1 Step -1
            If iFormCount <> 1 Then
                Unload Forms(iFormCount)
            End If
           Next iFormCount
        End If
        Load FrmParGen
        FrmParGen.Show
        Screen.MousePointer = 0
        
    Case 7
 '       Screen.MousePointer = 11
'        Load FrmTiempoReal
'        FrmTiempoReal.Show
'        Screen.MousePointer = 0
        Screen.MousePointer = 0
        Load frmGeneral
        frmGeneral.Show 1
     
    Case 8
        Screen.MousePointer = 11
        Load FrmTabla1
        FrmTabla1.Show
        Screen.MousePointer = 0
    Case 9
        Screen.MousePointer = 11
        Load FrmSub
        FrmSub.Show
        Screen.MousePointer = 0
    Case 10
        Screen.MousePointer = 11
        Load GENCAL
        GENCAL.Show
        Screen.MousePointer = 0
    'PAQUETE
    Case 11
     Load frmPaquete
     frmPaquete.Show
    Case Is = 12
     Load FrmIncentivos
     FrmIncentivos.Show
     Exit Sub
    Case 14
        Screen.MousePointer = 11
        Load FrmClave
        FrmClave.Show
        Screen.MousePointer = 0
    Case 15
       Dim iFrmCount As Integer
       Screen.MousePointer = 11
       For iFrmCount = Forms.count - 1 To 1 Step -1
         If iFrmCount <> 1 Then
            Unload Forms(iFrmCount)
         End If
       Next iFrmCount
       NOTPermisos
       For fila = 0 To 20
         MDIForm1.menuCia(fila).Checked = False
       Next fila
       PUB_FLAG = 1
       Screen.MousePointer = 11
       MDIForm_Load
       Screen.MousePointer = 0
    Case 16
        'REINICIAR SISTEMA
        WF_REINICIAR = 2 'PARA VERIFICAR EN MDIFORM_UNLOAD
        Screen.MousePointer = 11
        MDIForm1.Caption = "Descargando Modulos..."
        CN.Cancel
        CN.Close
        Pub_ConnAdo.Close
        MDIForm1.WindowState = 1
        MDIForm1.Hide
        Unload MDIForm1
        PUB_FLAG = 0
        Unload Splash
        Screen.MousePointer = 0
        Load Splash
    Case 17
        'SALIR
        If LK_EMP = "PLA" Then
          If XLL Is Nothing Then
          Else
                XLL.Workbooks.Close
                XLL.Application.Quit
            End If
          Set XLL = Nothing
        End If
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Pub_ConnAdo.Close
        Screen.MousePointer = 0
        End
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... ", 48, Pub_Titulo
End Select
Exit Sub
sale:
'MsgBox Err.Description
Resume Next
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit2_Click(Index As Integer)
On Error GoTo sale
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FORMGEN
        FORMGEN.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load FORM_COT
        FORM_COT.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frmletras
        Screen.MousePointer = 0
        frmletras.Show
        '  MsgBox "EN CONSTRUCCION...", 48, Pub_Titulo
        '  Exit Sub
        'End If
        'Load FORM_GRIFO
        'FORM_GRIFO.Show
    Case 4
        Screen.MousePointer = 11
        Load frmletras
        Screen.MousePointer = 0
        frmletras.Show

    Case 6
       Screen.MousePointer = 11
       Load Frmcospro
       Frmcospro.Show
       Screen.MousePointer = 0
    Case 8
       Screen.MousePointer = 11
       Load PRODIA
       Screen.MousePointer = 0
       PRODIA.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo
End Select
Exit Sub
sale:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If


End Sub


Private Sub SubmenuTit6_Click(Index As Integer)
On Error GoTo sale
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load KARDEX
        KARDEX.Show
        Screen.MousePointer = 0
    Case 1
        Screen.MousePointer = 11
        Load FrmCaa
        FrmCaa.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load frm_cheques
        frm_cheques.Show
        Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
        Load frmdocu
        frmdocu.Show
        Screen.MousePointer = 0
    Case 6
        Screen.MousePointer = 11
        Load frmManteLote
        frmManteLote.Show
        Screen.MousePointer = 0
    Case 8
        Screen.MousePointer = 11
        Load frmPlanS
        Screen.MousePointer = 0
        frmPlanS.Show 1
    Case 10
        Screen.MousePointer = 11
        Load FrmPedPro
        Screen.MousePointer = 0
        FrmPedPro.Show 1
    Case 11
        Screen.MousePointer = 11
        Load frmRegStock
        Screen.MousePointer = 0
        frmRegStock.Show
    Case 13
        Screen.MousePointer = 11
        Load frmConProAnul
        Screen.MousePointer = 0
        frmConProAnul.Show
    Case 15
        Screen.MousePointer = 11
        Unload frmCuotas
        Screen.MousePointer = 0
        frmCuotas.Show 1
        'Load frmEstructurasSC
        'Screen.MousePointer = 0
        'frmEstructurasSC.Show 1
    Case 17
        Screen.MousePointer = 11
        Load FrmCopyP
        Screen.MousePointer = 0
        FrmCopyP.Show 1
    Case 19
        Screen.MousePointer = 11
        Load frmPresup
        Screen.MousePointer = 0
        frmPresup.Show 1
End Select
Exit Sub
sale:
Screen.MousePointer = 0
If Err.Number = 53 Then
    MsgBox "NO esta Instalado  ...", 48, Pub_Titulo
    Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If

End Sub

Private Sub SubmenuTit5_Click(Index As Integer)
On Error GoTo sale
Select Case Index
    Case 0
       Call Shell("C:\MSSQL\BINN\SQLEW.EXE", 3) ' Ejecuta SQL
    Case 1
        Load frmeditor
        frmeditor.Show 1
    Case 2
       Screen.MousePointer = 11
       Unload FORMGEN
       Load frmConfigRepo
       Screen.MousePointer = 0
       frmConfigRepo.Show 1
    Case 3
        If LK_CODUSU <> "ADMIN" Then
          Load FrmGrupos
          FrmGrupos.Show 1
        Else
          pub_mensaje = "ADMIN , desea el Formulario de Imprimer BDATOS "
          Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
          If Pub_Respuesta = vbYes Then
            Load FrmImprime
            FrmImprime.Show 1
          Else
            Load FrmGrupos
            FrmGrupos.Show 1
          End If
       End If
    Case 4
        Load frmDiseno
        frmDiseno.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo
End Select

Exit Sub
sale:
If Err.Number = 53 Then
 Screen.MousePointer = 0
 MsgBox "Si Elijio Backup, Este No Procede en Terminales. Solo en su Servidor , Consulte a su Proveedor para mas Detalles", 48, Pub_Titulo
 Exit Sub
End If
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If
End Sub
Private Sub SubmenuTit3_Click(Index As Integer)
On Error GoTo sale
Select Case Index
    Case 0
        Screen.MousePointer = 11
        Load FrmCoPar
        Screen.MousePointer = 0
        FrmCoPar.Show 1
    Case 1
        Screen.MousePointer = 11
        Load FORM_CONT
        FORM_CONT.Show
        Screen.MousePointer = 0
    Case 2
        Screen.MousePointer = 11
        Load FORM_CONTA
        FORM_CONTA.Show
        Screen.MousePointer = 0
    Case 3
        Screen.MousePointer = 11
        Load frm_mayoriz
        Screen.MousePointer = 0
        frm_mayoriz.Show 1
    Case 4
        Screen.MousePointer = 11
        Load frm_mes_cierre
        Screen.MousePointer = 0
        frm_mes_cierre.Show 1
    
    Case 5
        Screen.MousePointer = 11
        Load GenContab
        Screen.MousePointer = 0
        GenContab.Show 1
        
    Case Else
        MsgBox "INDICE FUERA DE RANGO ...", 48, Pub_Titulo
End Select
Exit Sub
sale:
If Err.Number = 424 Then
 Screen.MousePointer = 0
 MsgBox "No tiene Acceso ", 48, Pub_Titulo
End If


End Sub

Private Sub submenutit7_Click(Index As Integer)
On Error GoTo sale
  Select Case Index
    Case 0
       
    Case 2
       Screen.MousePointer = 11
       Load FORM_COTV
       FORM_COTV.Show
       Screen.MousePointer = 0
    Case 4
        Screen.MousePointer = 11
       Load FORM_PED
       FORM_PED.Show
       Screen.MousePointer = 0
   Case 6
       Screen.MousePointer = 11
       Load frmOpcion
       Screen.MousePointer = 0
       frmOpcion.Show 1
    Case 8
       Screen.MousePointer = 11
'       Load Frmimport
       Screen.MousePointer = 0
 '      Frmimport.Show 1
    Case Else
        MsgBox "INDICE FUERA DE RANGO ... ", 48, Pub_Titulo
End Select
sale:
Screen.MousePointer = 0
Exit Sub

End Sub


Private Sub submenuConsultorio_Click()
On Error GoTo sale
  Screen.MousePointer = 11
  Load FrmCitasMedicas
  FrmCitasMedicas.Show
  Screen.MousePointer = 0
  Exit Sub
sale:
  Screen.MousePointer = 0
  MsgBox Err.Description, 48, Pub_Titulo
End Sub

Private Sub Timer1_Timer()
On Error GoTo sale
Timer1.Enabled = False
If Trim(MDIForm1.gfecha.Caption) <> "" Then
  MDIForm1.gfecha.Caption = ""
Else
  MDIForm1.gfecha.Caption = Format(LK_FECHA_DIA, "dd/mm/yyyy")
End If
PSMSG_LLAVE(0) = LK_CODUSU
msg_llave.Requery
If Trim(msg_llave!USU_MSGBOX) <> "" Then
   MsgBox "Usted. Tiene un Mensaje de :" & Chr(13) & Chr(13) & Trim(msg_llave!USU_MSGBOX), vbExclamation, "Mensajeria OSBusiness"
   msg_llave.Edit
   msg_llave!USU_MSGBOX = ""
   msg_llave.Update
End If
sale:

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RetVal
Dim opdato As String
Dim wclave As String
On Error GoTo sal
Dim iFormCount As Integer
Dim WCAMBIA
LogTrace "MDI: Toolbar Click - " & Button.key
Select Case Button.key
  Case "FORMGEN"
      Screen.MousePointer = 11
      Load FORMGEN
      FORMGEN.Show
      Screen.MousePointer = 0
  Case "Tablas"
      Screen.MousePointer = 11
      Load FrmTabla1
      FrmTabla1.Show
      Screen.MousePointer = 0
  Case "KARDEX"
      Screen.MousePointer = 11
      Load KARDEX
      KARDEX.Show
      Screen.MousePointer = 0
  Case "PRECIOS"
      If LK_FLAG_SOS = "A" Then
         MsgBox "En Contrucciro n", 48, Pub_Titulo
         Exit Sub
      End If

      Screen.MousePointer = 11
      Load FrmPreUnidad
      FrmPreUnidad.Show
      Screen.MousePointer = 0
  Case "ARTI"
      Screen.MousePointer = 11
      Load frmARTI
      frmARTI.Show
      Screen.MousePointer = 0
  Case "CLIENTES"
      Screen.MousePointer = 11
      Load frmCLI
      frmCLI.Show
      Screen.MousePointer = 0
  Case "CUENTAS"
      Load FrmCaa
      FrmCaa.Show
  Case "MSGBOX"
      Load frmsucesos
      frmsucesos.Show 1
  Case "CHEQUES"
     If LK_FLAG_SOS = "A" Then
        Screen.MousePointer = 0
        Exit Sub
      End If
      Load frm_cheques
      frm_cheques.Show
  Case "Password"
      Load FrmClave
      FrmClave.Show 1
  Case "CONSUL"
      Load frmsucesos
      frmsucesos.lblsuc.Caption = "Consultar Precios de Competencia"
      frmsucesos.lstprod.Visible = True
      frmsucesos.Show 1
  Case "AVISO"
      Load FrmDetener
      FrmDetener.Show 1
  Case "BANCOS"
      If LK_FLAG_SOS = "A" Then
        Screen.MousePointer = 0
        Exit Sub
      End If
      Load frmBancos
      frmBancos.Show
  Case "consultas"
      Screen.MousePointer = 11
      Load frmdocu
      frmdocu.Show
      Screen.MousePointer = 0
  Case "IA"
        LK_ACCESO_REPORT = ""
        Load frmclave2
        Screen.MousePointer = 0
        frmclave2.Show 1
        If LK_ACCESO_REPORT <> "A" Then
        Exit Sub
        End If

      Screen.MousePointer = 11
       IA_General
      Screen.MousePointer = 0
  
  Case "CAL"
        iFormCount = 0
        For fila = 1 To lk_OTROS_Count
         If Val(lk_OTROS(fila)) = 21 Then ' bloque de sumatorias
            iFormCount = 1
         End If
        Next fila
        Load FrmCal
        If iFormCount = 1 Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then FrmCal.lblacc.Visible = False
        FrmCal.Show 1
  Case "COTI"
      Screen.MousePointer = 11
      Load FORM_COT
      FORM_COT.Show
      Screen.MousePointer = 0
  Case "ORDEN"
      Screen.MousePointer = 11
      Load FORM_PED
      FORM_PED.Show
      Screen.MousePointer = 0
  Case "Salir"
        Screen.MousePointer = 11
        DoEvents
        MDIForm1.WindowState = 1
        MDIForm1.Caption = "Descargando Modulos..."
        EN.Close
        CN.Close
        Pub_ConnAdo.Close
        Screen.MousePointer = 0
        If LK_EMP = "PLA" Then
          If XLL Is Nothing Then
          Else
                XLL.Workbooks.Close
                XLL.Application.Quit
            End If
          Set XLL = Nothing
         End If
         End
  Case "ALTERNO"
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
      LK_FLAG_ALTERNO = " "
      Toolbar1.Buttons.Item(12).Image = 18 '********** tenr en cuent
    Else
      Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
      LK_FLAG_ALTERNO = "A"
      Toolbar1.Buttons.Item(12).Image = 5
    End If
    
      
   If Forms.count - 1 > 0 Then
     For iFormCount = Forms.count - 1 To 1 Step -1
        If iFormCount <> 1 Then
           If UCase(Forms(iFormCount).Name) = "FRMARTI" Then
                'MsgBox UCase(Forms(iFormCount).Name)
                frmARTI.PROCESO_CANCELAR
                frmARTI.PROCESO_ARTI
           ElseIf UCase(Forms(iFormCount).Name) = "KARDEX" Then
                KARDEX.ListView1.Visible = False
                KARDEX.i_codart2.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMPREUNIDAD" Then
                FrmPreUnidad.PRO_GOTFOCUS
                FrmPreUnidad.ListView1.Visible = False
                FrmPreUnidad.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FRMREPO" Then
                FrmRepo.ListView1.Visible = False
                FrmRepo.Txt_key.Text = ""
           ElseIf UCase(Forms(iFormCount).Name) = "FORMGEN" Then
                FORMGEN.LV_ART.Visible = False
                FORMGEN.TEXTOVAR.Text = ""
           End If
        End If
       Next iFormCount
   End If
 Case "tope"
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
    Exit Sub
    End If
    par_llave.Requery
    
    opdato = InputBox("Digitar : FA para cambio de Tope de Facturas y Otros Documentos " & Chr(13) & " LE para cambio de solo Tope de Letras y Letras Renovadas" & Chr(13) & " PR para cambio de solo Tope de Pagos Programados ", "seleccion")
    If opdato = "" Then Exit Sub
    If UCase(opdato) = "FA" Or UCase(opdato) = "LE" Or UCase(opdato) = "PR" Then
    Else
      MsgBox "Dato Incorrecto", 48, Pub_Titulo
      Exit Sub
    End If
    'PAR_SALDO_CAJA_ayer
    If UCase(opdato) = "FA" Then
         wclave = InputBox("Cambiar el Monto Maximo para pago por Día (Facturas y Otros Documentos):", "Monto en Nuevos Soles", par_llave!PAR_TASA_VENTA)
         If wclave = "" Then Exit Sub
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Or LK_CODCIA = "10" Then
             pub_cadena = "UPDATE PARGEN SET PAR_TASA_VENTA =  " & Val(wclave) & "  WHERE PAR_CODCIA IN ('01','03','07','10')"
             CN.Execute pub_cadena
         Else
             par_llave.Requery
             par_llave.Edit
             par_llave!PAR_TASA_VENTA = Val(wclave)
             par_llave.Update
        End If
   ElseIf UCase(opdato) = "LE" Then
         wclave = InputBox("Cambiar el Monto Maximo para pago por Día (Solo Letras y Renovadas):", "Monto en Nuevos Soles", par_llave!PAR_SALDO_CAJA_D_AYER)
         If wclave = "" Then Exit Sub
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Or LK_CODCIA = "10" Then
             pub_cadena = "UPDATE PARGEN SET PAR_SALDO_CAJA_D_AYER =  " & Val(wclave) & "  WHERE PAR_CODCIA IN ('01','03','07','10')"
             CN.Execute pub_cadena
         Else
             par_llave.Requery
             par_llave.Edit
             par_llave!PAR_SALDO_CAJA_D_AYER = Val(wclave)
             par_llave.Update
        End If
   ElseIf UCase(opdato) = "PR" Then
         wclave = InputBox("Cambiar el Monto Maximo para pago por Día (Solo Programadas):", "Monto en Nuevos Soles", par_llave!PAR_SALDO_CAJA_HOY)
         If wclave = "" Then Exit Sub
         If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Or LK_CODCIA = "10" Then
             pub_cadena = "UPDATE PARGEN SET PAR_SALDO_CAJA_HOY =  " & Val(wclave) & "  WHERE PAR_CODCIA IN ('01','03','07','10')"
             CN.Execute pub_cadena
         Else
             par_llave.Requery
             par_llave.Edit
             par_llave!PAR_SALDO_CAJA_HOY = Val(wclave)
             par_llave.Update
        End If
   
   End If
  Case "FECHAP"
    wclave = InputBox("Tiene 2 opciones : Digitar 1 para cambiar la Fecha del Dia / Digitar 2 para Colocar Porcentaje para la Lista Precios Cientes ", "opcion", "1")
    If wclave = "" Then Exit Sub
    If wclave = 1 Then
        wclave = InputBox("Cambiar de Fecha de Proceso :" & Chr(13) & "OJO, debe ser supervisado por el administrador del sistema." & Chr(13) & "Por Defecto Fecha del Día.", "FECHA DE PROCESO !!!! ", Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy"))
        If wclave = "" Then Exit Sub
        If Not IsDate(wclave) Then Exit Sub
        
        LK_FECHA_DIA = Format(wclave, "dd/mm/yyyy")
        If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(par_llave!PAR_FECHA_DIA, "dd/mm/yyyy") Then
          gfecha.ForeColor = QBColor(4) '4 ES ESTANDAR
          gfecha.BackColor = QBColor(15)
        Else
          gfecha.ForeColor = QBColor(15)
          gfecha.BackColor = QBColor(12)
        End If
    ElseIf wclave = 2 Then
        pub_cadena = "select  * from TABLAS WHERE TAB_CODCIA = '00' AND TAB_TIPREG = 23 AND TAB_NUMTAB = 1 " ' EL PORCENTAJE PARA LISTA DE PRECIOS CLIENTES
        Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
        If X.EOF Then
          X.AddNew
          X!tab_codcia = "00"
          X!tab_tipreg = 23
          X!TAB_NUMTAB = 1
          X!TAB_NOMLARGO = ""
          X!tab_nomcorto = "0"
          X.Update
        Else
         
        End If
        wclave = InputBox("CAMBIO DE PORCENTAJE ", " CAMBIAR", Val(X!tab_nomcorto))
        If wclave = "" Then Exit Sub
          X.Edit
          X!tab_nomcorto = wclave
          X.Update
      MsgBox "OK, CAMBIADO", 48, Pub_Titulo
      
    
    End If
    
     ' &H000000FF& ' OJO ENCENDIDO
  Case "Excel"
   '  If Trim(LK_FLAG_SOS) = "" Then
   '    LK_FLAG_SOS = "A"
   '    MsgBox "activo"
   '  Else
   '    LK_FLAG_SOS = ""
   '    MsgBox "Des activo"
   '   End If
   '   LK_ACTIVA = "A"
     On Error GoTo vercal
     RetVal = Shell("C:\WINDOWS\CALC.EXE", 1)    ' Ejecuta Calculadora.
     On Error GoTo 0
    Exit Sub
   If LK_EMP <> "CAM" Then
     ACT_TIPO_CAMBIO
     Exit Sub
   End If
    pub_cadena = "SELECT * FROM CONTROLL"
    Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
    GEN.Requery
    If Trim(LK_FLAG_SOS) <> Trim(UCase(GEN!gen_bloqueo)) Then
        con_llave.Close
        MsgBox "Listo!!!", 48, Pub_Titulo
        End
    End If
    If Trim(UCase(GEN!gen_bloqueo)) = "A" Then
      If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPERVISOR" Then
          wclave = InputBox("Desactivar:", "Activar/Desactivar", "**********")
          If wclave = Trim(GEN!GEN_PWD) Then
            GEN.Edit
            GEN!gen_bloqueo = " "
            GEN.Update
            con_llave.Close
            MsgBox "Control Desactivado", 48, Pub_Titulo
            End
          Else
            con_llave.Close
            Exit Sub
          End If
      Else
'         MsgBox "Control fue Activado!!!", 48, Pub_Titulo
         con_llave.Close
         Exit Sub
      End If
    End If
    GEN.Edit
    GEN!gen_bloqueo = "A"
    GEN.Update
    con_llave.Close
    If Trim(UCase(GEN!gen_bloqueo)) = "A" Then
       MsgBox "Activado Control", 48, Pub_Titulo
       End
    End If
  Case "INCENTIVO"
     GoSub incentivos
End Select
Exit Sub

sal:
  
  Select Case Err.Number
  Case 53
    MsgBox "No se encontro el Archivo ....!!!!..", 48, Pub_Titulo
  Case 424
    MsgBox "No existe el Formulario ....!!!!..r", 48, Pub_Titulo
  Case Else
    MsgBox Err.Number & "  " & Err.Description & "    !!!!... ", 48, Pub_Titulo
  End Select
  Screen.MousePointer = 0
  Exit Sub

Exit Sub
vercal:
If Err.Number = 76 Then
  On Error GoTo SALETOTAL
   RetVal = Shell("C:\WINNT\System32\calc.exe", 1)    ' Ejecuta Calculadora.
End If
Exit Sub
SALETOTAL:
Exit Sub

incentivos:
On Error GoTo sale
Screen.MousePointer = 13

MDIForm1.Reportes.Connect = PUB_ODBC
MDIForm1.Reportes.Destination = crptToWindow  '= crptToPrinter
MDIForm1.Reportes.WindowLeft = 350
MDIForm1.Reportes.WindowTop = 70
MDIForm1.Reportes.WindowWidth = 435
MDIForm1.Reportes.WindowHeight = 490
MDIForm1.Reportes.Formulas(1) = ""

'frmdocu.Reportes.Formulas(1) = "SON_EFECTIVO=  'SON: " & CONVER_LETRAS(PUB_NETO, rmoneda) & "'"
MDIForm1.Reportes.WindowTitle = "Listado de Incentivos "
MDIForm1.Reportes.ReportFileName = PUB_RUTA_OTRO + "REPINC.RPT"
pub_cadena = "{MAEINCENTIVO.MIN_CODCIA} = '01'" ' siempre los incentivos se originan en Dilvisa
MDIForm1.Reportes.SelectionFormula = pub_cadena
MDIForm1.Reportes.WindowTitle = MDIForm1.Reportes.WindowTitle & " Archivo: " & Trim(MDIForm1.Reportes.ReportFileName)

MDIForm1.Reportes.Action = 1
Screen.MousePointer = 0
Return

Exit Sub

sale:
MsgBox "Existe algun problema en el reporte.", 48, Pub_Titulo

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub


Private Sub TXTCIA_Click()
  MDIForm1.PopupMenu Menucompania, 0, 7400, True
End Sub
Public Function PRO_DES(WCFP As Currency) As Currency
Dim WRES As Currency
WRES = (WCFP * 9) + 9
PRO_DES = WRES
End Function


Public Sub VER_MENSAJES()
PSMEN_LLAVE(0) = ""
men_llave.Requery
End Sub
