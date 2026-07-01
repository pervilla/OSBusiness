VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmARTI 
   Caption         =   "Maestro de Articulos"
   ClientHeight    =   12075
   ClientLeft      =   690
   ClientTop       =   1185
   ClientWidth     =   11940
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   12075
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ListView1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   540
      Left            =   9600
      ScaleHeight     =   480
      ScaleWidth      =   1545
      TabIndex        =   260
      Top             =   7440
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame fraact6 
      Caption         =   "Afecta a la Lista de Precios :"
      Height          =   1095
      Left            =   5160
      TabIndex        =   134
      Top             =   11160
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ComboBox cmbtipo 
         Height          =   315
         ItemData        =   "Arti.frx":0000
         Left            =   4080
         List            =   "Arti.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   272
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox cmdlista 
         Height          =   315
         ItemData        =   "Arti.frx":0010
         Left            =   120
         List            =   "Arti.frx":0038
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdcan 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   140
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtnrolista 
         Height          =   285
         Left            =   3240
         MaxLength       =   2
         TabIndex        =   137
         Text            =   "1"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtporlista 
         Height          =   285
         Left            =   3240
         TabIndex        =   136
         Text            =   "0.00"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdact6 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   135
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Lista a Actualizar"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   157
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "N� Lista a Usar :"
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   139
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Adicione el (%) :"
         Height          =   255
         Left            =   1920
         TabIndex        =   138
         Top             =   600
         Width           =   1335
      End
   End
    End
      Caption         =   "Calcular Stock Minimos y Maximos por Promedio."
      Height          =   2895
      Left            =   2520
      TabIndex        =   142
      Top             =   8280
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtperiodo 
         Height          =   285
         Left            =   5880
         TabIndex        =   257
         Text            =   "4"
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chenueva 
         Caption         =   "Usar Nueva Formula"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   3720
         TabIndex        =   256
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtapor 
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   4800
         TabIndex        =   208
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox cheapor 
         Caption         =   "Porcentaje de Alerta."
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3360
         TabIndex        =   207
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdapor 
         Caption         =   "Actualizar Porcentajes"
         Height          =   375
         Left            =   3360
         TabIndex        =   206
         Top             =   2040
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.ListBox listlin 
         Height          =   1185
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   153
         Top             =   1440
         Width           =   3135
      End
      Begin VB.PictureBox pb 
         Appearance      =   0  'Flat
         Height          =   180
         Left            =   240
         ScaleHeight     =   150
         ScaleWidth      =   2865
         TabIndex        =   152
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtd 
         Height          =   285
         Left            =   3840
         TabIndex        =   150
         Text            =   "000"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdgen 
         Caption         =   "Generar Proceso"
         Height          =   615
         Left            =   4560
         TabIndex        =   149
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton ops 
         Caption         =   "Maximos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   148
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton ops 
         Caption         =   "Minimos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   147
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   2160
         TabIndex        =   143
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   2160
         TabIndex        =   144
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label16 
         Caption         =   "Periodos Estadisticos:"
         Height          =   495
         Left            =   5760
         TabIndex        =   258
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblapor 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   210
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblapor 
         Caption         =   "Asignar Porcentaje:"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   209
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Filtrar por Lineas:"
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   240
         TabIndex        =   154
         Top             =   1200
         Width           =   2145
      End
      Begin VB.Label Label11 
         Caption         =   "Dias Prom."
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   3600
         TabIndex        =   151
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblcampo2 
         Caption         =   "Fec. Fin :"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1200
         TabIndex        =   146
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblcampo1 
         Caption         =   "Fec. Inicio :"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1200
         TabIndex        =   145
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame fraactlinea 
      Caption         =   "Actualizar por Lineas :"
      Height          =   5175
      Left            =   12015
      TabIndex        =   158
      Top             =   1710
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox chere9 
         Caption         =   "Red. Especial 9"
         Height          =   195
         Left            =   5880
         TabIndex        =   205
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CheckBox chere8 
         Caption         =   "Red. Especial 8"
         Height          =   195
         Left            =   5880
         TabIndex        =   204
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox chere7 
         Caption         =   "Red. Especial 7"
         Height          =   195
         Left            =   5880
         TabIndex        =   203
         Top             =   3240
         Width           =   1455
      End
      Begin VB.CheckBox chere6 
         Caption         =   "Red. Especial 6"
         Height          =   195
         Left            =   5880
         TabIndex        =   202
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CheckBox chere5 
         Caption         =   "Red. Especial 5"
         Height          =   195
         Left            =   5880
         TabIndex        =   201
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox chere4 
         Caption         =   "Red. Especia 4"
         Height          =   195
         Left            =   5880
         TabIndex        =   200
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chere3 
         Caption         =   "Red. Especial 3"
         Height          =   195
         Left            =   5880
         TabIndex        =   199
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chere2 
         Caption         =   "Red. Especial 2"
         Height          =   195
         Left            =   5880
         TabIndex        =   198
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chere1 
         Caption         =   "Red. Especial 1"
         Height          =   195
         Left            =   5880
         TabIndex        =   197
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox R9 
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
         ItemData        =   "Arti.frx":00BD
         Left            =   4800
         List            =   "Arti.frx":00D0
         Style           =   2  'Dropdown List
         TabIndex        =   194
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R8 
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
         ItemData        =   "Arti.frx":00E3
         Left            =   4800
         List            =   "Arti.frx":00F6
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   3600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R7 
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
         ItemData        =   "Arti.frx":0109
         Left            =   4800
         List            =   "Arti.frx":011C
         Style           =   2  'Dropdown List
         TabIndex        =   192
         Top             =   3240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R6 
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
         ItemData        =   "Arti.frx":012F
         Left            =   4800
         List            =   "Arti.frx":0142
         Style           =   2  'Dropdown List
         TabIndex        =   191
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R5 
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
         ItemData        =   "Arti.frx":0155
         Left            =   4800
         List            =   "Arti.frx":0168
         Style           =   2  'Dropdown List
         TabIndex        =   190
         Top             =   2520
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R4 
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
         ItemData        =   "Arti.frx":017B
         Left            =   4800
         List            =   "Arti.frx":018E
         Style           =   2  'Dropdown List
         TabIndex        =   189
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R3 
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
         ItemData        =   "Arti.frx":01A1
         Left            =   4800
         List            =   "Arti.frx":01B4
         Style           =   2  'Dropdown List
         TabIndex        =   188
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R2 
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
         ItemData        =   "Arti.frx":01C7
         Left            =   4800
         List            =   "Arti.frx":01DA
         Style           =   2  'Dropdown List
         TabIndex        =   187
         Top             =   1440
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox R1 
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
         ItemData        =   "Arti.frx":01ED
         Left            =   4800
         List            =   "Arti.frx":0200
         Style           =   2  'Dropdown List
         TabIndex        =   186
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.ComboBox cmdptipo 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   184
         Top             =   360
         Width           =   4215
      End
      Begin VB.CommandButton cmdcanlin 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   183
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton cmdactlineas 
         Caption         =   "Actualizar Listas "
         Height          =   375
         Left            =   600
         TabIndex        =   181
         Top             =   4560
         Width           =   2295
      End
      Begin VB.PictureBox pbp 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   4305
         TabIndex        =   180
         Top             =   4200
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.TextBox tpo2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   177
         Text            =   "0.00"
         Top             =   1440
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   176
         Text            =   "0.00"
         Top             =   1800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   175
         Text            =   "0.00"
         Top             =   2160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   174
         Text            =   "0.00"
         Top             =   2520
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   173
         Text            =   "0.00"
         Top             =   2880
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   172
         Text            =   "0.00"
         Top             =   3240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   171
         Text            =   "0.00"
         Top             =   3600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   170
         Text            =   "0.00"
         Top             =   3960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox tpo1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         TabIndex        =   169
         Text            =   "0.00"
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CheckBox chep9 
         Caption         =   "Precio 9"
         Height          =   195
         Left            =   3120
         TabIndex        =   168
         Top             =   3960
         Width           =   975
      End
      Begin VB.CheckBox chep3 
         Caption         =   "Precio 3"
         Height          =   195
         Left            =   3120
         TabIndex        =   167
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox chep4 
         Caption         =   "Precio 4"
         Height          =   195
         Left            =   3120
         TabIndex        =   166
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chep2 
         Caption         =   "Precio 2"
         Height          =   195
         Left            =   3120
         TabIndex        =   165
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chep5 
         Caption         =   "Precio 5"
         Height          =   195
         Left            =   3120
         TabIndex        =   164
         Top             =   2520
         Width           =   975
      End
      Begin VB.CheckBox chep6 
         Caption         =   "Precio 6"
         Height          =   195
         Left            =   3120
         TabIndex        =   163
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chep7 
         Caption         =   "Precio 7"
         Height          =   195
         Left            =   3120
         TabIndex        =   162
         Top             =   3240
         Width           =   975
      End
      Begin VB.CheckBox chep8 
         Caption         =   "Precio 8"
         Height          =   195
         Left            =   3120
         TabIndex        =   161
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chep1 
         Caption         =   "Precio 1"
         Height          =   195
         Left            =   3120
         TabIndex        =   160
         Top             =   1080
         Width           =   975
      End
      Begin VB.ListBox listlineas 
         Height          =   3210
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   159
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Redondeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   196
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   195
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Porc.(%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   185
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Tipo de  Actualizaciion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   179
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label13 
         Caption         =   "Marcar Linea :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   178
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame fralstfiltro 
      Caption         =   "Familia para Filtro! Reportes"
      Height          =   4095
      Left            =   360
      TabIndex        =   246
      Top             =   8520
      Visible         =   0   'False
      Width           =   3135
      Begin VB.PictureBox PBF 
         Appearance      =   0  'Flat
         Height          =   135
         Left            =   240
         ScaleHeight     =   105
         ScaleWidth      =   2505
         TabIndex        =   250
         Top             =   3480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton cmdcancelfiltro 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1920
         TabIndex        =   249
         Top             =   3600
         Width           =   855
      End
      Begin VB.CommandButton cmdgrabafiltro 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   240
         TabIndex        =   248
         Top             =   3600
         Width           =   1215
      End
      Begin VB.ListBox lstfiltro 
         Height          =   3210
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   247
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame fracias 
      Caption         =   "Seleccione Cia Aplicar"
      Height          =   2895
      Left            =   8160
      TabIndex        =   243
      Top             =   4680
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdaplica 
         Caption         =   "Aplicar"
         Height          =   435
         Left            =   480
         TabIndex        =   245
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ListBox lstcias 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   244
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CheckBox cheotros 
      BackColor       =   &H00808000&
      Caption         =   "Incluir Productos DesActivados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   10440
      TabIndex        =   155
      Top             =   5370
      Width           =   1275
   End
   Begin VB.CommandButton cmdact 
      Caption         =   "Act. Lista de Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   133
      Top             =   5970
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton MANOS 
      Caption         =   "&Ant."
      Height          =   495
      Index           =   0
      Left            =   10440
      Picture         =   "Arti.frx":0213
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton MANOS 
      Caption         =   "&Sig"
      Height          =   495
      Index           =   1
      Left            =   11160
      Picture         =   "Arti.frx":0655
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   120
      Width           =   615
   End
   Begin VB.Frame F14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   64
      Top             =   7200
      Visible         =   0   'False
      Width           =   7965
      Begin VB.OptionButton Op 
         Caption         =   "Seleccionar uno de la Lista "
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   69
         Top             =   1800
         Width           =   2535
      End
      Begin VB.OptionButton Op 
         Caption         =   "Ignorar la Lista "
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   68
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton cmdconfirma 
         Caption         =   "Con&firmar Grabaci�n"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   67
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton CmdEscapa 
         Caption         =   "E&scapar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   66
         Top             =   1800
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid ListExiste 
         Height          =   1575
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   4
         BackColorBkg    =   8421376
      End
   End
   Begin VB.TextBox tcospro 
      DataField       =   "ART_COSPRO"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      MaxLength       =   11
      TabIndex        =   84
      Top             =   3000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame2 
      Height          =   700
      Left            =   2880
      TabIndex        =   105
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtnombre 
         DataField       =   "ART_NOMBRE"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         MaxLength       =   60
         TabIndex        =   2
         Top             =   300
         Width           =   6975
      End
      Begin VB.Label lblnomarti 
         AutoSize        =   -1  'True
         Caption         =   "Descripci�n del Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   106
         Top             =   120
         Width           =   2040
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2895
      Left            =   0
      TabIndex        =   37
      Top             =   840
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   5106
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Definici�n de Estructura"
      TabPicture(0)   =   "Arti.frx":0A97
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Fvarios"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cheasignarc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chebof"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cheoferta"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Porcentajes"
      TabPicture(1)   =   "Arti.frx":0AB3
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CERO"
      Tab(1).Control(1)=   "Fop"
      Tab(1).Control(2)=   "Fcomi"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Otras Opciones"
      TabPicture(2)   =   "Arti.frx":0ACF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "frarelacion"
      Tab(2).Control(3)=   "frmpro"
      Tab(2).ControlCount=   4
      Begin VB.CheckBox cheoferta 
         Caption         =   "Esta en Oferta !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   6480
         TabIndex        =   284
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chebof 
         Caption         =   "No Incluir en Lista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6420
         TabIndex        =   282
         Top             =   1240
         Width           =   2385
      End
      Begin VB.CheckBox cheasignarc 
         Caption         =   "Asignar al Convenio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   8280
         TabIndex        =   273
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         Caption         =   "Otras Opciones :"
         ForeColor       =   &H00808000&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   264
         Top             =   360
         Width           =   6855
         Begin VB.ListBox lfvta 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            Left            =   4320
            Style           =   1  'Checkbox
            TabIndex        =   286
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton cmdverifi 
            Caption         =   "Verificar Prod. Unid. Activa"
            Height          =   375
            Left            =   120
            TabIndex        =   279
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Frame fracuotas 
            Caption         =   "Ac/ Cuotas x Periodo :"
            ForeColor       =   &H00808000&
            Height          =   1575
            Left            =   2400
            TabIndex        =   274
            Top             =   600
            Width           =   1815
            Begin VB.TextBox txtper 
               Height          =   285
               Left            =   240
               TabIndex        =   276
               Text            =   "042007"
               Top             =   480
               Width           =   735
            End
            Begin VB.TextBox txtum 
               Height          =   285
               Left            =   240
               TabIndex        =   275
               Text            =   "0"
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Periodo del Mes:"
               ForeColor       =   &H00808000&
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   278
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Unidades Maxima :"
               ForeColor       =   &H00808000&
               Height          =   375
               Index           =   4
               Left            =   120
               TabIndex        =   277
               Top             =   840
               Width           =   1455
            End
         End
         Begin VB.TextBox txtCantMaxADD 
            Height          =   285
            Left            =   1560
            TabIndex        =   271
            Text            =   "0"
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtCantMax 
            Height          =   285
            Left            =   1560
            TabIndex        =   267
            Text            =   "0"
            Top             =   720
            Width           =   735
         End
         Begin VB.ComboBox cmbUsu 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1920
            TabIndex        =   266
            Text            =   "cmbUsu"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Acceso a Fza. Vta."
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   5
            Left            =   4320
            TabIndex        =   287
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "Cantidad Adicional Temporal"
            ForeColor       =   &H00808000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   270
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblum 
            Caption         =   "."
            ForeColor       =   &H00808000&
            Height          =   375
            Left            =   2040
            TabIndex        =   269
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "Cantidad Maxima por Pedido:"
            ForeColor       =   &H00808000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   268
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Usuario Responsable :"
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   265
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Otros"
         Height          =   2415
         Left            =   -67920
         TabIndex        =   251
         Top             =   360
         Width           =   3135
         Begin VB.CommandButton cmdofertas 
            Caption         =   "Actualizar Ofertas..."
            Height          =   480
            Left            =   360
            Picture         =   "Arti.frx":0AEB
            Style           =   1  'Graphical
            TabIndex        =   252
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label fechahora 
            Alignment       =   2  'Center
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   480
            TabIndex        =   259
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   6360
         TabIndex        =   110
         Top             =   315
         Width           =   3870
         Begin VB.TextBox txtmargen 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   285
            Left            =   960
            TabIndex        =   285
            Text            =   "0"
            Top             =   2040
            Width           =   495
         End
         Begin VB.CommandButton cmdqc 
            Caption         =   "Quitar Conv."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   280
            Top             =   1920
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmbStock 
            Caption         =   "Ver Dif. Stock"
            Height          =   1095
            Left            =   3120
            TabIndex        =   255
            Top             =   360
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame1 
            Height          =   420
            Left            =   120
            TabIndex        =   113
            Top             =   1440
            Width           =   3570
            Begin VB.OptionButton cheservi 
               Caption         =   "Servicio"
               Height          =   195
               Index           =   2
               Left            =   840
               TabIndex        =   116
               Top             =   840
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.OptionButton cheservi 
               Caption         =   "Pqte  Mercaderia"
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
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   1560
               TabIndex        =   115
               Top             =   120
               Width           =   1860
            End
            Begin VB.OptionButton cheservi 
               Caption         =   "Mercaderia"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   114
               Top             =   120
               Width           =   1815
            End
         End
         Begin VB.CheckBox chesub 
            Caption         =   "Prioridad para Vta."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   1200
            Width           =   2385
         End
         Begin VB.CheckBox exigv 
            Caption         =   "Exoneraci�n IGV"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   690
            Value           =   1  'Checked
            Width           =   2250
         End
         Begin VB.TextBox txtcospro 
            DataField       =   "ART_COSPRO"
            DataSource      =   "Data1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2385
            MaxLength       =   11
            TabIndex        =   14
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.CheckBox checambio 
            Caption         =   "Para Cambio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   2130
            TabIndex        =   12
            Top             =   270
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblequivarti 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Caption         =   "(*)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   254
            Top             =   2085
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "Partes d' Und.Min:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   120
            TabIndex        =   253
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "(%) Igv :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   10
            Left            =   2520
            TabIndex        =   111
            Top             =   600
            Visible         =   0   'False
            Width           =   720
         End
            Begin VB.CheckBox art_situacion 
               Caption         =   "DesActivado"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   255
               Left            =   7200
               TabIndex        =   263
               Top             =   1320
               Width           =   2175
            End
      End
      Begin VB.Frame frarelacion 
         Caption         =   "Articulo Relacionado con Almacen Defectuoso : "
         Height          =   615
         Left            =   -65280
         TabIndex        =   98
         Top             =   1800
         Visible         =   0   'False
         Width           =   1815
         Begin VB.TextBox txtcodigo2 
            Height          =   285
            Left            =   4800
            TabIndex        =   102
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Adicionar"
            Height          =   600
            Left            =   4680
            Picture         =   "Arti.frx":0F2D
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton cmdquitar 
            Caption         =   "&Quitar Relaci�n"
            Height          =   600
            Left            =   5640
            Picture         =   "Arti.frx":136F
            Style           =   1  'Graphical
            TabIndex        =   100
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.ComboBox cmbcal 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   99
            Top             =   600
            Width           =   2295
         End
         Begin MSFlexGridLib.MSFlexGrid gridrel 
            Height          =   1935
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   3413
            _Version        =   393216
            BackColorBkg    =   8421376
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Codigo de Relaci�n"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   4680
            TabIndex        =   104
            Top             =   960
            Visible         =   0   'False
            Width           =   1485
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblcal 
            Caption         =   "Calidad Relacionada para Agregar"
            Height          =   375
            Left            =   4560
            TabIndex        =   103
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame frmpro 
         Caption         =   "Relaci�n de Procesos"
         Height          =   375
         Left            =   -74640
         TabIndex        =   81
         Top             =   3360
         Width           =   8415
         Begin VB.Data dataO 
            Caption         =   "Data1"
            Connect         =   "Access"
            DatabaseName    =   ""
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   375
            Left            =   360
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   1560
            Width           =   1140
         End
         Begin VB.CommandButton cmdp 
            Caption         =   "Activar Relaci�n"
            Height          =   735
            Left            =   0
            TabIndex        =   83
            Top             =   0
            Width           =   1455
         End
         Begin MSFlexGridLib.MSFlexGrid gridp 
            Height          =   2055
            Left            =   1800
            TabIndex        =   82
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   4
         End
      End
      Begin VB.CommandButton CERO 
         Caption         =   "Producto 0 (Para ADMIN)"
         Height          =   255
         Left            =   -67320
         TabIndex        =   72
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Frame Fop 
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2415
         Left            =   -69480
         TabIndex        =   63
         Top             =   360
         Width           =   4695
         Begin VB.CommandButton cmdactLista 
            Caption         =   "Actualizar por Lineas (Precios)"
            Height          =   495
            Left            =   120
            TabIndex        =   182
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CommandButton Cmdmax 
            Caption         =   "Generar Stock Maximos/Minimos"
            Height          =   495
            Left            =   120
            TabIndex        =   141
            Top             =   1800
            Width           =   1815
         End
         Begin VB.ComboBox DS 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            ItemData        =   "Arti.frx":17B1
            Left            =   3540
            List            =   "Arti.frx":17B3
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   1560
            Width           =   840
         End
         Begin VB.ComboBox decimales 
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
            Height          =   315
            Left            =   3120
            TabIndex        =   34
            ToolTipText     =   "Decimales para la Cantidad Formulada"
            Top             =   600
            Width           =   1455
         End
         Begin VB.ComboBox art_codpro 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2040
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1200
            Width           =   2295
         End
         Begin VB.TextBox txtfechault 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3120
            TabIndex        =   33
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtMax 
            Height          =   285
            Left            =   360
            MaxLength       =   13
            TabIndex        =   31
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtMin 
            Height          =   285
            Left            =   360
            MaxLength       =   13
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Moneda Articulo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   18
            Left            =   2085
            TabIndex        =   128
            Top             =   1620
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Decimales:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   19
            Left            =   1920
            TabIndex        =   79
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Ultima Compra de:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   2040
            TabIndex        =   74
            Top             =   960
            Width           =   1680
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fec. Compra:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   73
            Top             =   240
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Stock Minimo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   71
            Top             =   165
            Width           =   1275
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Stock Maximo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   70
            Top             =   680
            Width           =   1320
         End
      End
      Begin VB.Frame Fvarios 
         Height          =   2445
         Left            =   90
         TabIndex        =   52
         Top             =   315
         Width           =   6135
         Begin VB.ComboBox art_grupop 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   330
            Left            =   60
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   3015
         End
         Begin VB.ComboBox CmbCalidad 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2060
            Width           =   1215
         End
         Begin VB.ComboBox art_plancha 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox art_marca 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1440
            Width           =   2775
         End
         Begin VB.ComboBox art_linea 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   900
            Width           =   2775
         End
         Begin VB.ComboBox art_numero 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   300
            Width           =   3015
         End
         Begin VB.ComboBox art_grupo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2055
            Width           =   3015
         End
         Begin VB.ComboBox art_subfam 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   80
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1455
            Width           =   3015
         End
         Begin VB.ComboBox art_familia 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   60
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   900
            Width           =   3015
         End
         Begin VB.Label lblgrupop 
            AutoSize        =   -1  'True
            Caption         =   "Grupo Principal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   281
            Top             =   120
            Width           =   1275
         End
         Begin VB.Label Lbl3 
            AutoSize        =   -1  'True
            Caption         =   "Calidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Left            =   4800
            TabIndex        =   109
            Top             =   1860
            Width           =   645
         End
         Begin VB.Label lblart 
            Caption         =   "Lote:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   6
            Left            =   3240
            TabIndex        =   85
            Top             =   1800
            Width           =   1365
         End
         Begin VB.Label lblart 
            Caption         =   "Clase:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   5
            Left            =   3240
            TabIndex        =   80
            Top             =   1245
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Marca:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   4
            Left            =   3240
            TabIndex        =   78
            Top             =   720
            Width           =   3000
         End
         Begin VB.Label lblart 
            Caption         =   "Sub Linea:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   3240
            TabIndex        =   77
            Top             =   120
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Linea Genericos::"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   1875
            Width           =   1965
         End
         Begin VB.Label lblart 
            Caption         =   "Familia:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   1275
            Width           =   3330
         End
         Begin VB.Label lblart 
            AutoSize        =   -1  'True
            Caption         =   "Divisi�n:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   750
         End
      End
      Begin VB.Frame Fcomi 
         Caption         =   "Porc. de Comisiones por Articulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2415
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   5295
         Begin VB.ComboBox cmdvend 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   240
            Top             =   240
            Width           =   3135
         End
         Begin VB.CommandButton cmdactvend 
            Caption         =   "Actualizar"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3720
            TabIndex        =   239
            Top             =   1800
            Width           =   1455
         End
         Begin VB.CheckBox cheespecial 
            Caption         =   "Actualizaci�n en Bloque"
            ForeColor       =   &H00808000&
            Height          =   495
            Left            =   3720
            TabIndex        =   238
            Top             =   480
            Width           =   1335
         End
         Begin VB.ComboBox cmdlineas 
            Height          =   315
            Left            =   3720
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   237
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtpor99 
            Height          =   285
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   218
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor88 
            Height          =   285
            Left            =   2520
            MaxLength       =   6
            TabIndex        =   217
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor77 
            Height          =   285
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   216
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor66 
            Height          =   285
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   215
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtpor9 
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   214
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor8 
            Height          =   285
            Left            =   600
            MaxLength       =   6
            TabIndex        =   213
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor7 
            Height          =   285
            Left            =   0
            MaxLength       =   6
            TabIndex        =   212
            Top             =   1995
            Width           =   495
         End
         Begin VB.TextBox txtpor6 
            Height          =   285
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   211
            Top             =   1530
            Width           =   495
         End
         Begin VB.TextBox txtpor11 
            Height          =   285
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   124
            Top             =   1050
            Width           =   495
         End
         Begin VB.TextBox txtpor22 
            Height          =   285
            Left            =   2520
            MaxLength       =   6
            TabIndex        =   123
            Top             =   1050
            Width           =   495
         End
         Begin VB.TextBox txtpor33 
            Height          =   285
            Left            =   3120
            MaxLength       =   6
            TabIndex        =   122
            Top             =   1050
            Width           =   495
         End
         Begin VB.TextBox txtpor44 
            Height          =   285
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   121
            Top             =   1530
            Width           =   495
         End
         Begin VB.TextBox txtpor55 
            Height          =   285
            Left            =   2520
            MaxLength       =   6
            TabIndex        =   120
            Top             =   1530
            Width           =   495
         End
         Begin VB.TextBox txtpor5 
            Height          =   285
            Left            =   600
            MaxLength       =   6
            TabIndex        =   29
            Top             =   1530
            Width           =   495
         End
         Begin VB.TextBox txtpor4 
            Height          =   285
            Left            =   0
            MaxLength       =   6
            TabIndex        =   28
            Top             =   1530
            Width           =   495
         End
         Begin VB.TextBox txtpor3 
            Height          =   285
            Left            =   4680
            MaxLength       =   6
            TabIndex        =   27
            Top             =   1050
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtpor2 
            Height          =   285
            Left            =   600
            MaxLength       =   6
            TabIndex        =   26
            Top             =   1050
            Width           =   1095
         End
         Begin VB.TextBox txtpor1 
            Height          =   285
            Left            =   0
            MaxLength       =   6
            TabIndex        =   25
            Top             =   1050
            Width           =   495
         End
         Begin VB.Label lblc 
            Caption         =   "Lineas "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   21
            Left            =   3720
            TabIndex        =   242
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "Vend :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   241
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   19
            Left            =   3240
            TabIndex        =   236
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   18
            Left            =   2640
            TabIndex        =   235
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   17
            Left            =   2040
            TabIndex        =   234
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   16
            Left            =   3240
            TabIndex        =   233
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   15
            Left            =   2640
            TabIndex        =   232
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   14
            Left            =   2040
            TabIndex        =   231
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   230
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   12
            Left            =   2640
            TabIndex        =   229
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   11
            Left            =   2040
            TabIndex        =   228
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   10
            Left            =   1320
            TabIndex        =   227
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   226
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   225
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   224
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   223
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   222
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   4
            Left            =   4800
            TabIndex        =   221
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "Valor "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   220
            Top             =   840
            Width           =   855
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   219
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "Cobranza"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   126
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblc 
            Alignment       =   2  'Center
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   125
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 5 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   4
            Left            =   2400
            TabIndex        =   62
            Top             =   2160
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 4 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   61
            Top             =   2160
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 3 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   2
            Left            =   3000
            TabIndex        =   60
            Top             =   2160
            Width           =   1395
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 2 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   1
            Left            =   3240
            TabIndex        =   59
            Top             =   2160
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.Label lblpor 
            AutoSize        =   -1  'True
            Caption         =   "% p'  Precios 1 :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   58
            Top             =   2160
            Visible         =   0   'False
            Width           =   1395
         End
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Correla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   23
         Left            =   -68640
         TabIndex        =   51
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "N�. Dir."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   -71040
         TabIndex        =   50
         Top             =   480
         Width           =   645
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Direc. Trabajo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   -74760
         TabIndex        =   49
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "SubZona Trab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   -70920
         TabIndex        =   48
         Top             =   1080
         Width           =   1290
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Zona : Trab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   -74760
         TabIndex        =   47
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prendas :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   -71880
         TabIndex        =   46
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prop. (1) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   -74760
         TabIndex        =   45
         Top             =   1680
         Width           =   870
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Prop. (2) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   44
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Rg.Pub.(1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   16
         Left            =   -71760
         TabIndex        =   43
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Rg.Pub.(2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   17
         Left            =   -71880
         TabIndex        =   42
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autovaluo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   18
         Left            =   -74760
         TabIndex        =   41
         Top             =   3480
         Width           =   990
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autos (1) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   15
         Left            =   -74760
         TabIndex        =   40
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label LblDatos 
         AutoSize        =   -1  'True
         Caption         =   "Autos (2) :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   -71760
         TabIndex        =   39
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "    Relaci�n      Cia  -  Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -67080
         TabIndex        =   38
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Fcomun 
      Height          =   3255
      Left            =   0
      TabIndex        =   86
      Top             =   3720
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdActBloque 
         Caption         =   "Precios Bloque"
         Height          =   255
         Left            =   8400
         TabIndex        =   283
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtlitro 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.TextBox txtvar 
         Height          =   285
         Left            =   6240
         MaxLength       =   9
         TabIndex        =   87
         Top             =   1035
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmddolares 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtpeso 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   120
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grid_unid 
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "[INSERT] Agrega, [DEL] Quitar"
         Top             =   1035
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3625
         _Version        =   393216
         Cols            =   4
         FixedCols       =   3
         BackColorBkg    =   8421376
         Enabled         =   0   'False
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Lst1 
         Caption         =   "St."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   4680
         TabIndex        =   262
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblstock1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5880
         TabIndex        =   261
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   6
         Left            =   9600
         TabIndex        =   132
         Tag             =   "5"
         Top             =   800
         Width           =   420
      End
      Begin VB.Label lblstock0 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   3480
         TabIndex        =   131
         Top             =   150
         Width           =   945
      End
      Begin VB.Label Lst0 
         Caption         =   "St."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   2475
         TabIndex        =   130
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Litros/Cilind. Unid. Act :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblcospro 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   9000
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label LOTRO 
         Caption         =   "C.Pro. Un.Act.: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   7680
         TabIndex        =   97
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidad Activa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   120
         TabIndex        =   96
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label LBLCOSTO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COSTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Left            =   2280
         TabIndex        =   95
         Top             =   800
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   4
         Left            =   8280
         TabIndex        =   94
         Tag             =   "5"
         Top             =   800
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   6960
         TabIndex        =   93
         Tag             =   "4"
         Top             =   795
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   5640
         TabIndex        =   92
         Tag             =   "3"
         Top             =   800
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   4320
         TabIndex        =   91
         Tag             =   "2"
         Top             =   800
         Width           =   1260
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRECIO. 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   3120
         TabIndex        =   90
         Tag             =   "1"
         Top             =   800
         Width           =   1185
      End
      Begin VB.Label LBLUNIDAD 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   89
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "P.(Kg) Un.Act.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Fdatos 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   700
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txt_key 
         DataField       =   "ART_KEY"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   300
         Width           =   1095
      End
      Begin VB.TextBox txt_alterno 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   1
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label lblalterno 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   1440
         TabIndex        =   56
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Interno:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   55
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Limpiar"
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
      Height          =   625
      Left            =   10440
      Picture         =   "Arti.frx":17B5
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3720
      Width           =   1300
   End
   Begin VB.Timer Parpadea 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   7320
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10440
      Picture         =   "Arti.frx":1BF7
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4650
      Width           =   1300
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificaci�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10440
      Picture         =   "Arti.frx":1D41
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   840
      Width           =   1300
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10440
      Picture         =   "Arti.frx":1E8B
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2760
      Width           =   1300
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Adicionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   625
      Left            =   10440
      Picture         =   "Arti.frx":22CD
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   1300
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   112
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   10440
      TabIndex        =   108
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Height          =   7095
      Index           =   5
      Left            =   10320
      TabIndex        =   107
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   4320
      TabIndex        =   36
      Top             =   6645
      Width           =   3285
   End
End
Attribute VB_Name = "frmARTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_orig_situ As String * 1
Dim FLAG_BUSCAR_MANITO As String * 1
Dim wflag_bloq_desactiva As String * 1
Dim flag_art_situacion As Integer
Dim Pres As String
'*************
Dim LOC_FECHA_INI  As Date
Dim LOC_FECHA_FIN  As Date

Dim llave_sum_arti As rdoResultset
Dim PSLOC_WARTI As rdoQuery

Dim PSCOMI_VEN As rdoQuery
Dim comivem_llave As rdoResultset
Dim CU As Integer
Dim loc_flag_bloq As String * 1
Dim PSART_LOC As rdoQuery
Dim PSART_LOC2 As rdoQuery
Dim artloc_llave As rdoResultset
Dim artloc_mayor As rdoResultset
Dim PSART_KEY As rdoQuery
Dim artloc_key As rdoResultset
Dim LOC_OPER As Currency
Dim loc_tipo As String * 1
Dim LOC_TIPREG As Integer
Dim temporal As String
Dim loc_key As Integer
Dim Flag_Consis As String * 1
Dim Flag_F2 As String * 1
Dim Flag_Bloq As String * 1
Dim Flag_Inicial As String * 1
Dim Flag_Change  As String * 1
Dim loc_fila  As Integer
Dim loc_colum  As Integer
Dim loc_unid As String * 1
Dim VAR_ACTIVAR As Integer
Dim LOC_ORIGINAL As Currency
Dim LOC_ALTERNO As String
Dim LOC_NOMBRE As String
Dim LOC_CALIDAD As Integer
Dim VAR_NEWCAL As Integer
Dim LOC_CODART2 As Currency
Dim LOC_CANCELA As Integer
Dim pasa  As Integer
Dim LOC_CTA_CLI(2) As String * 12
Dim LOC_DES_CLI(2) As String * 50
Dim PSART_RELA As rdoQuery
Dim art_rela As rdoResultset
Dim ws_flag_equiv As String * 1
Dim wflag_bloq_descrip As String
Dim Flag_Add_cant As String * 1
Dim loc_equiv_max As Integer
Dim loc_p As String
Dim loc_cheoferta As String * 1
Dim loc_acc_convenio As String * 1

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
End If
End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
End If
End Sub

Private Sub art_codpro_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_codpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If DS.Enabled Then
    DS.SetFocus
    SendKeys "%{UP}"
  Else
    DS_KeyPress 13
  End If
End If
End Sub

Private Sub art_codpro_LostFocus()
Fvarios.Refresh
End Sub

Private Sub art_familia_Click()
 art_subfam.Clear
' art_grupo.Clear
' art_numero.Clear
' art_linea.Clear
' art_marca.Clear

End Sub

Private Sub art_familia_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_familia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If loc_tipo = "V" Then
        DoEvents
        art_subfam.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End If
End Sub

Private Sub art_familia_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wpos
If KeyCode = vbKeyF5 Then
  GoTo muestra
End If
If KeyCode = vbKeyF6 Then
  GoTo muestra
End If


If KeyCode <> 45 Then
 Exit Sub
End If
Flag_Bloq = "A"
wpos = art_familia.ListIndex
PUB_TIPREG = Mid(art_familia.ToolTipText, 13, Len(art_familia.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "FAMILIAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
Flag_Bloq = ""
DoEvents
LLENADO_FAM
DoEvents
On Error GoTo sigue
art_familia.ListIndex = wpos
On Error GoTo 0
If art_familia.Enabled Then art_familia.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next
Exit Sub
muestra:
    lstfiltro.ToolTipText = "TAB_TIPREG = 122"
    lstfiltro.Clear
    lstfiltro.Tag = KeyCode
    pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 122 AND TAB_CODCIA = '" & LK_CODCIA & "' ORDER BY TAB_NOMLARGO"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  
    Do Until X.EOF
        lstfiltro.AddItem X!TAB_NOMLARGO & String(80, " ") & Trim(CStr(X!TAB_NUMTAB))
        If KeyCode = 116 Then
          If Nulo_Valor0(X!tab_codclie) = 1 Then
            lstfiltro.Selected(lstfiltro.ListCount - 1) = True
          End If
        Else
          If Nulo_Valor0(X!tab_codart) = 1 Then
            lstfiltro.Selected(lstfiltro.ListCount - 1) = True
          End If
        End If
        
        X.MoveNext
    Loop
    If KeyCode = 116 Then
    fralstfiltro.Caption = "Familia para Filtro! Reportes"
    Else
    fralstfiltro.Caption = "Familia Aceptan Cambios"
    End If
fralstfiltro.Visible = True
End Sub

Private Sub art_familia_LostFocus()
Exit Sub
txtnombre.Text = ARMA_NOMBRE
Dim wpos As Integer
Dim WFAMI2 As Integer
If Flag_Bloq = "A" Then
 Exit Sub
End If
If Trim(art_familia.Text) = "" Then
 art_subfam.Clear
 Exit Sub
End If
wpos = art_subfam.ListIndex
WFAMI2 = Val(Trim(Right(art_familia.Text, 6)))
PUB_TIPREG = 123
LLENADO_SUBFAM art_subfam, WFAMI2
Fvarios.Refresh
On Error GoTo sigue
art_subfam.ListIndex = wpos
Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_grupo_Click()
' art_numero.Clear
' art_linea.Clear
' art_marca.Clear
End Sub

Private Sub art_grupo_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub art_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_numero.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub art_grupo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wfami As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
On Error GoTo sigue
Dim wpos
Pres = "INSERT"
wpos = art_grupo.ListIndex
wfami = Val(Trim(Right(art_subfam.Text, 6)))
PUB_CODART = wfami
PUB_TIPREG = Mid(art_grupo.ToolTipText, 13, Len(art_grupo.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If

Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents

LLENADO_SUBFAM art_grupo, wfami

art_grupo.ListIndex = wpos
On Error GoTo 0
art_grupo.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_grupo_LostFocus()
Exit Sub
Fvarios.Refresh
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If

  If Trim(art_grupo.Text) = "" Then
   'art_numero.Clear
   Exit Sub
  End If
  wpos = art_numero.ListIndex
  WFAMI2 = Val(Trim(Right(art_grupo.Text, 6)))
  PUB_TIPREG = 130
  LLENADO_SUBFAM art_numero, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_numero.ListIndex = wpos
  Exit Sub

sigue:
Resume Next
End Sub


Private Sub art_grupop_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wswheregrupo As String
Dim wswherefami As String

Dim wpos
If KeyCode = 114 And LK_CODUSU = "ADMIN" Then
   wswheregrupo = InputBox("Colocar el codigo con que Grupo Reemplazar: ", "Grupo Aplicar.", 0)
   If wswheregrupo = "" Then Exit Sub
   wswherefami = InputBox("Colocar el codigo de Familia ", "Familia", 0)
   If wswherefami = "" Then Exit Sub
   pub_cadena = "update arti set art_grupop = " & Val(wswheregrupo) & " where art_familia = " & Val(wswherefami)
   CN.Execute pub_cadena
   MsgBox "Actulizado"
End If
If KeyCode <> 45 Then
 Exit Sub
End If
Flag_Bloq = "A"
wpos = art_grupop.ListIndex
PUB_TIPREG = Mid(art_grupop.ToolTipText, 13, Len(art_grupop.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPO PRINCIPAL -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
Flag_Bloq = ""
DoEvents
LLENADO_GRUPOP
DoEvents
On Error GoTo sigue
art_grupop.ListIndex = wpos
On Error GoTo 0
If art_grupop.Enabled Then art_grupop.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next
Exit Sub

End Sub

Private Sub art_linea_Click()
'  art_marca.Clear

End Sub

Private Sub art_linea_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wSubGrupo As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
Pres = "INSERT"
wpos = art_linea.ListIndex
wSubGrupo = Val(Trim(Right(art_numero.Text, 6)))
'**************************
'ojo
'**************************
If Not art_linea.ToolTipText = "" Then
 PUB_TIPREG = Mid(art_linea.ToolTipText, 13, Len(art_linea.ToolTipText))
End If
PUB_CODCIA = LK_CODCIA
PUB_CODART = wSubGrupo
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_linea, wSubGrupo
On Error GoTo sigue
art_linea.ListIndex = wpos
On Error GoTo 0
art_linea.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next


End Sub

Private Sub art_linea_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_marca.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub art_linea_LostFocus()
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If
 ' If Trim(art_linea.Text) = "" Then
 '  art_marca.Clear
 '  Exit Sub
 ' End If
  'wpos = art_marca.ListIndex
  'WFAMI2 = Val(Trim(Right(art_linea.Text, 6)))
  'PUB_TIPREG = 132
  'LLENADO_SUBFAM art_marca, WFAMI2
  'Fvarios.Refresh
  'On Error GoTo sigue
  'art_marca.ListIndex = wpos
  Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_plancha.SetFocus
 SendKeys "%{up}"
End If

End Sub
Private Sub art_marca_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
  OP_FORM = "C"
  Load frmprecol
  frmprecol.Show 1
  OP_FORM = ""
  Exit Sub
End If
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = art_marca.ListIndex
wfami = Val(Trim(Right(art_linea.Text, 6)))
PUB_TIPREG = Mid(art_marca.ToolTipText, 13, Len(art_marca.ToolTipText))
PUB_CODCIA = LK_CODCIA
PUB_CODART = wfami
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "Marca  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_marca, wfami
On Error GoTo sigue
art_marca.ListIndex = wpos
On Error GoTo 0
art_marca.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_marca_LostFocus()
txtnombre.Text = ARMA_NOMBRE
End Sub

Private Sub art_numero_Click()
'If Not art_numero.ListIndex = art_numero.ListIndex Then
' art_linea.Clear
' art_marca.Clear
'End If
End Sub

Private Sub art_numero_KeyDown(KeyCode As Integer, Shift As Integer)
Dim wgrupo As Integer
If KeyCode = 116 Then
  OP_FORM = "T"
  Load frmprecol
  frmprecol.Show 1
  OP_FORM = ""
  Exit Sub
End If
If KeyCode <> 45 Then
 Exit Sub
End If
Pres = "INSERT"
Dim wpos

wpos = art_numero.ListIndex
wgrupo = Val(Trim(Right(art_grupo.Text, 6)))
PUB_TIPREG = Mid(art_numero.ToolTipText, 13, Len(art_numero.ToolTipText))
PUB_CODCIA = LK_CODCIA
PUB_CODART = wgrupo
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "GRUPOS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_numero, wgrupo
On Error GoTo sigue
art_numero.ListIndex = wpos
On Error GoTo 0
art_numero.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_numero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_linea.SetFocus
 SendKeys "%{up}"
End If

End Sub

Private Sub art_numero_LostFocus()
Exit Sub
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If

  If Trim(art_numero.Text) = "" Then
   art_linea.Clear
   Exit Sub
  End If
  wpos = art_linea.ListIndex
  WFAMI2 = Val(Trim(Right(art_numero.Text, 6)))
  PUB_TIPREG = 131
  LLENADO_SUBFAM art_linea, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_linea.ListIndex = wpos
  Exit Sub
sigue:
Resume Next
End Sub

Private Sub art_plancha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 art_codpro.SetFocus
 SendKeys "%{up}"
End If
End Sub

Private Sub art_plancha_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = art_plancha.ListIndex
PUB_TIPREG = Mid(art_plancha.ToolTipText, 13, Len(art_plancha.ToolTipText))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
Load FrmDatArti
FrmDatArti.Caption = "Lote  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_PLANCHA
On Error GoTo sigue
art_plancha.ListIndex = wpos
On Error GoTo 0
art_plancha.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:

End Sub

Private Sub art_plancha_LostFocus()
txtnombre.Text = ARMA_NOMBRE
End Sub

Private Sub art_situacion_Click()
If Left(cmdModificar.Caption, 2) <> "&G" Then Exit Sub
If art_situacion.Value = 0 Then
  If flag_art_situacion = 1 Then
    If wflag_bloq_desactiva = "A" Then
    Else
       MsgBox "Solo puede Activar el Usuario ADMIN", 48, Pub_Titulo
       art_situacion.Value = 1
       Exit Sub
    End If
     cias_marca
  End If
End If
End Sub

Private Sub art_subfam_Click()
'art_grupo.Clear
'art_numero.Clear
'art_linea.Clear
'art_marca.Clear

End Sub

Private Sub art_subfam_GotFocus()
frmARTI.F14.Visible = False
End Sub
Private Sub art_subfam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   art_grupo.SetFocus
   SendKeys "%{up}"
   Exit Sub
End If
End Sub
Private Sub art_subfam_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wfami As Integer
If KeyCode <> 45 Then
 Exit Sub
End If
wfami = Val(Trim(Right(art_familia.Text, 6)))
If Mid(art_subfam.ToolTipText, 13, Len(art_subfam.ToolTipText)) = "" Then
  Exit Sub
End If
Pres = "INSERT"
'*******************
Dim wpos
wpos = art_subfam.ListIndex
PUB_TIPREG = Val(Mid(art_subfam.ToolTipText, 13, Len(art_subfam.ToolTipText)))
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
PUB_CODART = wfami
Load FrmDatArti
FrmDatArti.Caption = "SUB-FAMILIAS  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_SUBFAM art_subfam, wfami
On Error GoTo sigue
art_subfam.ListIndex = wpos
On Error GoTo 0
art_subfam.SetFocus
SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub art_subfam_LostFocus()
Exit Sub
Fvarios.Refresh
txtnombre.Text = ARMA_NOMBRE
If Pres = "INSERT" Then
 Pres = ""
 Exit Sub
End If
  If Trim(art_subfam.Text) = "" Then
   art_grupo.Clear
   Exit Sub
  End If
  wpos = art_grupo.ListIndex
  WFAMI2 = Val(Trim(Right(art_subfam.Text, 6)))
  PUB_TIPREG = 129
  LLENADO_SUBFAM art_grupo, WFAMI2
  Fvarios.Refresh
  On Error GoTo sigue
  art_grupo.ListIndex = wpos
  Exit Sub

Exit Sub
sigue:
Resume Next
End Sub

Private Sub CERO_Click()
PUB_KEY = 0
PUB_CODCIA = LK_CODCIA
LOC_OPER = 2
LEER_LOC
If artloc_key.EOF Then
  ARTI_CERO
  MsgBox "Producto Creado ..", 48, Pub_Titulo
End If
End Sub


Private Sub cheapor_Click()
If cheapor.Value = 1 Then
  cmdapor.Visible = True
  lblapor(0).Visible = True
  lblapor(1).Visible = True
  txtapor.Visible = True
  Azul txtapor, txtapor
Else
  cmdapor.Visible = False
  lblapor(0).Visible = False
  lblapor(1).Visible = False
  txtapor.Visible = False

End If
End Sub

Private Sub cheasignarc_Click()
 cmdqc.Visible = False
If cheasignarc.Value = 1 Then
  If loc_acc_convenio = "A" Then
    cmdqc.Visible = True
  Else
    cmdqc.Visible = False
     MsgBox "Usted. No tiene Acceso a Asignar ", 48, Pub_Titulo
     cheasignarc.Value = 0
     Exit Sub
  End If
  
End If


End Sub

Private Sub cheespecial_Click()
If cheespecial.Value = 1 Then
  cmdlineas.Visible = True
Else
  cmdlineas.Visible = False
End If

End Sub

Private Sub chep1_Click()
  If chep1.Value = 1 Then
  tpo1.Visible = True
  R1.Visible = True
  Azul tpo1, tpo1
  Else
  tpo1.Visible = False
  R1.Visible = False
  tpo1.Text = "0.00"
  End If
End Sub

Private Sub chep2_Click()
  If chep2.Value = 1 Then
  tpo2.Visible = True
  R2.Visible = True
  Azul tpo2, tpo2
  Else
  tpo2.Visible = False
  R2.Visible = False
  tpo2.Text = "0.00"
  End If
End Sub

Private Sub chep3_Click()
  If chep3.Value = 1 Then
  R3.Visible = True
  tpo3.Visible = True
  Azul tpo3, tpo3
  Else
  R3.Visible = False
  tpo3.Visible = False
  tpo3.Text = "0.00"
  End If

End Sub

Private Sub chep4_Click()
  If chep4.Value = 1 Then
  R4.Visible = True
  tpo4.Visible = True
  Azul tpo4, tpo4
  Else
  R4.Visible = False
  tpo4.Visible = False
  tpo4.Text = "0.00"
  End If

End Sub

Private Sub chep5_Click()
  If chep5.Value = 1 Then
  R5.Visible = True
  tpo5.Visible = True
  Azul tpo5, tpo5
  Else
  R5.Visible = False
  tpo5.Visible = False
  tpo5.Text = "0.00"
  End If

End Sub

Private Sub chep6_Click()
  If chep6.Value = 1 Then
  R6.Visible = True
  tpo6.Visible = True
  Azul tpo6, tpo6
  Else
  R6.Visible = True
  tpo6.Visible = False
  tpo6.Text = "0.00"
  End If
End Sub

Private Sub chep7_Click()
  If chep7.Value = 1 Then
  tpo7.Visible = True
  Azul tpo7, tpo7
  Else
  tpo7.Visible = False
  tpo7.Text = "0.00"
  End If

End Sub

Private Sub chep8_Click()
  If chep8.Value = 1 Then
  R8.Visible = True
  tpo8.Visible = True
  Azul tpo8, tpo8
  Else
  R8.Visible = False
  tpo8.Visible = False
  tpo8.Text = "0.00"
  End If

End Sub

Private Sub chep9_Click()
  If chep9.Value = 1 Then
    R9.Visible = True
   tpo9.Visible = True
  Azul tpo9, tpo9
  Else
  R9.Visible = False
  tpo9.Visible = False
  tpo9.Text = "0.00"
  End If

End Sub

Private Sub cmbactprecios_Click()
Load frmpreciocia
frmpreciocia.Show 1

End Sub

Private Sub cmbactprecios_DblClick()


End Sub

Private Sub cmbcal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdAdd_Click
End Sub

Private Sub CmbCalidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  art_codpro.SetFocus
  DoEvents
  SendKeys "%{up}"
End If

End Sub

Private Sub CmbCalidad_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 45 Then
 Exit Sub
End If
Dim wpos
wpos = CmbCalidad.ListIndex
PUB_TIPREG = Mid(CmbCalidad.ToolTipText, 13, Len(CmbCalidad.ToolTipText))
PUB_CODCIA = LK_CODCIA
Load FrmDatArti
FrmDatArti.Caption = "CALIDAD  -  TAB_TIPREG = " & PUB_TIPREG
FrmDatArti.Show 1
DoEvents
LLENADO_CAL
On Error GoTo sigue
CmbCalidad.ListIndex = wpos
On Error GoTo 0
'CmbCalidad.SetFocus
'SendKeys "%{up}"
Exit Sub
sigue:
Resume Next

End Sub

Private Sub CmbCalidad_LostFocus()
Fvarios.Refresh
End Sub

Private Sub cmbStock_Click()
Dim fg As String
Dim wWRITE As String
Dim RUTA  As String
Dim WfART_llave As rdoResultset
Dim WPSART_LLAVE As rdoQuery

Dim PSPRE_LOTES As rdoQuery
Dim pre_lotes As rdoResultset



pub_cadena = "SELECT SUM(FAR_CANTIDAD * FAR_SIGNO_ARM) AS STOCK FROM FACART WHERE FAR_CODART = ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E' AND FAR_ESTADO2 <> 'L' "
Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
PSPRE_LOTES(0) = 0
PSPRE_LOTES(1) = 0
Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT ARM_CODART,ARM_STOCK FROM ARTICULO WHERE ARM_CODCIA = ? AND ARM_CODART = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
RUTA = "C:\" & "Stock.txt"

Open RUTA For Output As #1
fg = ""
Do Until art_mayor.EOF
   If art_mayor!art_key = 0 Then GoTo SIG
    cmdAgregar.Caption = art_mayor.AbsolutePosition & " / " & art_mayor.RowCount
    DoEvents
    WPSART_LLAVE(0) = art_mayor!art_codcia
    WPSART_LLAVE(1) = art_mayor!art_key
    WfART_llave.Requery
  
    PSPRE_LOTES(0) = art_mayor!art_key
    PSPRE_LOTES(1) = art_mayor!art_codcia
    pre_lotes.Requery
    If Val(Nulo_Valor0(pre_lotes!stock)) <> Val(WfART_llave!arm_stock) Then
       Print #1, Format(art_mayor!art_key, "00000000") & " , " & art_mayor!art_nombre & ",  Dif. = " & (Val(WfART_llave!arm_stock) - Val(Nulo_Valor0(pre_lotes!stock)))
       fg = "A"
    End If
   DoEvents
SIG:
   art_mayor.MoveNext
Loop
If Trim(fg) = "" Then
  Print #1, "No Hay Diferencias"
End If
Close #1

MsgBox "TERMINO DE CHEQUEAR STOCK BUSCAR EL ARCHIVO C:\STOCK.TXT"

End Sub

Private Sub cmdact_Click()
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

If flag_cambio_pre <> "A" Then
  MsgBox "No tiene Acceso !!!", 48, Pub_Titulo
  Exit Sub
End If

fraact6.Visible = True
End Sub

Private Sub cmdact6_Click()
Dim WVA As Currency
Dim ARMAD  As String
Dim ARMAS As String
If Val(cmbtipo.Text) = 1 Or Val(cmbtipo.Text) = -1 Then
Else
  MsgBox "Selecione el Tipo de Operaci�n. ", 48, Pub_Titulo
  Exit Sub
End If
pub_mensaje = "Actualziar esta Lista �Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If



If Val(txtnrolista.Text) >= 1 And Val(txtnrolista.Text) <= 12 Then
Else
  MsgBox "No Procede solo para las Lista 1,2,3,4,5,6,7,8,9,10,11,12,25", 48, Pub_Titulo
End If
WVA = Val(txtporlista.Text)

If Val(cmbtipo.Text) = 1 Then
WVA = 1 + (WVA / 100)
Else
WVA = ((100 - WVA) / 100)
End If




If Val(Right(Trim(cmdlista.Text), 2)) = 1 Then
 c_lista_d = "11"
 c_lista_s = "1"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 2 Then
 c_lista_d = "22"
 c_lista_s = "2"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 3 Then
 c_lista_d = "33"
 c_lista_s = "3"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 4 Then
 c_lista_d = "44"
 c_lista_s = "4"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 5 Then
 c_lista_d = "55"
 c_lista_s = "5"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 6 Then
 c_lista_d = "66"
 c_lista_s = "6"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 7 Then
 c_lista_d = "77"
 c_lista_s = "7"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 8 Then
 c_lista_d = "88"
 c_lista_s = "8"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 9 Then
 c_lista_d = "99"
 c_lista_s = "9"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 10 Then
 c_lista_d = "CHI11"
 c_lista_s = "CHI1"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 11 Then
 c_lista_d = "C11"
 c_lista_s = "C1"
ElseIf Val(Right(Trim(cmdlista.Text), 2)) = 12 Then
 c_lista_d = "C22"
 c_lista_s = "C2"
End If

    
    
    If Val(txtnrolista.Text) = 1 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE11 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE1 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 2 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE22 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE2 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 3 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE33 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE3 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 4 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE44 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE4 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 5 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE55 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE5 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 6 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE66 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE6 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 7 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE77 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE7 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 8 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE88 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE8 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 9 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRE99 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRE9 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 10 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PRECHI11 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PRECHI1 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 11 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PREC11 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PREC1 * " & WVA & ")"
    ElseIf Val(txtnrolista.Text) = 12 Then
      ARMAD = "PRE_PRE" & c_lista_d & " = (PRE_PREC22 * " & WVA & ")"
      ARMAS = "PRE_PRE" & c_lista_s & " = (PRE_PREC2 * " & WVA & ")"
    End If
If LK_CODCIA = "25" Or LK_CODCIA = "27" Or LK_CODCIA = "28" Or LK_CODCIA = "29" Then
  MsgBox "Actualizara en bloque a los Codigos de Cias: 25,27,28,29", 48, Pub_Titulo
  Screen.MousePointer = 13
 If cmddolares.Tag = "D" Then
   CN.Execute "UPDATE PRECIOS " & ARMAD & "  WHERE PRE_CODCIA IN ('25','27','28','29')"
 Else
   CN.Execute "UPDATE PRECIOS SET " & ARMAS & "  WHERE PRE_CODCIA  IN ('25','27','28','29')"
 End If
 Screen.MousePointer = 0
Else
 If cmddolares.Tag = "D" Then
   CN.Execute "UPDATE PRECIOS " & ARMAD & "  WHERE PRE_CODCIA =  '" & LK_CODCIA & "' " ' AND PRE_CODART = " & Trim(txt_key.Text) & " "
 Else
   CN.Execute "UPDATE PRECIOS SET " & ARMAS & "  WHERE PRE_CODCIA =  '" & LK_CODCIA & "' " ' AND PRE_CODART = " & Trim(txt_key.Text) & " "
 End If
End If
 MsgBox "Actualizacion Terminada", 48, Pub_Titulo

End Sub

Private Sub cmdActBloque_Click()
Load frmpreciocia
frmpreciocia.Show 1

End Sub

Private Sub cmdactlineas_Click()
Dim WS_PRE As String
Dim WS_VALOR_PRE As String
Dim llave_listart As rdoResultset
Dim PS_LISTART As rdoQuery
Dim CADE_LINEA  As String
If LK_CODCIA = "01" Or LK_CODCIA = "03" Or LK_CODCIA = "07" Or LK_CODCIA = "10" Then
Else
  MsgBox "No Procede en estas Sedes", 48, Pub_Titulo
  Exit Sub
End If
MsgBox "Usted a escogido la opcion de : " & Chr(13) & cmdptipo.Text, 48, Pub_Titulo
pub_mensaje = "Confirmar la actualizacion de Precios...   �Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If


CADE_LINEA = ""
For fila = 0 To listlineas.ListCount - 1
  If listlineas.Selected(fila) Then
     CADE_LINEA = CADE_LINEA + Trim(Right(listlineas.List(fila), 8)) + ", "
  End If
Next fila
If Trim(CADE_LINEA) <> "" Then
  CADE_LINEA = " AND ART_FAMILIA IN (" & Mid(CADE_LINEA, 1, Len(CADE_LINEA) - 2) & ")"
End If

pub_cadena = "SELECT ART_CODCIA, ART_KEY, ARM_COSPRO FROM ARTI, ARTICULO WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_CODCIA = '" & LK_CODCIA & "' " & CADE_LINEA
Set PS_LISTART = CN.CreateQuery("", pub_cadena)
Set llave_listart = PS_LISTART.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

If Val(Left(cmdptipo.Text, 2)) = 1 Then GoTo AUMENTO_COSTO
If Val(Left(cmdptipo.Text, 2)) = 2 Then GoTo DISMINUCION_LISTA

Exit Sub

'--------------
AUMENTO_COSTO:
'--------------
pbpValue = 0
pbpMin = 0

llave_listart.Requery
If Not llave_listart.EOF Then
   pbpMax = llave_listart.RowCount
   pbp.Visible = True
End If
SQ_OPER = 2
pu_codcia = LK_CODCIA
Do Until llave_listart.EOF
   pbpValue = pbpValue + 1
   If chep1.Value = 1 Then
      cade_multi = 1 + (Val(tpo1.Text) / 100)
      If chere1.Value = 1 Then R1.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR1 = " & tpo1.Text & " ,PRE_PRE1 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R1.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      
      pub_cadena = "UPDATE PRECIOS SET PRE_POR1 = " & tpo1.Text & " ,PRE_PRE1 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R1.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      
      If chere1.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualizo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_PRE = Format(pre_mayor!PRE_PRE1, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
            Else
              WS_VALOR_PRE = 0
            End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE1 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE1 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep2.Value = 1 Then
      cade_multi = 1 + (Val(tpo2.Text) / 100)
      If chere2.Value = 1 Then R2.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR2 = " & tpo2.Text & " , PRE_PRE2 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R2.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      
      pub_cadena = "UPDATE PRECIOS SET PRE_POR2 = " & tpo2.Text & " ,PRE_PRE2 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R2.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      
      If chere2.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_PRE = Format(pre_mayor!PRE_PRE2, "0.0000")
            WS_VALOR_PRE = pre_mayor!PRE_PRE2
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
            Else
             WS_VALOR_PRE = 0
            End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE2 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE2 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep3.Value = 1 Then
      cade_multi = 1 + (Val(tpo3.Text) / 100)
      If chere3.Value = 1 Then R3.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR3 = " & tpo3.Text & " , PRE_PRE3 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R3.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      
      pub_cadena = "UPDATE PRECIOS SET PRE_POR3 = " & tpo3.Text & " ,PRE_PRE3 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R3.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      
      If chere3.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE3
            WS_PRE = Format(pre_mayor!PRE_PRE3, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
            WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE3 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE3 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   
   If chep4.Value = 1 Then
      cade_multi = 1 + (Val(tpo4.Text) / 100)
      If chere4.Value = 1 Then R4.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR4 = " & tpo4.Text & " , PRE_PRE4 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R4.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR4 = " & tpo4.Text & " ,PRE_PRE4 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R4.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere4.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
            If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE4
            WS_PRE = Format(pre_mayor!PRE_PRE4, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
            Else
             WS_VALOR_PRE = 0
            End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE4 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE4 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep5.Value = 1 Then
      cade_multi = 1 + (Val(tpo5.Text) / 100)
      If chere5.Value = 1 Then R5.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR5 = " & tpo5.Text & " , PRE_PRE5 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R5.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR5 = " & tpo5.Text & " ,PRE_PRE5 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R5.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere5.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE5
            WS_PRE = Format(pre_mayor!PRE_PRE5, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
            WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE5 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE5 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep6.Value = 1 Then
      cade_multi = 1 + (Val(tpo6.Text) / 100)
      If chere6.Value = 1 Then R6.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR6 = " & tpo6.Text & " , PRE_PRE6 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R6.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR6 = " & tpo6.Text & " ,PRE_PRE6 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R6.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere6.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE6
            WS_PRE = Format(pre_mayor!PRE_PRE6, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
           WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE6 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE6 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep7.Value = 1 Then
      cade_multi = 1 + (Val(tpo7.Text) / 100)
      If chere7.Value = 1 Then R7.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR7 = " & tpo7.Text & " , PRE_PRE7 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R7.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR7 = " & tpo7.Text & " ,PRE_PRE7 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R7.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere7.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE7
            WS_PRE = Format(pre_mayor!PRE_PRE7, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
            WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE7 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE7 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep8.Value = 1 Then
      cade_multi = 1 + (Val(tpo8.Text) / 100)
      If chere8.Value = 1 Then R8.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR8 = " & tpo8.Text & " , PRE_PRE8 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R8.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR8 = " & tpo8.Text & " ,PRE_PRE8 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R8.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere8.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE8
            WS_PRE = Format(pre_mayor!PRE_PRE8, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
            WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE8 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE8 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
   
   If chep9.Value = 1 Then
      cade_multi = 1 + (Val(tpo9.Text) / 100)
      If chere9.Value = 1 Then R9.ListIndex = 4
      'pub_cadena = "UPDATE PRECIOS SET PRE_POR9 = " & tpo9.Text & " , PRE_PRE9 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R9.Text) & ") WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND  PRE_CODART = " & llave_listart!ART_KEY
      'CN.Execute pub_cadena, rdExecDirect
      pub_cadena = "UPDATE PRECIOS SET PRE_POR9 = " & tpo9.Text & " ,PRE_PRE9 = round( (" & llave_listart!ARM_COSPRO & " * PRE_EQUIV ) * " & cade_multi & "," & Trim(R9.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
      If chere8.Value = 1 Then
       PUB_CODART = llave_listart!art_key
       LEER_PRE_LLAVE
       If pre_mayor.EOF Then
          MsgBox "No Actualiazo en codigo :" & llave_listart!art_key
       Else
         Do Until pre_mayor.EOF
           If pre_mayor!PRE_EQUIV = 1 And pre_mayor.RowCount <> 1 Then
            WS_VALOR_PRE = pre_mayor!PRE_PRE9
            WS_PRE = Format(pre_mayor!PRE_PRE9, "0.0000")
            If Val(Mid(WS_PRE, Len(Trim(WS_PRE)) - 2, Len(Trim(WS_PRE)))) <> 0 Then
                 WS_PRE = Mid(WS_PRE, 1, Len(Trim(WS_PRE)) - 3)
                 WS_VALOR_PRE = Val(WS_PRE) + 0.1
            End If
           Else
             WS_VALOR_PRE = 0
           End If
            pre_mayor.Edit
            pre_mayor!PRE_PRE9 = WS_VALOR_PRE
            pre_mayor.Update
            pub_cadena = "UPDATE PRECIOS SET PRE_PRE9 = " & WS_VALOR_PRE & " WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key & " AND PRE_SECUENCIA = " & pre_mayor!PRE_SECUENCIA
            CN.Execute pub_cadena, rdExecDirect
            pre_mayor.MoveNext
         Loop
       End If
      End If
   End If
 llave_listart.MoveNext
Loop
pbp.Visible = False
MsgBox "Lista Actualizada.", 48, Pub_Titulo
fraactlinea.Visible = False
Exit Sub
'--------------
DISMINUCION_LISTA:
'-------------
pbpValue = 0
pbpMin = 0

llave_listart.Requery
If Not llave_listart.EOF Then
   pbpMax = llave_listart.RowCount
   pbp.Visible = True
End If

Do Until llave_listart.EOF
   pbpValue = pbpValue + 1
   If chep1.Value = 1 Then
      cade_multi = Val(tpo1.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE1 = Round( (PRE_PRE1 * (100 - " & cade_multi & "))/100 ," & Trim(R1.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
  If chep2.Value = 1 Then
      cade_multi = Val(tpo2.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE2 = Round( (PRE_PRE2 * (100 - " & cade_multi & "))/100 ," & Trim(R2.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep3.Value = 1 Then
      cade_multi = Val(tpo3.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE3 = Round( (PRE_PRE3 * (100 - " & cade_multi & "))/100 ," & Trim(R3.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep4.Value = 1 Then
      cade_multi = Val(tpo4.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE4 = Round( (PRE_PRE4 * (100 - " & cade_multi & "))/100 ," & Trim(R4.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep5.Value = 1 Then
      cade_multi = Val(tpo5.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE5 = Round( (PRE_PRE5 * (100 - " & cade_multi & "))/100 ," & Trim(R5.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep6.Value = 1 Then
      cade_multi = Val(tpo6.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE6 = Round( (PRE_PRE6 * (100 - " & cade_multi & "))/100 ," & Trim(R6.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep7.Value = 1 Then
      cade_multi = Val(tpo7.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE7 = Round( (PRE_PRE7 * (100 - " & cade_multi & "))/100 ," & Trim(R7.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10')  AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep8.Value = 1 Then
      cade_multi = Val(tpo8.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE8 = Round( (PRE_PRE8 * (100 - " & cade_multi & "))/100 ," & Trim(R8.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
   If chep9.Value = 1 Then
      cade_multi = Val(tpo9.Text)
      pub_cadena = "UPDATE PRECIOS SET PRE_PRE9 = Round( (PRE_PRE9 * (100 - " & cade_multi & "))/100 ," & Trim(R9.Text) & ") WHERE PRE_CODCIA IN ('01','03','07','10') AND  PRE_CODART = " & llave_listart!art_key
      CN.Execute pub_cadena, rdExecDirect
   End If
 llave_listart.MoveNext
Loop
pbp.Visible = False
MsgBox "Lista Actualizada.", 48, Pub_Titulo
fraactlinea.Visible = False

Exit Sub

End Sub

Private Sub cmdactLista_Click()
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

If flag_cambio_pre <> "A" Then
  MsgBox "No tiene Acceso !!!", 48, Pub_Titulo
  Exit Sub
End If

chep1.Value = 0
chep2.Value = 0
chep3.Value = 0
chep4.Value = 0
chep5.Value = 0
chep6.Value = 0
chep7.Value = 0
chep8.Value = 0
chep9.Value = 0

cmdptipo.Clear
cmdptipo.AddItem "1 =Porcentaje Adicional al Costo Promedio C/IGV"
cmdptipo.AddItem "2 =Porcentaje de Disminuci�n al Precio C/IGV"
cmdptipo.ListIndex = 0
fraactlinea.Visible = True
End Sub

Private Sub cmdactvend_Click()
Dim var_codven As Integer
Dim var_codart  As Currency
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If


var_codven = Val(Left(cmdvend.Text, 3))
var_codart = Val(txt_key.Text)
If var_codven = 0 Then Exit Sub
If var_codart = 0 Then
  MsgBox "No procede, especifique un producto y un Vendedor ", 48, Pub_Titulo
  Exit Sub
End If


If cheespecial.Value = 1 Then GoTo act_especicial
PSCOMI_VEN.rdoParameters(0) = LK_CODCIA
PSCOMI_VEN.rdoParameters(1) = var_codart
PSCOMI_VEN.rdoParameters(2) = var_codven
comivem_llave.Requery
If comivem_llave.EOF Then
  comivem_llave.AddNew
  comivem_llave!COA_CODCIA = LK_CODCIA
  comivem_llave!COA_CODART = var_codart
  comivem_llave!COA_CODVEN = var_codven
Else
  comivem_llave.Edit
End If
comivem_llave!COA_COMV1 = Val(txtpor1.Text)
comivem_llave!COA_COMV2 = Val(txtpor2.Text)
comivem_llave!COA_COMV3 = Val(txtpor3.Text)
comivem_llave!COA_COMV4 = Val(txtpor4.Text)
comivem_llave!COA_COMV5 = Val(txtpor5.Text)
comivem_llave!COA_COMV6 = Val(txtpor6.Text)
comivem_llave!COA_COMV7 = Val(txtpor7.Text)
comivem_llave!COA_COMV8 = Val(txtpor8.Text)
comivem_llave!COA_COMV9 = Val(txtpor9.Text)


comivem_llave!COA_COMC1 = Val(txtpor11.Text)
comivem_llave!COA_COMC2 = Val(txtpor22.Text)
comivem_llave!COA_COMC3 = Val(txtpor33.Text)
comivem_llave!COA_COMC4 = Val(txtpor44.Text)
comivem_llave!COA_COMC5 = Val(txtpor55.Text)
comivem_llave!COA_COMC6 = Val(txtpor66.Text)
comivem_llave!COA_COMC7 = Val(txtpor77.Text)
comivem_llave!COA_COMC8 = Val(txtpor88.Text)
comivem_llave!COA_COMC9 = Val(txtpor99.Text)

comivem_llave.Update
cmdactvend.Enabled = False

Exit Sub
act_especicial:
Dim WPSCOMI_LLAVE As rdoQuery
Dim WCOMI_llave As rdoResultset


Dim WPSART_LLAVE As rdoQuery
Dim WfART_llave As rdoResultset

Dim llave_data As rdoResultset
Dim ps_data As rdoQuery

pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? AND VEM_CODVEN = " & var_codven
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMIARTI WHERE COA_CODCIA = ? AND COA_CODART = ? AND COA_CODVEN = ? "
Set WPSCOMI_LLAVE = CN.CreateQuery("", pub_cadena)
WPSCOMI_LLAVE(0) = LK_CODCIA
WPSCOMI_LLAVE(1) = 0
WPSCOMI_LLAVE(2) = 0
Set WCOMI_llave = WPSCOMI_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

If Val(Trim(Right(cmdlineas, 8))) = 0 Then
  pub_cadena = "SELECT * from arti where art_codcia = ? " 'And art_familia = " & Val(Trim(Right(cmdlineas, 8)))"
Else
  pub_cadena = "SELECT * from arti where art_codcia = ? And art_familia = " & Val(Trim(Right(frmARTI.cmdlineas.Text, 8)))
End If
Set ps_data = CN.CreateQuery("", pub_cadena)
ps_data(0) = LK_CODCIA
Set llave_data = ps_data.OpenResultset(rdOpenKeyset, rdConcurValues)
llave_data.Requery

Do Until llave_data.EOF
  cmdactvend.Caption = llave_data.AbsolutePosition & "/" & llave_data.RowCount
  DoEvents
   ' If llave_data!ART_KEY = 72170 Then Stop
    WfART_llave.Requery
    Do Until WfART_llave.EOF
       WPSCOMI_LLAVE(0) = LK_CODCIA
       WPSCOMI_LLAVE(1) = llave_data!art_key
       WPSCOMI_LLAVE(2) = WfART_llave!VEM_CODVEN
       WCOMI_llave.Requery
       If WCOMI_llave.EOF Then
           WCOMI_llave.AddNew
       Else
           WCOMI_llave.Edit
       End If
       WCOMI_llave!COA_CODCIA = LK_CODCIA
       WCOMI_llave!COA_CODART = llave_data!art_key
       WCOMI_llave!COA_CODVEN = WfART_llave!VEM_CODVEN
       WCOMI_llave!COA_COMV1 = Val(txtpor1.Text)
       WCOMI_llave!COA_COMV2 = Val(txtpor2.Text)
       WCOMI_llave!COA_COMV3 = Val(txtpor3.Text)
       WCOMI_llave!COA_COMV4 = Val(txtpor4.Text)
       WCOMI_llave!COA_COMV5 = Val(txtpor5.Text)
       WCOMI_llave!COA_COMV6 = Val(txtpor6.Text)
       WCOMI_llave!COA_COMV7 = Val(txtpor7.Text)
       WCOMI_llave!COA_COMV8 = Val(txtpor8.Text)
       WCOMI_llave!COA_COMV9 = Val(txtpor9.Text)
       
       WCOMI_llave!COA_COMC1 = Val(txtpor11.Text)
       WCOMI_llave!COA_COMC2 = Val(txtpor22.Text)
       WCOMI_llave!COA_COMC3 = Val(txtpor33.Text)
       WCOMI_llave!COA_COMC4 = Val(txtpor44.Text)
       WCOMI_llave!COA_COMC5 = Val(txtpor55.Text)
       WCOMI_llave!COA_COMC6 = Val(txtpor66.Text)
       WCOMI_llave!COA_COMC7 = Val(txtpor77.Text)
       WCOMI_llave!COA_COMC8 = Val(txtpor88.Text)
       WCOMI_llave!COA_COMC9 = Val(txtpor99.Text)
       WCOMI_llave.Update
      WfART_llave.MoveNext
    Loop
llave_data.MoveNext
Loop

cmdactvend.Caption = "Actualizar"
MsgBox "Actualizaci�n Terminada "
cheespecial.Value = 0
cmdactvend.Enabled = False


End Sub

Private Sub cmdAdd_Click()
Dim wnombre
Dim wrellena As String
Dim WART_COSPRO As Currency
Dim WART_COSPRO_ANT As Currency
Dim WART_COSTO_ULT As Currency

If Trim(cmbcal.Text) = "" Then
 MsgBox " Seleccione su Calidad.", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If
If Left(cmbcal.Text, 1) = "<" Then
 MsgBox "No Existe mas Calidades ", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If

wnombre = InputBox("Ingrese la Descripci�n del Articulo :", Pub_Titulo, Trim(txtnombre.Text))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
If Trim(wnombre) = "" Then
  Screen.MousePointer = 0
  MsgBox "Descripci�n NO Validad.", 48, Pub_Titulo
  Exit Sub
End If
LOC_NOMBRE = wnombre
VAR_NEWCAL = 1
LOC_CALIDAD = Val(Right(cmbcal.Text, 3))
wrellena = String(LOC_CALIDAD - 1, "*")
LOC_ALTERNO = Trim(txt_alterno.Text) + wrellena
pu_codcia = LK_CODCIA
PUB_CODART = artloc_llave!art_key
SQ_OPER = 1
LEER_ARM_LLAVE
WART_COSPRO = arm_llave!ARM_COSPRO
WART_COSTO_ULT = arm_llave!ARM_COSTO_ULT

On Error GoTo ESCAPA
CN.Execute "Begin Transaction", rdExecDirect
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
LOC_ORIGINAL = GENERA_CODI()
LOC_CODART2 = LOC_ORIGINAL
GRABAR_ARTI
PUB_CODART = LOC_ORIGINAL
SQ_OPER = 1
LEER_ARM_LLAVE
If arm_llave.EOF Then
   Screen.MousePointer = 0
   arm_llave.AddNew
   arm_llave!ARM_CODART = LOC_ORIGINAL
   arm_llave!ARM_CODCIA = LK_CODCIA
   arm_llave!arm_stock = 0
   arm_llave!ARM_INGRESOS = 0
   arm_llave!ARM_SALIDAS = 0
   arm_llave!ARM_COSPRO = WART_COSPRO
   arm_llave!arm_stock2 = 0
   arm_llave!ARM_Saldo_n = 0
   arm_llave!arm_saldo_n2 = 0
   arm_llave!ARM_saldo_s = 0
   arm_llave!arm_saldo_s2 = 1
   arm_llave!ARM_COSTO_ULT = WART_COSTO_ULT
   arm_llave.Update
Else
  MsgBox "Codigo Existe en tabla: Articulo verificar ...", 48, Pub_Titulo
  GoTo ESCAPA
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   pu_alterno = Trim(txt_alterno.Text)
Else
   PUB_KEY = Val(txt_key.Text)
End If
PUB_CODCIA = LK_CODCIA
LOC_OPER = 1
LEER_LOC
If artloc_llave.EOF Then
  GoTo ESCAPA
 Exit Sub
End If

artloc_llave.Edit
artloc_llave!ART_CODART2 = LOC_CODART2
'artloc_llave!ART_COSPRO = Val(tcospro.text)
artloc_llave.Update
gridrel.TextMatrix(1, 0) = LOC_ORIGINAL
gridrel.TextMatrix(1, 1) = LOC_ALTERNO
gridrel.TextMatrix(1, 2) = LOC_NOMBRE
gridrel.TextMatrix(1, 3) = Left(cmbcal.Text, 40)
txtcodigo2.Text = LOC_CODART2
cmbcal.Visible = False
lblcal.Visible = False
cmdAdd.Visible = False
cmdquitar.Visible = True
con_llave.Close
CN.Execute "Commit Transaction", rdExecDirect
VAR_NEWCAL = 0
On Error GoTo 0
Exit Sub
ESCAPA:
VAR_NEWCAL = 0
If con_llave Is Nothing Then
 con_llave.Close
 CN.Execute "Rollback Transaction", rdExecDirect
End If
MsgBox " Intente Nuevamente..", 48, Pub_Titulo

     
End Sub

Private Sub AgregarUnidad_Click()
    If grid_unid.Rows > 1 Then
        grid_unid.Rows = grid_unid.Rows + 1
    End If
    grid_unid.Row = grid_unid.Rows - 1
    grid_unid.Col = 0
    grid_unid.SetFocus
End Sub

Private Sub cmdAgregar_Click()
Dim loc_flag As String
If Trim(txt_key.Text) = "1" Then
     MENSAJE_ARTI "No Procede. .."
     Exit Sub
End If
If Left(cmdAgregar.Caption, 2) = "&A" Then

    loc_flag = ""
    For fila = 1 To lk_OTROS_Count
       If Val(lk_OTROS(fila)) = 35 Then loc_flag = "A"
    Next fila


    If loc_flag = "A" Then
    Else
       MsgBox "Usted. No tiene Acceso a Crear Productos.", 48, Pub_Titulo
       Exit Sub
    End If
    ws_flag_equiv = "A"
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_ARTI
    frmARTI.decimales.ListIndex = 1
    If frmARTI.CmbCalidad.ListCount <> 0 Then
       frmARTI.CmbCalidad.ListIndex = 0
    End If
    If frmARTI.art_grupo.ListCount <> 0 Then
       frmARTI.art_grupo.ListIndex = 0
    End If
    frmARTI.txt_key = GENERA_CODI
    frmARTI.txt_alterno.Text = frmARTI.txt_key.Text
    DESBLOQUEA_TEXT txtnombre, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    DESBLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    DESBLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    DESBLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    DESBLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    BLOQUEA_TEXT txt_key
    If LK_MONEDA = "D" Then
        frmARTI.DS.ListIndex = 1
        frmARTI.DS.Enabled = False
    ElseIf LK_MONEDA = "S" Then
        frmARTI.DS.ListIndex = 0
        frmARTI.DS.Enabled = False
    ElseIf LK_MONEDA = "A" Then
        frmARTI.DS.ListIndex = -1
    End If
    Flag_Inicial = "A"
    CABEZA_UNID
    loc_equiv_max = 1
    'grid_unid.Rows = grid_unid.Rows + 1
    grid_unid.RowHeight(1) = 285
    If LK_EMP = "PIU" Then
       grid_unid.TextMatrix(1, 0) = "PARES"
    ElseIf LK_EMP = "CAM" Then
       grid_unid.TextMatrix(1, 0) = "Kg."
    Else
       grid_unid.TextMatrix(1, 0) = "UND"
    End If
    grid_unid.TextMatrix(1, 1) = "1.00"
    grid_unid.TextMatrix(1, 3) = "0.00"
    grid_unid.COL = 4
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 4) = "0.00"
    grid_unid.TextMatrix(1, 5) = "0.0000"
    grid_unid.COL = 6
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 6) = "0.00"
    grid_unid.TextMatrix(1, 7) = "0.0000"
    grid_unid.COL = 8
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 8) = "0.00"
    grid_unid.TextMatrix(1, 9) = "0.0000"
    grid_unid.COL = 10
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 10) = "0.00"
    grid_unid.TextMatrix(1, 11) = "0.0000"
    grid_unid.COL = 12
    grid_unid.CellForeColor = QBColor(9)
    grid_unid.TextMatrix(1, 12) = "0.00"
    grid_unid.TextMatrix(1, 13) = "0.0000"
    grid_unid.TextMatrix(1, 16) = "A"
    Flag_Inicial = ""
    grid_unid.COL = 0
    LBLUNIDAD.Caption = "UNIDAD"
    DESBLOQUEA_TEXT txt_alterno
    pasa = 1
    If txt_alterno.Visible And txt_alterno.Enabled Then
      txt_alterno.SetFocus
    ElseIf txtnombre.Visible And txtnombre.Enabled Then
     txtnombre.SetFocus
    End If
    cheservi(0).Value = True
    MANOS(0).Enabled = False
    MANOS(1).Enabled = False
    Fcomun.Refresh
    Fvarios.Refresh
Else
    If frmARTI.DS.ListIndex = -1 Then
       MsgBox "Determinar la Moneda del Articulo ..", 48, Pub_Titulo
       frmARTI.DS.SetFocus
       SendKeys "%{UP}"
       Exit Sub
    End If
    If Trim(frmARTI.CmbCalidad.Text) = "" Then
       MsgBox "Definir Calidad en,  Tablas del Sistema ", 48, Pub_Titulo
       Exit Sub
    End If
    If Trim(cmbUsu.Text) = "" Then
       MsgBox "Definir Usuario Responsable. (FICHA OTRAS OPCIONES)", 48, Pub_Titulo
       Exit Sub
    End If
    If Not CONSIS_ARTI Then
       Exit Sub
    End If
    If Not CONSIS_UNIDAD Then
       Screen.MousePointer = 0
       MsgBox "Verificar Datos de unidad no valido..", 48, Pub_Titulo
       grid_unid.SetFocus
       Exit Sub
    End If
    If LK_FLAG_ORIGINAL <> "A" Then
       If Trim(txt_alterno.Text) = "" Then
          MsgBox "Codigo alterno no valido ..!!! Verificar ", 48, Pub_Titulo
          Azul frmARTI.txt_alterno, frmARTI.txt_alterno
          GoTo fin
       End If
       SQ_OPER = 3
       pu_alterno = txt_alterno.Text
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       If Not art_llave_alt.EOF Then
          MsgBox "Codigo Alterno  EXISTE ...!!! Verificar " & Chr(13) & Trim(art_llave_alt!ART_alterno) & " : " & Trim(art_llave_alt!art_nombre), 48, Pub_Titulo
          Azul frmARTI.txt_alterno, frmARTI.txt_alterno
          Debug.Print frmARTI.txt_alterno
          GoTo fin
       End If
    End If
    If pasa = 1 Then
    '   If EXISTE_ART(txtnombre.Text, Trim(Txt_key.Text)) Then
    '      MENSAJE_ARTI "Existen algunos Articulos con estos NOMBRES .."
    '      frmARTI.ListExiste.SetFocus
    '      Exit Sub
    '   End If
    End If
    pasa = 0
     Screen.MousePointer = 11
'     On Error GoTo ESCAPA
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     CN.Execute "Begin Transaction", rdExecDirect
     pub_cadena = "SELECT * FROM CONTROLL"
     Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
     frmARTI.txt_key = GENERA_CODI
     PUB_KEY = Val(frmARTI.txt_key)
     If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
'          If pu_codcia = "05" And LK_EMP = "CAM" Then GoTo SALE
          If Trim(pu_codcia) = "" Then Exit For
             GRABAR_ARTI
            'GoSub IR_POR_CIA
sale:
          xcuenta = xcuenta + 2
        Next fila
     Else
        GRABAR_ARTI
     End If
     If cheasignarc.Value = 1 Then add_convenio Val(frmARTI.txt_key)
     con_llave.Close
     CN.Execute "Commit Transaction", rdExecDirect
     On Error GoTo 0
     cmdAgregar.Caption = "&Adicionar"
     cmdCancelar.Enabled = True
     cmdEliminar.Enabled = True
     cmdModificar.Enabled = True
     LIMPIA_ARTI
     BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
     BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
     BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
     BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
     BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
     If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       DESBLOQUEA_TEXT txt_alterno
       BLOQUEA_TEXT txt_key
       txt_alterno.SetFocus
     Else
       DESBLOQUEA_TEXT txt_key
       BLOQUEA_TEXT txt_alterno
       txt_key.SetFocus
     End If
     Screen.MousePointer = 0
     MANOS(0).Enabled = True
     MANOS(1).Enabled = True
     MENSAJE_ARTI "Articulo   AGREGADO ... "
End If
Exit Sub
ESCAPA:
'    If con_llave Is Nothing Then
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
 '   End If
    If Err.Number = 40002 Then
        MsgBox "Hay Error en la LLave ..Intente Nuevamente. "
    ElseIf Err.Number <> 0 Then
        MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    End If
    cmdAgregar.Caption = "&Adicionar"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT txt_key
     txt_key.SetFocus
    End If
    Screen.MousePointer = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
   Exit Sub
fin:
End Sub

Private Sub cmdagregar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
'    frmARTI.txt_key.SetFocus
End If

End Sub



Private Sub cmdaplica_Click()
fracias.Visible = False
End Sub

Private Sub cmdapor_Click()
Dim llave_LISTARTI As rdoResultset
Dim PSLOC_LISTAARTI As rdoQuery
  
Dim WS_CODART As Currency
Dim REP_FECHA1
Dim REP_FECHA2

CADENITA = ""
For fila = 0 To listlin.ListCount - 1
  If listlin.Selected(fila) = True Then
    CADENITA = CADENITA + Trim(Right(listlin.List(fila), 8)) + " ,"
  End If
Next fila
If CADENITA <> "" Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1)
End If



If CADENITA <> "" Then
  pub_cadena = "SELECT ART_KEY , ART_COSTO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY > 0 AND ART_FAMILIA  IN( " & CADENITA & ")"
Else
  pub_cadena = "SELECT ART_KEY , ART_COSTO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY > 0 "
End If
Set PSLOC_LISTAARTI = CN.CreateQuery("", pub_cadena)
PSLOC_LISTAARTI(0) = ""
Set llave_LISTARTI = PSLOC_LISTAARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
PSLOC_LISTAARTI(0) = LK_CODCIA
llave_LISTARTI.Requery
pb.Visible = True
pbMin = 0
pbValue = 0
If Not llave_LISTARTI.EOF Then pbMax = llave_LISTARTI.RowCount
Do Until llave_LISTARTI.EOF
  pbValue = pbValue + 1
  llave_LISTARTI.Edit
  llave_LISTARTI!ART_COSTO = Val(txtapor.Text)
  llave_LISTARTI.Update
  llave_LISTARTI.MoveNext
Loop
pb.Visible = False
pbMin = 0
cheapor.Value = 0
MsgBox "Actualizaci�n Terminada.", 48, Pub_Titulo

End Sub

Private Sub cmdcan_Click()
fraact6.Visible = False
End Sub

Private Sub cmdcancelar_Click()
PROCESO_CANCELAR
End Sub

Private Sub cmdCancelar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
End Sub

Private Sub cmdcancelfiltro_Click()
fralstfiltro.Visible = False
End Sub

Private Sub cmdcanlin_Click()
fraactlinea.Visible = False

End Sub

Private Sub cmdCerrar_Click()
cmdcancelar_Click
frmARTI.Hide
End Sub

Private Sub cmdCerrar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
End Sub

Private Sub cmdconfirma_Click()
  If Op(1).Value Then
     pasa = 0
     frmARTI.F14.Visible = False
     If Left(cmdModificar.Caption, 2) = "&G" Then
       CmdModificar_Click
     Else
       cmdAgregar_Click
     End If
     Exit Sub
  End If
  MsgBox "Seleccione una de las dos Opciones ..", 48, Pub_Titulo
End Sub

Private Sub cmddolares_Click()
Exit Sub
Dim WSPOR As Currency
'If LK_MONEDA = "S" Or LK_MONEDA = "D" Then
 ' grid_unid.SetFocus
  'Exit Sub
'End If
'Exit Sub
If LK_EMP = "3AA" Then
  If cmddolares.Tag = "D" Then
    cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
    cmddolares.Tag = "S"
  Else
    cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
    cmddolares.Tag = "D"
  End If
Else
  If cmddolares.Tag = "D" Then
    cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
    cmddolares.Tag = "S"
    DS.ListIndex = 0
  Else
    cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
    cmddolares.Tag = "D"
    DS.ListIndex = 1
  End If
End If
For fila = 1 To grid_unid.Rows - 1
    If cmddolares.Tag = "D" Then
     grid_unid.TextMatrix(fila, 5) = grid_unid.TextMatrix(fila, 18)
     grid_unid.TextMatrix(fila, 7) = grid_unid.TextMatrix(fila, 19)
     grid_unid.TextMatrix(fila, 9) = grid_unid.TextMatrix(fila, 20)
     grid_unid.TextMatrix(fila, 11) = grid_unid.TextMatrix(fila, 21)
     grid_unid.TextMatrix(fila, 13) = grid_unid.TextMatrix(fila, 22)
     grid_unid.TextMatrix(fila, 15) = grid_unid.TextMatrix(fila, 31)
    Else
     grid_unid.TextMatrix(fila, 5) = grid_unid.TextMatrix(fila, 23)
     grid_unid.TextMatrix(fila, 7) = grid_unid.TextMatrix(fila, 24)
     grid_unid.TextMatrix(fila, 9) = grid_unid.TextMatrix(fila, 25)
     grid_unid.TextMatrix(fila, 11) = grid_unid.TextMatrix(fila, 26)
     grid_unid.TextMatrix(fila, 13) = grid_unid.TextMatrix(fila, 27)
     grid_unid.TextMatrix(fila, 15) = grid_unid.TextMatrix(fila, 32)
    End If
    If cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(fila, 3) = redondea(Val(grid_unid.TextMatrix(fila, 29)) / LK_TIPO_CAMBIO)
     Else
         grid_unid.TextMatrix(fila, 3) = redondea(Val(grid_unid.TextMatrix(fila, 29)))
     End If
    
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 5)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 4) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 7)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 6) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 9)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 8) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 11)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 10) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 13)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 12) = Format(WSPOR, "0.00")
     If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
       WSPOR = (Val(grid_unid.TextMatrix(fila, 15)) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
     End If
     grid_unid.TextMatrix(fila, 14) = Format(WSPOR, "0.00")
    
 Next fila
'If Trim(DS.Text) = "S" Then
'   If cmddolares.Tag = "D" Then
'     grid_unid.Enabled = False
'   Else
     grid_unid.Enabled = True
     If grid_unid.Enabled Then grid_unid.SetFocus
'   End If
'Else
'   If cmddolares.Tag = "D" Then
'    grid_unid.Enabled = True
'    grid_unid.SetFocus
'   Else
'    grid_unid.Enabled = False
'   End If
'End If


End Sub

Private Sub cmdEliminar_Click()
Dim ws_codcia As String
Dim WS_CODART As Currency
Dim flag_puntos As String * 1
On Error GoTo ESCAPA
If Len(txt_key) = 0 Or Len(txtnombre.Text) = 0 Then
    If Not Trim(txt_key) = "1" Then
       Screen.MousePointer = 0
       MENSAJE_ARTI "NO a seleccionado ningun Articulo... !"
'       txt_key.SetFocus
       Exit Sub
    End If
End If
  Dim PS_REP01 As rdoQuery
  Dim llave_rep01 As rdoResultset
  Dim OpenForms
  
  WS_CODART = artloc_llave!art_key
  If Trim(par_llave!par_art_cias) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
            ws_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
            If Trim(ws_codcia) = "" Then Exit For
            SQ_OPER = 1
            PUB_CODART = WS_CODART
            pu_codcia = ws_codcia
            LEER_ARM_LLAVE
            If Not arm_llave.EOF Then
                If arm_llave!arm_stock = 0 And arm_llave!ARM_INGRESOS = 0 And arm_llave!ARM_SALIDAS = 0 Then
                Else
                    LblMensaje.Visible = False
                    Screen.MousePointer = 0
                    MsgBox "NO se Puede Eliminar ...  ARTICULO CON HISTORIA " & Chr(13) & "Compa�ia : " & ws_codcia, 48, Pub_Titulo
                    Exit Sub
                End If
            End If
            xcuenta = xcuenta + 2
        Next fila
  End If
  
  Screen.MousePointer = 11
  LblMensaje.Visible = True
  LblMensaje.Caption = "Verificando Data.  un Momento..."
  WS_CODART = artloc_llave!art_key
  SQ_OPER = 1
  PUB_CODART = artloc_llave!art_key
  pu_codcia = LK_CODCIA
  LEER_ARM_LLAVE
  If Not arm_llave.EOF Then
      If arm_llave!arm_stock = 0 And arm_llave!arm_stock = 0 And arm_llave!arm_stock = 0 Then
      Else
          LblMensaje.Visible = False
          Screen.MousePointer = 0
          MsgBox "NO se Puede Eliminar ...  ARTICULO CON HISTORIA ", 48, Pub_Titulo
          Exit Sub
      End If
  End If
  pub_cadena = "SELECT FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_ESTADO <> 'E' "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.rdoParameters(0) = " "
  PS_REP01.rdoParameters(1) = 0
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = PUB_CODART
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  ARTICULO  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  
  If LK_EMP_PTO = "A" Then
    If LK_CODCIA <> "00" Then
      Screen.MousePointer = 0
      LblMensaje.Visible = False
      MsgBox "No Procede la Eliminaci�n.  Punto de Venta no permitido!!(solo en la Cia. central)", 48, Pub_Titulo
      Exit Sub
    End If
  End If
  LblMensaje.Visible = False
  pub_mensaje = " �Desea Eliminar el Articulo... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
      LblMensaje.Visible = False
      Screen.MousePointer = 0
      Exit Sub
  End If
  LblMensaje.Visible = True
  LblMensaje.Caption = "Eliminando.  un Momento..."
  flag_puntos = ""
  CN.Execute "Begin Transaction", rdExecDirect
  pub_cadena = "SELECT * FROM CONTROLL"
  Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
 flag_puntos = ""
  If Trim(par_llave!par_art_cias) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
            ws_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
            If ws_codcia = "00" Then
               flag_puntos = "A"
            Else
               flag_puntos = ""
            End If
            If Trim(ws_codcia) = "" Then Exit For
            GoSub BORRA_ART_CIA
            xcuenta = xcuenta + 2
        Next fila
  Else
    ws_codcia = LK_CODCIA
 '   flag_puntos = "A"
    GoSub BORRA_ART_CIA
  End If

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close

cmdcancelar_Click
MENSAJE_ARTI "Articulo   ELIMINADO ... "
Screen.MousePointer = 0
Exit Sub

BORRA_ART_CIA:
If flag_puntos = "A" Then
 SQ_OPER = 1
 PUB_KEY = WS_CODART
 pu_codcia = ws_codcia
 LEER_ART_LLAVE
 art_LLAVE.Delete
End If
 SQ_OPER = 1
 PUB_KEY = WS_CODART
 pu_codcia = ws_codcia
 LEER_ART_LLAVE
 art_LLAVE.Delete
 SQ_OPER = 1
 PUB_CODART = WS_CODART
 pu_codcia = ws_codcia
 LEER_ARM_LLAVE
 arm_llave.Delete
 pub_cadena = "DELETE PRECIOS WHERE PRE_CODCIA = '" & ws_codcia & "' AND PRE_CODART = " & WS_CODART
 CN.Execute pub_cadena, rdExecDirect
Return

Exit Sub

ESCAPA:
    con_llave.Close
    CN.Execute "Rollback Transaction", rdExecDirect
    Screen.MousePointer = 0
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    LblMensaje.Visible = False
    DoEvents
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    cmdAgregar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT art_linea, art_numero, art_marca, art_plancha, checambio, txtlitro, chebof, chesub
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, exigv, txtcospro, cmddolares, txtfechault
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, cmddolares, txtpeso
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT txt_key
     txt_key.SetFocus
    End If
    Screen.MousePointer = 0

End Sub

Private Sub cmdEliminar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
End Sub

Private Sub cmdEliminar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
'    frmARTI.txt_key.SetFocus
End If

End Sub

Private Sub CmdEscapa_Click()
  frmARTI.txtnombre.SetFocus
  frmARTI.F14.Visible = False
End Sub

Private Sub cmdgen_Click()
Dim llave_LISTARTI As rdoResultset
Dim PSLOC_LISTAARTI As rdoQuery
  
  Dim WS_CODART As Currency
Dim REP_FECHA1
Dim REP_FECHA2
If Val(txtd.Text) = 0 Then
  MsgBox "Dias para promediar, no procede ", 48, Pub_Titulo
  Exit Sub
End If
If Right(frmARTI.txtCampo1.Text, 2) = "__" Then
  REP_FECHA1 = Left(frmARTI.txtCampo1.Text, 8)
Else
  REP_FECHA1 = Trim(frmARTI.txtCampo1.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 frmARTI.txtCampo1, frmARTI.txtCampo1
    GoTo fin
End If
If Right(frmARTI.txtCampo2.Text, 2) = "__" Then
  REP_FECHA2 = Left(frmARTI.txtCampo2.Text, 8)
Else
  REP_FECHA2 = Trim(frmARTI.txtCampo2.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 Azul2 frmARTI.txtCampo2, frmARTI.txtCampo2
 GoTo fin
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fechas Invalidadas ..", 48, Pub_Titulo
 Azul2 frmARTI.txtCampo1, frmARTI.txtCampo1
 GoTo fin
End If

LOC_FECHA_INI = REP_FECHA1
LOC_FECHA_FIN = REP_FECHA2

CADENITA = ""
For fila = 0 To listlin.ListCount - 1
  If listlin.Selected(fila) = True Then
    CADENITA = CADENITA + Trim(Right(listlin.List(fila), 8)) + " ,"
  End If
Next fila
If CADENITA <> "" Then
  CADENITA = Left(CADENITA, Len(CADENITA) - 1)
End If



If CADENITA <> "" Then
  pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY > 0 AND ART_FAMILIA  IN( " & CADENITA & ")"
Else
  pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY > 0 "
End If
Set PSLOC_LISTAARTI = CN.CreateQuery("", pub_cadena)
PSLOC_LISTAARTI(0) = ""
Set llave_LISTARTI = PSLOC_LISTAARTI.OpenResultset(rdOpenKeyset, rdConcurValues)
PSLOC_LISTAARTI(0) = LK_CODCIA
llave_LISTARTI.Requery
pb.Visible = True
pbMin = 0
pbValue = 0
If Not llave_LISTARTI.EOF Then pbMax = llave_LISTARTI.RowCount

' listlin
Do Until llave_LISTARTI.EOF
   pbValue = pbValue + 1
   WS_CODART = llave_LISTARTI!art_key
'    If llave_LISTARTI!art_key = 106912 Then Stop
'   Print llave_LISTARTI!art_nombre
   If chenueva.Value = 1 Then
     WS_CANTIDAD = Pedido_Estimado(llave_LISTARTI!art_key, txtd.Text, CDate(REP_FECHA1), CDate(REP_FECHA2))
   Else
     WS_CANTIDAD = PROCESA_CANTIDAD(WS_CODART, Val(txtd.Text))
     WS_CANTIDAD = Val(Format(WS_CANTIDAD * Val(txtd.Text), "#"))
   End If
   
   llave_LISTARTI.Edit
   If ops(0).Value Then
     llave_LISTARTI!ART_STOCK_MIN = WS_CANTIDAD
   Else
     llave_LISTARTI!ART_STOCK_MAX = WS_CANTIDAD
   End If
  llave_LISTARTI.Update
  DoEvents
SIG:
 llave_LISTARTI.MoveNext
Loop
pb.Visible = False
MsgBox "Proceso de Calculo de Stock a sido Terminado.", 48, Pub_Titulo
LLENADO_FAM
fragen.Visible = False

Exit Sub
fin:

End Sub


Private Sub cmdgrabafiltro_Click()
If lstfiltro.ListCount <= 0 Then Exit Sub
PBFMin = 0
PBFmax = lstfiltro.ListCount - 1
PBFValue = 0
PBFVisible = True
DoEvents
For fila = 0 To lstfiltro.ListCount - 1
        PBFValue = fila
        DoEvents
        PUB_TIPREG = 122
        PUB_CODCIA = LK_CODCIA
        SQ_OPER = 1
        PUB_NUMTAB = Right(lstfiltro.List(fila), 8)
        LEER_TAB_LLAVE
        If Not tab_llave.EOF Then
            tab_llave.Edit
            If lstfiltro.Tag = 116 Then
              If lstfiltro.Selected(fila) Then
              tab_llave!tab_codclie = 1
              Else
              tab_llave!tab_codclie = 0
              End If
            Else
              If lstfiltro.Selected(fila) Then
                tab_llave!tab_codart = 1
              Else
               tab_llave!tab_codart = 0
              End If
            End If
            tab_llave.Update
        End If
Next fila
PBFVisible = False
fralstfiltro.Visible = False
End Sub

Private Sub Cmdmax_Click()
MsgBox "Proceso esta en Servidor"
Exit Sub
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

 txtCampo1.Mask = "##/##/####"
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 fragen.Visible = True
 
End Sub


Private Sub CmdModificar_Click()
'On Error GoTo ESCAPA
If Trim(txt_key.Text) = "1" Then
     MENSAJE_ARTI "No Procede. .."
     Exit Sub
End If
If Len(txt_key) = 0 Or Trim(txtnombre.Text) = "" Then
   MENSAJE_ARTI "NO a seleccionado ningun Articulo... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    ws_flag_equiv = ""
    LK_ACCESO_REPORT = ""
    Load frmclave2
    Screen.MousePointer = 0
    frmclave2.Show 1
    If LK_ACCESO_REPORT <> "A" Then
       ws_flag_equiv = ""
       MsgBox "No Tiene el Acceso a Modificar las Equivalencias, ni Descripciones!!", 48, Pub_Titulo
       wflag_bloq_descrip = "A"
    Else
       ws_flag_equiv = "A"
       wflag_bloq_descrip = ""
    End If
    If wflag_bloq_descrip = "A" Then
        Fvarios.Enabled = True
        Frame3.Enabled = False
    Else
        Fvarios.Enabled = True
        Frame3.Enabled = True
    End If
    
    cmdModificar.Caption = "&Grabar"
    cmdEliminar.Enabled = False
    cmdAgregar.Enabled = False
    cmdCancelar.Enabled = True
    If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
       DESBLOQUEA_TEXT txt_alterno
    Else
       BLOQUEA_TEXT txt_alterno
    End If
    BLOQUEA_TEXT txt_key
    DESBLOQUEA_TEXT txtnombre, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    DESBLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    DESBLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    DESBLOQUEA_TEXT txtmargen, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    If LK_CODUSU = "ADMIN" Then DESBLOQUEA_TEXT cmbUsu
    If loc_cheoferta = "A" Then DESBLOQUEA_TEXT cheoferta
    If LK_EMP_PTO = "A" Then
      If LK_CODCIA <> "00" Then
        BLOQUEA_TEXT decimales, art_grupo, art_familia, art_subfam, art_codpro, txtcodigo2
        BLOQUEA_TEXT art_situacion, DS, cheservi(0), cheservi(1), cheservi(2), txtMin, txtMax
      End If
    End If
    DESBLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    'MODIFICADO PARA PODER MODIFICAR LA MONEDA
    MANOS(0).Enabled = False
    MANOS(1).Enabled = False
    '07/12/2001
    If LK_MONEDA = "D" Then
        frmARTI.DS.Enabled = Not False
    ElseIf LK_MONEDA = "S" Then
        frmARTI.DS.Enabled = Not False
    End If
    If loc_flag_bloq = "A" Then
      Fcomun.Enabled = False
    Else
      Fcomun.Enabled = True
    End If
    If wflag_bloq_desactiva = "A" Then
      art_situacion.Enabled = True
    Else
      art_situacion.Enabled = False
    End If
    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "OPER11" Then
    Else
      chebof.Enabled = False
    End If


    If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then GoTo ava_acc
    txtCantMaxADD.Enabled = False
    txtCantMax.Enabled = False
    If Flag_Add_cant = "A" Then txtCantMaxADD.Enabled = True
    If loc_p = "A" Then
      txtmargen.Enabled = True
    Else
      txtmargen.Enabled = False
    End If
    

ava_acc:
    pasa = 1
    frmARTI.txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If Not IsDate(frmARTI.txtfechault.Text) Then
       MsgBox "Fecha de ultima compra no es correcta.", 48, Pub_Titulo
      Exit Sub
    End If
    Screen.MousePointer = 11
    If Val(loc_orig_situ) = 1 And art_situacion.Value = 0 Then
       If Trim(cmbUsu.Text) = "" Then
            cmbUsu.Enabled = True
            Screen.MousePointer = 0
            MsgBox "Definir Usuario Responsable. (FICHA OTRAS OPCIONES)", 48, Pub_Titulo
            cmbUsu.SetFocus
            Exit Sub
       End If
    End If
    
    If Trim(DS.Text) = "D" Or Trim(DS.Text) = "S" Then
    Else
      Screen.MousePointer = 0
      MsgBox "Solo Acepta  D = Dolares,  o  S = Soles ... ", 48, Pub_Titulo
      DS.SetFocus
      Exit Sub
    End If
    If Val(decimales.Text) < 0 Then
      Screen.MousePointer = 0
      MsgBox "Escoger con cuantos decimales de precision...", 48, Pub_Titulo
      decimales.SetFocus
      Exit Sub
    End If
    If Not CONSIS_ARTI Then
       Screen.MousePointer = 0
       Exit Sub
    End If
    If Not CONSIS_UNIDAD Then
       Screen.MousePointer = 0
       grid_unid.SetFocus
       Exit Sub
    End If
    
    If art_situacion.Value = 1 Then
     If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
          If Trim(pu_codcia) = "" Then Exit For
            SQ_OPER = 1
            PUB_CODART = Val(txt_key.Text)
            pu_codcia = pu_codcia
            LEER_ARM_LLAVE
            If Not arm_llave.EOF Then
              If Val(arm_llave!arm_stock) <> 0 Then
                 Screen.MousePointer = 0
                 MsgBox "No Procede al desactivar, Producto Tiene Stock." & Chr(13) & "Compa�ia: " & pu_codcia, 48, Pub_Titulo
                 Exit Sub
               End If
            End If
          xcuenta = xcuenta + 2
        Next fila
     Else
        SQ_OPER = 1
        PUB_CODART = Val(txt_key.Text)
        pu_codcia = LK_CODCIA
        LEER_ARM_LLAVE
        If Not arm_llave.EOF Then
          If Val(arm_llave!arm_stock) <> 0 Then
            Screen.MousePointer = 0
            MsgBox "No Procede al desactivar, Producto Tiene Stock.", 48, Pub_Titulo
            Exit Sub
          End If
        End If
     End If
    End If
    If pasa = 1 Then
      If EXISTE_ART(txtnombre.Text, Trim(txt_key.Text)) Then
          Screen.MousePointer = 0
          MENSAJE_ARTI "Existen algunos Articulos con estos NOMBRES .."
          frmARTI.ListExiste.SetFocus
          Exit Sub
      End If
    End If
    pasa = 0
    CN.Execute "Begin Transaction", rdExecDirect
    If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
        xcuenta = 1
        For fila = 1 To 30
          pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
          If Trim(pu_codcia) = "" Then Exit For
             PUB_KEY = Val(frmARTI.txt_key.Text)
             PUB_CODCIA = pu_codcia
             LOC_OPER = 1
             LEER_LOC
             GRABAR_ARTI
            'GoSub IR_POR_CIA
          xcuenta = xcuenta + 2
        Next fila
    Else
      GRABAR_ARTI
    End If
    If cheasignarc.Value = 1 Then add_convenio Val(frmARTI.txt_key)
    CN.Execute "Commit Transaction", rdExecDirect
    ' Actuzliar Varios Producto amarre por Sub Familia
    If Val(Right(art_subfam.Text, 8)) <> 0 Then
        LK_ACCESO_REPORT = ""
        Load frmclave2
        Screen.MousePointer = 0
        frmclave2.Show 1
        If LK_ACCESO_REPORT <> "A" Then
        Else
          Load frmPreciosArti
          frmPreciosArti.Show 1
        End If
    End If
    cmdModificar.Caption = "&Modificaci�n"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdAgregar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    MENSAJE_ARTI "Articulo,  MODIFICADO... "
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      DESBLOQUEA_TEXT txt_alterno
      txt_alterno.SetFocus
    Else
      DESBLOQUEA_TEXT txt_key
      txt_key.SetFocus
    End If
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
    Screen.MousePointer = 0
End If
Exit Sub
ESCAPA:
    If Err.Number = 40002 Then
        MsgBox "Hay Error en la LLave ..Intente Nuevamente. "
    Else
        MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    End If
    CN.Execute "Rollback Transaction", rdExecDirect
    cmdModificar.Caption = "&Modificaci�n"
    cmdCancelar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    cmdAgregar.Enabled = True
    LIMPIA_ARTI
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     DESBLOQUEA_TEXT txt_alterno
     txt_alterno.SetFocus
    Else
     DESBLOQUEA_TEXT txt_key
     txt_key.SetFocus
    End If
    Screen.MousePointer = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True

End Sub

Private Sub cmdModificar_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
End Sub

Private Sub cmdModificar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 '   frmARTI.txt_key.SetFocus
End If
End Sub


Private Sub cmdofertas_Click()
Load frmstockgen

frmstockgen.Show 1
End Sub

Private Sub cmdp_Click()
Dim wnombre As String
Dim WORIGINAL As Currency
Dim WORIGINAL1 As Currency
Dim WORIGINAL2 As Currency
Dim walterno  As String
Dim WCALIDAD  As Integer
Dim WFAMILIA  As Integer
Dim WSUBFAMI  As Integer
Dim WNUMERO As Integer
Dim wgrupo As Integer
Dim WLINEA As Integer
Dim wMARCA As Integer

If Trim(cmbcal.Text) = "" Then
 MsgBox " Seleccione su Calidad.", 48, Pub_Titulo
 cmbcal.SetFocus
 Exit Sub
End If
If Val(Right(art_familia.Text, 8)) <> 4 Then
  MsgBox "solo se Activa para los 4 - Producto Terminado", 48, Pub_Titulo
  Exit Sub
End If

For fila = 2 To 3
    WORIGINAL = artloc_llave!art_key
    If fila = 2 Then
      WORIGINAL1 = GENERA_CODI()
    Else
      WORIGINAL2 = GENERA_CODI()
    End If
'    walterno = WORIGINAL
    WCALIDAD = 1
    If fila = 2 Then
      WFAMILIA = 3
    Else
      WFAMILIA = 2
    End If
    WSUBFAMI = 1
    wnombre = artloc_llave!art_nombre
    For I = 0 To art_familia.ListCount - 1
      art_familia.ListIndex = I
      If WFAMILIA = Val(Right(art_familia.Text, 6)) Then
         wnombre = Trim(Left(art_familia.Text, 6)) & " " & Trim(Left(art_grupo.Text, 15)) & " " & Trim(Left(art_numero.Text, 15)) & "-" & Trim(Left(art_marca.Text, 15))  ' & " " & Left(art_linea.Text, 3) ' Trim(Left(art_linea.Text, 5))
         Exit For
      End If
    Next I
    WNUMERO = artloc_llave!art_numero
    wgrupo = artloc_llave!art_subgru
    WLINEA = artloc_llave!art_linea
    wMARCA = artloc_llave!art_marca
    GoSub grabar
Next fila
PSART_RELA.rdoParameters(0) = LK_CODCIA
PSART_RELA.rdoParameters(1) = WORIGINAL1
art_rela.Requery
If art_rela.EOF Then
  MsgBox "NO ESTA BIEN LA RELACION."
  Exit Sub
End If
art_rela.Edit
art_rela!ART_CODART2 = WORIGINAL2
art_rela.Update

artloc_llave.Edit
artloc_llave!ART_CODART2 = WORIGINAL1
artloc_llave.Update
cmdcancelar_Click
MsgBox "Se Activo.", 48, Pub_Titulo
Exit Sub


grabar:

art_LLAVE.AddNew
If fila = 2 Then
    art_LLAVE!art_key = WORIGINAL1
    art_LLAVE!ART_alterno = Trim(Str(WORIGINAL1))
    art_LLAVE!ART_CODART2 = 0
Else
    art_LLAVE!art_key = WORIGINAL2
    art_LLAVE!ART_alterno = Trim(Str(WORIGINAL2))
    art_LLAVE!ART_CODART2 = 0
End If

art_LLAVE!art_familia = WFAMILIA
art_LLAVE!art_subfam = WSUBFAMI
art_LLAVE!ART_CALIDAD = WCALIDAD
art_LLAVE!art_subgru = wgrupo
art_LLAVE!art_linea = WLINEA
art_LLAVE!art_numero = WNUMERO
art_LLAVE!art_marca = wMARCA

art_LLAVE!art_nombre = wnombre
    
art_LLAVE!ART_POR_IGV = artloc_llave!ART_POR_IGV
'art_LLAVE!ART_ORDEN = artloc_llave!ART_ORDEN
art_LLAVE!art_tipo = loc_tipo
art_LLAVE!art_codclie = artloc_llave!art_codclie
art_LLAVE!art_codcia = PUB_CODCIA
art_LLAVE!ART_DECIMALES = artloc_llave!ART_DECIMALES
art_LLAVE!ART_MONEDA = artloc_llave!ART_MONEDA
art_LLAVE!art_situacion = artloc_llave!art_situacion
art_LLAVE!ART_STOCK_MIN = artloc_llave!ART_STOCK_MIN
art_LLAVE!ART_STOCK_MAX = artloc_llave!ART_STOCK_MAX
art_LLAVE!ART_POR_IGV = artloc_llave!ART_POR_IGV
art_LLAVE!ART_EX_IGV = ""
art_LLAVE!ART_COSPRO = 0
art_LLAVE!ART_EX_IGV = "A"
art_LLAVE!art_flag_stock = ""
art_LLAVE!art_flag_stock = "M"
art_LLAVE!ART_POR1 = 0
art_LLAVE!ART_POR2 = 0
art_LLAVE!ART_POR3 = 0
art_LLAVE!ART_POR4 = 0
art_LLAVE!ART_POR5 = 0
art_LLAVE.Update

'pu_codcia = PUB_CODCIA
'PUB_CODART = WORIGINAL
'SQ_OPER = 1
'LEER_ARM_LLAVE
'If arm_llave.EOF Then
    arm_llave.AddNew
If fila = 2 Then
    arm_llave!ARM_CODART = WORIGINAL1
Else
    arm_llave!ARM_CODART = WORIGINAL2
End If



    arm_llave!ARM_CODCIA = LK_CODCIA
    arm_llave!arm_stock = 0
    arm_llave!ARM_INGRESOS = 0
    arm_llave!ARM_SALIDAS = 0
    arm_llave!ARM_COSPRO = 0
    arm_llave!arm_stock2 = 0
    arm_llave!ARM_COSTO_ULT = 0
    arm_llave!ARM_FECHA_ULT = LK_FECHA_DIA
    arm_llave.Update
'End If

pre_mayor.AddNew
pre_mayor!PRE_CODCIA = LK_CODCIA
If fila = 2 Then
pre_mayor!PRE_CODART = WORIGINAL1
Else
pre_mayor!PRE_CODART = WORIGINAL2
End If

pre_mayor!PRE_SECUENCIA = 0
pre_mayor!PRE_UNIDAD = Left(grid_unid.TextMatrix(1, 0), 17)
pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(1, 1))
pre_mayor!PRE_PRE11 = Val(grid_unid.TextMatrix(1, 18))
pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(1, 19))
pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(1, 20))
pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(1, 21))
pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(1, 22))
pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(1, 23))
pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(1, 24))
pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(1, 25))
pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(1, 26))
pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(1, 27))
pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(1, 26))
pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(1, 28))
pre_mayor!pre_FLAG_UNIDAD = grid_unid.TextMatrix(1, 14)
pre_mayor.Update
Return
End Sub

Private Sub cmdqc_Click()
pub_cadena = "delete from tabconvenio where tco_codart = " & Val(txt_key.Text)
CN.Execute pub_cadena
MsgBox "Producto Quitado del Convenio.", 48, Pub_Titulo
cheasignarc.Value = 1
cheasignarc.Value = 0

End Sub

Private Sub cmdquitar_Click()
On Error GoTo ESCAPA
SQ_OPER = 1
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
PUB_KEY = PUB_CODART
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_LLAVE.EOF Then
  MsgBox "Arti no Existe ,,", 48, Pub_Titulo
  Exit Sub
End If

SQ_OPER = 1
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
If arm_llave.EOF Then
  MsgBox "Arti no Existe ,,", 48, Pub_Titulo
  Exit Sub
End If
If arm_llave!ARM_INGRESOS = 0 And arm_llave!ARM_SALIDAS = 0 Then
Else
  MsgBox "NO Procede - Articulo Defectuoso Tiene Historia ", 48, Pub_Titulo
  Exit Sub
End If
If art_LLAVE!ART_CALIDAD <> 2 Then
  MsgBox "Articulo no Es Defectuoso Verificar !!!!! ", 48, Pub_Titulo
  Exit Sub
End If

pub_cadena = "SELECT FAR_CODART FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_ESTADO <> 'E' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01.rdoParameters(0) = " "
PS_REP01.rdoParameters(1) = 0
PS_REP01.MaxRows = 1
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = PUB_CODART
llave_rep01.Requery
If Not llave_rep01.EOF Then
   Screen.MousePointer = 0
   MsgBox "NO se Puede Eliminar ...  ARTICULO  TIENE H I S T O R I A.. ", 48, Pub_Titulo
   Exit Sub
End If
  

pub_mensaje = " Desea Eliminar la Relacion y el Articulo. :" & Trim(gridrel.TextMatrix(1, 2)) & " ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  Exit Sub
End If


CN.Execute "Begin Transaction", rdExecDirect
PUB_CODART = Val(gridrel.TextMatrix(1, 0))
CN.Execute "DELETE ARTI WHERE ART_KEY = " & PUB_CODART & " AND ART_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "DELETE PRECIOS WHERE PRE_CODART = " & PUB_CODART & " AND PRE_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "DELETE ARTICULO WHERE ARM_CODART = " & PUB_CODART & " AND ARM_CODCIA = '" & LK_CODCIA & "'", rdExecDirect
CN.Execute "Commit Transaction", rdExecDirect
On Error GoTo 0
gridrel.TextMatrix(1, 0) = ""
gridrel.TextMatrix(1, 1) = ""
gridrel.TextMatrix(1, 2) = ""
gridrel.TextMatrix(1, 3) = ""
cmdquitar.Visible = False
cmdAdd.Visible = True
cmbcal.Visible = True
lblcal.Visible = True
Exit Sub
ESCAPA:
CN.Execute "Rollback Transaction", rdExecDirect
MsgBox "Intente Nuevamente.", 48, titulo

End Sub



Private Sub cmdvend_Click()
Dim var_codven As Integer
Dim var_codart  As Currency

var_codven = Val(Left(cmdvend.Text, 3))
var_codart = Val(txt_key.Text)
If var_codart = 0 Then Exit Sub
If var_codven = 0 Then Exit Sub
PSCOMI_VEN.rdoParameters(0) = LK_CODCIA
PSCOMI_VEN.rdoParameters(1) = var_codart
PSCOMI_VEN.rdoParameters(2) = var_codven
comivem_llave.Requery
If comivem_llave.EOF Then
txtpor1.Text = ""
txtpor2.Text = ""
txtpor3.Text = ""
txtpor4.Text = ""
txtpor5.Text = ""
txtpor6.Text = ""
txtpor7.Text = ""
txtpor8.Text = ""
txtpor9.Text = ""

txtpor11.Text = ""
txtpor22.Text = ""
txtpor33.Text = ""
txtpor44.Text = ""
txtpor55.Text = ""
txtpor66.Text = ""
txtpor77.Text = ""
txtpor88.Text = ""
txtpor99.Text = ""

cmdactvend.Enabled = False
  MsgBox "Vendedor no tiene su tabla de Comision por Producto", 48, Pub_Titulo
  If txtpor1.Enabled Then txtpor1.SetFocus
  Exit Sub
End If
txtpor1.Text = Format(comivem_llave!COA_COMV1, "0.00")
txtpor2.Text = Format(comivem_llave!COA_COMV2, "0.00")
txtpor3.Text = Format(comivem_llave!COA_COMV3, "0.00")
txtpor4.Text = Format(comivem_llave!COA_COMV4, "0.00")
txtpor5.Text = Format(comivem_llave!COA_COMV5, "0.00")
txtpor6.Text = Format(comivem_llave!COA_COMV6, "0.00")
txtpor7.Text = Format(comivem_llave!COA_COMV7, "0.00")
txtpor8.Text = Format(comivem_llave!COA_COMV8, "0.00")
txtpor9.Text = Format(comivem_llave!COA_COMV9, "0.00")

txtpor11.Text = Format(comivem_llave!COA_COMC1, "0.00")
txtpor22.Text = Format(comivem_llave!COA_COMC2, "0.00")
txtpor33.Text = Format(comivem_llave!COA_COMC3, "0.00")
txtpor44.Text = Format(comivem_llave!COA_COMC4, "0.00")
txtpor55.Text = Format(comivem_llave!COA_COMC5, "0.00")
txtpor66.Text = Format(comivem_llave!COA_COMC6, "0.00")
txtpor77.Text = Format(comivem_llave!COA_COMC7, "0.00")
txtpor88.Text = Format(comivem_llave!COA_COMC8, "0.00")
txtpor99.Text = Format(comivem_llave!COA_COMC9, "0.00")

cmdactvend.Enabled = False
Azul txtpor1, txtpor1

End Sub

Private Sub cmdverifi_Click()
On Error GoTo sale
pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '01'  AND ARM_CODART NOT IN (SELECT PRE_CODART FROM PRECIOS WHERE PRE_CODCIA = '01' AND PRE_FLAG_UNIDAD = 'A')"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
  MsgBox "Verificar Producto sin unidad Activa : " & X!ARM_CODART & Chr(13)
X.MoveNext
Loop
MsgBox "Proceso Terminado.", 48, Pub_Titulo
sale:

End Sub

Private Sub decimales_GotFocus()
frmARTI.F14.Visible = False
End Sub

Private Sub decimales_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Flag_Inicial = "A"
 grid_unid.COL = 0
 grid_unid.Row = 1
 Flag_Inicial = ""
 grid_unid.SetFocus
End If

End Sub

Private Sub decimales_LostFocus()
Fvarios.Refresh
End Sub

Private Sub DS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Flag_Inicial = "A"
 grid_unid.COL = 0
 grid_unid.Row = 1
 Flag_Inicial = ""
 If Trim(frmARTI.DS.Text) = "S" Then
     cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
     cmddolares.Tag = "S"
 Else
     cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
     cmddolares.Tag = "D"
 End If
 'cmddolares_Click
 If grid_unid.Enabled Then grid_unid.SetFocus
  
End If
End Sub

Private Sub exigv_Click()
If Left(cmdModificar.Caption, 2) = "&M" And Left(cmdAgregar.Caption, 2) = "&A" Then Exit Sub
  
If exigv.Value = 0 Then
'  txtcospro.Enabled = False
  txtcospro.Text = ""
Else
'  txtcospro.Enabled = True
  txtcospro.SetFocus
End If
End Sub

Private Sub Form_Activate()
frmARTI.SSTab1.tab = 0
End Sub

Private Sub Form_DblClick()
'Load FrmNC
'FrmNC.Show 1
'LEVANTE_ZONA
'LEVANTE_ZONA_ORIGINAL
'LEVANTE_ZONA
'opcional
 busca
End Sub

Private Sub Form_Load()
CenterMe Arti
Dim loc_flag_listas As String
Dim psvend As rdoResultset

flag_art_situacion = 0
lstcias.Clear
If LK_CODUSU = "SUPERVISOR" Or LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then
 cmdverifi.Visible = True
 Cmdmax.Visible = True
Else
Cmdmax.Visible = False
End If
loc_equiv_max = 1
Flag_Add_cant = ""
loc_flag_bloq = ""
wflag_bloq_desactiva = ""
loc_flag_listas = ""
loc_p = ""
loc_acc_convenio = ""
loc_cheoferta = ""
For fila = 1 To lk_OTROS_Count

   If Val(lk_OTROS(fila)) = 6 Then ' bloque de precios en mastros de articulos
    loc_flag_bloq = "A"
   End If
   If Val(lk_OTROS(fila)) = 27 Then ' bloque de precios en mastros de articulos
      wflag_bloq_desactiva = "A"
   End If
   If Val(lk_OTROS(fila)) = 31 Then loc_flag_listas = "A"
   If Val(lk_OTROS(fila)) = 30 Then Flag_Add_cant = "A"
   If Val(lk_OTROS(fila)) = 36 Then loc_p = "A"
   If Val(lk_OTROS(fila)) = 38 Then loc_acc_convenio = "A"
   If Val(lk_OTROS(fila)) = 41 Then loc_cheoferta = "A" ' ACCESO A CHEOFERTA
Next fila

pub_cadena = "SELECT ART_NOMBRE, ART_ALTERNO, ART_KEY , ART_codart2  FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? "
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA.rdoParameters(0) = 0
PSART_RELA.rdoParameters(1) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY = ? ORDER BY ART_KEY"
Set PSART_KEY = CN.CreateQuery("", pub_cadena)
PSART_KEY.rdoParameters(0) = "  "
PSART_KEY.rdoParameters(1) = 0
Set artloc_key = PSART_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT SUM(FAR_CANTIDAD)AS CANTIDAD FROM FACART WHERE (FAR_CODCIA = ? or FAR_CODCIA = ? or FAR_CODCIA = ?) AND FAR_CODART = ? and (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_TIPMOV = 10  AND FAR_ESTADO  <> 'E' AND FAR_ESTADO2 <> 'L' AND FAR_FBG <> 'G' GROUP BY FAR_CODART"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI.rdoParameters(0) = 0
PSLOC_WARTI.rdoParameters(1) = 0
PSLOC_WARTI.rdoParameters(2) = 0
PSLOC_WARTI.rdoParameters(3) = 0
PSLOC_WARTI.rdoParameters(4) = LK_FECHA_DIA
PSLOC_WARTI.rdoParameters(5) = LK_FECHA_DIA
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)



'If LK_EMP = "HER" Then
  Fcomi.Visible = True
'Else
'  Fop.Left = 120
'End If
    VAR_NEWCAL = 0
    frmARTI.F14.Left = 90
    frmARTI.F14.Top = 3360
    frmARTI.F14.Height = 2415
    frmARTI.F14.Width = 9380
    pasa = 0
    frmARTI.ListExiste.Cols = 4
    For fila = 0 To frmARTI.ListExiste.COL - 1
      frmARTI.ListExiste.COL = fila
      frmARTI.ListExiste.FixedAlignment(fila) = 2
    Next fila
    pasa = 0
    frmARTI.ListExiste.ColWidth(0) = 350
    frmARTI.ListExiste.ColWidth(1) = 800
    frmARTI.ListExiste.ColWidth(2) = 4000
    frmARTI.ListExiste.ColWidth(3) = 1500
    Screen.MousePointer = 11
    frmARTI.DS.AddItem "S"
    frmARTI.DS.AddItem "D"
    DS.ListIndex = 0
    frmARTI.decimales.AddItem "1"
    frmARTI.decimales.AddItem "2"
    frmARTI.decimales.AddItem "3"
    frmARTI.decimales.AddItem "4"
    frmARTI.DS.ListIndex = -1
    frmARTI.decimales.ListIndex = -1
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, art_grupop, lfvta
    If LK_FLAG_ORIGINAL = "A" Then
       txt_alterno.Visible = False
       lblalterno.Visible = False
       txtnombre.Left = 2160
       lblnomarti.Left = 2160
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       BLOQUEA_TEXT txt_key
       pub_cadena = "SELECT ART_ALTERNO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_alterno"
    Else
      BLOQUEA_TEXT txt_alterno
       pub_cadena = "SELECT ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_NOMBRE"
    End If
    Set PSMANO_CODI = CN.CreateQuery("", pub_cadena)
    PSMANO_CODI.rdoParameters(0) = 0
    Set mano_CODI = PSMANO_CODI.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSMANO_CODI(0) = LK_CODCIA
    mano_CODI.Requery
    LLENADO_GRUPOP
    Fvarios.Visible = False
    Fcomun.Visible = True
    Fdatos.Visible = True
    PROCESO_ARTI
    loc_tipo = "V"
    LLENADO_FAM
    'LLENADO_GRUPO
    LLENADO_CAL
    LLENADO_NUMERO 130
    LLENADO_MARCA
    LLENADO_LINEA 131
    LLENADO_PLANCHA
    Fvarios.Visible = True
    Screen.MousePointer = 0
    LBLUNIDAD.Caption = ""
    frmARTI.Fcomun.Visible = True
    frmARTI.Fcomun.Enabled = True
    grid_unid.Enabled = False
    cmdCancelar.Enabled = True
    SQ_OPER = 2
    PUB_TIPREG = 45
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
      Label3(tab_mayor!TAB_NUMTAB - 1).Caption = Trim(tab_mayor!TAB_NOMLARGO)
      lblpor(tab_mayor!TAB_NUMTAB - 1).Caption = Left(lblpor(tab_mayor!TAB_NUMTAB - 1).Caption, 5) & Trim(tab_mayor!TAB_NOMLARGO) & " :"
      tab_mayor.MoveNext
   Loop
   pasa = 0
   loc_tipo = "V"
   PROCESA_PROV
   frmARTI.frarelacion.Enabled = False
   cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
   cmddolares.Tag = "S"
   SQ_OPER = 1
   PUB_CODCIA = "00"
   PUB_TIPREG = 340
   PUB_NUMTAB = 0
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(0).Caption = Trim(tab_llave!TAB_NOMLARGO)
   PUB_NUMTAB = 1
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(1).Caption = Trim(tab_llave!TAB_NOMLARGO)
   PUB_NUMTAB = 2
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(2).Caption = Trim(tab_llave!TAB_NOMLARGO)
   PUB_NUMTAB = 3
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(3).Caption = Trim(tab_llave!TAB_NOMLARGO)
   PUB_NUMTAB = 4
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(4).Caption = Trim(tab_llave!TAB_NOMLARGO)
   PUB_NUMTAB = 5
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then lblart(5).Caption = Trim(tab_llave!TAB_NOMLARGO)
   'PUB_NUMTAB = 6
   'LEER_TAB_LLAVE
   'If Not tab_llave.EOF Then lblart(5).Caption = Trim(tab_llave!TAB_NOMLARGO)
   
   If Trim(loc_flag_listas) = "A" Then cmdact.Visible = True
   
   pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = '" & LK_CODCIA & "' ORDER BY VEM_CODVEN"
   Set psvend = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   psvend.Requery
   cmdvend.Clear
   Do Until psvend.EOF
     cmdvend.AddItem Format(psvend!VEM_CODVEN, "000") & " - " & Trim(psvend!VEM_NOMBRE)
     psvend.MoveNext
   Loop
   pub_cadena = "SELECT *  FROM COMIARTI WHERE COA_CODCIA = ? AND COA_CODART = ? AND COA_CODVEN = ? "
   Set PSCOMI_VEN = CN.CreateQuery("", pub_cadena)
   PSCOMI_VEN.rdoParameters(0) = 0
   PSCOMI_VEN.rdoParameters(1) = 0
   PSCOMI_VEN.rdoParameters(2) = 0
   Set comivem_llave = PSCOMI_VEN.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "SUPER" Then cmbStock.Visible = True

ws_flag_equiv = ""
FLAG_BUSCAR_MANITO = ""
cmbUsu.Clear
pub_cadena = "SELECT  * FROM USUARIOS "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
  cmbUsu.AddItem Trim(X!usu_key)
X.MoveNext
Loop
Dim i As Integer
For i = 0 To cmbUsu.ListCount - 1
    If Trim(cmbUsu.List(i)) = Trim(LK_CODUSU) Then
        cmbUsu.ListIndex = i
        Exit For
    End If
Next i
If LK_CODUSU = "ADMIN" Or LK_CODUSU = "OPER08" Then
  fracuotas.Enabled = True
Else
  fracuotas.Enabled = False
End If
txtper.Text = Format(LK_FECHA_DIA, "mmyyyy")
End Sub

Public Sub LLENADO_FAM()
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 122
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_familia.ToolTipText = "TAB_TIPREG = 122"
    art_familia.Clear
    listlin.Clear
    listlineas.Clear
    cmdlineas.Clear
    cmdlineas.AddItem "(...)"
    Do Until tab_mayor.EOF
        art_familia.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        listlin.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        cmdlineas.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        listlineas.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
    
        
End Sub
Public Sub LLENADO_GRUPOP()

    PUB_TIPREG = 121
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_grupop.ToolTipText = "TAB_TIPREG = 121"
    art_grupop.Clear
    Do Until tab_mayor.EOF
        art_grupop.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub LLENADO_SUBFAM(ctlCombo As ComboBox, ByVal wfami As Integer)
On Error GoTo sale
Dim CONTA As Integer
    CONTA = -1
'    Select Case ctlCombo.Name
'      Case Is = "art_subfam"
'       PUB_TIPREG = 123
'      Case Is = "art_grupo"
'       PUB_TIPREG = 129
'      Case Is = "art_numero"
'       PUB_TIPREG = 130
'      Case Is = "art_linea"
'       PUB_TIPREG = 131
'    End Select
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    Select Case ctlCombo.Name
      Case Is = "art_subfam"
       ctlCombo.ToolTipText = "TAB_TIPREG = 123"
      Case Is = "art_grupo"
       ctlCombo.ToolTipText = "TAB_TIPREG = 129"
      Case Is = "art_numero"
       ctlCombo.ToolTipText = "TAB_TIPREG = 130"
      Case Is = "art_linea"
       ctlCombo.ToolTipText = "TAB_TIPREG = 131"
      Case Is = "art_marca"
       ctlCombo.ToolTipText = "TAB_TIPREG = 132"
    End Select
    'ctlCombo.ToolTipText = "TAB_TIPREG = 123"
    ctlCombo.Clear
    Do Until tab_menor.EOF
        ctlCombo.AddItem tab_menor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
Exit Sub
sale:
Resume Next
End Sub
Public Sub LLENADO_GRUPO(ByVal wSubFam As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 129
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    PUB_CODART = wSubFam
    LEER_TAB_LLAVE
    art_grupo.ToolTipText = "TAB_TIPREG = 129"
    art_grupo.Clear
    Do Until tab_mayor.EOF
        art_grupo.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
'PROCEDIMIENTO PARA LLENAR GRUPO 2
Public Sub LLENADO_NUMERO(ByVal wgrupo As Integer)
    PUB_TIPREG = 130
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
'    PUB_CODART = wgrupo
    LEER_TAB_LLAVE
    art_numero.ToolTipText = "TAB_TIPREG = 130"
    art_numero.Clear
    Do Until tab_mayor.EOF
       art_numero.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_MARCA()
    PUB_TIPREG = 132
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_marca.ToolTipText = "TAB_TIPREG = 132"
    art_marca.Clear
    Do Until tab_mayor.EOF
       art_marca.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_PLANCHA()
    PUB_TIPREG = 133
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    art_plancha.ToolTipText = "TAB_TIPREG = 133"
    art_plancha.Clear
    Do Until tab_mayor.EOF
       art_plancha.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
'PROCEDIMEINTO PARA LLENAR LAS LINEAS
Public Sub LLENADO_LINEA(ByVal wSubGrupo As Integer)
    PUB_TIPREG = 131
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    PUB_CODART = wSubGrupo
    LEER_TAB_LLAVE
    art_linea.ToolTipText = "TAB_TIPREG = 131"
    art_linea.Clear
    Do Until tab_mayor.EOF
       art_linea.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & Trim(CStr(tab_mayor!TAB_NUMTAB))
       tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADO_CAL()
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 2
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
       PUB_CODCIA = "00"
    End If
    SQ_OPER = 2
    LEER_TAB_LLAVE
    CmbCalidad.ToolTipText = "TAB_TIPREG = 2"
    CmbCalidad.Clear
    Do Until tab_mayor.EOF
        CmbCalidad.AddItem tab_mayor!TAB_NOMLARGO & String(80, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub ASIGNA(WCONTROL As ComboBox, txt As String)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Trim(WCONTROL.List(C)) = Trim(txt) Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub LEER_LOC()
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If
If LOC_OPER = 1 Then
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      PSART_LOC.rdoParameters(0) = pu_alterno
    Else
      PSART_LOC.rdoParameters(0) = PUB_KEY
    End If
    PSART_LOC.rdoParameters(1) = PUB_CODCIA
    PSART_LOC.rdoParameters(2) = loc_tipo
    artloc_llave.Requery
ElseIf LOC_OPER = 2 Then
  PSART_KEY.rdoParameters(0) = PUB_CODCIA
  PSART_KEY.rdoParameters(1) = PUB_KEY
  artloc_key.Requery
End If

End Sub

Public Sub CAB_ARTI()
grid1.TextMatrix(0, 0) = "Cia."
grid1.TextMatrix(0, 3) = "Grupo"
grid1.TextMatrix(0, 2) = "Articulo"

End Sub
Public Sub ASIGNA_CHAR(WCONTROL As ComboBox, txt As String)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Trim(Left(WCONTROL.List(C), 1)) = Trim(txt) Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub
Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Currency)
Dim C As Currency
For C = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(C), 6))) = txt Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub

Private Sub grid_unid_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
 Azul txtpeso, txtpeso
End If
End Sub

Private Sub Label3_DblClick(Index As Integer)
If LK_CODUSU <> "SUPER" And Trim(LK_CODUSU) <> "ADMIN" And Trim(LK_CODUSU) <> "SUPERVISOR" Then
 Exit Sub
End If
If Trim(Label3(Index).Tag) = "" Then
 Exit Sub
End If
Dim wnombre
wnombre = InputBox("Ingrese la Nueva Descripci�n para este Campo :", Pub_Titulo, Trim(Label3(Index).Caption))
If wnombre = "" Then
  Screen.MousePointer = 0
  Exit Sub
End If
Screen.MousePointer = 11
SQ_OPER = 1
PUB_TIPREG = 45
PUB_NUMTAB = Val(Label3(Index).Tag)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_llave.EOF Then
  tab_llave.AddNew
Else
  tab_llave.Edit
End If
  tab_llave!tab_codcia = LK_CODCIA
  tab_llave!tab_tipreg = 45
  tab_llave!TAB_NUMTAB = Val(Label3(Index).Tag)
  tab_llave!TAB_NOMLARGO = Left(wnombre, 40)
  tab_llave!tab_nomcorto = Left(wnombre, 10)
  tab_llave.Update
  Label3(Index).Caption = Left(wnombre, 40)
  lblpor(Index).Caption = Left(lblpor(Index).Caption, 5) & Trim(wnombre) & " :"
Screen.MousePointer = 0

End Sub

Private Sub lblcospro_DblClick()
Dim wcos  As String
If LK_CODUSU = "OPERCH01" Or LK_CODUSU = "ADMIN" Then
Else
Exit Sub
End If
If Val(txt_key.Text) <> 0 Then
  SQ_OPER = 1
  PUB_CODART = txt_key.Text
  pu_codcia = LK_CODCIA
  LEER_ARM_LLAVE
  If arm_llave.EOF Then
   MsgBox "Productos No Existe"
  End If
  wcos = InputBox("COSTO DE LA UNIDAD MINIMA : ", "EDITAR COSTO", arm_llave!ARM_COSPRO)
  If wcos = "" Then Exit Sub
  
  arm_llave.Edit
  arm_llave!ARM_COSPRO = Val(wcos)
  arm_llave.Update
  MsgBox "Cambiado"
End If
End Sub

Private Sub lblequivarti_DblClick()
LK_ACCESO_REPORT = ""
Load frmclave2
Screen.MousePointer = 0
frmclave2.Show 1
If LK_ACCESO_REPORT <> "A" Then
   Exit Sub
End If

Load frmEquivArti
frmEquivArti.Show 1
End Sub

Private Sub ListExiste_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And frmARTI.txtnombre.Enabled Then
    frmARTI.txtnombre.SetFocus
    frmARTI.F14.Visible = False
    Exit Sub
End If

End Sub

Private Sub ListExiste_LostFocus()
If frmARTI.ListExiste.Visible = False Then
    Exit Sub
End If

End Sub

Private Sub ListView1_DblClick()
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 loc_key = ListView1.SelectedItem.Index
 txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_alterno_KeyPress 13
Else
 loc_key = ListView1.SelectedItem.Index
 txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 txt_key_KeyPress 13
End If
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 Else
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
 End If
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 ListView1.Visible = False
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" And txt_alterno.Enabled Then
  txt_alterno.Text = ""
  txt_alterno.SetFocus
 ElseIf LK_FLAG_ALTERNO <> "A" And txt_key.Enabled Then
  txt_key.Text = ""
  txt_key.SetFocus
 End If
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub


Private Sub MANOS_Click(Index As Integer)
If FLAG_BUSCAR_MANITO = "" Then
    FLAG_BUSCAR_MANITO = "A"
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      pub_cadena = "SELECT ART_KEY FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_NOMBRE"
    Else
      pub_cadena = "SELECT ART_ALTERNO FROM ARTI WHERE ART_CODCIA = ? AND ART_KEY <> 0 order by art_alterno"
    End If
    Set PSMANO_CODI = CN.CreateQuery("", pub_cadena)
    PSMANO_CODI.rdoParameters(0) = 0
    Set mano_CODI = PSMANO_CODI.OpenResultset(rdOpenKeyset, rdConcurValues)
    PSMANO_CODI(0) = LK_CODCIA
    mano_CODI.Requery
End If


mano_CODI.MoveFirst
If mano_CODI.RowCount > 0 Then
   Do Until mano_CODI.EOF
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       If Trim(mano_CODI!ART_alterno) = Trim(txt_alterno.Text) Then Exit Do
    Else
       If Val(txt_key.Text) = Val(mano_CODI!art_key) Then Exit Do
    End If
    mano_CODI.MoveNext
  Loop
End If
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
    If txt_alterno.Text = "" Then
     mano_CODI.MoveFirst
    GoTo SALT
    End If
Else
    If txt_key.Text = "" Then
     mano_CODI.MoveFirst
     GoTo SALT
    End If
End If


If Index = 0 Then
  If Not mano_CODI.BOF Then mano_CODI.MovePrevious
Else
  If Not mano_CODI.EOF Then mano_CODI.MoveNext
End If
SALT:
If mano_CODI.EOF Or mano_CODI.BOF Then Exit Sub
art_plancha.ListIndex = -1
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  txt_alterno.Text = Trim(mano_CODI!ART_alterno)
  txt_alterno_KeyPress 13
Else
 txt_key.Text = Trim(mano_CODI!art_key)
 txt_key_KeyPress 13

End If


End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = Not LblMensaje.Visible
 If CU > 4 Then
   CU = 0
   Parpadea.Enabled = False
   LblMensaje.Visible = False
 End If

End Sub

Public Sub LLENA_ARTI(ban As Integer)
Dim ws_flag As String * 1
Dim WFAMI2 As Integer
Dim WS_FLAG_UNIDAD As Integer
Dim WSPOR As Currency
Dim psfv As rdoResultset
If ban <> 1 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
       Else
         If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          If VAR_ACTIVAR = 99 Then
           txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text)
          End If
          pu_alterno = Trim(txt_alterno.Text)
         Else
          txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
          PUB_KEY = Val(txt_key.Text)
         End If
       End If
       PUB_CODCIA = LK_CODCIA
       LOC_OPER = 1
       LEER_LOC
       If artloc_llave.EOF Then
        Exit Sub
       End If
End If
'LLENADO
frmARTI.frarelacion.Enabled = True
frmARTI.txt_key.Text = artloc_llave!art_key
frmARTI.txtnombre.Text = RTrim(Nulo_Valors(artloc_llave!art_nombre))
ASIGNA_INT art_familia, Nulo_Valor0(artloc_llave!art_familia)

ASIGNA_INT art_grupop, Nulo_Valor0(artloc_llave!art_grupop)

PUB_TIPREG = 123
WFAMI2 = Val(Trim(Right(art_familia.Text, 6)))
LLENADO_SUBFAM art_subfam, WFAMI2
ASIGNA_INT art_subfam, Nulo_Valor0(artloc_llave!art_subfam)

PUB_TIPREG = 129
WFAMI2 = Val(Trim(Right(art_subfam.Text, 6)))
LLENADO_SUBFAM art_grupo, WFAMI2
ASIGNA_INT art_grupo, Nulo_Valor0(artloc_llave!art_subgru)

'PUB_TIPREG = 130
'WFAMI2 = Val(Trim(Right(art_grupo.Text, 6)))
'LLENADO_SUBFAM art_numero, WFAMI2
ASIGNA_INT art_numero, Nulo_Valor0(artloc_llave!art_numero)

'PUB_TIPREG = 131
'WFAMI2 = Val(Trim(Right(art_numero.Text, 6)))
'LLENADO_SUBFAM art_linea, WFAMI2
ASIGNA_INT art_linea, Nulo_Valor0(artloc_llave!art_linea)

'PUB_TIPREG = 132
'WFAMI2 = Val(Trim(Right(art_linea.Text, 6)))
'LLENADO_SUBFAM art_marca, WFAMI2
ASIGNA_INT art_marca, Nulo_Valor0(artloc_llave!art_marca)

ASIGNA_INT CmbCalidad, Nulo_Valor0(artloc_llave!ART_CALIDAD)

ASIGNA_INT art_plancha, Nulo_Valor0(artloc_llave!art_plancha)
ASIGNA_INT art_codpro, Nulo_Valor0(artloc_llave!art_codclie)
frmARTI.txtcospro.Text = Nulo_Valor0(artloc_llave!ART_POR_IGV)
SQ_OPER = 1
PUB_CODART = artloc_llave!art_key
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
frmARTI.lblcospro.Caption = Nulo_Valor0(arm_llave!ARM_COSPRO)
frmARTI.txtfechault.Text = Format(arm_llave!ARM_FECHA_ULT, "dd/mm/yyyy")
frmARTI.DS.Text = Trim(Nulo_Valors(artloc_llave!ART_MONEDA))
frmARTI.DS.Text = Trim(Nulo_Valors(artloc_llave!ART_MONEDA))
frmARTI.cmbUsu.Text = Trim(Nulo_Valors(artloc_llave!ART_CODUSU_ORIG))
frmARTI.decimales.Text = Val(Nulo_Valor0(artloc_llave!ART_DECIMALES))
frmARTI.txt_alterno.Text = Nulo_Valors(artloc_llave!ART_alterno)
frmARTI.txtMin.Text = Nulo_Valors(artloc_llave!ART_STOCK_MIN)
frmARTI.txtMax.Text = Nulo_Valors(artloc_llave!ART_STOCK_MAX)
LLENA_CALREL Nulo_Valor0(artloc_llave!ART_CALIDAD)
LLENA_RELACION Nulo_Valor0(artloc_llave!ART_CODART2)
txtcodigo2.Text = Nulo_Valor0(artloc_llave!ART_CODART2)
art_situacion.Value = Val((artloc_llave!art_situacion))
loc_orig_situ = Trim((artloc_llave!art_situacion))
chebof.Value = Val((artloc_llave!ART_ORDEN))
If Nulo_Valors(artloc_llave!ART_EX_IGV) = "A" Then
   exigv.Value = 1
Else
 exigv.Value = 0
End If

checambio.Value = 0
If Nulo_Valors(artloc_llave!art_flag_cambio) = "A" Then
   checambio.Value = 1
End If


pub_cadena = "Select  * from tabcuo where cuo_codcia = '" & LK_CODCIA & "' and cuo_periodo = '" & Format(LK_FECHA_DIA, "mmyyyy") & "' and cuo_codart = " & artloc_llave!art_key
Set rs_regcuotas = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
If Not rs_regcuotas.EOF Then
  pub_cadena = "SELECT   top 1 pre_equiv FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND PRE_CODART = " & artloc_llave!art_key & " order by pre_equiv desc"
  Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   txtper.Text = rs_regcuotas!CUO_PERIODO
   If X.EOF Then
   Else
     txtum.Text = Format(rs_regcuotas!CUO_CANTIDAD / X!PRE_EQUIV, "0.00")
   End If
End If




cheservi(0).Value = False
cheservi(1).Value = False
cheservi(2).Value = False
If Nulo_Valors(artloc_llave!art_flag_stock) = "M" Then
 cheservi(0).Value = True
ElseIf Nulo_Valors(artloc_llave!art_flag_stock) = "P" Then
 cheservi(1).Value = True
ElseIf Nulo_Valors(artloc_llave!art_flag_stock) = "S" Then
 cheservi(2).Value = True
End If
If Trim(gridrel.TextMatrix(1, 0)) = "" Then
 cmbcal.Visible = True
 lblcal.Visible = True
 cmdAdd.Visible = True
 cmdquitar.Visible = False
Else
 cmbcal.Visible = False
 lblcal.Visible = False
 cmdAdd.Visible = False
 cmdquitar.Visible = True
End If
txtmargen.Text = Format(Nulo_Valor0(artloc_llave!ART_MARGEN), "0")
fechahora.Caption = Trim(artloc_llave!art_fechahora)
txtpor1.Text = Nulo_Valor0(artloc_llave!ART_POR1)
txtpor2.Text = Nulo_Valor0(artloc_llave!ART_POR2)
txtpor3.Text = Nulo_Valor0(artloc_llave!ART_POR3)
txtpor4.Text = Nulo_Valor0(artloc_llave!ART_POR4)
txtpor5.Text = Nulo_Valor0(artloc_llave!ART_POR5)
txtpor11.Text = Nulo_Valor0(artloc_llave!ART_POR11)
txtpor22.Text = Nulo_Valor0(artloc_llave!ART_POR22)
txtpor33.Text = Nulo_Valor0(artloc_llave!ART_POR33)
txtpor44.Text = Nulo_Valor0(artloc_llave!ART_POR44)
txtpor55.Text = Nulo_Valor0(artloc_llave!ART_POR55)
chesub.Value = Val(Nulo_Valor0(artloc_llave!art_estado))
flag_art_situacion = Nulo_Valor0(artloc_llave!art_situacion)
cheoferta.Value = Nulo_Valor0(artloc_llave!ART_POR6)   ' FLAG 1 = ESTA EN OFERTA
If Trim(frmARTI.DS.Text) = "S" Then
 cmddolares.Caption = "Lista de Precios en S/. (Nuevos Soles)"
 cmddolares.Tag = "S"
Else
 cmddolares.Caption = "Lista de Precios en US$. (Dolares Americanos)"
 cmddolares.Tag = "D"
End If

' ACCESO A MODIFICAR DESCRIP PRODUCTO
ws_flag = ""
For fila = 1 To lk_OTROS_Count
   If Val(lk_OTROS(fila)) = 8 Then 'modificar Venta
     ws_flag = "A"
   End If
Next fila
If ws_flag = "A" Then
  frmARTI.txtnombre.Locked = False
Else
  frmARTI.txtnombre.Locked = True
End If

llena_pre Trim(frmARTI.DS.Text)
frmARTI.txtCantMax.Text = Format(Nulo_Valor0(arm_llave!arm_stock2) / loc_equiv_max, "0")
frmARTI.txtCantMaxADD.Text = Format(Nulo_Valor0(arm_llave!arm_saldo_n2) / loc_equiv_max, "0")

lfvta.Clear
pub_cadena = "SELECT  * FROM TABLAS WHERE TAB_TIPREG = 39 AND TAB_CODCIA = '" & LK_CODCIA & "' ORDER BY TAB_NOMLARGO"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
  lfvta.AddItem X!TAB_NOMLARGO & String(80, " ") & X!TAB_NUMTAB
  pub_cadena = "SELECT * FROM FVENTAS WHERE FV_CODCIA = '" & LK_CODCIA & "' AND FV_TIPO = 2 AND FV_KEYFV = " & X!TAB_NUMTAB & " AND FV_CODIGO = " & (artloc_llave!art_key)
  Set psfv = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
  If Not psfv.EOF Then
    lfvta.Selected(lfvta.ListCount - 1) = True
  End If
X.MoveNext
Loop


If LK_EMP = "CAM" Then
  PROD_PROC
End If
ws_flag_equiv = ""
frmARTI.SSTab1.tab = 0
VAR_ACTIVAR = 0
End Sub
Public Sub LIMPIA_ARTI()
Dim I As Integer
lfvta.Clear
art_grupop.ListIndex = -1
txtum.Text = ""
txtper.Text = ""
cheasignarc.Value = 0
frmARTI.txtCantMax.Text = ""
frmARTI.txtCantMaxADD.Text = ""
cmbUsu.ListIndex = -1
fechahora.Caption = ""
txtmargen.Text = ""
flag_art_situacion = 0
lstcias.Clear
fragen.Visible = False
frmARTI.Lst0.Caption = ""
frmARTI.Lst1.Caption = ""
frmARTI.lblstock0.Caption = "0.000"
frmARTI.lblstock1.Caption = "0.000"
txtMin.Text = ""
txtMax.Text = ""
frmARTI.txt_alterno.Text = ""
frmARTI.txt_key.Text = ""
frmARTI.txtnombre.Text = ""
DS.ListIndex = -1
LBLUNIDAD.Caption = ""
cheoferta.Value = 0
decimales.ListIndex = -1
CmbCalidad.ListIndex = -1
art_familia.ListIndex = -1
art_subfam.ListIndex = -1
art_grupo.ListIndex = -1
art_codpro.ListIndex = -1
art_linea.ListIndex = -1
art_numero.ListIndex = -1
art_marca.ListIndex = -1
'txtcospro.Text = ""
tcospro.Text = ""
DS.ListIndex = -1
art_plancha.ListIndex = -1
decimales.Text = ""
grid_unid.Clear
grid_unid.Cols = 1
grid_unid.Rows = 1
frmARTI.SSTab1.tab = 0
cheservi(0).Value = False
cheservi(1).Value = False
cheservi(2).Value = False
checambio.Value = 0
txtpeso.Text = ""
txtlitro.Text = ""
gridrel.Clear
txtpor1.Text = ""
txtpor2.Text = ""
txtpor3.Text = ""
txtpor4.Text = ""
txtpor5.Text = ""
txtpor6.Text = ""
txtpor7.Text = ""
txtpor8.Text = ""
txtpor9.Text = ""

txtpor11.Text = ""
txtpor22.Text = ""
txtpor33.Text = ""
txtpor44.Text = ""
txtpor55.Text = ""
txtpor66.Text = ""
txtpor77.Text = ""
txtpor88.Text = ""
txtpor99.Text = ""


  txtcodigo2.Text = ""
'  exigv.Value = 0
VAR_ACTIVAR = 0
art_situacion.Value = 0
chebof.Value = 0
lblcospro.Caption = ""
chesub.Value = 0
frmARTI.txtnombre.Locked = False
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.tab = 1 And txtMin.Enabled = True Then
 ' If LK_EMP = "HER" Then
    Azul txtpor1, txtpor1
 ' Else
 '   Azul txtMin, txtMin
 ' End If
ElseIf SSTab1.tab = 0 And art_familia.Enabled = True Then
  art_familia.SetFocus
End If
End Sub


Private Sub SSTab1_GotFocus()
If ListView1.Visible Then
  ListView1.Visible = False
  txt_key.Text = ""
End If
If SSTab1.tab = 1 And txtpor1.Enabled = True Then
  If LK_EMP = "HER" Then
    Azul txtpor1, txtpor1
  End If
ElseIf SSTab1.tab = 0 And art_familia.Enabled = True Then
  art_familia.SetFocus
End If

End Sub

Private Sub txtcospro_GotFocus()
    Azul txtcospro, txtcospro
    frmARTI.F14.Visible = False
End Sub

Private Sub txtcospro_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtcospro, KeyAscii
End Sub

Private Sub txtlitro_Change()
  If grid_unid.Rows > 1 Then
    grid_unid.TextMatrix(grid_unid.Row, 30) = Val(txtlitro.Text)
  End If

End Sub

Private Sub txtlitro_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtlitro, KeyAscii
If KeyAscii = 13 Then
  grid_unid.SetFocus
End If

End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Left(cmdAgregar.Caption, 2) = "&G" Then
        cmdAgregar.SetFocus
    Else
        cmdModificar.SetFocus
    End If
End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Azul txtMax, txtMax
End If
End Sub

Private Sub txtnombre_GotFocus()
'    Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
  KeyAscii = 0
  Exit Sub
End If
If KeyAscii = 13 Then
    If loc_tipo = "V" Then
    If frmARTI.SSTab1.tab = 1 Then
       If LK_EMP = "HER" Then
        Azul txtpor1, txtpor1
        Exit Sub
       End If
       art_situacion.SetFocus
       Exit Sub
    End If
       If art_familia.Enabled Then art_familia.SetFocus
        DoEvents
        SendKeys "%{UP}"
        Exit Sub
    End If
End If
  
End Sub

Public Sub GRABAR_ARTI()
Dim psfv As rdoResultset
Dim rs_regcuotas As rdoResultset
Dim wSTOCK  As Currency
Dim ws_igv As Currency
Dim I As Integer
Dim WS_IMPORTE As Currency
Dim WS_FLAG_UNIDAD As Integer
Dim WORIGINAL As Currency
Dim walterno As String
Dim wnombre As String
Dim WCODART2 As Currency
Dim ws_codcia As String * 2
Dim xcuenta As Integer
WS_FLAG_UNIDAD = 0
WS_IMPORTE = 0
ws_igv = 0
WORIGINAL = Val(frmARTI.txt_key.Text)
If LK_FLAG_ORIGINAL = "A" Then
 walterno = frmARTI.txt_key.Text
Else
 walterno = Trim(frmARTI.txt_alterno.Text)
End If
wnombre = frmARTI.txtnombre.Text
WCALIDAD = Val(Right(CmbCalidad.Text, 3))
WCODART2 = Val(txtcodigo2.Text)
PUB_CODCIA = pu_codcia 'LK_CODCIA
If LK_EMP_PTO = "A" Then
  PUB_CODCIA = "00"
End If

If VAR_NEWCAL = 1 Then
  WORIGINAL = LOC_ORIGINAL
  walterno = LOC_ALTERNO
  wnombre = LOC_NOMBRE
  WCALIDAD = LOC_CALIDAD
  WCODART2 = 0
  GoTo IR_GRABA1
End If
If Left(cmdModificar.Caption, 2) = "&G" Then
    artloc_llave.Edit
     PUB_CODART = WORIGINAL
     SQ_OPER = 1
     LEER_ARM_LLAVE
     arm_llave.Edit
     arm_llave!ARM_FECHA_ULT = frmARTI.txtfechault.Text
     If pu_codcia = "03" And pu_codcia = "05" And pu_codcia = "07" And pu_codcia = "09" And pu_codcia = "50" And pu_codcia = "25" Then
        arm_llave!arm_stock2 = Val(txtCantMax.Text) * loc_equiv_max
     End If
     arm_llave!arm_saldo_n2 = Val(txtCantMaxADD.Text) * loc_equiv_max
     arm_llave.Update
   
Else
IR_GRABA1:
    artloc_llave.AddNew
    artloc_llave!art_key = WORIGINAL
    If LK_FLAG_ORIGINAL = "A" Then
     artloc_llave!ART_alterno = Trim(Str(WORIGINAL))
    Else
     artloc_llave!ART_alterno = walterno
    End If
    artloc_llave!ART_POR_IGV = 0
    artloc_llave!ART_STOCK_MIN = 0
    artloc_llave!ART_STOCK_MAX = 0
    artloc_llave!art_situacion = 0
End If
If LK_FLAG_ORIGINAL = "A" Then
 artloc_llave!ART_alterno = WORIGINAL ' frmARTI.Txt_key.Text
Else
  artloc_llave!ART_alterno = walterno ' Trim(txt_alterno.Text)
End If

If pu_codcia = LK_CODCIA Then
   artloc_llave!ART_STOCK_MIN = Nulo_Valor0(frmARTI.txtMin.Text)
   artloc_llave!ART_STOCK_MAX = Nulo_Valor0(frmARTI.txtMax.Text)
   artloc_llave!art_marca = Val(Right(art_marca.Text, 3))
   'CUO_CODCIA CUO_PERIODO CUO_CODART  CUO_CANTIDAD
   pub_cadena = "select  * from tabcuo where cuo_codcia = '" & LK_CODCIA & "' and cuo_periodo = '" & txtper.Text & "' and cuo_codart = " & artloc_llave!art_key
   Set rs_regcuotas = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If rs_regcuotas.EOF Then
      rs_regcuotas.AddNew
   Else
      rs_regcuotas.Edit
   End If
   rs_regcuotas!CUO_CODCIA = LK_CODCIA
   rs_regcuotas!CUO_PERIODO = txtper.Text
   rs_regcuotas!cuo_codart = Val(artloc_llave!art_key)
   pub_cadena = "SELECT   top 1 pre_equiv FROM PRECIOS WHERE PRE_CODCIA = '" & LK_CODCIA & "' AND PRE_CODART = " & artloc_llave!art_key & " order by pre_equiv desc"
   Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If X.EOF Then
   Else
     rs_regcuotas!CUO_CANTIDAD = Val(txtum.Text) * Val(X!PRE_EQUIV)
   End If
   rs_regcuotas.Update
   
   ' SOLO COMPA�IA ACTUAL
    pub_cadena = "DELETE FROM FVENTAS WHERE FV_CODCIA = '" & LK_CODCIA & "' AND FV_TIPO = 2 AND FV_CODIGO = " & (artloc_llave!art_key)
    CN.Execute pub_cadena
    pub_cadena = "SELECT * FROM FVENTAS WHERE FV_CODCIA = '" & LK_CODCIA & "' AND FV_TIPO = 2 AND FV_CODIGO = " & (artloc_llave!art_key)
    Set psfv = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    For fila = 0 To lfvta.ListCount - 1
        If lfvta.Selected(fila) Then
        psfv.AddNew
        psfv!FV_CODCIA = LK_CODCIA
        psfv!FV_TIPO = 2
        psfv!FV_CODIGO = (artloc_llave!art_key)
        psfv!FV_KEYFV = Val(Right(lfvta.List(fila), 10))
        psfv.Update
        End If
    Next fila
   
End If

If Left(cmdModificar.Caption, 2) = "&G" Then
 If frmARTI.art_situacion.Value = 0 And flag_art_situacion = 1 Then
   For int_i = 0 To lstcias.ListCount - 1
     If Left(lstcias.List(int_i), 2) = pu_codcia Then
       If lstcias.Selected(int_i) = True Then
          artloc_llave!art_situacion = frmARTI.art_situacion.Value
       End If
     End If
   Next int_i
 Else
   artloc_llave!art_situacion = frmARTI.art_situacion.Value
 End If
End If
artloc_llave!art_numero = Val(Right(art_numero.Text, 3))
artloc_llave!ART_CUENTA_CONTAB_C = txtClase
'artloc_llave!ART_CUENTA_CONTAB = 0 ' txtLab

artloc_llave!art_tipo = loc_tipo
artloc_llave!ART_CODUSU_ORIG = Trim(cmbUsu.Text)
artloc_llave!art_familia = Val(Right(art_familia.Text, 5))
artloc_llave!art_grupop = Val(Right(art_grupop.Text, 10))
artloc_llave!art_subfam = Val(Right(art_subfam.Text, 5))
artloc_llave!ART_CALIDAD = WCALIDAD
artloc_llave!art_subgru = Val(Right(art_grupo.Text, 5))
artloc_llave!art_linea = Val(Right(art_linea.Text, 5))
artloc_llave!art_plancha = Val(Right(art_plancha.Text, 6))
artloc_llave!art_codclie = Val(Right(art_codpro.Text, 6))
artloc_llave!art_nombre = wnombre
artloc_llave!art_codcia = pu_codcia 'PUB_CODCIA
artloc_llave!ART_DECIMALES = Val(frmARTI.decimales.Text)
artloc_llave!ART_MONEDA = "S" 'Trim(frmARTI.DS.Text)
artloc_llave!ART_ORDEN = frmARTI.chebof.Value
artloc_llave!ART_POR_IGV = Nulo_Valor0(frmARTI.txtcospro.Text)
artloc_llave!ART_CODART2 = WCODART2
artloc_llave!ART_EX_IGV = ""
artloc_llave!ART_CP = ""
'artloc_llave!ART_COSPRO = Nulo_Valor0(tcospro.Text)
If exigv.Value = 1 Then
  artloc_llave!ART_EX_IGV = "A"
End If

artloc_llave!art_flag_cambio = " "
If checambio.Value = 1 Then
   artloc_llave!art_flag_cambio = "A"
End If
artloc_llave!art_fechahora = Format(Now, "dd/mm/yy") & Chr(13) & Format(Now, "hh:mm AMPM") & Chr(13) & LK_CODUSU
' ICA
If Val(artloc_llave!ART_POR_IGV) <> 0 And exigv.Value = 0 Then
   artloc_llave!ART_EX_IGV = "E"
End If
artloc_llave!art_flag_stock = ""
If cheservi(0).Value = True Then
  artloc_llave!art_flag_stock = "M"
ElseIf cheservi(1).Value = True Then
  artloc_llave!art_flag_stock = "P"
ElseIf cheservi(2).Value = True Then
  artloc_llave!art_flag_stock = "S"
End If
'If LK_EMP = "HER" Then
 artloc_llave!ART_POR1 = Val(txtpor1.Text)
 artloc_llave!ART_POR2 = Val(txtpor2.Text)
 artloc_llave!ART_POR3 = Val(txtpor3.Text)
 artloc_llave!ART_POR4 = Val(txtpor4.Text)
 artloc_llave!ART_POR5 = Val(txtpor5.Text)
 artloc_llave!ART_POR11 = Val(txtpor11.Text)
 artloc_llave!ART_POR22 = Val(txtpor22.Text)
 artloc_llave!ART_POR33 = Val(txtpor33.Text)
 artloc_llave!ART_POR44 = Val(txtpor44.Text)
 artloc_llave!ART_POR55 = Val(txtpor55.Text)
 artloc_llave!ART_MARGEN = Val(txtmargen.Text)
 artloc_llave!ART_POR6 = Val(cheoferta.Value)  ' FLAG 1 = ESTA EN OFERTA
 artloc_llave!art_estado = chesub.Value

'End If
artloc_llave.Update
pu_codcia = pu_codcia 'PUB_CODCIA
PUB_CODART = WORIGINAL
SQ_OPER = 1
LEER_ARM_LLAVE
'If Not arm_llave.EOF Then
' arm_llave.Edit
' arm_llave!arm_cospro = Nulo_Valor0(tcospro.Text)
' arm_llave.Update
'End If

 
ws_codcia = pu_codcia 'LK_CODCIA
If Left(cmdModificar.Caption, 2) = "&G" Then
 GoSub IR_POR_CIA
 Exit Sub
End If
If LK_EMP_PTO = "A" Then
  If Trim(GEN!gen_ART_CIAS) <> "" Then
     xcuenta = 1
     For fila = 1 To 30
         ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
         If Trim(ws_codcia) = "" Then Exit For
         GoSub IR_POR_CIA
         xcuenta = xcuenta + 2
     Next fila
  Else
    GoSub IR_POR_CIA
  End If
Else
  GoSub IR_POR_CIA
End If

Exit Sub


IR_POR_CIA: ' Actualiza Cias o Cia Actual

pu_codcia = ws_codcia
PUB_CODART = WORIGINAL
SQ_OPER = 2
LEER_PRE_LLAVE
If VAR_NEWCAL = 1 Then
 GoTo IR_GRABA2
End If
If Left(cmdModificar.Caption, 2) = "&G" Then
    fila = 0
    Flag_Inicial = "A"
    Do Until pre_mayor.EOF
       fila = fila + 1
      If fila >= grid_unid.Rows Then
        pre_mayor.Delete
        GoTo OTRO
      End If
      If pre_mayor!PRE_SECUENCIA <> Val(grid_unid.TextMatrix(fila, 17)) Then
        pre_mayor.Delete
        fila = fila - 1
        GoTo OTRO
      End If
       pre_mayor.Edit
       pre_mayor!PRE_SECUENCIA = Val(grid_unid.TextMatrix(fila, 17))
       pre_mayor!PRE_UNIDAD = Left(grid_unid.TextMatrix(fila, 0), 15)
       pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(fila, 1))
       If LK_CODCIA = "01" Or LK_CODCIA = "04" Then ' OTRA EMPRESA
       Else
        If LK_CODCIA = "05" Or LK_CODCIA = "09" Or LK_CODCIA = "29" Or LK_CODCIA = "50" Then
         If ws_codcia = LK_CODCIA Then
         Else
         GoTo pasa_otro_pre
         End If
        Else
         If ws_codcia = "05" Or ws_codcia = "09" Or ws_codcia = "29" Or ws_codcia = "50" Then GoTo pasa_otro_pre
        End If
       End If
       
       pre_mayor!PRE_PRE11 = Val(grid_unid.TextMatrix(fila, 18))
       pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(fila, 19))
       pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(fila, 20))
       pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(fila, 21))
       pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(fila, 22))
       pre_mayor!PRE_PRE66 = Val(grid_unid.TextMatrix(fila, 31))
       'If ws_codcia = LK_CODCIA Then
        pre_mayor!PRE_POR1 = Val(grid_unid.TextMatrix(fila, 4))
        pre_mayor!PRE_POR2 = Val(grid_unid.TextMatrix(fila, 6))
        pre_mayor!PRE_POR3 = Val(grid_unid.TextMatrix(fila, 8))
        pre_mayor!PRE_POR4 = Val(grid_unid.TextMatrix(fila, 10))
        pre_mayor!PRE_POR5 = Val(grid_unid.TextMatrix(fila, 12))
        pre_mayor!PRE_POR6 = Val(grid_unid.TextMatrix(fila, 14))
       'End If
       pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(fila, 23))
       pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(fila, 24))
       pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(fila, 25))
       pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(fila, 26))
       pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(fila, 27))
       pre_mayor!PRE_PRE6 = Val(grid_unid.TextMatrix(fila, 32))
       
       pre_mayor!PRE_PREC1 = Val(grid_unid.TextMatrix(fila, 46))
       pre_mayor!PRE_PREC11 = Val(grid_unid.TextMatrix(fila, 47))
       
       pre_mayor!PRE_PREC2 = Val(grid_unid.TextMatrix(fila, 49))
       pre_mayor!PRE_PREC22 = Val(grid_unid.TextMatrix(fila, 50))
       
       pre_mayor!PRE_PRECHI1 = Val(grid_unid.TextMatrix(fila, 52))
       pre_mayor!PRE_PRECHI11 = Val(grid_unid.TextMatrix(fila, 53))
       pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(fila, 28))
       pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(fila, 30))
pasa_otro_pre:
       pre_mayor!pre_FLAG_UNIDAD = grid_unid.TextMatrix(fila, 16)
       pre_mayor.Update
OTRO:
       pre_mayor.MoveNext
    Loop
    If fila > grid_unid.Rows - 1 Then
    
    ElseIf fila <> grid_unid.Rows - 1 Then
     VARFILA = fila + 1
      GoTo AGREGA
    End If

Else
IR_GRABA2:
    VARFILA = 1
AGREGA:
    fila = 0
    Flag_Inicial = "A"
    For fila = VARFILA To grid_unid.Rows - 1
       pre_mayor.AddNew
       pre_mayor!PRE_CODCIA = ws_codcia
       pre_mayor!PRE_CODART = WORIGINAL
       pre_mayor!PRE_SECUENCIA = fila - 1
       pre_mayor!PRE_UNIDAD = Left(grid_unid.TextMatrix(fila, 0), 15)
       pre_mayor!PRE_EQUIV = Val(grid_unid.TextMatrix(fila, 1))
       
       pre_mayor!PRE_PRE11 = Val(grid_unid.TextMatrix(fila, 18))
       pre_mayor!PRE_PRE22 = Val(grid_unid.TextMatrix(fila, 19))
       pre_mayor!PRE_PRE33 = Val(grid_unid.TextMatrix(fila, 20))
       pre_mayor!PRE_PRE44 = Val(grid_unid.TextMatrix(fila, 21))
       pre_mayor!PRE_PRE55 = Val(grid_unid.TextMatrix(fila, 22))
       
       pre_mayor!PRE_POR1 = Val(grid_unid.TextMatrix(fila, 4))
       pre_mayor!PRE_POR2 = Val(grid_unid.TextMatrix(fila, 6))
       pre_mayor!PRE_POR3 = Val(grid_unid.TextMatrix(fila, 8))
       pre_mayor!PRE_POR4 = Val(grid_unid.TextMatrix(fila, 10))
       pre_mayor!PRE_POR5 = Val(grid_unid.TextMatrix(fila, 12))
       
       pre_mayor!PRE_PRE1 = Val(grid_unid.TextMatrix(fila, 23))
       pre_mayor!PRE_PRE2 = Val(grid_unid.TextMatrix(fila, 24))
       pre_mayor!PRE_PRE3 = Val(grid_unid.TextMatrix(fila, 25))
       pre_mayor!PRE_PRE4 = Val(grid_unid.TextMatrix(fila, 26))
       pre_mayor!PRE_PRE5 = Val(grid_unid.TextMatrix(fila, 27))
       pre_mayor!pre_PESO = Val(grid_unid.TextMatrix(fila, 28))
       pre_mayor!PRE_LITRO = Val(grid_unid.TextMatrix(fila, 30))
       pre_mayor!pre_FLAG_UNIDAD = grid_unid.TextMatrix(fila, 16)
       
       pre_mayor!PRE_PREC1 = Val(grid_unid.TextMatrix(fila, 46))
       pre_mayor!PRE_PREC11 = Val(grid_unid.TextMatrix(fila, 47))
       
       pre_mayor!PRE_PREC2 = Val(grid_unid.TextMatrix(fila, 49))
       pre_mayor!PRE_PREC22 = Val(grid_unid.TextMatrix(fila, 50))
       
       pre_mayor!PRE_PRECHI1 = Val(grid_unid.TextMatrix(fila, 52))
       pre_mayor!PRE_PRECHI11 = Val(grid_unid.TextMatrix(fila, 53))
   
       pre_mayor.Update
    Next fila
End If
If Left(cmdAgregar.Caption, 2) = "&G" Then
    pu_codcia = ws_codcia
    PUB_CODART = PUB_KEY
    SQ_OPER = 1
    LEER_ARM_LLAVE
    If arm_llave.EOF Then
     Screen.MousePointer = 0
     'If LK_EMP = "CAM" Then
     '   wSTOCK = InputBox("STOCK :", "", "")
     'End If
     arm_llave.AddNew
     arm_llave!ARM_CODART = PUB_KEY
     arm_llave!ARM_CODCIA = ws_codcia
     arm_llave!arm_stock = 0
     arm_llave!ARM_INGRESOS = 0
     arm_llave!ARM_SALIDAS = 0
     arm_llave!ARM_COSPRO = 0
     arm_llave!arm_stock2 = 0
     arm_llave!ARM_COSTO_ULT = 0
     arm_llave!ARM_saldo_s = 0
     arm_llave!arm_saldo_s2 = 1
     arm_llave!ARM_Saldo_n = 0
     arm_llave!arm_saldo_n2 = 0
     arm_llave!ARM_FECHA_ULT = #1/1/1900#
     arm_llave.Update
     MENSAJE_ARTI "Articulo Nuevo en Compa�ia . . ."
    Else
      MsgBox "Codigo Existe en tabla: Articulo verificar ...", 48, Pub_Titulo
    End If
End If
Return
End Sub

Public Function GENERA_CODI() As Double
Dim NUMCAD, FIJO As String
Dim DIGI As String * 2
Dim I, VINT1, VINT2, VINT3, VINT4 As Double
Dim VSTR1, VSTR2, VSTR3, VSTR4 As String
Dim VFIJO As Double
Dim VVARI As Integer
Dim STRpub_cadena As String
Dim INTpub_cadena As Double

PUB_KEY = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
LEER_ART_LLAVE

If art_mayor.EOF Then
    NUMCAD = "1"
Else
    art_mayor.MoveLast
    NUMCAD = art_mayor!art_key
End If

VINT2 = 0
NUMCAD = Trim(NUMCAD)
VINT1 = Len(NUMCAD)
If NUMCAD = "1" Or NUMCAD = "2" Or NUMCAD = "0" Then
  VINT2 = 1
  VINT1 = 2
End If
If VINT1 > 1 Then
    VSTR4 = Val(Mid(NUMCAD, 1, VINT1 - 2)) + 1
End If

For I = 1 To VINT1 - 2
   VSTR1 = Mid(VSTR4, I, 1)
   VINT2 = VINT2 + Val(VSTR1)
Next I
VINT3 = VINT2 * 7

VSTR3 = Right(CStr(VINT3), 2)
If Len(VSTR3) = 1 Then
  VSTR3 = "0" & VSTR3
End If
FIJO = VSTR4
STRpub_cadena = FIJO & VSTR3
INTpub_cadena = Val(STRpub_cadena)

GENERA_CODI = INTpub_cadena

End Function

Public Function CONSIS_ARTI() As Boolean
Dim wval As Currency
Dim WRES As Currency
Dim WSEND As String
If loc_tipo = "V" Then
    If Trim(art_familia.Text) = "" Then
        CONSIS_ARTI = False
        MsgBox "Falta seleccionar algun dato...", 48, Pub_Titulo
        art_familia.SetFocus
        SendKeys "%{UP}"
        GoTo ESCAPA
    End If
    If LK_FLAG_ORIGINAL <> "A" Then
     If txt_alterno.Text = "" Then
        CONSIS_ARTI = False
        MsgBox "Favor de Ingresar el Codigo Alterno...", 48, Pub_Titulo
        txt_alterno.SetFocus
        GoTo ESCAPA
     End If
    End If
    If txtnombre.Text = "" Then
        CONSIS_ARTI = False
        MsgBox "Falta Ingresar el NOMBRE...", 48, Pub_Titulo
        Azul txtnombre, txtnombre
        GoTo ESCAPA
    End If
    If Not IsNumeric(frmARTI.txtMin.Text) And Trim(txtMin.Text) <> "" Or Val(txtMin.Text) > 999999999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido Stock Minimo ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtMin, txtMin
        GoTo ESCAPA
    End If
    
    If Not IsNumeric(frmARTI.txtMax.Text) And Trim(txtMax.Text) <> "" Or Val(txtMax.Text) > 999999999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido Stock Maximo ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtMax, txtMax
        GoTo ESCAPA
    End If
    
    If Val(txtmargen.Text) <> 0 Then
       If grid_unid.Rows = 3 Then
          wval = grid_unid.TextMatrix(2, 1)
       ElseIf grid_unid.Rows = 2 Then
          wval = grid_unid.TextMatrix(1, 1)
       End If
        WRES = Int(wval / Val(txtmargen.Text))
        If (WRES - Val(Format(wval / Val(txtmargen.Text), "0.00"))) <> 0 Then
          CONSIS_ARTI = False
          MsgBox "El Factor para la Unidades Minimas no corresponde a la unidad Mayor ", 48, Pub_Titulo
          Azul txtmargen, txtmargen
          GoTo ESCAPA
        End If
    End If

    
    If Not IsNumeric(txtpor1) And Trim(txtpor1.Text) <> "" Or Val(txtpor1.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 1 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor1, txtpor1
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor2) And Trim(txtpor2.Text) <> "" Or Val(txtpor2.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 2 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor2, txtpor2
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor3) And Trim(txtpor3.Text) <> "" Or Val(txtpor3.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 3 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor3, txtpor3
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor4) And Trim(txtpor4.Text) <> "" Or Val(txtpor4.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 4 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor4, txtpor4
        GoTo ESCAPA
    ElseIf Not IsNumeric(txtpor5) And Trim(txtpor5.Text) <> "" Or Val(txtpor5.Text) > 999.99 Then
        CONSIS_ARTI = False
        MsgBox "Dato Invalido % p' 5 ", 48, Pub_Titulo
        frmARTI.SSTab1.tab = 1
        Azul txtpor5, txtpor5
        GoTo ESCAPA
    End If
End If

CONSIS_ARTI = True

Exit Function

ESCAPA:
 'msgbox WSEND, 48, pub_titulo

End Function

Public Sub MENSAJE_ARTI(TEXTO As String)
  LblMensaje.Caption = TEXTO
  Parpadea.Enabled = True
End Sub
Public Sub SOLO_PORCEBTAJE(Optional tecla)
'CONVIERTE TODA A MAYUSCULAS LETRAS
Dim car As String, Longt As Integer
car = Chr$(tecla)
car = UCase$(Chr$(tecla))
tecla = Asc(car)
If car < "0" Or car > "9" Then
    If tecla <> 8 And tecla <> 13 And tecla <> 32 And car <> "." Then
        tecla = 0
        
    End If
End If
End Sub

Private Sub txt_key_GotFocus()
 If ListView1.Visible Then
  ListView1.Visible = False
 End If
 txt_key.Text = ""
 frmARTI.F14.Visible = False
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem

If KeyAscii = 27 Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
  pu_codclie = Val(txt_key.Text)
On Error GoTo 0
If Len(txt_key.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   LOC_OPER = 1
   PUB_CODCIA = LK_CODCIA
   On Error GoTo ERROR_CODIGO
    PUB_KEY = pu_codclie
    LEER_LOC
   On Error GoTo 0
   If artloc_llave.EOF Then
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_key, txt_key
     GoTo fin
   Else
     If pu_codclie = 1 Then
       MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
       Azul txt_key, txt_key
       GoTo fin
     End If
     LLENA_ARTI 1
     BLOQUEA_TEXT frmARTI.txt_key
     frmARTI.cmdModificar.SetFocus
     BLOQUEA_TEXT txtnombre
     cmdCancelar.Enabled = True
   End If
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(VALOR, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   LLENA_ARTI 0
   BLOQUEA_TEXT frmARTI.txt_key
   frmARTI.cmdModificar.SetFocus
   BLOQUEA_TEXT txtnombre
   cmdCancelar.Enabled = True
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_key, txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If Len(txt_key.Text) = 0 Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    If txt_key.Text = "" Then txt_key.Text = " "
    var = Asc(txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    numarchi = 0
    If cheotros.Value = 1 Then
     archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    Else
     archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV, ART_SITUACION FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_SITUACION <> '1' AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
    End If
    PROC_LISVIEW ListView1, 1000
    loc_key = 0
    If ListView1.Visible Then
    loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As Object ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Public Function CONSIS_UNIDAD() As Boolean
Dim I As Integer
Dim QUIEN As String
QUIEN = 0
For fila = 1 To grid_unid.Rows - 1
   If Trim(grid_unid.TextMatrix(fila, 0)) <> "" And Val(grid_unid.TextMatrix(fila, 1)) <> 0 Then
   Else
    MsgBox "La ultima unidad ingresada no procede ...", 48, Pub_Titulo
    CONSIS_UNIDAD = False
    Exit Function
   End If
   If Val(grid_unid.TextMatrix(fila, 4)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 6)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 8)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 10)) > 999.99 Or Val(grid_unid.TextMatrix(fila, 12)) > 999.99 Then
'     MsgBox " El Procentaje debe ser menor o igual a  999.99 .", 48, Pub_Titulo
'     CONSIS_UNIDAD = False
'     Exit Function
   End If
   If Val(grid_unid.TextMatrix(fila, 4)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 6)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 8)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 10)) < -999.99 Or Val(grid_unid.TextMatrix(fila, 12)) < -999.99 Then
'     MsgBox " El Procentaje debe ser menor o igual a  -999.99 .", 48, Pub_Titulo
'     CONSIS_UNIDAD = False
'     Exit Function
   End If
Next fila
CONSIS_UNIDAD = True
End Function


' sigue todo para el grid
Public Sub CABEZA_UNID()
Flag_Inicial = "A"
grid_unid.Cols = 54
grid_unid.Rows = 2
grid_unid.FixedCols = 0
grid_unid.FixedRows = 1
grid_unid.ColWidth(0) = 800 ' unidad
grid_unid.ColWidth(1) = 700  ' equivalencia
grid_unid.ColWidth(2) = 1  'c. Repos.
grid_unid.ColWidth(3) = 800  ' cos. base
grid_unid.ColWidth(4) = 600   ' % 1
grid_unid.ColWidth(5) = 750   ' p 1
grid_unid.ColWidth(6) = 600   ' % 2
grid_unid.ColWidth(7) = 750   ' p 2
grid_unid.ColWidth(8) = 600   ' % 3
grid_unid.ColWidth(9) = 750   ' p 3
grid_unid.ColWidth(10) = 600   ' % 4
grid_unid.ColWidth(11) = 750   ' p 4
grid_unid.ColWidth(12) = 600   ' % 5
grid_unid.ColWidth(13) = 750   ' p 5
grid_unid.ColWidth(14) = 600   ' % 6
grid_unid.ColWidth(15) = 750   ' p 6

grid_unid.ColWidth(16) = 1   ' FLAG
grid_unid.ColWidth(17) = 1   ' SECUENCIA
grid_unid.ColWidth(18) = 1   ' Guarda p.digitado 11(S)
grid_unid.ColWidth(19) = 1   ' Guarda p.digitado 22(S)
grid_unid.ColWidth(20) = 1   ' Guarda p.digitado 33(S)
grid_unid.ColWidth(21) = 1   ' Guarda p.digitado 44(S)
grid_unid.ColWidth(22) = 1   ' Guarda p.digitado 55(S)
grid_unid.ColWidth(23) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(24) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(25) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(26) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(27) = 1   ' Guarda p.digitado 55(D)
grid_unid.ColWidth(28) = 1   ' Guarda el peso x unidad
grid_unid.ColWidth(29) = 1   ' Guarda el costo Rep. soles
grid_unid.ColWidth(30) = 1   ' Guarda el costo Rep. soles
grid_unid.ColWidth(31) = 1   ' Guarda precio dolares p6
grid_unid.ColWidth(32) = 1   ' Guarda precio soles p 6

grid_unid.ColWidth(33) = 600   ' Lista %  7
grid_unid.ColWidth(34) = 750   ' Lista   7
grid_unid.TextMatrix(0, 33) = "( % )  "
grid_unid.TextMatrix(0, 34) = "Prec. 7 "

grid_unid.ColWidth(35) = 1   ' Guarda precio dolares p7
grid_unid.ColWidth(36) = 1   ' Guarda precio soles p7

grid_unid.ColWidth(37) = 600   ' Lista %  8
grid_unid.ColWidth(38) = 750   ' Lista   8
grid_unid.TextMatrix(0, 37) = "( % )  "
grid_unid.TextMatrix(0, 38) = "Prec. 8 "

grid_unid.ColWidth(39) = 1   ' Guarda precio dolares p8
grid_unid.ColWidth(40) = 1   ' Guarda precio soles p8

grid_unid.ColWidth(41) = 600   ' Lista %  9
grid_unid.ColWidth(42) = 750   ' Lista   9
grid_unid.TextMatrix(0, 41) = "( % )  "
grid_unid.TextMatrix(0, 42) = "Prec. 9 "
grid_unid.ColWidth(43) = 1   ' Guarda precio dolares p9
grid_unid.ColWidth(44) = 1   ' Guarda precio soles p9

grid_unid.ColWidth(45) = 600   ' Lista %  9
grid_unid.ColWidth(46) = 750   ' Lista   9
grid_unid.ColWidth(47) = 0   ' Lista   9
grid_unid.TextMatrix(0, 45) = "( % )  "
grid_unid.TextMatrix(0, 46) = "P.11. "

grid_unid.ColWidth(48) = 600   ' Lista %  9
grid_unid.ColWidth(49) = 750   ' Lista   9
grid_unid.ColWidth(50) = 0   ' Lista   9
grid_unid.TextMatrix(0, 48) = "( % )  "
grid_unid.TextMatrix(0, 49) = "P.12. "

grid_unid.ColWidth(51) = 600   ' Lista 10 DE CHICLAYO
grid_unid.ColWidth(52) = 750   ' Lista 10 DE CHICLAYO
grid_unid.ColWidth(53) = 0   ' Lista 10 DE CHICLAYO
grid_unid.TextMatrix(0, 51) = "( % )  "
grid_unid.TextMatrix(0, 52) = "Prec.10"

grid_unid.TextMatrix(0, 0) = "Unidades"
grid_unid.TextMatrix(0, 1) = "Equiv."
grid_unid.TextMatrix(0, 2) = "COSTO."
grid_unid.TextMatrix(0, 3) = "Rep."
grid_unid.TextMatrix(0, 4) = "( % )  "
grid_unid.TextMatrix(0, 5) = " Prec. 1"
grid_unid.TextMatrix(0, 6) = "( % )  "
grid_unid.TextMatrix(0, 7) = "Prec. 2"
grid_unid.TextMatrix(0, 8) = "( % )  "
grid_unid.TextMatrix(0, 9) = "Prec. 3"
grid_unid.TextMatrix(0, 10) = "( % )  "
grid_unid.TextMatrix(0, 11) = "Prec. 4"
grid_unid.TextMatrix(0, 12) = "( % )  "
grid_unid.TextMatrix(0, 13) = "Prec. 5"

grid_unid.TextMatrix(0, 14) = "( % )  "
grid_unid.TextMatrix(0, 15) = "Prec. 6"

Flag_Inicial = ""
End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
If wsGrid.CellWidth < 0 Then
 Exit Sub
End If
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Visible = False

End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
If wsGrid.CellWidth < 0 Then
 Exit Sub
End If
If wsGrid.COL = 0 Then
 wsTexto.MaxLength = 20
Else
 wsTexto.MaxLength = 9
End If

'wsGrid.CellFontBold = True

wsGrid.CellBackColor = QBColor(7)
'wsGrid.CellForeColor = QBColor(15)

wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.COL = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.COL = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.COL = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.COL = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.COL = Bloq5 Then
    Flag_Bloq = "A"
  End If
End If

End Sub
Public Sub ElGrid_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode)
Flag_F2 = ""
If Flag_Bloq = "A" Then
  wsKeyCode = 0
  Exit Sub
End If

'If wsKeyCode <> 113 Then
 Exit Sub
'End If
If wsTexto.Visible = False Then
  Flag_F2 = "A"
  ElGrid_DblClick wsGrid, wsTexto
End If
End Sub
Public Sub ElGrid_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii, Optional SaltaCol)
If wsKeyAscii = 27 Then
 Exit Sub
End If
If wsKeyAscii = 9 Or wsKeyAscii = 13 Then
  If Not IsMissing(SaltaCol) Then
    If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.COL = wsGrid.FixedCols
       Exit Sub
    End If
    
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    If wsGrid.COL = 1 Then
     wsGrid.COL = wsGrid.COL + 2
    ElseIf wsGrid.COL > 13 Then
     wsGrid.COL = wsGrid.COL + 1
    End If
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade
'wsTexto.FontBold = True

' wsTexto.ForeColor = QBColor(1)
wsTexto.Text = ""
wsTexto.Visible = True
cade = UCase(Chr(wsKeyAscii))
'wsTexto.text = cade
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
Flag_Change = "A"
'cade = Chr(wsKeyAscii)
SendKeys cade, True
wsTexto.SelStart = Len(wsTexto)

End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 
 'wsTexto.ForeColor = QBColor(12)
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
'If Left(Trim(wsGrid.text), 1) = "-" Then
 'wsGrid.CellForeColor = QBColor(12)
 'wsGrid.CellBackColor = QBColor(15)
'Else
 wsGrid.CellBackColor = QBColor(15)
'wsGrid.CellForeColor = QBColor(0)
'End If
'wsGrid.CellFontBold = False
End Sub
Private Sub ElGrid_DblClick(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Bloq = "A" Then
  Exit Sub
End If
wsTexto.FontBold = True
'wsTexto.ForeColor = QBColor(12)
wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
End Sub
Private Sub ElGrid_GotFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Private Sub TEXTO_LosFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Public Sub TEXTO_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode As Integer, Optional SaltaCol)
If wsKeyCode = 40 Or wsKeyCode = 37 Or wsKeyCode = 39 Or wsKeyCode = 38 Then
 If Flag_F2 = "A" Then
   Exit Sub
 End If
 If Flag_Consis = "A" Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
   Beep
   Exit Sub
 End If
 Flag_Change = ""
 If wsGrid.COL <> 0 Then
  If Trim(wsTexto.Text) = "." Or Trim(wsTexto.Text) = "" Then
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = "0.000"
  Else
   If wsGrid.COL = 3 Then
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.00")
   Else
   If wsGrid.COL = 6 Or wsGrid.COL = 8 Or wsGrid.COL = 10 Or wsGrid.COL = 12 Or wsGrid.COL = 14 Then
     wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.00") ' CAMBIO DE DECIMANES DE PRECIOS
    Else
     wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000") ' CAMBIO DE DECIMANES DE PRECIOS
    End If
   End If
   Select Case wsGrid.COL
     Case 5
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 18) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 23) = Format(wsTexto.Text, "0.000")
       End If
     Case 7
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 19) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 24) = Format(wsTexto.Text, "0.000")
       End If
     Case 9
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 20) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 25) = Format(wsTexto.Text, "0.000")
       End If
     Case 11
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 21) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 26) = Format(wsTexto.Text, "0.000")
       End If
     Case 13
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 22) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 27) = Format(wsTexto.Text, "0.000")
       End If
     Case 15
       If frmARTI.cmddolares.Tag = "D" Then
         wsGrid.TextMatrix(wsGrid.Row, 31) = Format(wsTexto.Text, "0.000")
       Else
         wsGrid.TextMatrix(wsGrid.Row, 32) = Format(wsTexto.Text, "0.000")
       End If
   End Select
  End If
 Else
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = wsTexto.Text
 End If
 Flag_Bloq = ""
 wsTexto.Visible = False
 If wsKeyCode = 40 Then ' ABAJO
  If wsGrid.Row <> wsGrid.Rows - 1 Then
     wsGrid.Row = wsGrid.Row + 1
  End If
 End If
 If wsKeyCode = 38 Then ' arriba
  If wsGrid.Row <> wsGrid.FixedRows Then
     wsGrid.Row = wsGrid.Row - 1
  End If
 End If
 If wsKeyCode = 37 Then ' isquierda
  If wsGrid.COL <> wsGrid.FixedCols Then
     If wsGrid.COL = 3 Then
        wsGrid.COL = wsGrid.COL - 2
     Else
        wsGrid.COL = wsGrid.COL - 1
     End If
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
     If wsGrid.COL = SaltaCol Then
        If wsGrid.Row <> wsGrid.Rows - 1 Then
          wsGrid.Row = wsGrid.Row + 1
          wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
        ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
         If Trim(wsGrid.TextMatrix(wsGrid.Row, 0)) <> "" And Val(wsGrid.TextMatrix(wsGrid.Row, 1)) <> 0 And Val(wsGrid.TextMatrix(wsGrid.Row, 3)) <> 0 Then
          ' wsGrid.Rows = wsGrid.Rows + 1
           'wsGrid.Row = wsGrid.Row + 1
         '  wsGrid.Col = wsGrid.FixedCols
         '  GoTo wsfinal
         Else
           wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
         End If
        End If
     ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
        wsGrid.COL = wsGrid.FixedCols
        GoTo wsfinal
     End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
     If wsGrid.COL = 1 Then
         wsGrid.COL = wsGrid.COL + 2
     ElseIf wsGrid.COL >= 13 Then
         wsGrid.COL = wsGrid.FixedCols
         GoTo wsfinal
     Else
         If wsGrid.COL = 5 And LK_EMP = "3AA" Then
            wsGrid.COL = 13
         Else
           wsGrid.COL = wsGrid.COL + 1
         End If
     End If
  End If
 End If
wsfinal:
 'wsTexto.FontBold = False
 'wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
 wsGrid.SetFocus
End If
'Exit Sub

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10, Optional ConsisCol11, Optional ConsisVal11, Optional ConsisCol12, Optional ConsisVal12)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  Flag_Change = "A"
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.COL = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol4) And Not IsMissing(ConsisVal4) Then
  If wsGrid.COL = ConsisCol4 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal4, ConsisCol4
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol5) And Not IsMissing(ConsisVal5) Then
  If wsGrid.COL = ConsisCol5 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal5, ConsisCol5
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol6) And Not IsMissing(ConsisVal6) Then
  If wsGrid.COL = ConsisCol6 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal6, ConsisCol6
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol7) And Not IsMissing(ConsisVal7) Then
  If wsGrid.COL = ConsisCol7 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal7, ConsisCol7
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol8) And Not IsMissing(ConsisVal8) Then
  If wsGrid.COL = ConsisCol8 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal8, ConsisCol8
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol9) And Not IsMissing(ConsisVal9) Then
  If wsGrid.COL = ConsisCol9 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal9, ConsisCol9
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol10) And Not IsMissing(ConsisVal10) Then
  If wsGrid.COL = ConsisCol10 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal10, ConsisCol10
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol11) And Not IsMissing(ConsisVal11) Then
  If wsGrid.COL = ConsisCol11 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal11, ConsisCol11
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol12) And Not IsMissing(ConsisVal12) Then
  If wsGrid.COL = ConsisCol12 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal12, ConsisCol12
   Exit Sub
  End If
End If

End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static VALOR
  Dim car As String
  Flag_Consis = ""
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
      If car < "0" Or car > "9" Then
        If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "-" Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub


Public Sub CALCULAR(wsCosto As Currency)
Dim VALOR As Currency
Dim tC As Integer
Flag_Inicial = "A"
tC = grid_unid.COL
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 4)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 5
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 5
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 5) = Format(VALOR, "0.0000") ' PRECIO 1
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 6)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 7
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 7
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 7) = Format(VALOR, "0.0000") ' PRECIO 2
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 8)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 9
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 9
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 9) = Format(VALOR, "0.0000") ' PRECIO 3
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 10)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 11
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 11
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 11) = Format(VALOR, "0.0000") ' PRECIO 4
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 12)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 13) = Format(VALOR, "0.0000") ' PRECIO 5
VALOR = wsCosto * (1 + (Val(grid_unid.TextMatrix(grid_unid.Row, 14)) / 100))
If VALOR < 0 Then
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = 13
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.TextMatrix(grid_unid.Row, 15) = Format(VALOR, "0.0000") ' PRECIO 6


grid_unid.COL = tC
Flag_Inicial = ""

End Sub

Public Sub CALCULAR_OTRO(wsCosto As Currency)
Dim VALOR As Currency
Dim tC As Integer
Flag_Inicial = "A"
tC = grid_unid.COL
If wsCosto = 0 Then
  GoTo CERO
End If
VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 5) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 200
If Val(grid_unid.TextMatrix(grid_unid.Row, 5)) = 0 Then
'   grid_unid.TextMatrix(grid_unid.Row, 4) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 4
  'grid_unid.CellForeColor = QBColor(12)
Else
   grid_unid.TextMatrix(grid_unid.Row, 4) = Format(VALOR, "0.0000") ' PRECIO 1
  'grid_unid.Col = 4
  'grid_unid.CellForeColor = QBColor(0)
End If
VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 7) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 7)) = 0 Then
 '  grid_unid.TextMatrix(grid_unid.Row, 6) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 6
  'grid_unid.CellForeColor = QBColor(12)
Else
   grid_unid.TextMatrix(grid_unid.Row, 6) = Format(VALOR, "0.0000") ' PRECIO 1
  'grid_unid.Col = 6
  'grid_unid.CellForeColor = QBColor(0)
End If
VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 9) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 9)) = 0 Then
  ' grid_unid.TextMatrix(grid_unid.Row, 8) = Format(0, "0.00") ' PRECIO 1
  'grid_unid.Col = 8
  'grid_unid.CellForeColor = QBColor(12)
Else
  grid_unid.TextMatrix(grid_unid.Row, 8) = Format(VALOR, "0.0000") ' PRECIO 1
  'grid_unid.Col = 8
  'grid_unid.CellForeColor = QBColor(0)
End If
VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 11) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 11)) = 0 Then
  'grid_unid.Col = 10
  'grid_unid.CellForeColor = QBColor(12)
  'grid_unid.TextMatrix(grid_unid.Row, 10) = Format(0, "0.00") ' PRECIO 1
Else
 ' grid_unid.Col = 10
 ' grid_unid.CellForeColor = QBColor(0)
 grid_unid.TextMatrix(grid_unid.Row, 10) = Format(VALOR, "0.0000") ' PRECIO 1
End If
VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 13) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 13)) = 0 Then
  'grid_unid.Col = 12
  'grid_unid.CellForeColor = QBColor(12)
  'grid_unid.TextMatrix(grid_unid.Row, 12) = Format(0, "0.00") ' PRECIO 1
Else
  'grid_unid.Col = 12
 ' grid_unid.CellForeColor = QBColor(0)
  grid_unid.TextMatrix(grid_unid.Row, 12) = Format(VALOR, "0.0000") ' PRECIO 1
End If

VALOR = (Val(grid_unid.TextMatrix(grid_unid.Row, 15) * 100)) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
If Val(grid_unid.TextMatrix(grid_unid.Row, 13)) = 0 Then
Else
  grid_unid.TextMatrix(grid_unid.Row, 14) = Format(VALOR, "0.0000") ' PRECIO 6
End If

grid_unid.COL = tC
Flag_Inicial = ""
Exit Sub
CERO:
   grid_unid.TextMatrix(grid_unid.Row, 5) = Format(0, "0.0000") ' PRECIO 1
   grid_unid.TextMatrix(grid_unid.Row, 7) = Format(0, "0.0000") ' PRECIO 2
   grid_unid.TextMatrix(grid_unid.Row, 9) = Format(0, "0.0000") ' PRECIO 3
   grid_unid.TextMatrix(grid_unid.Row, 11) = Format(0, "0.0000") ' PRECIO 4
   grid_unid.TextMatrix(grid_unid.Row, 13) = Format(0, "0.0000") ' PRECIO 5
   grid_unid.TextMatrix(grid_unid.Row, 15) = Format(0, "0.0000") ' PRECIO 6
  Flag_Inicial = ""
End Sub



Public Sub CALCULAR_POR(WSPRE As Currency, WSCOL As Integer)
Dim VALOR As Currency
If Val(grid_unid.TextMatrix(grid_unid.Row, 3)) <> 0 Then
  VALOR = (WSPRE * 100) / Val(grid_unid.TextMatrix(grid_unid.Row, 3)) - 100
Else
  VALOR = 0
End If
If grid_unid.Row = 2 And Val(grid_unid.TextMatrix(grid_unid.Row, 1)) <> 0 Then
  grid_unid.TextMatrix(grid_unid.Row - 1, WSCOL) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  If WSCOL = 5 Then ' pre 1
    grid_unid.TextMatrix(grid_unid.Row - 1, 23) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  ElseIf WSCOL = 7 Then ' pre 2
    grid_unid.TextMatrix(grid_unid.Row - 1, 24) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  ElseIf WSCOL = 9 Then ' pre 3
    grid_unid.TextMatrix(grid_unid.Row - 1, 25) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  ElseIf WSCOL = 11 Then ' pre4
    grid_unid.TextMatrix(grid_unid.Row - 1, 26) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  ElseIf WSCOL = 13 Then ' pre 5
    grid_unid.TextMatrix(grid_unid.Row - 1, 27) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  ElseIf WSCOL = 15 Then ' pre 6
    grid_unid.TextMatrix(grid_unid.Row - 1, 32) = Format(Val(grid_unid.TextMatrix(grid_unid.Row, WSCOL)) / Val(grid_unid.TextMatrix(grid_unid.Row, 1)), "0.000")
  End If
End If
Flag_Inicial = "A"
If VALOR < 0 Then
  'grid_unid.Col = WSCOL - 1
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = WSCOL - 1
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.COL = WSCOL
Flag_Inicial = ""

grid_unid.TextMatrix(grid_unid.Row, WSCOL - 1) = Format(VALOR, "0.0000")
If grid_unid.Row = 2 Then
  grid_unid.TextMatrix(grid_unid.Row - 1, WSCOL - 1) = grid_unid.TextMatrix(grid_unid.Row, WSCOL - 1)
End If


End Sub

Public Sub CALCULAR_PRE(WSPOR As Currency, WSCOL As Integer)
Dim VALOR As Currency
VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 + (WSPOR / 100))
'VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100))
Flag_Inicial = "A"
If VALOR < 0 Then
  'grid_unid.Col = WSCOL + 1
  'grid_unid.CellForeColor = QBColor(12)
Else
  'grid_unid.Col = WSCOL + 1
  'grid_unid.CellForeColor = QBColor(0)
End If
grid_unid.COL = WSCOL
Flag_Inicial = ""
grid_unid.TextMatrix(grid_unid.Row, WSCOL + 1) = Format(VALOR, "0.000")
  Select Case WSCOL
    Case 4
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 18) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 23) = Format(VALOR, "0.000")
       End If
    Case 6
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 19) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 24) = Format(VALOR, "0.000")
       End If
    Case 8
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 20) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 25) = Format(VALOR, "0.000")
       End If
    Case 10
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 21) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 26) = Format(VALOR, "0.000")
       End If
    Case 12
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 22) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 27) = Format(VALOR, "0.000")
       End If
    Case 14
       If frmARTI.cmddolares.Tag = "D" Then
         grid_unid.TextMatrix(grid_unid.Row, 31) = Format(VALOR, "0.000")
       Else
         grid_unid.TextMatrix(grid_unid.Row, 32) = Format(VALOR, "0.000")
       End If
End Select


End Sub


Private Sub grid_UNID_Click()
ElGrid_Click grid_unid, txtvar
End Sub

Private Sub grid_UNID_DblClick()

If Flag_Inicial = "A" Then
 Exit Sub
End If
If grid_unid.COL = 1 Or grid_unid.COL = 3 Then
  Exit Sub
End If
If grid_unid.COL = 0 And wflag_bloq_descrip = "A" Then Exit Sub
ElGrid_DblClick grid_unid, txtvar
End Sub

Private Sub grid_UNID_EnterCell()
Fcomun.Refresh
txtpeso.Text = Format(grid_unid.TextMatrix(grid_unid.Row, 28), "0.00")
txtlitro.Text = Format(grid_unid.TextMatrix(grid_unid.Row, 30), "0.00")
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_EnterCell grid_unid, txtvar

End Sub

Private Sub grid_UNID_GotFocus()
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_GotFocus grid_unid, txtvar
'grid_UNID.Row = loc_fila
'grid_UNID.Col = loc_colum

End Sub

Private Sub grid_UNID_KeyDown(KeyCode As Integer, Shift As Integer)
If grid_unid.COL = 1 And ws_flag_equiv <> "A" Then
  KeyCode = 0
  Exit Sub
End If
If KeyCode = 32 Then
 Exit Sub
End If
If grid_unid.COL = 0 And wflag_bloq_descrip = "A" Then Exit Sub
If KeyCode = 46 Then
  If wflag_bloq_descrip = "A" Then Exit Sub
  If grid_unid.Row <> 1 Then
    If Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) <> "" And Val(grid_unid.TextMatrix(grid_unid.Row, 1)) <> 0 Then
      pub_mensaje = " Eliminar la Unidad de : " & Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) & " �Desea Continuar... ?"
      Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
      If Pub_Respuesta = vbNo Then
        grid_unid.SetFocus
        Exit Sub
      End If
      If Trim(grid_unid.TextMatrix(grid_unid.Row, 14)) = "A" Then
         grid_unid.TextMatrix(1, 14) = "A"
         LBLUNIDAD.Caption = Trim(grid_unid.TextMatrix(1, 0))
      End If
      grid_unid.RemoveItem grid_unid.Row
      grid_unid.SetFocus
    Else
     grid_unid.RemoveItem grid_unid.Row
     grid_unid.SetFocus
     Exit Sub
    End If
   Exit Sub
  End If
End If

If KeyCode = 45 Then
  If wflag_bloq_descrip = "A" Then Exit Sub
  If grid_unid.Row = grid_unid.Rows - 1 Then
    If Trim(grid_unid.TextMatrix(grid_unid.Row, 0)) <> "" And Val(grid_unid.TextMatrix(grid_unid.Row, 1)) <> 0 Then
      Flag_Inicial = "A"
      grid_unid.CellBackColor = QBColor(15)
      grid_unid.Rows = grid_unid.Rows + 1
      grid_unid.Row = grid_unid.Row + 1
      grid_unid.RowHeight(grid_unid.Row) = 285
      grid_unid.TextMatrix(grid_unid.Row, 0) = "CAJA" 'POR JL
      grid_unid.TextMatrix(grid_unid.Row, 1) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 3) = "0.000"
      grid_unid.COL = 4
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 4) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 5) = "0.000"
      grid_unid.COL = 6
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 6) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 7) = "0.000"
      grid_unid.COL = 8
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 8) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 9) = "0.000"
      grid_unid.COL = 10
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 10) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 11) = "0.000"
      grid_unid.COL = 12
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 12) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 13) = "0.000"
      grid_unid.COL = 14
      grid_unid.CellForeColor = QBColor(9)
      grid_unid.TextMatrix(grid_unid.Row, 14) = "0.000"
      grid_unid.TextMatrix(grid_unid.Row, 15) = "0.000"
      
      Flag_Inicial = ""
      grid_unid.COL = 0
      Exit Sub
    End If
  End If
End If
If grid_unid.COL = 1 And grid_unid.Row = 1 Then
 Exit Sub
End If

ElGrid_KeyDown grid_unid, txtvar, KeyCode
End Sub

Private Sub grid_UNID_KeyPress(KeyAscii As Integer)

If grid_unid.COL = 1 And ws_flag_equiv <> "A" Then
  KeyAscii = 0
  Exit Sub
End If
If grid_unid.COL = 0 And wflag_bloq_descrip = "A" Then Exit Sub
If KeyAscii = 32 And grid_unid.COL = 0 Then
 For fila = 1 To grid_unid.Rows - 1
    grid_unid.TextMatrix(fila, 16) = " "
 Next fila
 grid_unid.TextMatrix(grid_unid.Row, 16) = "A"
 LBLUNIDAD.Caption = grid_unid.TextMatrix(grid_unid.Row, 0)
 'FREDY 23/11/2002
 arm_llave.Requery
 If Not arm_llave.EOF Then
  frmARTI.lblcospro.Caption = Format(arm_llave!ARM_COSPRO * grid_unid.TextMatrix(grid_unid.Row, 1), "###,##0.0000")
 End If
 
 Exit Sub
End If
If grid_unid.COL = 1 And grid_unid.Row = 1 Then
 'aqui se expulsa para que no pueda modificar la unnidad minima = 1
 Exit Sub
End If
If grid_unid.COL = 3 Then Exit Sub
'If grid_unid.COL = 4 Then Exit Sub
ElGrid_KeyPress grid_unid, txtvar, KeyAscii, 13
End Sub

Private Sub grid_UNID_LeaveCell()
If Flag_Inicial = "A" Then
 Exit Sub
End If
If grid_unid.COL = 0 And wflag_bloq_descrip = "A" Then Exit Sub
If Flag_Change <> "A" Then
  If grid_unid.COL = 3 Then ' costo base
    If LK_FLAG_CALCULO = "A" Then
      CALCULAR Val(grid_unid.TextMatrix(grid_unid.Row, 3))
    Else
      CALCULAR_OTRO Val(grid_unid.TextMatrix(grid_unid.Row, 3))
    End If
  End If
  If grid_unid.COL = 5 Or grid_unid.COL = 7 Or grid_unid.COL = 9 Or grid_unid.COL = 11 Or grid_unid.COL = 13 Or grid_unid.COL = 15 Then ' costo PORCENTAJE
    CALCULAR_POR Val(grid_unid.TextMatrix(grid_unid.Row, grid_unid.COL)), grid_unid.COL
  End If
  If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Or grid_unid.COL = 14 Then ' costo PORCENTAJE
    CALCULAR_PRE Val(grid_unid.TextMatrix(grid_unid.Row, grid_unid.COL)), grid_unid.COL
  End If
  Flag_Change = "A"
End If
ElGrid_LeaveCell grid_unid, txtvar
End Sub

Private Sub txtper_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pub_cadena = "Select  * from tabcuo where cuo_codcia = '" & LK_CODCIA & "' and cuo_periodo = '" & txtper.Text & "' and cuo_codart = " & Val(txt_key.Text)
    Set rs_regcuotas = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
    If Not rs_regcuotas.EOF Then
       txtper.Text = rs_regcuotas!CUO_PERIODO
       txtum.Text = Format(rs_regcuotas!CUO_CANTIDAD)
    Else
       txtum.Text = "0.00"
    End If
End If
End Sub

Private Sub txtperiodo_Change()
Dim wrango As Integer
Dim fecha1
wrango = Val(txtd.Text) * Val(txtperiodo.Text) * -1
fecha1 = DateAdd("d", wrango, LK_FECHA_DIA)

txtCampo1.Text = Format(fecha1, "dd/mm/yyyy")
txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")

End Sub

Private Sub txtpeso_Change()
  If grid_unid.Rows > 1 Then
    grid_unid.TextMatrix(grid_unid.Row, 28) = Val(txtpeso.Text)
  End If
End Sub

Private Sub txtpeso_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL txtpeso, KeyAscii
If KeyAscii = 13 Then
  grid_unid.SetFocus
End If

End Sub

Private Sub txtpor1_Change()
cmdactvend.Enabled = True
End Sub

Private Sub txtpor11_Change()
cmdactvend.Enabled = True
End Sub

Private Sub txtpor11_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor22, txtpor22
End Sub
Private Sub txtpor22_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor33, txtpor33
End Sub
Private Sub txtpor33_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor44, txtpor44
End Sub

Private Sub txtpor44_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor55, txtpor55
End Sub

Private Sub txtpor55_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor66, txtpor66
End Sub

Private Sub txtpor66_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor77, txtpor77
End Sub
Private Sub txtpor77_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor88, txtpor88
End Sub
Private Sub txtpor88_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor99, txtpor99
End Sub


Private Sub txtpor22_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor33_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor44_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor55_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor66_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor77_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor88_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor99_Change()
cmdactvend.Enabled = True
End Sub

Private Sub txtpor2_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor3_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor4_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor5_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor6_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor7_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor8_Change()
cmdactvend.Enabled = True
End Sub
Private Sub txtpor9_Change()
cmdactvend.Enabled = True
End Sub

Private Sub txtpor1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor2, txtpor2
End Sub

Private Sub txtpor2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor3, txtpor3

End Sub

Private Sub txtpor3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor4, txtpor4
End Sub

Private Sub txtpor4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor5, txtpor5
End Sub

Private Sub txtpor5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor6, txtpor6
End Sub
Private Sub txtpor6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor7, txtpor7
End Sub
Private Sub txtpor7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul txtpor8, txtpor8
End Sub
Private Sub txtpor8_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor9, txtpor9
End Sub
Private Sub txtpor9_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Azul txtpor11, txtpor11
End Sub


Private Sub txtvar_Change()
Dim VALOR As Currency
If grid_unid.COL = 5 And txtvar.Visible Then
  Exit Sub
  grid_unid.TextMatrix(grid_unid.Row, 3) = Format(Val(txtvar), "0.0000")
 '  CALCULAR_PRE Val(grid_unid.TextMatrix(grid_unid.Row, 6)), 6 JLPV
  WSPOR = 0
  VALOR = Format(Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100)), "0.000")
  
  grid_unid.TextMatrix(grid_unid.Row, 5) = VALOR
  grid_unid.TextMatrix(grid_unid.Row, 23) = grid_unid.TextMatrix(grid_unid.Row, 5)
  WSPOR = Val(grid_unid.TextMatrix(grid_unid.Row, 6))
  VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100))
  grid_unid.TextMatrix(grid_unid.Row, 7) = Format(VALOR, "0.0000")
  grid_unid.TextMatrix(grid_unid.Row, 24) = grid_unid.TextMatrix(grid_unid.Row, 7)
  WSPOR = Val(grid_unid.TextMatrix(grid_unid.Row, 8))
  VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100))
  grid_unid.TextMatrix(grid_unid.Row, 9) = Format(VALOR, "0.0000")
  grid_unid.TextMatrix(grid_unid.Row, 25) = grid_unid.TextMatrix(grid_unid.Row, 9)
  WSPOR = Val(grid_unid.TextMatrix(grid_unid.Row, 10))
  VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100))
  grid_unid.TextMatrix(grid_unid.Row, 11) = Format(VALOR, "0.0000")
  grid_unid.TextMatrix(grid_unid.Row, 26) = grid_unid.TextMatrix(grid_unid.Row, 11)
  WSPOR = Val(grid_unid.TextMatrix(grid_unid.Row, 12))
  VALOR = Val(grid_unid.TextMatrix(grid_unid.Row, 3)) * (1 - (WSPOR / 100))
  grid_unid.TextMatrix(grid_unid.Row, 13) = Format(VALOR, "0.0000")
  grid_unid.TextMatrix(grid_unid.Row, 27) = grid_unid.TextMatrix(grid_unid.Row, 13)
  grid_unid.TextMatrix(grid_unid.Row, 15) = Format(VALOR, "0.0000")
  grid_unid.TextMatrix(grid_unid.Row, 32) = grid_unid.TextMatrix(grid_unid.Row, 15)
End If
Exit Sub
End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Then
 If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Then
   If Val(txtvar.Text) > 999.99 Or Val(txtvar.Text) < 0 Then
     txtvar.Visible = True
     Exit Sub
    End If
  End If
  grid_unid.SetFocus
End If

If KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Or KeyCode = 13 Then
 If grid_unid.COL = 1 Then
   If Val(grid_unid.TextMatrix(grid_unid.Row, 3)) = 0 Then
    grid_unid.TextMatrix(grid_unid.Row, 3) = Format(Val(grid_unid.TextMatrix(1, 3)) * Val(txtvar.Text), "0.00")
     If LK_FLAG_CALCULO = "A" Then
       CALCULAR Val(grid_unid.TextMatrix(grid_unid.Row, 3))
     Else
       CALCULAR_OTRO Val(grid_unid.TextMatrix(grid_unid.Row, 3))
     End If
   End If
 End If
End If

TEXTO_KeyDown grid_unid, txtvar, KeyCode, 13
If KeyCode = 13 Or KeyCode = 27 Then
 grid_unid.SetFocus
End If
End Sub

Private Sub txtvar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grid_unid.COL = 4 Or grid_unid.COL = 6 Or grid_unid.COL = 8 Or grid_unid.COL = 10 Or grid_unid.COL = 12 Then
  If Val(txtvar.Text) > 999.99 Or Val(txtvar.Text) < 0 Then
    txtvar.Visible = True
    Exit Sub
  End If
End If
End If

TEXTO_KeyPress grid_unid, txtvar, KeyAscii, 13, 1, 2, 3, 2, 4, 2, 5, 2, 6, 2, 7, 2, 8, 2, 9, 2, 10, 2, 11, 2, 12, 2, 13, 2
End Sub

Private Sub txtvar_LostFocus()
TEXTO_LosFocus grid_unid, txtvar
End Sub

Private Sub txt_alterno_GotFocus()
If Left(cmdAgregar.Caption, 2) = "&A" Or Left(cmdModificar.Caption, 2) = "&M" Then
 Exit Sub
End If
 If ListView1.Visible Then
  ListView1.Visible = False
 End If
 Azul txt_alterno, txt_alterno
 frmARTI.F14.Visible = False
End Sub
Private Sub txt_alterno_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If Not ListView1.Visible Or Left(cmdAgregar.Caption, 2) = "&G" Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_alterno.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_alterno.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_alterno.SelStart = Len(txt_alterno.Text)
  DoEvents
fin:

End Sub
Private Sub txt_alterno_KeyPress(KeyAscii As Integer)
Dim VALOR As String
Dim tf As Integer
Dim I
Dim itmFound As MSComctlLib.ListItem
VALOR = Chr(KeyAscii)

KeyAscii = Asc(UCase(VALOR))
VALOR = ""
If KeyAscii = 27 Then
  ListView1.Visible = False
 txt_alterno.Text = ""
End If
If Left(cmdAgregar.Caption, 2) = "&G" Then
 If KeyAscii = 13 Then
   txtnombre.SetFocus
   Exit Sub
 End If
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
If VAR_ACTIVAR <> 99 Then
 LOC_OPER = 1
 pu_alterno = txt_alterno.Text
 PUB_CODCIA = LK_CODCIA
 LEER_LOC
 If artloc_llave.EOF Then
   MsgBox "Codigo No Existe ...", 48, Pub_Titulo
   Azul txt_alterno, txt_alterno
   Exit Sub
 End If
 LLENA_ARTI 0
 BLOQUEA_TEXT txt_alterno
 frmARTI.cmdModificar.SetFocus
 BLOQUEA_TEXT txtnombre
 cmdCancelar.Enabled = True
 Exit Sub
End If
pu_alterno = Trim(txt_alterno.Text)
If Len(txt_alterno.Text) = 0 Then
   Exit Sub
End If
If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
  Exit Sub
End If
VALOR = UCase(ListView1.ListItems.Item(loc_key).Text)
If Trim(UCase(txt_alterno.Text)) = Left(VALOR, Len(Trim(txt_alterno.Text))) Then
Else
   Exit Sub
End If
LLENA_ARTI 0
BLOQUEA_TEXT txt_alterno
frmARTI.cmdModificar.SetFocus
BLOQUEA_TEXT txtnombre
cmdCancelar.Enabled = True
dale:
ListView1.Visible = False

fin:
End Sub

Private Sub txt_alterno_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If KeyCode = 13 Or KeyCode = 27 Then VAR_ACTIVAR = 0: Exit Sub
If Left(cmdAgregar.Caption, 2) = "&G" Then Exit Sub
If txt_alterno.Text = "*" And KeyCode = 106 Then
 VAR_ACTIVAR = 99
 Exit Sub
ElseIf txt_alterno.Text = "" Then
  ListView1.Visible = False
  VAR_ACTIVAR = 0
  Exit Sub
ElseIf txt_alterno.Text = "* " And KeyCode = 106 Then
 VAR_ACTIVAR = 99
 txt_alterno.Text = "*"
 txt_alterno.SelStart = Len(txt_alterno.Text)
 KeyCode = 0
 Exit Sub
End If

If VAR_ACTIVAR <> 99 Then
 Exit Sub
End If
If txt_alterno.Text = "*" Then
 Exit Sub
ElseIf Left(txt_alterno.Text, 1) = "*" Then
 txt_alterno.Text = Mid(txt_alterno.Text, 2, Len(txt_alterno.Text))
 txt_alterno.SelStart = Len(txt_alterno.Text)
End If
If Len(txt_alterno.Text) = 0 Or txt_alterno.Text = "" Or Left(cmdAgregar.Caption, 2) = "&G" Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_alterno.Text) = 1 Then
    If txt_alterno.Text = "" Then txt_alterno.Text = " "
    var = Asc(txt_alterno.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    'ElseIf var = 50 Then
    '   var = " 1"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
    ws_codcia = "00"
    End If
    numarchi = 3
    ' archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK FROM ARTI, ARTICULO WHERE (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND  ART_KEY <> 0 AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_alterno.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    archi = "SELECT TOP 2000 ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK ,PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS WHERE (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_CODCIA = ARM_CODCIA) AND (ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_KEY  <> 1 and ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_alterno.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    PROC_LISVIEW ListView1, 1000
    loc_key = 0
    If ListView1.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As MSComctlLib.ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_alterno.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
  End If
  Exit Sub
End If

End Sub

Public Sub PROCESO_ARTI()
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
 cade = "SELECT * FROM ARTI WHERE ART_ALTERNO = ? AND ART_CODCIA = ?  AND ART_TIPO = ? ORDER BY ART_CODCIA, ART_KEY"
Else
 cade = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ?  AND ART_TIPO = ? ORDER BY ART_CODCIA, ART_KEY"
End If
Set PSART_LOC = CN.CreateQuery("", cade)
 If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PSART_LOC.rdoParameters(0) = " "
 Else
  PSART_LOC.rdoParameters(0) = 0
 End If
 PSART_LOC.rdoParameters(1) = " "
 PSART_LOC.rdoParameters(2) = " "
Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
End Sub

Public Sub PROCESO_CANCELAR()
    If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
        frmARTI.frarelacion.Enabled = False
        LIMPIA_ARTI
        cmdCancelar.Enabled = True
        BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
        BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
        BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
        BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
        BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
        frmARTI.SSTab1.tab = 0
        If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          DESBLOQUEA_TEXT txt_alterno
          BLOQUEA_TEXT txt_key
          If frmARTI.txt_alterno.Visible Then frmARTI.txt_alterno.SetFocus
        Else
          BLOQUEA_TEXT txt_alterno
          DESBLOQUEA_TEXT txt_key
          If frmARTI.txt_key.Visible Then frmARTI.txt_key.SetFocus
        End If
        MANOS(0).Enabled = True
        MANOS(1).Enabled = True
        pasa = 0
        Exit Sub
    End If
    Screen.MousePointer = 11
    If Left(cmdModificar.Caption, 2) = "&G" Then
       cmdModificar.Caption = "&Modificaci�n"
       LLENA_ARTI 1
       BLOQUEA_TEXT txt_key
       BLOQUEA_TEXT txt_alterno
    Else
       frmARTI.frarelacion.Enabled = False
       cmdAgregar.Caption = "&Adicionar"
       LIMPIA_ARTI
       If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
          DESBLOQUEA_TEXT txt_alterno
          BLOQUEA_TEXT txt_key
          frmARTI.txt_alterno.SetFocus
        Else
          DESBLOQUEA_TEXT txt_key
          BLOQUEA_TEXT txt_alterno
          frmARTI.txt_key.SetFocus
        End If
    End If
    cmdAgregar.Enabled = True
    cmdEliminar.Enabled = True
    cmdModificar.Enabled = True
    BLOQUEA_TEXT txtnombre, CmbCalidad, decimales, DS, txtcospro, art_situacion, art_linea, art_numero, art_marca, art_plancha
    BLOQUEA_TEXT art_grupo, art_familia, art_subfam, grid_unid, txtMin, txtMax, art_codpro, txtcodigo2, chebof, chesub
    BLOQUEA_TEXT cheservi(0), cheservi(1), cheservi(2), exigv, txtcospro, cmddolares, txtpeso, txtfechault, checambio, txtlitro
    BLOQUEA_TEXT txtpor1, txtpor2, txtpor3, txtpor4, txtpor5, txtpor11, txtpor22, txtpor33, txtpor44, txtpor55
    BLOQUEA_TEXT cheoferta, txtmargen, cmbUsu, txtCantMax, txtCantMaxADD, cheasignarc, art_grupop, lfvta
    pasa = 0
    MANOS(0).Enabled = True
    MANOS(1).Enabled = True
    MENSAJE_ARTI "Proceso Cancelado... !!!    "
    frmARTI.SSTab1.tab = 0
    Screen.MousePointer = 0
End Sub


Public Sub PROCESA_PROV()
   SQ_OPER = 2
   pu_cp = "P"
   pu_codclie = 0
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   art_codpro.Clear
   Do Until cli_mayor.EOF
     art_codpro.AddItem cli_mayor!cli_nombre & String(20, " ") & Trim(CStr(cli_mayor!cli_codclie))
     cli_mayor.MoveNext
   Loop
End Sub

Public Function EXISTE_ART(WARTI As String, WCODI As String) As Boolean
Dim var
Dim tempo
tempo = Left(Trim(WARTI), Len(WARTI) - 1)
var = Asc(Right(Trim(WARTI), 1))
var = var + 1
If var = 91 Then
  var = "ZZZZZZZZ"
Else
  var = Chr(var)
End If
tempo = tempo + var
archi = "SELECT * FROM ARTI WHERE  ART_KEY <> " & WCODI & " AND ART_KEY  <> 1 and ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & WARTI & "' AND  '" & tempo & "' ORDER BY ART_NOMBRE "
ListExiste.Clear
F14.Caption = "Lista de Articulos  Parecidos ... "
frmARTI.ListExiste.TextMatrix(0, 0) = "Cia"
frmARTI.ListExiste.TextMatrix(0, 1) = "Codigo "
frmARTI.ListExiste.TextMatrix(0, 2) = "Articulo"
If LK_FLAG_ORIGINAL <> "A" Then
 frmARTI.ListExiste.TextMatrix(0, 3) = "Cod.Alterno"
End If

EXISTE_ART = False
Set PSX = CN.CreateQuery("", archi)
Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
If X.EOF Then
 frmARTI.ListExiste.Clear
 GoTo fin
End If

fila = 0
frmARTI.ListExiste.Rows = 2
Do Until X.EOF
    fila = fila + 1
    frmARTI.ListExiste.TextMatrix(fila, 0) = Nulo_Valors(X!art_codcia)
    frmARTI.ListExiste.TextMatrix(fila, 1) = Nulo_Valor0(X!art_key)
    frmARTI.ListExiste.TextMatrix(fila, 2) = Nulo_Valors(X!art_nombre)
    If LK_FLAG_ORIGINAL <> "A" Then
     frmARTI.ListExiste.TextMatrix(fila, 3) = Nulo_Valors(X!ART_alterno)
    End If
    frmARTI.ListExiste.Rows = frmARTI.ListExiste.Rows + 1
    X.MoveNext
Loop
EXISTE_ART = True
If EXISTE_ART Then
    frmARTI.ListExiste.Rows = frmARTI.ListExiste.Rows - 1
    Op(0).Value = False
    Op(0).Enabled = False
    Op(1).Value = True
    frmARTI.F14.Visible = True
    frmARTI.ListExiste.Row = 1
    frmARTI.ListExiste.COL = 1
    frmARTI.ListExiste.SetFocus
End If
GoTo fin
Exit Function

CHECKERROR:
MsgBox Err.Description
fin:

End Function


Public Sub CABE_RELACION()
gridrel.Cols = 4
gridrel.Rows = 1
gridrel.TextMatrix(0, 0) = "Cod.Orig."
gridrel.TextMatrix(0, 1) = "Cod.Alterno"
gridrel.TextMatrix(0, 2) = "Descripci�n"
gridrel.TextMatrix(0, 3) = "Calidad"
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
gridrel.ColWidth(1) = 1000
Else
gridrel.ColWidth(1) = 1
End If
gridrel.ColWidth(0) = 1000
gridrel.ColWidth(2) = 3100
gridrel.ColWidth(3) = 2000

End Sub

Public Sub LLENA_RELACION(Wkey_Rela As Currency)
gridrel.Clear
CABE_RELACION
gridrel.Rows = 2
SQ_OPER = 1
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
PUB_KEY = Wkey_Rela
PUB_CODCIA = LK_CODCIA
LOC_OPER = 2
LEER_LOC
If artloc_key.EOF Then
   artloc_llave.Edit
   artloc_llave!ART_CODART2 = 0
   artloc_llave.Update
   Exit Sub
End If
 If artloc_key!art_key = 0 Then Exit Sub
 gridrel.RowHeight(1) = 300
 gridrel.TextMatrix(1, 0) = artloc_key!art_key
 gridrel.TextMatrix(1, 1) = artloc_key!ART_alterno
 gridrel.TextMatrix(1, 2) = artloc_key!art_nombre
 SQ_OPER = 1
 PUB_TIPREG = 2
 PUB_NUMTAB = artloc_key!ART_CALIDAD
 PUB_CODCIA = LK_CODCIA
 LEER_TAB_LLAVE
 If tab_llave.EOF Then
  gridrel.TextMatrix(1, 3) = ""
 Else
  gridrel.TextMatrix(1, 3) = Trim(tab_llave!TAB_NOMLARGO)
 End If

End Sub

Public Sub LLENA_CALREL(wcla_actual As Integer)
Dim wa As Integer
PUB_TIPREG = 2
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
cmbcal.Clear
wa = 0
Do Until tab_mayor.EOF
    If tab_mayor!TAB_NUMTAB > wcla_actual Then
       cmbcal.AddItem tab_mayor!TAB_NOMLARGO & String(50, " ") & tab_mayor!TAB_NUMTAB
    End If
    wa = 1
    tab_mayor.MoveNext
Loop
If cmbcal.ListCount = 0 And wa = 1 Then
  cmbcal.AddItem "<Ninguno>"
End If
If cmbcal.ListCount > 0 Then cmbcal.ListIndex = 0
End Sub

Public Sub llena_pre(wlista As String)
Dim WT_ENT As Currency
pu_codcia = LK_CODCIA
PUB_CODART = Val(txt_key.Text)
SQ_OPER = 2
LEER_PRE_LLAVE
If pre_mayor.EOF Then
 MsgBox "Error de Unidades NO existe......", 48, Pub_Titulo
 Exit Sub
End If
fila = 0
Flag_Inicial = "A"
grid_unid.Clear
CABEZA_UNID
Do Until pre_mayor.EOF
   fila = fila + 1
   grid_unid.Rows = fila + 1
   grid_unid.Row = fila
   grid_unid.RowHeight(fila) = 285
   grid_unid.TextMatrix(fila, 0) = Trim(pre_mayor!PRE_UNIDAD)
   grid_unid.TextMatrix(fila, 1) = pre_mayor!PRE_EQUIV
   grid_unid.TextMatrix(fila, 2) = ""
  ' grid_unid.TextMatrix(fila, 3) = Format(Nulo_Valor0(pre_mayor!PRE_COSTO), "0.0000")
  ' grid_unid.TextMatrix(fila, 27) = Format(Nulo_Valor0(pre_mayor!PRE_COSTO), "0.00")
   grid_unid.TextMatrix(fila, 3) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * pre_mayor!PRE_EQUIV, "0.0000")
   grid_unid.TextMatrix(fila, 29) = Format(Nulo_Valor0(arm_llave!ARM_COSPRO) * pre_mayor!PRE_EQUIV, "0.0000")
   
   grid_unid.COL = 4
   grid_unid.CellForeColor = QBColor(9)
   If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then
  '   WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE1) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   End If
   WSPOR = pre_mayor!PRE_POR1
   grid_unid.TextMatrix(fila, 4) = Format(WSPOR, "0.00")
   If wlista = "S" Then
     grid_unid.TextMatrix(fila, 5) = Nulo_Valor0(pre_mayor!PRE_PRE1)
   Else
     grid_unid.TextMatrix(fila, 5) = Nulo_Valor0(pre_mayor!PRE_PRE11)
   End If
   
   'grid_unid.TextMatrix(fila, 3) = Format(grid_unid.TextMatrix(fila, 5), "0.00")
   
   grid_unid.TextMatrix(fila, 18) = Nulo_Valor0(pre_mayor!PRE_PRE11)
   grid_unid.TextMatrix(fila, 23) = Nulo_Valor0(pre_mayor!PRE_PRE1)
   grid_unid.COL = 6
   grid_unid.CellForeColor = QBColor(9)
   WSPOR = pre_mayor!PRE_POR2
   grid_unid.TextMatrix(fila, 6) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 7) = Nulo_Valor0(pre_mayor!PRE_PRE2)
   Else
    grid_unid.TextMatrix(fila, 7) = Nulo_Valor0(pre_mayor!PRE_PRE22)
   End If
   grid_unid.TextMatrix(fila, 19) = Nulo_Valor0(pre_mayor!PRE_PRE22)
   grid_unid.TextMatrix(fila, 24) = Nulo_Valor0(pre_mayor!PRE_PRE2)
   grid_unid.COL = 8
   grid_unid.CellForeColor = QBColor(9)
   ' If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE3) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   WSPOR = pre_mayor!PRE_POR3
   grid_unid.TextMatrix(fila, 8) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 9) = Nulo_Valor0(pre_mayor!PRE_PRE3)
   Else
    grid_unid.TextMatrix(fila, 9) = Nulo_Valor0(pre_mayor!PRE_PRE33)
   End If
   grid_unid.TextMatrix(fila, 20) = Nulo_Valor0(pre_mayor!PRE_PRE33)
   grid_unid.TextMatrix(fila, 25) = Nulo_Valor0(pre_mayor!PRE_PRE3)
   grid_unid.COL = 10
   grid_unid.CellForeColor = QBColor(9)
   ' If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE4) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   WSPOR = pre_mayor!PRE_POR4
   grid_unid.TextMatrix(fila, 10) = Format(WSPOR, "0.00")

   If wlista = "S" Then
     grid_unid.TextMatrix(fila, 11) = Nulo_Valor0(pre_mayor!PRE_PRE4)
   Else
     grid_unid.TextMatrix(fila, 11) = Nulo_Valor0(pre_mayor!PRE_PRE44)
   End If
   
   grid_unid.TextMatrix(fila, 21) = Nulo_Valor0(pre_mayor!PRE_PRE44)
   grid_unid.TextMatrix(fila, 26) = Nulo_Valor0(pre_mayor!PRE_PRE4)
   
   grid_unid.COL = 12
   grid_unid.CellForeColor = QBColor(9)
   ' If Val(grid_unid.TextMatrix(fila, 3)) <> 0 Then WSPOR = (Nulo_Valor0(pre_mayor!PRE_PRE5) * 100) / Val(grid_unid.TextMatrix(fila, 3)) - 100
   WSPOR = pre_mayor!PRE_POR5
   grid_unid.TextMatrix(fila, 12) = Format(WSPOR, "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 13) = Nulo_Valor0(pre_mayor!PRE_PRE5)
   Else
    grid_unid.TextMatrix(fila, 13) = Nulo_Valor0(pre_mayor!PRE_PRE55)
   End If
   grid_unid.TextMatrix(fila, 22) = Nulo_Valor0(pre_mayor!PRE_PRE55)
   grid_unid.TextMatrix(fila, 27) = Nulo_Valor0(pre_mayor!PRE_PRE5)
   
   grid_unid.TextMatrix(fila, 14) = Format(Nulo_Valor0(pre_mayor!PRE_POR6), "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 15) = Nulo_Valor0(pre_mayor!PRE_PRE6)
   Else
    grid_unid.TextMatrix(fila, 15) = Nulo_Valor0(pre_mayor!PRE_PRE66)
   End If
   grid_unid.TextMatrix(fila, 32) = Nulo_Valor0(pre_mayor!PRE_PRE6)
   grid_unid.TextMatrix(fila, 31) = Nulo_Valor0(pre_mayor!PRE_PRE66)
   
   grid_unid.TextMatrix(fila, 33) = Format(Nulo_Valor0(pre_mayor!PRE_POR7), "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 34) = Nulo_Valor0(pre_mayor!PRE_PRE7)
   Else
    grid_unid.TextMatrix(fila, 34) = Nulo_Valor0(pre_mayor!PRE_PRE77)
   End If
   grid_unid.TextMatrix(fila, 35) = Nulo_Valor0(pre_mayor!PRE_PRE7)
   grid_unid.TextMatrix(fila, 36) = Nulo_Valor0(pre_mayor!PRE_PRE77)
   
   grid_unid.TextMatrix(fila, 37) = Format(Nulo_Valor0(pre_mayor!PRE_POR8), "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 38) = Nulo_Valor0(pre_mayor!PRE_PRE8)
   Else
    grid_unid.TextMatrix(fila, 38) = Nulo_Valor0(pre_mayor!PRE_PRE88)
   End If
   grid_unid.TextMatrix(fila, 39) = Nulo_Valor0(pre_mayor!PRE_PRE8)
   grid_unid.TextMatrix(fila, 40) = Nulo_Valor0(pre_mayor!PRE_PRE88)
   
   grid_unid.TextMatrix(fila, 41) = Format(Nulo_Valor0(pre_mayor!PRE_POR9), "0.00")
   If wlista = "S" Then
    grid_unid.TextMatrix(fila, 42) = Nulo_Valor0(pre_mayor!PRE_PRE9)
   Else
    grid_unid.TextMatrix(fila, 42) = Nulo_Valor0(pre_mayor!PRE_PRE99)
   End If
   grid_unid.TextMatrix(fila, 43) = Nulo_Valor0(pre_mayor!PRE_PRE9)
   grid_unid.TextMatrix(fila, 44) = Nulo_Valor0(pre_mayor!PRE_PRE99)
   ' precio exacto
   grid_unid.TextMatrix(fila, 45) = 0
   grid_unid.TextMatrix(fila, 46) = Nulo_Valor0(pre_mayor!PRE_PREC1)
   grid_unid.TextMatrix(fila, 47) = Nulo_Valor0(pre_mayor!PRE_PREC11)
   
   grid_unid.TextMatrix(fila, 48) = 0
   grid_unid.TextMatrix(fila, 49) = Nulo_Valor0(pre_mayor!PRE_PREC2)
   grid_unid.TextMatrix(fila, 50) = Nulo_Valor0(pre_mayor!PRE_PREC22)
   grid_unid.TextMatrix(fila, 51) = 0
   grid_unid.TextMatrix(fila, 52) = Nulo_Valor0(pre_mayor!PRE_PRECHI1)
   grid_unid.TextMatrix(fila, 53) = Nulo_Valor0(pre_mayor!PRE_PRECHI11)
   
   
   
   grid_unid.TextMatrix(fila, 16) = Nulo_Valor0(pre_mayor!pre_FLAG_UNIDAD)
   If Trim(pre_mayor!pre_FLAG_UNIDAD) = "A" Then loc_equiv_max = pre_mayor!PRE_EQUIV
   grid_unid.TextMatrix(fila, 17) = Nulo_Valor0(pre_mayor!PRE_SECUENCIA)

   grid_unid.TextMatrix(fila, 28) = Nulo_Valor0(pre_mayor!pre_PESO)
   grid_unid.TextMatrix(fila, 30) = Nulo_Valor0(pre_mayor!PRE_LITRO)
   If pre_mayor!PRE_SECUENCIA = 0 Then
     Lst0.Caption = "St. " & Trim(pre_mayor!PRE_UNIDAD)
     frmARTI.lblstock0.Caption = Format(Val(arm_llave!arm_stock) / Val(pre_mayor!PRE_EQUIV), "0.000")
   ElseIf pre_mayor!PRE_SECUENCIA = 1 Then
     WT_ENT = Int(Val(arm_llave!arm_stock) / Val(pre_mayor!PRE_EQUIV))
     Lst1.Caption = "St. " & Trim(pre_mayor!PRE_UNIDAD)
     frmARTI.lblstock1.Caption = Format(WT_ENT, "0")
     WT_ENT = Val(arm_llave!arm_stock) - (WT_ENT * Val(pre_mayor!PRE_EQUIV))
     frmARTI.lblstock0.Caption = Format(WT_ENT, "0")
     lblum.Caption = "U. Maxima:" & Chr(13) & Trim(pre_mayor!PRE_UNIDAD)
   End If
     
   If Nulo_Valor0(pre_mayor!pre_FLAG_UNIDAD) = "A" Then
     LBLUNIDAD.Caption = Trim(pre_mayor!PRE_UNIDAD)
     frmARTI.lblcospro.Caption = Format(Val(frmARTI.lblcospro.Caption) * Trim(pre_mayor!PRE_EQUIV), "###,##0.0000")
   End If
   pre_mayor.MoveNext
Loop

Flag_Inicial = ""
grid_unid.Row = 1
grid_unid.COL = 0
If LK_EMP = "3AA" Then
'  cmddolares_Click
End If

End Sub

Public Sub ARTI_CERO()

PUB_KEY = 0
artloc_llave.AddNew
artloc_llave!art_key = PUB_KEY
artloc_llave!ART_alterno = PUB_KEY
artloc_llave!ART_POR_IGV = 0
artloc_llave!ART_ORDEN = 0
artloc_llave!art_tipo = loc_tipo
artloc_llave!art_familia = 0
artloc_llave!art_subfam = 0
artloc_llave!ART_CALIDAD = 1
artloc_llave!art_subgru = 0
artloc_llave!art_codclie = 0
artloc_llave!art_nombre = "Productos"
artloc_llave!art_codcia = LK_CODCIA
artloc_llave!ART_DECIMALES = 2
artloc_llave!ART_MONEDA = "S"
artloc_llave!art_situacion = 0
artloc_llave!ART_STOCK_MIN = 0
artloc_llave!ART_STOCK_MAX = 0
artloc_llave!ART_POR_IGV = 0
artloc_llave!ART_CODART2 = 0
artloc_llave!ART_EX_IGV = 0
artloc_llave!ART_COSPRO = 0
artloc_llave!ART_EX_IGV = "A"
artloc_llave!art_flag_stock = "M"
artloc_llave!ART_POR1 = 0
artloc_llave!ART_POR2 = 0
artloc_llave!ART_POR3 = 0
artloc_llave!ART_POR4 = 0
artloc_llave!ART_POR5 = Val(txtpor5.Text)
artloc_llave.Update

pre_mayor.AddNew
pre_mayor!PRE_CODCIA = LK_CODCIA
pre_mayor!PRE_CODART = PUB_KEY
pre_mayor!PRE_SECUENCIA = 1
pre_mayor!PRE_UNIDAD = "UND"
pre_mayor!PRE_EQUIV = 1

pre_mayor!PRE_PRE11 = 0
pre_mayor!PRE_PRE22 = 0
pre_mayor!PRE_PRE33 = 0
pre_mayor!PRE_PRE44 = 0
pre_mayor!PRE_PRE55 = 0

pre_mayor!PRE_PRE1 = 0
pre_mayor!PRE_PRE2 = 0
pre_mayor!PRE_PRE3 = 0
pre_mayor!PRE_PRE4 = 0
pre_mayor!PRE_PRE5 = 0
pre_mayor!pre_PESO = 0
pre_mayor!PRE_LITRO = 0
pre_mayor!pre_FLAG_UNIDAD = "A"
pre_mayor.Update

arm_llave.AddNew
arm_llave!ARM_CODART = PUB_KEY
arm_llave!ARM_CODCIA = LK_CODCIA
arm_llave!arm_stock = 0
arm_llave!ARM_STOCKS = 0
arm_llave!ARM_STOCKS2 = 0
arm_llave!ARM_STOCKN = 0
arm_llave!ARM_STOCKN2 = 0
arm_llave!ARM_INGRESOS = 0
arm_llave!ARM_SALIDAS = 0
arm_llave!ARM_COSPRO = 0
arm_llave!arm_stock2 = 0
arm_llave!ARM_COSTO_ULT = 0
arm_llave.Update

End Sub


Public Function ARMA_NOMBRE() As String
ARMA_NOMBRE = Trim(txtnombre.Text)
If LK_EMP = "CAM" Then
  If Val(Right(art_familia.Text, 6)) = 1 Then
  ''& " " & Trim(Left(art_marca.Text, 10)) & "-" & Trim(Left(art_numero.Text, 40)) & " " & Trim(Left(art_linea.Text, 5))
    ARMA_NOMBRE = Left(art_familia.Text, 1) & "." & Trim(Left(art_grupo.Text, 10)) & " " & Trim(Left(art_subfam.Text, 15))
    If Trim(Left(art_marca.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_marca.Text, 10))
    End If
    If Trim(Left(art_plancha.Text, 40)) <> "" Then
       ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_plancha.Text, 40))
    End If
    If Trim(Left(art_linea.Text, 5)) <> "" Then
        ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_linea.Text, 5))
    End If
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
End If
If Val(Right(art_familia.Text, 6)) = 4 Then
    ARMA_NOMBRE = Trim(Left(art_numero.Text, 15))
    If Trim(Left(art_grupo.Text, 15)) <> "" Then
      ARMA_NOMBRE = ARMA_NOMBRE & " " & Trim(Left(art_grupo.Text, 15))
    End If
    ARMA_NOMBRE = ARMA_NOMBRE & " " & Trim(Left(art_marca.Text, 15))
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
End If
  If Val(Right(art_familia.Text, 6)) = 2 Then
    'ARMA_NOMBRE = Trim(Left(art_numero.Text, 15)) & " " & Trim(Left(art_grupo.Text, 15))
    ARMA_NOMBRE = Trim(Left(art_numero.Text, 15)) & " " & Trim(Left(art_grupo.Text, 10)) '& " " & Trim(Left(art_subfam.Text, 15))
    If Trim(Left(art_marca.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_marca.Text, 10))
    End If
    If Trim(Left(art_plancha.Text, 10)) <> "" Then
      ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_plancha.Text, 10))
    End If
    If Trim(Left(art_linea.Text, 5)) <> "" Then
        ARMA_NOMBRE = Trim(ARMA_NOMBRE) & " " & Trim(Left(art_linea.Text, 5))
    End If
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
    
  
    ARMA_NOMBRE = UCase(ARMA_NOMBRE)
  End If
  Exit Function
End If

End Function

Public Sub PROD_PROC()

gridp.Clear
gridp.Cols = 3
gridp.Rows = 1
gridp.TextMatrix(0, 0) = "Descripci�n"
gridp.TextMatrix(0, 1) = "Codigo"
gridp.TextMatrix(0, 2) = "Productos"

gridp.ColWidth(0) = 1800
gridp.ColWidth(1) = 900
gridp.ColWidth(2) = 2500


PUB_TIPREG = 122
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
   PUB_CODCIA = "00"
End If
SQ_OPER = 2
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
 gridp.Rows = gridp.Rows + 1
 gridp.RowHeight(fila) = 285
 gridp.TextMatrix(fila + 1, 0) = Left(tab_mayor!TAB_NOMLARGO, 40)
 fila = fila + 1
 If fila = 3 Then
   Exit Do
 End If
 tab_mayor.MoveNext
Loop
WCODART2 = Val(artloc_llave!ART_CODART2)
fila = 4
Do Until WCODART2 = 0
PSART_RELA.rdoParameters(0) = LK_CODCIA
PSART_RELA.rdoParameters(1) = WCODART2
art_rela.Requery
If art_rela.EOF Then
 MsgBox "Verificar la relacion ."
 Exit Sub
End If
fila = fila - 1
gridp.TextMatrix(fila, 1) = art_rela!ART_alterno
gridp.TextMatrix(fila, 2) = art_rela!art_nombre
WCODART2 = Val(art_rela!ART_CODART2)
Loop

End Sub

Public Sub opcional()
Dim WDIRCLI As Integer

Dim VALOR As Currency

Dim WR_SUMA As Currency
Dim WWF As String
Dim ST_ACTUAL As Currency
Dim ps_llave_precios As rdoResultset
Dim WfART_llave As rdoResultset
Dim WPSART_LLAVE As rdoQuery

Dim PSFAR_PEDIDOS As rdoQuery
Dim far_pedidos As rdoResultset

Dim PSPRE_LOTES As rdoQuery
Dim pre_lotes As rdoResultset

Dim wcanti_unid As Currency
Dim WS_FILA As Integer
Dim xl  As Object
Dim E As Integer
Dim WS_SALDO_FINAL As Currency
Stop
Dim ADDPRE As rdoResultset
Dim WCOSTO As Currency
Dim sp_lot As rdoResultset
Dim ws_dif As Currency
Dim cto_fila As Integer

Stop

pub_cadena = "SELECT  * FROM ARTICULO WHERE ARM_CODCIA IN ('25','27','28','29','15','10')   "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
  cmdAgregar.Caption = X.RowCount & "/" & X.AbsolutePosition
  DoEvents
  'pub_cadena = "SELECT  * FROM LOTE WHERE LOT_CODCIA = '" & X!arm_codcia & "' AND LOT_SALDOs >= " & Abs(X!arm_ajuste) & " AND LOT_CODART = " & X!ARM_CODART & " "
  'Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
  'If Not sp_lot.EOF Then
  'sp_lot.Edit
  'sp_lot!lot_saldos = sp_lot!lot_saldos + Val(X!arm_ajuste)
  'sp_lot.Update
  'X.Edit
  'X!arm_ajuste = 0
  'X.Update
  
 ' End If
  
  
    
    'pub_cadena = "SELECT  * FROM LOTE WHERE LOT_CODCIA = '" & X!arm_codcia & "' AND LOT_CODART = " & X!ARM_CODART & " ORDER BY LOT_FECHA_VCTO DESC"
    'Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    'If sp_lot.EOF Then
    '    sp_lot.AddNew
    '    sp_lot!LOT_CODCIA = X!arm_codcia
    '    sp_lot!LOT_NROLOTE = "(*)(*)"
    '    sp_lot!lot_codart = X!ARM_CODART
    '    sp_lot!lot_fecha_vcto = "01/01/2012"
    '    sp_lot!LOT_SALDOS = Val(X!ARM_AJUSTA)
    '    sp_lot!lot_equiv = 1
    '    sp_lot!lot_fecha_ing = LK_FECHA_DIA
    '    sp_lot!lot_codclie = 0
    '    sp_lot!lot_numfac = 0
    '    sp_lot!lot_numguia = 0
    '    sp_lot!LOT_CODUSU = "ADMIN"
    '    sp_lot.Update
    '
    'Else'
'
 '   sp_lot.Edit
 '   sp_lot!LOT_SALDOS = sp_lot!LOT_SALDOS + Val(X!ARM_AJUSTA)
 '   sp_lot.Update
 '   End If
  
    pub_cadena = "SELECT  SUM(LOT_SALDOS) AS AJUSTE  FROM LOTE WHERE LOT_CODCIA = '" & X!ARM_CODCIA & "' AND LOT_CODART = " & X!ARM_CODART & " "
    Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    ws_dif = 0
    If Not sp_lot.EOF Then
     If Nulo_Valor0(sp_lot!AJUSTE) <> Val(X!arm_stock) Then ''
       X.Edit
       X!ARM_AJUSTE = Val(X!arm_stock) - Val(Nulo_Valor0(sp_lot!AJUSTE))
       X.Update
     End If
    End If
    
X.MoveNext
Loop

Stop



pub_cadena = "SELECT  * FROM ARTICULO WHERE ARM_CODCIA IN ('25','27','28','29') AND ARM_STOCK <> 0 "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
    pub_cadena = "select  sum(lot_saldos) as stock  from lotenuevo where lot_codcia in  ('" & X!ARM_CODCIA & "') and   lot_codart = " & X!ARM_CODART & " "
    Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    If Val(Format(Nulo_Valor0(X!arm_stock), "0.00")) <> Val(Format(Nulo_Valor0(sp_lot!stock), "0.00")) Then
        If Val(Format(Nulo_Valor0(X!arm_stock), "0.00")) - Val(Format(Nulo_Valor0(sp_lot!stock), "0.00")) < 0 Then
            pub_cadena = "select top 1 * from lotenuevo where lot_codcia in  ('" & X!ARM_CODCIA & "') and   lot_codart = " & X!ARM_CODART & " order by lot_fecha_vcto desc "
            Set ps_llave_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
            ps_llave_precios.Edit
            ps_llave_precios!LOT_NROLOTE = Trim(ps_llave_precios!LOT_NROLOTE) & " [.]"
            ps_llave_precios!LOT_SALDOS = ps_llave_precios!LOT_SALDOS + (Val(Nulo_Valor0(X!arm_stock)) - Val(Nulo_Valor0(sp_lot!stock)))
            ps_llave_precios.Update
        Else
           pub_cadena = "select * from lotenuevo where lot_codcia in  ('" & X!ARM_CODCIA & "') and  lot_codart = " & X!ARM_CODART & " "
           Set ps_llave_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    
            ps_llave_precios.AddNew
            ps_llave_precios!LOT_CODCIA = X!ARM_CODCIA
            ps_llave_precios!LOT_NROLOTE = "[.]"
            ps_llave_precios!lot_codart = X!ARM_CODART
            ps_llave_precios!lot_fecha_vcto = "01/01/2099"
            ps_llave_precios!LOT_SALDOS = (Val(Nulo_Valor0(X!arm_stock)) - Val(Nulo_Valor0(sp_lot!stock)))
            ps_llave_precios!lot_equiv = 1
            ps_llave_precios!lot_fecha_ing = LK_FECHA_DIA
            ps_llave_precios!lot_codclie = 0
            ps_llave_precios!lot_numfac = 0
            ps_llave_precios!lot_numguia = 0
            ps_llave_precios!LOT_CODUSU = ""
            ps_llave_precios.Update
        End If
    End If
    X.MoveNext
Loop
Stop

Stop
pub_cadena = "SELECT  * from prelista "
Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)


pub_cadena = "select  * from precios , arti where (art_codcia = pre_codcia) and (art_key = pre_codart) and pre_codcia in ('25')"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
'On Error GoTo pasasigue
Do Until X.EOF
'  On Error GoTo pasa
    pub_cadena = "SELECT  * from articulo where  arm_codcia = '25' and arm_codart = " & X!PRE_CODART
    Set arm_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
'   sp_lot.CancelUpdate
   ' 1
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 1
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE1
    sp_lot!p_precioIGV = X!PRE_PRE1
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE1 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE1
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 2
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 2
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE2
    sp_lot!p_precioIGV = X!PRE_PRE2
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE2 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE2
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 3
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 3
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE3
    sp_lot!p_precioIGV = X!PRE_PRE3
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE3 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE3
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 4
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 4
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE4
    sp_lot!p_precioIGV = X!PRE_PRE4
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE4 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE4
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   
   ' 5
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 5
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE5
    sp_lot!p_precioIGV = X!PRE_PRE5
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE5 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE5
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 6
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 6
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE6
    sp_lot!p_precioIGV = X!PRE_PRE6
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE6 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE6
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 7
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 7
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE7
    sp_lot!p_precioIGV = X!PRE_PRE7
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE7 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE7
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 8
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 8
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE8
    sp_lot!p_precioIGV = X!PRE_PRE8
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE8 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE8
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 9
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 9
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRE9
    sp_lot!p_precioIGV = X!PRE_PRE9
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRE9 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRE9
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 10
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 10
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PREC1
    sp_lot!p_precioIGV = X!PRE_PREC1
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PREC1 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PREC1
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 11
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 11
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PREC2
    sp_lot!p_precioIGV = X!PRE_PREC2
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PREC2 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PREC2
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 12
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 12
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_PRECHI1
    sp_lot!p_precioIGV = X!PRE_PRECHI1
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_PRECHI1 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_PRECHI1
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   ' 13
   sp_lot.AddNew
   sp_lot!p_codcia = X!PRE_CODCIA
   sp_lot!p_codart = X!PRE_CODART
   sp_lot!p_equiv = X!PRE_EQUIV
   sp_lot!p_codlista = 13
   sp_lot!p_SECUENCIA = X!PRE_SECUENCIA
   sp_lot!p_canti = X!PRE_CANT
   If X!ART_EX_IGV = "A" Then
    sp_lot!p_precio = X!PRE_OP2
    sp_lot!p_precioIGV = X!PRE_OP2
    sp_lot!p_costo = Format(arm_llave!ARM_COSPRO * X!PRE_EQUIV, "0.0000")
   Else
    sp_lot!p_precio = Format(X!PRE_OP2 / 1.19, "0.0000")
    sp_lot!p_precioIGV = X!PRE_OP2
    sp_lot!p_costo = Format((arm_llave!ARM_COSPRO * X!PRE_EQUIV) / 1.19, "0.0000")
   End If
   sp_lot.Update
   
   
   cmdAgregar.Caption = X.AbsolutePosition & "/" & X.RowCount
   DoEvents
X.MoveNext
Loop

Stop
Stop

pasasigue:
Debug.Print X!PRE_CODART
sp_lot.CancelUpdate
Return

Exit Sub
pub_cadena = "SELECT  par_codcia , par_art_cias FROM PARGEN WHERE par_art_cias <> ''"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
    X.Edit
    X!par_art_cias = "25" & Trim(X!par_art_cias)
    X.Update
  X.MoveNext
Loop





pub_cadena = "SELECT  par_codcia , par_art_cias FROM PARGEN WHERE par_art_cias <> ''"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
    X.Edit
    X!par_art_cias = "25" & Trim(X!par_art_cias)
    X.Update
  X.MoveNext
Loop

Stop
Stop
Stop

pub_cadena = "SELECT *  FROM PRECIOS WHERE PRE_EQUIV = 1 AND PRE_CODCIA = '09'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
  cmdAgregar.Caption = X.AbsolutePosition & " - " & X.RowCount
  DoEvents
   pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '09' AND PRE_EQUIV <> 1  AND PRE_CODART = " & X!PRE_CODART & ""
   Set far_pedidos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
   If far_pedidos.EOF Then
   Else
       wvalor = X!PRE_PRE1 * far_pedidos!PRE_EQUIV
      CN.Execute "update  PRECIOS SET PRE_PRE1 = " & wvalor & ",PRE_PRE2 = " & wvalor & " WHERE PRE_CODCIA = '09' AND PRE_CODART = " & X!PRE_CODART & " AND PRE_EQUIV <> 1 "
   End If
X.MoveNext
Loop

Stop
Stop
pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '" & LK_CODCIA & "' AND ART_SITUACION <> '1'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
cmdAgregar.Caption = X.AbsolutePosition & " - " & X.RowCount
   DoEvents
   pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '" & LK_CODCIA & "' AND ARM_CODART = " & X!art_key
   Set far_pedidos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
   
   pub_cadena = "SELECT TOP 1 FAR_FECHA_COMPRA , FAR_CODART, FAR_PRECIO FROM FACART , CLIENTES WHERE " & _
   "(FAR_CODCIA = CLI_CODCIA ) AND (FAR_CODCLIE = CLI_CODCLIE )  " & _
   "AND (FAR_CP = CLI_CP) AND FAR_CODCIA = '" & LK_CODCIA & "' AND FAR_TIPMOV = 20 AND FAR_PRECIO > 0.05 AND " & _
   " FAR_ESTADO <> 'E' AND FAR_FECHA >= '01/01/2005' AND CLI_LETRA <> 1  AND FAR_CODART = " & X!art_key & " ORDER BY FAR_FECHA_COMPRA DESC"
   Set ps_llave_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
   If Not ps_llave_precios.EOF Then
    far_pedidos.Edit
    far_pedidos!ARM_COSTO_ULT = ps_llave_precios!FAR_PRECIO
    far_pedidos.Update
   End If
  
X.MoveNext
Loop

MsgBox "OK"


Exit Sub

CN.Execute "SELECT * INTO mpresup_copyData from mpresup where mpr_codcia in ('01','03','05','07','09') and mpr_periodo = '072006'"
CN.Execute "update  mpresup_copyData set mpr_periodo = '082006'"
CN.Execute "insert into mpresup select  * from mpresup_copyData"
CN.Execute "drop table mpresup_copyData"

If xl Is Nothing Then
  Set xl = CreateObject("Excel.Application")
End If
DoEvents
xl.Workbooks.Open "C:\PROCO.xls", 0, True
fila = 1

xl.Application.Visible = True
Do Until Trim(xl.Cells(fila, 1)) = ""
    pub_cadena = "SELECT * FROM BDATOS.DBO.ARTI WHERE BDATOS.DBO.ARTI.ART_CODCIA = '05' AND BDATOS.DBO.ARTI.ART_ALTERNO = '" & Trim(xl.Cells(fila, 1)) & "'"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
    
    pub_cadena = "SELECT * FROM BDATOS.DBO.PRECIOS WHERE BDATOS.DBO.PRECIOS.PRE_CODCIA = '05' AND BDATOS.DBO.PRECIOS.PRE_CODART = " & X!art_key
    Set ps_llave_precios = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
    
    pub_cadena = "SELECT * FROM ARTI WHERE ARTI.ART_CODCIA = '05' AND ARTI.ART_ALTERNO = '" & Trim(xl.Cells(fila, 1)) & "'"
    Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
    
    pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '05' AND PRE_CODART = " & X!art_key
    Set far_pedidos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
    If far_pedidos.RowCount < ps_llave_precios.RowCount Then
    
    ElseIf far_pedidos.RowCount <> ps_llave_precios.RowCount Then
        Debug.Print Trim(xl.Cells(fila, 1))
'    Stop
    End If
 
fila = fila + 1
Loop
MsgBox "TEMRINO"

Exit Sub

loc_cp = "C"
pub_cadena = "SELECT CAA_CODCIA, CAA_CP,CAA_CODCLIE,CAA_TIPDOC,CAA_SERDOC, CAA_NUMDOC, sum(caa_importe) AS PROVI FROM CARACU WHERE CAA_CODCIA = '01' AND CAA_ESTADO <> 'E' AND CAA_CP = 'C' Group by CAA_CODCIA, CAA_CP,CAA_CODCLIE,CAA_TIPDOC,CAA_SERDOC, CAA_NUMDOC"
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
far_pedidos.Requery
Do Until far_pedidos.EOF
    cmdAgregar.Caption = far_pedidos.AbsolutePosition & "/" & far_pedidos.RowCount
    DoEvents
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = far_pedidos!CAA_CODCLIE
    pu_codcia = "01"
    PUB_SERDOC = far_pedidos!CAA_SERDOC
    PUB_NUMDOC = far_pedidos!CAA_NUMDOC
    PUB_TIPDOC = far_pedidos!CAA_TIPDOC
    LEER_CAR_LLAVE
    If Val(car_llave!car_importe) <> Val(far_pedidos!PROVI) Then
      Debug.Print far_pedidos!CAA_CODCLIE
    End If
        
  far_pedidos.MoveNext
Loop
xl.Application.Visible = True
MsgBox "OK"
Exit Sub




Exit Sub
'SELECT * FROM TABLAS WHERE TAB_CODCIA = '01'
'SELECT  * FROM PARGEN WHERE PAR_CODCIA = '01'
'SELECT  * FROM COPARAM WHERE COP_CODCIA = '01'
PUB_FECHA = LK_FECHA_DIA
PUB_NUMSER = 1
PUB_FBG = "F"
PUB_TIPMOV = 10
PUB_CODVEN = 1
PUB_CP = "C"
PUB_CODCLIE = 13
WDIRCLI = 19132

pub_cadena = "SELECT  FAR_CODART,  SUM(FAR_CANTIDAD) AS Cantidad, SUM((FAR_CANTIDAD/FAR_EQUIV) * FAR_PRECIO) AS VENTA FROM FACART WHERE FAR_TIPMOV = 10 AND FAR_CODCIA = '01' AND FAR_FBG = 'G' AND FAR_FECHA>= '01/11/2005' AND FAR_FECHA<= '30/11/2005' AND FAR_ESTADO2 <> 'L' AND FAR_ESTADO <> 'E' GROUP BY FAR_CODART ORDER BY VENTA"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
WS_SALDO_FINAL = 0
cto_fila = 1
PUB_NUMFAC = 93902
WS_NUMSEC = 0
Do Until X.EOF
 cmdAgregar.Caption = "Posicionde inicio : " & X.AbsolutePosition & " / " & X.RowCount
' X.AbsolutePosition = 1693
DoEvents
 If cto_fila >= 4 Then
  PUB_IMPTO = Format(WS_SALDO_FINAL - (WS_SALDO_FINAL / 1.19), "0.00")
  PUB_SUBTOTAL = Format((WS_SALDO_FINAL / 1.19), "0.00")
  CN.Execute "UPDATE FACART SET FAR_BRUTO = " & PUB_SUBTOTAL & " , FAR_IMPTO = " & PUB_IMPTO & "    WHERE FAR_CODCIA = '10' AND FAR_FBG = 'F' AND FAR_TIPMOV = 10 AND FAR_NUMSER = 1  AND FAR_NUMFAC = " & PUB_NUMFAC
  If PUB_NUMFAC = 93740 Then
    Stop
    MsgBox "ANOTAR"
  End If

  PUB_NUMFAC = PUB_NUMFAC + 1
  cto_fila = 0
  WS_NUMSEC = 0
  WS_SALDO_FINAL = 0
 End If
 WS_SALDO_FINAL = WS_SALDO_FINAL + Val(X!VENTA)
 
 If WS_SALDO_FINAL > 21215000 Then
  If cto_fila = 1 Then
  GoTo PIDEOTRO
  End If
  WS_SALDO_FINAL = WS_SALDO_FINAL - Val(X!VENTA)
  PUB_IMPTO = Format(WS_SALDO_FINAL - (WS_SALDO_FINAL / 1.19), "0.00")
  PUB_SUBTOTAL = Format((WS_SALDO_FINAL / 1.19), "0.00")
  CN.Execute "UPDATE FACART SET FAR_BRUTO = " & PUB_SUBTOTAL & " , FAR_IMPTO = " & PUB_IMPTO & "    WHERE FAR_CODCIA = '10' AND FAR_FBG = 'F' AND FAR_TIPMOV = 10 AND FAR_NUMSER = 1  AND FAR_NUMFAC = " & PUB_NUMFAC
  If PUB_NUMFAC = 93740 Then
    Stop
    MsgBox "ANOTAR"
  End If
  PUB_NUMFAC = PUB_NUMFAC + 1
  
  cto_fila = 0
  WS_NUMSEC = 0
  WS_SALDO_FINAL = 0
  WS_SALDO_FINAL = WS_SALDO_FINAL + Val(X!VENTA)
 End If
 cto_fila = cto_fila + 1
 
WS_NUMSEC = WS_NUMSEC + 1
pub_cadena = "SELECT  TOP 1 FAR_CODART,  FAR_CODLOT  AS LOTE FROM FACART WHERE FAR_TIPMOV = 10 AND FAR_CODCIA = '01' AND " & _
" FAR_FBG = 'G' AND FAR_FECHA>= '01/11/2005' AND FAR_FECHA<= '30/11/2005' AND " & _
" FAR_ESTADO2 <> 'L' AND FAR_ESTADO <> 'E' AND FAR_CODART = " & X!far_codart
Set pre_lotes = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

' incio
      far_llave.AddNew
        far_llave!FAR_tipmov = PUB_TIPMOV
        far_llave!FAR_CODCIA = LK_CODCIA
        far_llave!far_fbg = PUB_FBG
        far_llave!far_numser = PUB_NUMSER
        far_llave!FAR_CODVEN = PUB_CODVEN
        far_llave!far_numfac = PUB_NUMFAC
        far_llave!FAR_cp = PUB_CP
        far_llave!far_codclie = PUB_CODCLIE
        far_llave!FAR_dircli = WDIRCLI
        far_llave!FAR_RUC = ""
        far_llave!FAR_MONEDA = "S"
 
        far_llave!far_numsec = WS_NUMSEC
        far_llave!far_codart = Val(X!far_codart)
        far_llave!FAR_cantidad = Val(X!cantidad)
        far_llave!far_cantidad_p = Val(X!cantidad)
        far_llave!FAR_PRECIO = Format(Val(X!VENTA) / Val(X!cantidad), "0.0000")
        far_llave!far_codlot = pre_lotes!LOTE
        far_llave!far_impto = PUB_IMPTO
        far_llave!far_bruto = PUB_SUBTOTAL
        far_llave!far_subtotal = Val(X!VENTA)
        far_llave!FAR_fecha_compra = PUB_FECHA
        far_llave!far_precio_neto = Format(Val(X!VENTA) / Val(X!cantidad), "0.0000")
        SQ_OPER = 1
        pu_codcia = LK_CODCIA
        PUB_CODART = Val(X!far_codart)
        PUB_SECUEN = 0
        LEER_PRE_LLAVE
        far_llave!FAR_equiv = 1
        far_llave!far_descri = Trim(pre_llave!PRE_UNIDAD)
        
        
        far_llave!far_cod_sunat = "01" ' SIEMPRE FACTURA  PARA COMPRAS DE MERCADERIA
        far_llave!far_cliente = ""
        far_llave!FAR_DOCCLI = ""
        far_llave!FAR_STOCK = 0
        far_llave!far_signo_car = 0
        far_llave!far_signo_arm = -1
        far_llave!FAR_EX_IGV = ""
        far_llave!far_estado = "N"
        far_llave!FAR_ESTADO2 = "N"
        far_llave!FAR_COSPRO = 0
        far_llave!FAR_TOT_FLETE = 0
        far_llave!FAR_FLETE = 0
        far_llave!FAR_DESCTO = 0
        far_llave!far_PORDESCTOS = 0
        far_llave!FAR_TOT_DESCTO = 0
        far_llave!FAR_GASTOS = 0
        far_llave!FAR_NUMDOC = 0
        far_llave!far_numguia = 0
        far_llave!far_serguia = 0
        far_llave!FAR_pordescto1 = 0
        
        far_llave!FAR_tipo_cambio = 0
        far_llave!FAR_COSTEO = ""
        far_llave!FAR_COSTEO_REAL = ""
        far_llave!FAR_DIAS = 0
        far_llave!far_fecha = PUB_FECHA
        far_llave!FAR_NUMSER_C = 0
        far_llave!FAR_NUMFAC_C = 0
        far_llave!FAR_NUMOPER = 0
        far_llave!FAR_NUMOPER2 = 0
        
        far_llave!far_subtra = "VENTA"
        far_llave!far_transito = ""
        far_llave!far_otra_cia = ""
        
        far_llave!far_JABAS = 0
        far_llave!far_UNIDADES = 0
        far_llave!far_mortal = 0
        far_llave!far_num_precio = 0
        far_llave!far_PESO = 0
        far_llave!FAR_LITRO = 0
        far_llave!FAR_TIPDOC = "FA"
        far_llave!FAR_ORDEN_UNIDADES = 0
        far_llave!far_concepto = ""
        far_llave!far_codusu = PUB_CODUSU
        far_llave!far_hora = ""
        far_llave!FAR_NUM_LOTE = 0
      
        far_llave!FAR_PEDSER = 0
        far_llave!FAR_PEDFAC = 0
        far_llave!far_turno = 0
        far_llave!far_pedsec = 0
        far_llave!FAR_FLAG_SO = ""
        far_llave!FAR_NUMFAC_C = 0
        far_llave!FAR_NUMSER_C = 0
        far_llave!far_key_dircli = 0
        far_llave!FAR_OC = ""
        far_llave!FAR_fecha_pro = PUB_FECHA
        far_llave!FAR_fecha_can = PUB_FECHA
        far_llave!far_turno = PUB_TURNO
        far_llave.Update
        CN.Execute "UPDATE FACART SET FAR_CONCEPTO = '" & PUB_FBG & "/" & PUB_NUMSER & "-" & PUB_NUMFAC & "'   WHERE FAR_CODCIA = '01' AND FAR_FBG = 'G' AND FAR_FECHA>= '01/11/2005' AND FAR_FECHA<= '30/11/2005' AND  FAR_ESTADO2 <> 'L' AND FAR_ESTADO <> 'E' AND FAR_CODART = " & X!far_codart
  
PIDEOTRO:

X.MoveNext
Loop


MsgBox "TERMINO"
Stop


Exit Sub

pub_cadena = "SELECT  * FROM RELCOMPRA WHERE REL_LIQUIDO <> 1 AND  REL_CODCIA = '" & LK_CODCIA & "'"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
Do Until X.EOF
  Print X.AbsolutePosition
    
    pub_cadena = "SELECT  * FROM ALLOG WHERE ALL_CODCIA = '" & LK_CODCIA & "' AND ALL_TIPMOV = 20 AND ALL_CODTRA = 1401 AND ALL_CODCLIE = " & X!rel_codpro & "  AND ALL_CP = 'P' AND ALL_FLAG_EXT <> 'E' AND ALL_NUMSER = " & X!rel_numser & "  AND ALL_NUMFAC = " & X!rel_numfac & ""
    Set far_pedidos = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)  ' rdConcurReadOnly) ', rdConcurLock)
    If far_pedidos.EOF Then
       X.Edit
       X!rel_LIQUIDO = "9"
       X.Update
    Else

'    Stop
    End If

X.MoveNext
Loop
  Stop
pub_cadena = "SELECT  * FROM PRECIOS WHERE PRE_CODCIA = '01'"
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
far_pedidos.Requery
Do Until far_pedidos.EOF
    cmdAgregar.Caption = far_pedidos.AbsolutePosition & " / " & far_pedidos.RowCount
    DoEvents
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = far_pedidos!PRE_CODART
    PUB_SECUEN = far_pedidos!PRE_SECUENCIA
    LEER_PRE_LLAVE
    pre_llave.Edit
    For fila = 3 To far_pedidos.rdoColumns.count - 1
       If Not IsNull(far_pedidos(fila)) Then pre_llave(fila) = far_pedidos(fila)
    Next fila
    pre_llave.Update
far_pedidos.MoveNext
Loop

MsgBox "ok"
Exit Sub


pub_cadena = "SELECT * FRMO FACART WHERE FAR_CODCIA = '01' AND FAR_TIPMOV = 5 AND FAR_SUBTRA LIKE '%inciner%' AND FAR_FECHA>= '01/01/2005'"
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
far_pedidos.Requery
Do Until far_pedidos.EOF
   pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = '01' AND ALL_FECHA_DIA = '" & 0 & "' AND ALL_NUM_OPER = " & far_pedidos!FAR_NUMOPER
   Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If sp_lot.EOF Then
   Stop
   Else
     If far_pedidos!far_codclie <> sp_lot!ALL_CODCLIE Then
       Stop
       far_pedidos.Edit
       far_pedidos!far_codcli = sp_lot!ALL_CODCLIE
       far_pedidos.Update
       
     End If
   End If
   

far_pedidos.MoveNext
Loop



Exit Sub



pub_cadena = "SELECT  * FROM ARTI"
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
far_pedidos.Requery


Exit Sub


If xl Is Nothing Then
  Set xl = CreateObject("Excel.Application")
End If
DoEvents
 xl.Workbooks.Open "C:\DATOS.xls", 0, True
 fila = 0
  
  
loc_cp = "C"
pub_cadena = "SELECT CAA_CODCIA, CAA_CP,CAA_CODCLIE,CAA_TIPDOC,CAA_SERDOC, CAA_NUMDOC, sum(caa_importe) AS PROVI FROM CARACU WHERE CAA_CODCIA = '01' AND CAA_ESTADO <> 'E' AND CAA_CP = 'C' Group by CAA_CODCIA, CAA_CP,CAA_CODCLIE,CAA_TIPDOC,CAA_SERDOC, CAA_NUMDOC"
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)
far_pedidos.Requery
Do Until far_pedidos.EOF
    cmdAgregar.Caption = far_pedidos.AbsolutePosition & "/" & far_pedidos.RowCount
    DoEvents
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = far_pedidos!CAA_CODCLIE
    pu_codcia = "01"
    PUB_SERDOC = far_pedidos!CAA_SERDOC
    PUB_NUMDOC = far_pedidos!CAA_NUMDOC
    PUB_TIPDOC = far_pedidos!CAA_TIPDOC
    LEER_CAR_LLAVE
    If Val(car_llave!car_importe) <> Val(far_pedidos!PROVI) Then
      Debug.Print far_pedidos!CAA_CODCLIE
    End If
        
  far_pedidos.MoveNext
Loop
xl.Application.Visible = True
MsgBox "OK"
Exit Sub













' COPIAR OTRA CIA
' pub_cadena = "SELECT  * FROM COPARAM WHERE COP_CODCIA = '01'"
pub_cadena = "SELECT  * FROM PARGEN WHERE PAR_CODCIA = '01'"
'pub_cadena = "SELECT  * FROM TABLAS WHERE TAB_CODCIA = '01'"
pub_cadena = "SELECT  * FROM CLIENTES WHERE CLI_CODCIA = '01' AND CLI_CP = 'C' AND CLI_CODCLIE = 1 "
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)

'pub_cadena = "SELECT * FROM ARTI WHERE ART_FAMILI"
' pub_cadena = "SELECT  * FROM COPARAM WHERE COP_CODCIA = '01'"
'pub_cadena = "SELECT  * FROM TABLAS WHERE TAB_CODCIA = '01'"
pub_cadena = "SELECT  * FROM PARGEN WHERE PAR_CODCIA = '01'"
pub_cadena = "SELECT  * FROM CLIENTES WHERE CLI_CODCIA = '01' AND CLI_CP = 'C' AND CLI_CODCLIE = 1 "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery
Do Until WfART_llave.EOF
    far_pedidos.AddNew
    For fila = 0 To WfART_llave.rdoColumns.count - 1
       If IsNull(WfART_llave(fila)) Then
       Else
         far_pedidos(fila) = WfART_llave(fila)
       End If
    Next fila
    far_pedidos!CLI_CODCIA = "07"
    far_pedidos.Update
  WfART_llave.MoveNext
Loop
MsgBox "TE"
Exit Sub
'**************************************************
' DIFERECNIAS EN LOTES
pub_cadena = "SELECT ARM_CODART,ARM_STOCK, SUM(LOT_SALDOS) AS SALLOT FROM ARTICULO, LOTE  WHERE " & _
             "ARM_CODCIA = '" & LK_CODCIA & "'  AND (ARM_CODCIA = LOT_CODCIA) AND (ARM_CODART = LOT_CODART) " & _
             "GROUP BY ARM_CODART,ARM_STOCK Having ((ARM_STOCK <> Sum(LOT_SALDOS))) AND Sum(LOT_SALDOS) <> 0 "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery
Do Until WfART_llave.EOF
   ws_dif = WfART_llave!arm_stock - WfART_llave!SALLOT
   pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & WfART_llave!ARM_CODART & " "
   Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If sp_lot.EOF Then
     GoTo pasa_lo800
   Else
     Do Until sp_lot.EOF
     ' despues pasar con los lotes que tiene 0 y aumentarles
        If sp_lot!LOT_SALDOS >= Abs(ws_dif) Then
           sp_lot.Edit
           sp_lot!LOT_SALDOS = sp_lot!LOT_SALDOS + ws_dif
           sp_lot.Update
           Exit Do
        End If
      sp_lot.MoveNext
     Loop
   End If
   
pasa_lo800:
 WfART_llave.MoveNext
Loop
 
MsgBox "termino"

Exit Sub

'**************************************************
' AJUSTAR LOS LOTES NEGATIVOS
'**************************************************

pub_cadena = "SELECT DISTINCT LOT_CODART , LOT_SALDOS FROM LOTE WHERE LOT_SALDOS < 0 AND LOT_CODCIA = '" & LK_CODCIA & "'"
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery
Do Until WfART_llave.EOF
   pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & WfART_llave!lot_codart & " AND LOT_SALDOS > 0"
   Set sp_lot = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
   If sp_lot.EOF Then
'    Stop
'    MsgBox "revisar "
    GoTo pasa_caba
   Else
    Do Until sp_lot.EOF
       If Val(sp_lot!LOT_SALDOS) >= Val(Abs(WfART_llave!LOT_SALDOS)) Then
         sp_lot.Edit
         sp_lot!LOT_SALDOS = sp_lot!LOT_SALDOS - Val(Abs(WfART_llave!LOT_SALDOS))
         sp_lot.Update
         Exit Do
       Else
        Debug.Print pub_cadena
       End If
      sp_lot.MoveNext
    Loop
   End If
pasa_caba:
WfART_llave.MoveNext
Loop
MsgBox "TERMINO"
Exit Sub








'**************************************************
'**************************************************
pub_cadena = "SELECT DISTINCT FAR_CODCLIE , FAR_CODVEN  FROM FACART WHERE FAR_CODCIA = '01' " & _
             " AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND FAR_FECHA >= '01/10/2003' and far_codclie <> 1 " & _
             "GROUP BY FAR_CODCLIE ,   FAR_CODVEN"
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery

Do Until WfART_llave.EOF
SQ_OPER = 1
pu_cp = "C"
pu_codclie = WfART_llave!far_codclie
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE
If cli_llave.EOF Then
Stop
Else
cli_llave.Edit
cli_llave!cli_dia_visita = WfART_llave!FAR_CODVEN
cli_llave.Update
End If


WfART_llave.MoveNext
Loop

MsgBox ""
Exit Sub
'************************

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '01' AND ART_SITUACION <> '1' "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = '03' AND ART_KEY = ? "
Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
PSPRE_LOTES(0) = 0
Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)
fila = 0
Do Until WfART_llave.EOF
cmdAgregar.Caption = WfART_llave.AbsolutePosition & "/" & WfART_llave.RowCount
DoEvents
 PSPRE_LOTES(0) = WfART_llave!art_key
 pre_lotes.Requery
 If Not pre_lotes.EOF Then
   If pre_lotes!art_situacion = "1" Then
     pre_lotes.Edit
     pre_lotes!art_situacion = "0"
     pre_lotes.Update
        fila = fila + 1
     Debug.Print "Act." & fila
   Else
   End If
 Else
 Stop
 End If


WfART_llave.MoveNext
Loop


MsgBox "listo"

Exit Sub

pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '03' AND ARM_STOCK <> 0"
Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)
pre_lotes.Requery

pub_cadena = "SELECT FAR_CODCIA FROM FACART WHERE FAR_CODCIA = ?  AND FAR_CODART = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
WPSART_LLAVE.MaxRows = 1
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


Do Until pre_lotes.EOF
WPSART_LLAVE(0) = LK_CODCIA
WPSART_LLAVE(1) = pre_lotes!ARM_CODART
WfART_llave.Requery
If WfART_llave.EOF Then
  Debug.Print "sTOCK " & pre_lotes!arm_stock & " cODIGO " & pre_lotes!ARM_CODART
End If
DoEvents
cmdAgregar.Caption = " " & pre_lotes.AbsolutePosition & "/" & pre_lotes.RowCount
pre_lotes.MoveNext
Loop



Exit Sub

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ?  AND PRE_CODART = 1117505  "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WfART_llave.Requery
Do Until WfART_llave.EOF
cmdAgregar.Caption = WfART_llave.AbsolutePosition & "/" & WfART_llave.RowCount
DoEvents
SQ_OPER = 1
PUB_CODART = WfART_llave!PRE_CODART
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE

WCOSTO = arm_llave!ARM_COSPRO * WfART_llave!PRE_EQUIV

WfART_llave.Edit

If WCOSTO = 0 Then GoTo COLOCAR_CERO

VALOR = (WfART_llave!PRE_PRE1 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR1 = Format(VALOR, "0.000")

VALOR = (WfART_llave!PRE_PRE2 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR2 = Format(VALOR, "0.000")

VALOR = (WfART_llave!PRE_PRE3 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR3 = Format(VALOR, "0.000")

VALOR = (WfART_llave!PRE_PRE4 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR4 = Format(VALOR, "0.000")

VALOR = (WfART_llave!PRE_PRE5 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR5 = Format(VALOR, "0.000")

VALOR = (WfART_llave!PRE_PRE6 * 100) / Val(WCOSTO) - 100
If VALOR = -100 Then VALOR = 0
WfART_llave!PRE_POR6 = Format(VALOR, "0.000")
GoTo ESTABIENOK

COLOCAR_CERO:
WfART_llave!PRE_POR1 = 0
WfART_llave!PRE_POR2 = 0
WfART_llave!PRE_POR3 = 0
WfART_llave!PRE_POR4 = 0
WfART_llave!PRE_POR5 = 0
WfART_llave!PRE_POR6 = 0

ESTABIENOK:
WfART_llave.Update

WfART_llave.MoveNext
Loop

MsgBox "STOP"



Exit Sub


Exit Sub

SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE

'pub_cadena = "SELECT FAR_NUMFAC, far_codart, FAR_DESCRI , FAR_EQUIV FROM FACART WHERE FAR_CODCIA = '01' AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 20 AND " & _
'             "FAR_CODART = ?"
'Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
'PSFAR_PEDIDOS(0) = 0
'Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurValues)

Do Until art_mayor.EOF
 SQ_OPER = 1
 PUB_CODART = art_mayor!art_key
 pu_codcia = LK_CODCIA
 LEER_ARM_LLAVE
 ACTLIVES = "PRE_COSTO = (" & arm_llave!ARM_COSPRO & "* PRE_EQUIV)"
          pub_cadena = "UPDATE PRECIOS SET " & ACTLIVES & " WHERE  PRE_CODCIA = '03' AND PRE_CODART = " & art_mayor!art_key
          CN.Execute pub_cadena, rdExecDirect

art_mayor.MoveNext
Loop

MsgBox "listo"
Exit Sub
'pub_cadena = "SELECT FAR_CODART FROM FACART WHERE FAR_CODCIA = '01' AND FAR_ESTADO <> 'E' AND " & _
'             "FAR_DESCRI = ? AND FAR_EQUIV = ?  AND FAR_CODART = ? "
'Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
'PSPRE_LOTES(0) = 0
'PSPRE_LOTES(1) = 0
'PSPRE_LOTES(2) = 0
'PSPRE_LOTES.MaxRows = 1
'Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = " select * from precios where pre_unidad = ? and " & _
             " pre_equiv = ? and pre_codart = ? and pre_codcia = '01' "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
WPSART_LLAVE(2) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)



Do Until art_mayor.EOF
   PSFAR_PEDIDOS(0) = art_mayor!art_key
   far_pedidos.Requery
   Do Until far_pedidos.EOF
        WPSART_LLAVE(0) = Trim(far_pedidos!far_descri)
        WPSART_LLAVE(1) = far_pedidos!FAR_equiv
        WPSART_LLAVE(2) = art_mayor!art_key
        WfART_llave.Requery
        If WfART_llave.EOF Then
            Debug.Print Left(art_mayor!art_nombre, 40) & " Nro COmpra : " & far_pedidos!far_numfac & " COD: " & far_pedidos!far_codart
        End If
    far_pedidos.MoveNext
   Loop
   art_mayor.MoveNext
Loop


MsgBox TERMINO
Exit Sub
SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE

pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? and FAR_CODART = ? ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA,FAR_NUMOPER"
Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
PSPRE_LOTES(0) = 0
PSPRE_LOTES(1) = 0
PSPRE_LOTES.MaxRows = 1
Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)




Do Until art_mayor.EOF
   PSPRE_LOTES(0) = LK_CODCIA
   PSPRE_LOTES(1) = art_mayor!art_key
   pre_lotes.Requery
   DoEvents
   cmdAgregar.Caption = art_mayor.AbsolutePosition & " " & art_mayor.RowCount
   If Not pre_lotes.EOF Then
     If pre_lotes!FAR_STOCK = 0 Then GoTo saltat
     If pre_lotes!FAR_STOCK <> pre_lotes!FAR_cantidad Then
         Debug.Print art_mayor!ART_alterno & " " & pre_lotes!FAR_fecha_compra & " " & pre_lotes!FAR_STOCK & " " & art_mayor!art_nombre
         Stop
         pre_lotes.Edit
         pre_lotes!FAR_STOCK = 0
         pre_lotes.Update
         
        
     End If
   End If
saltat:
   art_mayor.MoveNext
Loop

Exit Sub


pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = ? and LOT_CODART = ? and lot_nrolote = ? ORDER BY LOT_FECHA_VCTO DESC "
Set PSPRE_LOTES = CN.CreateQuery("", pub_cadena)
PSPRE_LOTES(0) = 0
PSPRE_LOTES(1) = 0
PSPRE_LOTES(2) = 0
Set pre_lotes = PSPRE_LOTES.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODCIA = '" & LK_CODCIA & "'"
Set ADDPRE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
ADDPRE.Requery

pub_cadena = "SELECT * FROM LOTE WHERE LOT_CODCIA = ? and LOT_CODART = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


'pub_cadena = "UPDATE LOTE SET LOT_SALDOS = 0 WHERE  LOT_CODCIA = '" & LK_CODCIA & "' "
'CN.Execute pub_cadena, rdExecDirect

Print ADDPRE.RowCount
fila = 0
Do Until ADDPRE.EOF
    WPSART_LLAVE(0) = LK_CODCIA
    WPSART_LLAVE(1) = ADDPRE!ARM_CODART
    WfART_llave.Requery
    WR_SUMA = 0
    Do Until WfART_llave.EOF
      WR_SUMA = WR_SUMA + Val(Nulo_Valor0(WfART_llave!LOT_SALDOS))
      WfART_llave.MoveNext
    Loop
    If Val(ADDPRE!arm_stock) <> Val(WR_SUMA) Then
      fila = fila + 1
      'pub_cadena = "UPDATE LOTE SET LOT_SALDOS = 0 WHERE  LOT_CODCIA = '" & LK_CODCIA & "' AND LOT_CODART = " & ADDPRE!ARM_CODART
      'CN.Execute pub_cadena, rdExecDirect
      'pub_cadena = "select far_codlot from facart where far_tipmov = 20 and far_fecha_compra >= '01/05/2004' and far_estado <> 'E' and far_codart <> 0  and far_codart = " & ADDPRE!ARM_CODART & " order by far_fecha_compra desc"
      'Set usu = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
      'If usu.EOF Then
      ' Stop
      'End If
dale_unamas:
      'PSPRE_LOTES(0) = LK_CODCIA
      'PSPRE_LOTES(1) = ADDPRE!ARM_CODART
      'PSPRE_LOTES(2) = usu!far_codlot
      'pre_lotes.Requery
      'If pre_lotes.EOF Then
      '  Stop
      '  usu.MoveNext
      '  GoTo dale_unamas
      'Else
      '  pre_lotes.Edit
      '  pre_lotes!LOT_SALDOS = Val(ADDPRE!arm_stock)
      '  pre_lotes.Update
      'End If
      
      Debug.Print fila & " - " & ADDPRE!ARM_CODART & " " & ADDPRE!arm_stock & " " & WR_SUMA
    ' Stop
    End If
      Print ADDPRE.RowCount / ADDPRE.AbsolutePosition
    ADDPRE.MoveNext
Loop
MsgBox "Ya"
Exit Sub










pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = '99'"
Set ADDPRE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

Dim WUN As String
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? and PRE_SECUENCIA = 1 "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

Print WfART_llave.RowCount
Do Until WfART_llave.EOF
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = WfART_llave!PRE_CODART
    LEER_PRE_LLAVE
    WUN1 = pre_mayor!PRE_UNIDAD
    pre_mayor.MoveNext
    WUN2 = pre_mayor!PRE_UNIDAD
    pre_mayor.MoveFirst
    pre_mayor.Edit
    pre_mayor!PRE_UNIDAD = WUN2
    pre_mayor.Update
    pre_mayor.MoveNext
    pre_mayor.Edit
    pre_mayor!PRE_UNIDAD = "CAJA"
    pre_mayor.Update
    
WfART_llave.MoveNext
Loop
MsgBox "Ya"
Exit Sub
pub_cadena = "SELECT PED_FECHA,PED_NUMSER ,PED_NUMFAC, PED_CONDI,PED_CODVEN , PED_CODART , PED_CANTIDAD, PED_EQUIV , PED_CODCLIE FROM PEDIDOS " & _
"WHERE PED_CODCIA = ? AND PED_TIPMOV = 201 AND PED_CODART = ? AND PED_SITUACION not in ('P','E')  "
Set PSFAR_PEDIDOS = CN.CreateQuery("", pub_cadena)
PSFAR_PEDIDOS(0) = 0
PSFAR_PEDIDOS(1) = 0
Set far_pedidos = PSFAR_PEDIDOS.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT far_cantidad, Far_signo_arm  FROM FACART WHERE FAR_CODCIA = ? and FAR_CODART = ? ORDER BY FAR_FECHA , FAR_NUMOPER"
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
WPSART_LLAVE(1) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
WS_SALDO_FINAL = 0
Do Until art_mayor.EOF
   WS_SALDO_FINAL = 0
   ' If art_mayor!ART_KEY <> 28177 Then GoTo VA_STOCK
   WPSART_LLAVE(0) = LK_CODCIA
   WPSART_LLAVE(1) = Val(art_mayor!art_key)
   WfART_llave.Requery
   Do Until WfART_llave.EOF
      WS_SALDO_FINAL = WS_SALDO_FINAL + (Val(WfART_llave!FAR_cantidad) * Val(WfART_llave!far_signo_arm))
      WfART_llave.MoveNext
   Loop
   
   PSFAR_PEDIDOS(0) = LK_CODCIA
   PSFAR_PEDIDOS(1) = Val(art_mayor!art_key)
   far_pedidos.Requery
   Do Until far_pedidos.EOF
      WS_SALDO_FINAL = WS_SALDO_FINAL + (Val(far_pedidos!PED_CANTIDAD) * -1)
      far_pedidos.MoveNext
   Loop
   If WS_SALDO_FINAL = 0 Then GoTo VA_STOCK
   SQ_OPER = 1
   PUB_CODART = Val(art_mayor!art_key)
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   arm_llave.Edit
   If Val(arm_llave!arm_stock) <> Val(WS_SALDO_FINAL) Then
     arm_llave!arm_stock = WS_SALDO_FINAL
   End If
   arm_llave.Update
VA_STOCK:
art_mayor.MoveNext

Loop
MsgBox "STOCK"
Exit Sub
Stop
Stop
Stop

If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents

WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\ARTICULO.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2
Do Until WS_FILA = 965000
 If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
 WS_FILA = WS_FILA + 1
 cmdAgregar_Click
 art_familia.ListIndex = 0
 txt_alterno.Text = Format(Val(Trim(xl.Cells(WS_FILA, 1))), "#######")
 ' Txt_key.Text = Trim(xl.Cells(WS_FILA, 1))
 txtnombre.Text = Trim(xl.Cells(WS_FILA, 2))
 PUB_IMPORTE = Val(xl.Cells(WS_FILA, 4))
 PUB_IMPORTE_AMORT = Val(xl.Cells(WS_FILA, 5))
 PUB_NUMFAC = Val(xl.Cells(WS_FILA, 6))
 
 
   If Trim(xl.Cells(WS_FILA, 8)) > 1 Then
     grid_UNID_KeyDown 45, 0
     grid_unid.TextMatrix(1, 0) = "UND"
'     grid_unid.TextMatrix(1, 1) = 1
     grid_unid.TextMatrix(1, 8) = 3
     grid_unid.TextMatrix(1, 14) = ""
     grid_unid.TextMatrix(1, 5) = Format((Val(xl.Cells(WS_FILA, 14)) * 1.18), "0.0000")
     grid_unid.TextMatrix(1, 21) = Format((Val(xl.Cells(WS_FILA, 14)) * 1.18), "0.0000") ' DOLARES
      grid_unid.TextMatrix(1, 28) = Trim(xl.Cells(WS_FILA, 7)) ' CANTIDAD DE LITROS
      grid_unid.TextMatrix(1, 26) = Trim(xl.Cells(WS_FILA, 7)) ' CANTIDAD DE PESO
      ' UNIDAD MAX
     grid_unid.TextMatrix(2, 8) = 3
     grid_unid.TextMatrix(2, 14) = "A"
     grid_unid.TextMatrix(2, 0) = Trim(xl.Cells(WS_FILA, 3))
     grid_unid.TextMatrix(2, 1) = Trim(xl.Cells(WS_FILA, 8))
     grid_unid.TextMatrix(2, 5) = Format(Val(grid_unid.TextMatrix(1, 5)) * Val(grid_unid.TextMatrix(2, 1)), "0.00")
     ' grid_unid.TextMatrix(1, 21) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(2, 21) = Format(Val(grid_unid.TextMatrix(1, 5)) * Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(2, 28) = Format(Val(grid_unid.TextMatrix(1, 7)) * Val(grid_unid.TextMatrix(2, 1)), "0.00")
     grid_unid.TextMatrix(2, 26) = Format(Val(grid_unid.TextMatrix(1, 7)) * Val(grid_unid.TextMatrix(2, 1)), "0.00")
    Else
      grid_unid.TextMatrix(1, 0) = Trim(xl.Cells(WS_FILA, 3))
      grid_unid.TextMatrix(1, 8) = 3
      ' grid_unid.TextMatrix(1, 1) = 1
      grid_unid.TextMatrix(1, 5) = Format((Val(xl.Cells(WS_FILA, 14)) * 1.18), "0.00")
      grid_unid.TextMatrix(1, 21) = Format((Val(xl.Cells(WS_FILA, 14)) * 1.18), "0.00") ' DOLARES
      grid_unid.TextMatrix(1, 28) = Trim(xl.Cells(WS_FILA, 7)) ' CANTIDAD DE LITROS
      grid_unid.TextMatrix(1, 26) = Trim(xl.Cells(WS_FILA, 7)) ' CANTIDAD DE PESO

    End If
 
 
  cmdAgregar_Click
 
  
  
AVAAva99:
Loop

MsgBox "TERMINO PROCESO"
' TERMINO DE PROCESO
Exit Sub


' CODCIA DIRCLI      CODCLI     CP   DIREC
' REF                                                CLI_LUGAR_TRAB CLI_TRAB_ZONA CLI_CASA_SUBZONA CLI_TRAB_SUBZONA NUMERO      DIRCOMP
pub_cadena = "SELECT  * FROM FACART " & _
             "WHERE FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' AND FAR_NUM_PRECIO = '0' AND FAR_CODCIA ='01' " & _
             "ORDER BY FAR_CODVEN ,FAR_CODART, FAR_FECHA "
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)
art_rela.Requery

Do Until art_rela.EOF
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_CODART = art_rela!far_codart
  PUB_SECUEN = 0
  LEER_PRE_LLAVE
  If pre_llave.EOF Then
  Else
    art_rela.Edit
    If Val(art_rela!FAR_PRECIO) = Val(pre_llave!PRE_PRE1) Then
      art_rela!far_num_precio = 1
    ElseIf Val(art_rela!FAR_PRECIO) = Val(pre_llave!PRE_PRE2) Then
      art_rela!far_num_precio = 2
    ElseIf Val(art_rela!FAR_PRECIO) = Val(pre_llave!PRE_PRE3) Then
      art_rela!far_num_precio = 3
    ElseIf Val(art_rela!FAR_PRECIO) = Val(pre_llave!PRE_PRE4) Then
      art_rela!far_num_precio = 4
    ElseIf Val(art_rela!FAR_PRECIO) = Val(pre_llave!PRE_PRE5) Then
      art_rela!far_num_precio = 5
    Else
     If Val(art_rela!FAR_PRECIO) = 0 Then
       art_rela!far_num_precio = 4
     Else
       Stop
     End If
    Print art_rela!far_PORDESCTOS

    End If
    art_rela.Update
  End If
  art_rela.MoveNext
Loop

MsgBox "LISTO"


Exit Sub

pub_cadena = "select * from DIRCLI where CODCIA = ? and CODCLI = ? AND CP = ?"
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA(0) = 0
PSART_RELA(1) = 0
PSART_RELA(2) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_RELA(0) = LK_CODCIA

SQ_OPER = 2
pu_codclie = 0
pu_cp = "C"
pu_codcia = LK_CODCIA
LEER_CLI_LLAVE

Do Until cli_mayor.EOF
    PSART_RELA(0) = LK_CODCIA
    PSART_RELA(1) = cli_mayor!cli_codclie
    PSART_RELA(2) = "C"
    art_rela.Requery
    art_rela.Edit
    If Val(art_rela!cli_TRAB_SUBZONA) = 1 Then
        art_rela!cli_TRAB_SUBZONA = cli_mayor!cli_zona_new
        SQ_OPER = 1
        PUB_TIPREG = 35
        PUB_NUMTAB = cli_mayor!cli_zona_new
        PUB_CODCIA = "00"
        LEER_TAB_LLAVE
        If Not tab_llave.EOF Then
          art_rela!DIRCOMP = art_rela!DIRCOMP & " - " & Trim(tab_llave!TAB_NOMLARGO)
        Else
          Stop
        End If
    End If
    art_rela.Update
    cmdAgregar.Caption = cli_mayor.RowCount & "/" & cli_mayor.AbsolutePosition
    DoEvents
 cli_mayor.MoveNext
Loop


MsgBox " termino "
Exit Sub


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents

WPAS = "131296"
xl.Workbooks.Open "C:\LLEVO\CLIENTES.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2
Do Until WS_FILA = 965000
 If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
 SQ_OPER = 3
 pu_alterno = xl.Cells(WS_FILA, 1)
 pu_codcia = LK_CODCIA
 LEER_ART_LLAVE
 If art_llave_alt.EOF Then
   MsgBox "Notar codigo No Existe ...." & pu_alterno & " " & WS_FILA
   GoTo Ava99
 End If
 SQ_OPER = 1
 PUB_CODART = art_llave_alt!art_key
 pu_codcia = LK_CODCIA
 LEER_ARM_LLAVE
 If arm_llave.EOF Then
   MsgBox " FALLO "
 End If
' GoTo Ava99
 SQ_OPER = 2
 pu_codcia = LK_CODCIA
 PUB_CODART = arm_llave!ARM_CODART
 LEER_PRE_LLAVE
 Do Until pre_mayor.EOF
   pre_mayor.Edit
   If Val(pre_mayor!PRE_EQUIV) = 1 Then
   
        pre_mayor!PRE_PRE1 = Format(Val(xl.Cells(WS_FILA, 2)), "0.00")
        pre_mayor!PRE_PRE2 = Format(Val(xl.Cells(WS_FILA, 3)), "0.00")
        pre_mayor!PRE_PRE3 = Format(Val(xl.Cells(WS_FILA, 4)), "0.00")
        pre_mayor!PRE_PRE4 = Format(Val(xl.Cells(WS_FILA, 5)), "0.00")
        pre_mayor!PRE_PRE5 = 0
        pre_mayor!PRE_POR1 = 0
        pre_mayor!PRE_POR2 = 2
        pre_mayor!PRE_POR3 = 3
        pre_mayor!PRE_POR4 = 3
        pre_mayor!PRE_POR5 = 0
   Else
     ' MsgBox art_llave_alt!art_nombre
        pre_mayor!PRE_PRE1 = Format(Val(xl.Cells(WS_FILA, 2)) * Val(pre_mayor!PRE_EQUIV), "0.00")
        pre_mayor!PRE_PRE2 = Format(Val(xl.Cells(WS_FILA, 3)) * Val(pre_mayor!PRE_EQUIV), "0.00")
        pre_mayor!PRE_PRE3 = Format(Val(xl.Cells(WS_FILA, 4)) * Val(pre_mayor!PRE_EQUIV), "0.00")
        pre_mayor!PRE_PRE4 = Format(Val(xl.Cells(WS_FILA, 5)) * Val(pre_mayor!PRE_EQUIV), "0.00")
        pre_mayor!PRE_PRE5 = 0
        pre_mayor!PRE_POR1 = 0
        pre_mayor!PRE_POR2 = 2
        pre_mayor!PRE_POR3 = 3
        pre_mayor!PRE_POR4 = 3
        pre_mayor!PRE_POR5 = 0
   End If
   pre_mayor.Update
 pre_mayor.MoveNext
 Loop
 
 
 
 
 
Ava99:
 WS_FILA = WS_FILA + 1
Loop

MsgBox "TERMINO PROCESO"
' TERMINO DE PROCESO
Exit Sub




pub_cadena = "select pa_codpa, sum(pa_prom) as tot  from paquetes where pa_codpa <> 0 and pa_codcia = ? group by pa_codpa"
Set PSART_RELA = CN.CreateQuery("", pub_cadena)
PSART_RELA(0) = 0
Set art_rela = PSART_RELA.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_RELA(0) = LK_CODCIA
art_rela.Requery

Do Until art_rela.EOF
SQ_OPER = 1
pu_codcia = LK_CODCIA
If art_rela!pa_codpa = 22242 Then Stop
PUB_CODART = art_rela!pa_codpa
LEER_ARM_LLAVE
If arm_llave.EOF Then
  MsgBox "no hay relacion  =  " & art_rela!pa_codpa
Else
  If Val(Nulo_Valor0(art_rela!TOT)) <> 0 Then
     arm_llave.Edit
     arm_llave!arm_saldo_s2 = Val(Nulo_Valor0(art_rela!TOT))
     arm_llave.Update
   End If
End If

art_rela.MoveNext
Loop


MsgBox " termino "
Exit Sub
Do Until art_rela.EOF
  SQ_OPER = 1
  PUB_TIPREG = 333
  PUB_NUMTAB = Val(art_rela!CLI_SUBGRUPO)
  PUB_CODCIA = LK_CODCIA
  LEER_TAB_LLAVE
  If tab_llave.EOF Then
      MsgBox "NO AMARRA CODIGO :  " & art_rela!cli_codclie
   '    art_rela.Edit
   '    art_rela!CLI_SUBGRUPO = 7
   '    art_rela!CLI_CASA1 = "07"
   '    art_rela.Update
'     GoTo salER
  End If
  If art_rela!CLI_MONEDA = "D" Then
      MsgBox "dOLARES  OJO " & art_rela!cli_codclie
      Stop
      GoTo salER
  End If
  
  If UCase(Left(Trim(art_rela!cli_nombre), 3)) = "AMB" Then
     art_rela.Edit
     art_rela!CLI_CASA1 = "06"
     art_rela!CLI_SUBGRUPO = 6
     art_rela.Update
  'Else
  '   art_rela!CLI_CASA1 = "07"
  End If
  
  
salER:
art_rela.MoveNext
Loop
MsgBox "acabo"
Exit Sub


Stop


SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
Do Until art_mayor.EOF
 SQ_OPER = 2
 pu_codcia = LK_CODCIA
 PUB_CODART = art_mayor!art_key
 LEER_PRE_LLAVE
 If pre_mayor.RowCount > 1 Then
 If Trim(pre_mayor!pre_FLAG_UNIDAD) <> "" Then
   pre_mayor.Edit
   pre_mayor!pre_FLAG_UNIDAD = " "
   pre_mayor.Update
  
   pre_mayor.MoveNext
   pre_mayor.Edit
   pre_mayor!pre_FLAG_UNIDAD = "A"
   pre_mayor.Update
  End If
 End If
art_mayor.MoveNext
Loop

Exit Sub
' RECALCULO DE LOS STOCK PARA TOMA DE INVENTARIO
MsgBox "LISTO PARA EMPEZAR"
'If xl Is Nothing Then
'   Set xl = CreateObject("Excel.Application")
'End If
DoEvents
pub_cadena = "SELECT DISTINCT FAR_NUMGUIA FROM FACART WHERE FAR_CODCIA = ? AND FAR_SERGUIA = ?  AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' ORDER BY  FAR_NUMGUIA DESC"
pub_cadena = "SELECT DISTINCT FAR_FBG, FAR_NUMSER, FAR_NUMFAC FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ?  AND FAR_EQUIV = 1 AND  FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = 0
WPSART_LLAVE(1) = 0
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA


PUB_KEY = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
LEER_ART_LLAVE
WWF = ""
WS_FILA = 2
Do Until art_mayor.EOF

SQ_OPER = 2
PUB_CODART = art_mayor!art_key
LEER_PRE_LLAVE
If pre_mayor.RowCount > 1 Then
 WPSART_LLAVE(1) = pre_mayor!PRE_CODART
 WfART_llave.Requery
 If Not WfART_llave.EOF Then
   Debug.Print WfART_llave!far_fbg & "/" & WfART_llave!far_numser & " " & WfART_llave!far_numfac
 End If
 
End If


Ava00:
 
 art_mayor.MoveNext
Loop

MsgBox "LISTO"


Exit Sub


WS_FILA = 2
Do Until WS_FILA = 65000
 If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
 SQ_OPER = 3
 pu_alterno = xl.Cells(WS_FILA, 1)
 pu_codcia = LK_CODCIA
 LEER_ART_LLAVE
 If art_llave_alt.EOF Then
   MsgBox "Notar codigo No Existe ...." & pu_alterno & " " & WS_FILA
   GoTo Ava0001
 End If
 SQ_OPER = 1
 PUB_CODART = art_llave_alt!art_key
 pu_codcia = LK_CODCIA
 LEER_ARM_LLAVE
 If arm_llave.EOF Then
   MsgBox " FALLO "
 End If
 WPSART_LLAVE(2) = arm_llave!ARM_CODART
 WfART_llave.Requery
 If WfART_llave.EOF Then
  ' MsgBox " NO SE AGREGO EN SALDO INCIAL "
   xl.Cells(WS_FILA, 8) = "MANUAL"
   GoTo Ava0001
 End If
 SQ_OPER = 2
 pu_codcia = LK_CODCIA
 PUB_CODART = arm_llave!ARM_CODART
 LEER_PRE_LLAVE
 Do Until pre_mayor.EOF
    If Trim(pre_mayor!pre_FLAG_UNIDAD) = "A" Then Exit Do
    
 pre_mayor.MoveNext
 Loop
 
 
' ST_ACTUAL = Format(pre_mayor!pre_equiv * Val(xl.Cells(WS_FILA, 5)), "0.00")
' If ST_ACTUAL <> Val(arm_llave!arm_stock) Then
'   WDIF = ST_ACTUAL - Val(arm_llave!arm_stock)
'   WfART_llave.Edit
'   WfART_llave!far_cantidad_p = WfART_llave!far_cantidad
'   WfART_llave!far_cantidad = WfART_llave!far_cantidad + WDIF
'   WfART_llave.Update
'   arm_llave.Edit
'   arm_llave!arm_stock2 = arm_llave!arm_stock
'   arm_llave!arm_stock = ST_ACTUAL 'Format(Val(xl.Cells(WS_FILA, 2)) * Val(xl.Cells(WS_FILA, 3)), "0.0000")
'
'   arm_llave.Update
 'End If
 
 
 
 
Ava0001:
 WS_FILA = WS_FILA + 1
Loop

MsgBox "TERMINO PROCESO"
' TERMINO DE PROCESO
Exit Sub

' ACTUALIZA PRECIOS
Stop
pub_cadena = "SELECT * FROM ARTI , PRECIOS WHERE (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND ART_FAMILIA = 1 AND PRE_SECUENCIA = 1 AND ART_CODCIA = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
WPSART_LLAVE(0) = LK_CODCIA
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA
WfART_llave.Requery

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_SECUENCIA = 0 AND PRE_CODCIA = ? AND PRE_CODART = ? "
Set PSART_KEY = CN.CreateQuery("", pub_cadena)
PSART_KEY.rdoParameters(0) = "  "
PSART_KEY.rdoParameters(1) = 0
Set artloc_key = PSART_KEY.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_KEY.rdoParameters(0) = LK_CODCIA
Do Until WfART_llave.EOF
   PSART_KEY.rdoParameters(1) = WfART_llave!PRE_CODART
   artloc_key.Requery
   artloc_key.Edit
   artloc_key!PRE_PRE1 = Format(WfART_llave!PRE_PRE1 / WfART_llave!PRE_EQUIV, "0.00")
   artloc_key!PRE_LITRO = Format(Nulo_Valor0(WfART_llave!PRE_LITRO) / WfART_llave!PRE_EQUIV, "0.000")
   artloc_key!pre_PESO = Format(WfART_llave!pre_PESO / WfART_llave!PRE_EQUIV, "0.00")
   artloc_key!pre_FLAG_UNIDAD = " "
   artloc_key.Update
   WfART_llave.Edit
   WfART_llave!pre_FLAG_UNIDAD = "A"
   WfART_llave.Update
   WfART_llave.MoveNext
Loop
0 MsgBox "FIN"
Exit Sub

' FIN DE ACTUALIZAR




Stop
Stop
Stop

' ***************
' VENDEDORES
' ***************

If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents

WPAS = "131296"
xl.Workbooks.Open "C:\LLEVO\TABVEN.xls"  ', 0, True, 4,  WPAS, WPAS

xl.Application.Visible = True
WS_FILA = 2
Stop
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do

ven_llave.AddNew
ven_llave!vem_codcia = LK_CODCIA
ven_llave!VEM_CODVEN = Val(xl.Cells(WS_FILA, 1))
ven_llave!VEM_NOMBRE = Trim(xl.Cells(WS_FILA, 4))
ven_llave!VEM_FECHA_ING = Trim(xl.Cells(WS_FILA, 8))
ven_llave!VEM_TELE_CASA = " "
ven_llave!VEM_TELE_CELU = " "
ven_llave.Update

Ava0022:
WS_FILA = WS_FILA + 1
Loop


MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
MsgBox "PARE STOP"
Exit Sub
Stop

WS_FILA = 2
'Do Until dataO.Recordset.EOF
PUB_NUMSER = 1
PUB_NUMFAC = 1
WS_NUMSEC = 0
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
If Val(xl.Cells(WS_FILA, 7)) = 0 Then GoTo Ava
If Val(xl.Cells(WS_FILA, 25)) = 0 Then GoTo Ava
wcanti_unid = Format(Val(xl.Cells(WS_FILA, 7)) * Val(xl.Cells(WS_FILA, 25)), "0.0000")
  
SQ_OPER = 3
pu_alterno = xl.Cells(WS_FILA, 1)
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
If art_llave_alt.EOF Then
  MsgBox "Notar codigo No Existe ...." & pu_alterno
  GoTo Ava
End If
SQ_OPER = 1
PUB_CODART = art_llave_alt!art_key
pu_codcia = LK_CODCIA
LEER_ARM_LLAVE
arm_llave.Edit
arm_llave!arm_stock = Abs(wcanti_unid)
arm_llave.Update
arm_llave.Requery

far_llave.AddNew
      far_llave!FAR_tipmov = 6
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_cod_sunat = 0 'Val(Right(i_codsunat.Text, 5))
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_CODVEN = 0
      far_llave!far_numfac = PUB_NUMFAC + 1
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!far_numsec = WS_NUMSEC
      far_llave!FAR_STOCK = Abs(wcanti_unid)
      far_llave!far_codart = Val(arm_llave!ARM_CODART)
      far_llave!FAR_cantidad = Abs(wcanti_unid)
      far_llave!FAR_PRECIO = Format(Val(xl.Cells(WS_FILA, 14)) / Val(xl.Cells(WS_FILA, 7)), "0.0000")
      far_llave!FAR_equiv = 1 'Val(xl.Cells(WS_FILA, 7))
      far_llave!far_descri = "PZA" 'Trim(xl.Cells(WS_FILA, 6))
      far_llave!far_PESO = 0
      far_llave!far_signo_car = 0
      far_llave!far_signo_car = 0
      If wcanti_unid < 0 Then
        far_llave!far_signo_arm = -1
      Else
        far_llave!far_signo_arm = 1
      End If
      far_llave!far_codclie = 0
      far_llave!FAR_MONEDA = "S"
      far_llave!FAR_EX_IGV = 0
      far_llave!FAR_cp = " "
      far_llave!FAR_fecha_compra = LK_FECHA_DIA
      far_llave!far_estado = "N"
      far_llave!FAR_ESTADO2 = "N"
      far_llave!FAR_COSPRO = 0
      far_llave!FAR_COSPRO_ANT = 0
      
      far_llave!far_fbg = " "
      far_llave!far_impto = 0
      far_llave!FAR_TOT_FLETE = 0
      far_llave!FAR_FLETE = 0
      far_llave!FAR_DESCTO = 0
      far_llave!FAR_TOT_DESCTO = 0
      far_llave!FAR_GASTOS = 0
      far_llave!far_bruto = 0
      far_llave!FAR_NUMDOC = 1
      far_llave!far_numguia = 0
      far_llave!far_serguia = 0
      far_llave!FAR_pordescto1 = 0
      far_llave!FAR_COSTEO = "A"
      far_llave!FAR_COSTEO_REAL = "A"
      far_llave!FAR_tipo_cambio = 1
      far_llave!FAR_DIAS = 0
      far_llave!far_fecha = LK_FECHA_DIA
      far_llave!FAR_NUMSER_C = 0
      far_llave!FAR_NUMFAC_C = 1
      far_llave!FAR_NUMOPER = 1
      far_llave!far_precio_neto = 0
'      far_llave!FAR_CONSIG = 0
      far_llave!far_otra_cia = " "
      far_llave!far_subtra = 1
      far_llave!far_transito = " "
      far_llave!far_subtra = " "
      far_llave!far_otra_cia = " "
      far_llave!far_transito = " "
      far_llave!far_subtra = " "
      far_llave!far_JABAS = 0
      far_llave!far_UNIDADES = 0
      far_llave!far_mortal = 0
      far_llave!far_num_precio = 0
      far_llave!FAR_ORDEN_UNIDADES = 0
      far_llave!far_subtotal = 0
      'far_llave!far_ISLA = 0
      far_llave!far_turno = 0
      far_llave!far_concepto = " "
      far_llave!far_concepto = "Saldo Inicial de Inventario"
      far_llave!far_codusu = LK_CODUSU
      far_llave!far_hora = Format(Now, "hh:mm:ss AMPM")
      far_llave!FAR_NUM_LOTE = 0
      far_llave!FAR_PEDSER = 0
      far_llave!FAR_PEDFAC = 0
      far_llave!far_pedsec = PUB_VISITA
      'far_llave!far_fbg2 = 0
      far_llave!FAR_TIPDOC = "IN"
      far_llave.Update
Ava:
WS_FILA = WS_FILA + 1
Loop

MsgBox "TERMNO"
Exit Sub




'Dim BDRUTA
'BDRUTA = "E:\SALDO\bd2.mdb"
'dataO.DatabaseName = BDRUTA
'dataO.RecordSource = "STOCKCOL"
'dataO.Refresh
'If dataO.Recordset.EOF Then
'  MsgBox "VERIFICAR ...NO HAY DATOS EN dataO..", 48, Pub_Titulo
'  Exit Sub
'End If


'pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = '55'  ORDER BY COM_CUENTA"
pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? "
Set WPSART_LLAVE = CN.CreateQuery("", pub_cadena)
Set WfART_llave = WPSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
WPSART_LLAVE(0) = LK_CODCIA
WfART_llave.Requery
' AGREGAR POR MESES !!!!!!!!!!!!!!!
Stop
BDRUTA = "C:\ST.mdb"
PUB_FECHA = CDate("30/06/2000")
dataO.DatabaseName = BDRUTA
dataO.RecordSource = "STOCKCOL"
dataO.Refresh
If dataO.Recordset.EOF Then
  MsgBox "VERIFICAR ...NO HAY DATOS EN dataO..", 48, Pub_Titulo
  Exit Sub
End If
Barra.Visible = True
Barra.Min = 0
Barra.max = dataO.Recordset.RecordCount
Barra.Value = 0
ww_conta = 0
flag_otro = ""
WS_NRO_MOV = -1
'WCVOUCHER = Val(Mid(Trim(DataO.Recordset!NUMVOU), 3, Len(Trim(DataO.Recordset!NUMVOU))))
Stop
WC_TIPMOV = -1
Do Until dataO.Recordset.EOF
    DoEvents
    If UCase(Trim(dataO.Recordset!DETMOV)) = "APERTURA" Then GoTo PASA_APERTURA
    ww_conta = ww_conta + 1
PASA_APERTURA:
dataO.Recordset.MoveNext
Loop

Screen.MousePointer = 0
MsgBox "TERMINO ..... SIGUE CHEQUEO DE CUENTAS"

Exit Sub
Exit Sub


'PARA AGREGAR ARTICULOS
'' *************************
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
'lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents
WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\ALMACEN.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2

Do Until Trim(xl.Cells(WS_FILA, 1)) = ""
    If Trim(xl.Cells(WS_FILA, 7)) = 0 Then
      MsgBox "no agregado"
      GoTo SALTA_ARTI_001
    End If
    
SALTA_ARTI_001:
    WS_FILA = WS_FILA + 1
Loop
MsgBox " TEWRMINO "

Exit Sub

Stop
WS_FILA = 156
xl.Application.Visible = True


Exit Sub

'PARA AGREGAR ARTICULOS
'' *************************
If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
'lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents
WPAS = "131296"
xl.Workbooks.Open "C:\CARGA\PRODUCTOS.xls"  ', 0, True, 4,  WPAS, WPAS
xl.Application.Visible = True
WS_FILA = 2

Do Until Trim(xl.Cells(WS_FILA, 1)) = ""
    
    If Trim(xl.Cells(WS_FILA, 7)) = 0 Then
     MsgBox "no agregado"
     GoTo SALTA_ART33I
    End If
    
    pu_alterno = Trim(xl.Cells(WS_FILA, 2))
    pu_codcia = LK_CODCIA
    art_llave_alt.Requery
    
    cmdAgregar_Click

    art_familia.ListIndex = 0
    
    DS.ListIndex = 0
    txtnombre.Text = Trim(xl.Cells(WS_FILA, 4))
    txt_alterno.Text = Trim(xl.Cells(WS_FILA, 2))
    txtClase = Trim(xl.Cells(WS_FILA, 15))
    txtLab = Trim(xl.Cells(WS_FILA, 1))
    
    'If Trim(xl.Cells(WS_FILA, 7)) > 1 Then
    ' grid_UNID_KeyDown 45, 0
    ' grid_unid.TextMatrix(2, 0) = Trim(xl.Cells(WS_FILA, 5))
    ' grid_unid.TextMatrix(2, 1) = Trim(xl.Cells(WS_FILA, 7))
    ' grid_unid.TextMatrix(2, 5) = Format((Val(xl.Cells(WS_FILA, 9)) * 1.18), "0.00")
   '   grid_unid.TextMatrix(2, 21) = Trim(xl.Cells(WS_FILA, 11)) ' SOLES
    ' grid_unid.TextMatrix(2, 16) = Format((Val(xl.Cells(WS_FILA, 9)) * 1.18), "0.00") ' DOLARES
    ' grid_unid.TextMatrix(2, 28) = Val(Trim(xl.Cells(WS_FILA, 8))) ' CANTIDAD DE LITROS
    '  grid_unid.TextMatrix(2, 26) = Val(Trim(xl.Cells(WS_FILA, 9))) ' CANTIDAD DE PESO
      ' UNIDAD MINIMA
     'grid_unid.TextMatrix(1, 5) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     ' grid_unid.TextMatrix(1, 21) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     'grid_unid.TextMatrix(1, 16) = Format(Val(grid_unid.TextMatrix(2, 5)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     'grid_unid.TextMatrix(1, 28) = Format(Val(grid_unid.TextMatrix(2, 28)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
     'grid_unid.TextMatrix(1, 26) = Format(Val(grid_unid.TextMatrix(2, 26)) / Val(grid_unid.TextMatrix(2, 1)), "0.00")
    
    'Else
     grid_unid.TextMatrix(1, 0) = Trim(xl.Cells(WS_FILA, 6))
     grid_unid.TextMatrix(1, 1) = Trim(xl.Cells(WS_FILA, 7))
     grid_unid.TextMatrix(1, 5) = Format((Val(xl.Cells(WS_FILA, 9)) * 1.18), "0.00")
     
     'grid_unid.TextMatrix(1, 21) = Trim(xl.Cells(WS_FILA, 11)) ' SOLES
     grid_unid.TextMatrix(1, 16) = Trim(xl.Cells(WS_FILA, 16)) ' dolares
     grid_unid.TextMatrix(1, 28) = Val(Trim(xl.Cells(WS_FILA, 8))) ' CANTIDAD DE LITROS
     grid_unid.TextMatrix(1, 26) = Val(Trim(xl.Cells(WS_FILA, 9))) ' CANTIDAD DE PESO

    'End If
    
    cmdAgregar_Click
SALTA_ART33I:
    WS_FILA = WS_FILA + 1
Loop

MsgBox " TEWRMINO "

Exit Sub
Set xl = Nothing
xl.Close
Stop


If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If
DoEvents
' lblproceso.Caption = "Abriendo , Archivo Saldos.xls . . . "
DoEvents


xl.Workbooks.Open "C:\CARGA\ZONAS.xls"  ', 0, True, 4,  WPAS, WPAS

xl.Application.Visible = True
WS_FILA = 2
Stop
Do Until WS_FILA = 65000
If Trim(xl.Cells(WS_FILA, 1)) = "" Then Exit Do
  tab_llave.Requery
  tab_llave.AddNew
  tab_llave!tab_codcia = LK_CODCIA
  tab_llave!tab_tipreg = 35
  tab_llave!TAB_NUMTAB = Val(Trim(xl.Cells(WS_FILA, 1)))
  tab_llave!TAB_NOMLARGO = Trim(xl.Cells(WS_FILA, 2))
  tab_llave!tab_nomcorto = Trim(xl.Cells(WS_FILA, 2))
  tab_llave!tab_codart = 0
  tab_llave!tab_codclie = 0
  tab_llaveTAB_CONTABLE2 = 0

  tab_llave.Update


WS_FILA = WS_FILA + 1

Loop


MsgBox " TERMINO"

End Sub
Public Function PROCESA_CANTIDAD(WCODART As Currency, wdias As Integer) As Currency
Dim fx As Integer
Dim fecha1
Dim fecha2
Dim wdividir As Integer
Dim wflag  As String * 1
Dim WARTI As rdoQuery
Dim wllave_arti As rdoResultset
Dim wrango As Integer
Dim sum_cantidad As Currency
Dim WQ_cANTIDAD As Currency
PSLOC_WARTI(0) = ""
PSLOC_WARTI(1) = ""
PSLOC_WARTI(2) = ""

If Val(PSLOC_WARTI(0)) = 0 And Val(PSLOC_WARTI(1)) = 0 And Val(PSLOC_WARTI(2)) = 0 Then
  PSLOC_WARTI(0) = LK_CODCIA
End If

PSLOC_WARTI(3) = WCODART
wdividir = 0
'wrango = Val(tpromedio.Text) * -1
PROCESA_CANTIDAD = 0
fecha1 = LOC_FECHA_INI ' DateAdd("d", wrango, LK_FECHA_DIA)
fecha2 = LOC_FECHA_FIN  ' DateAdd("d", WDIAS, fecha1)
PSLOC_WARTI(4) = fecha1
PSLOC_WARTI(5) = fecha2
llave_sum_arti.Requery
If Not llave_sum_arti.EOF Then
  If Nulo_Valor0(llave_sum_arti!cantidad) <> 0 Then
    sum_cantidad = sum_cantidad + Nulo_Valor0(llave_sum_arti!cantidad)
  End If
End If
wdividir = DateDiff("d", fecha1, fecha2) + 1
If sum_cantidad <> 0 And wdividir <> 0 Then
  WQ_cANTIDAD = Format((sum_cantidad / wdividir), "0.0000")
Else
  sum_cantidad = 0
End If
PROCESA_CANTIDAD = WQ_cANTIDAD

End Function


Public Sub LEVANTE_ZONA()
Dim w_equiv As Currency
Dim W_STOCK As Currency
Dim posab As Integer
Dim ZON_CODCIA As String
Dim codi  As String
Dim pszona_facart As rdoResultset
Dim PSPZONA_FACART As rdoQuery

Dim pszona_llave1  As rdoResultset
Dim PSPZONA_LLAVE As rdoQuery

Dim pszona_llave1_F  As rdoResultset
Dim PSPZONA_LLAVE_F As rdoQuery

Dim ps_llave_precios  As rdoResultset
Dim PSPRECIOS_LLAVE As rdoQuery

Dim ps_local_precios As rdoResultset
Dim PSPRECIOS_LOCAL As rdoQuery
MsgBox "Empezar"
'Stop

ZON_CODCIA = "05"

'        pub_cadena = "SELECT * FROM BDCHEPEN.DBO.ARTI WHERE BDCHEPEN.DBO.ARTI.ART_CODCIA = '05' AND BDCHEPEN.DBO.ARTI.ART_TIPO <> 'X' "
'        Set PSPRECIOS_LLAVE = CN.CreateQuery("", pub_cadena)
'        Set ps_llave_precios = PSPRECIOS_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
'        ps_llave_precios.Requery
'        Do Until ps_llave_precios.EOF
'             pub_cadena = "INSERT INTO FACART SELECT * FROM BDCHEPEN.DBO.FACART WHERE BDCHEPEN.DBO.FACART.FAR_CODCIA = '05' AND BDCHEPEN.DBO.FACART.FAR_CODART = " & 0
'            CN.Execute pub_cadena, rdExecDirect
'            ps_llave_precios.MoveNext
'        Loop
'        Exit Sub

pub_cadena = "SELECT * FROM BDCHEPEN.DBO.PRECIOS WHERE BDCHEPEN.DBO.PRECIOS.PRE_CODCIA = '05' AND BDCHEPEN.DBO.PRECIOS.PRE_CODART = ? AND BDCHEPEN.DBO.PRECIOS.PRE_EQUIV = ?  "
Set PSPRECIOS_LLAVE = CN.CreateQuery("", pub_cadena)
PSPRECIOS_LLAVE(0) = 0
PSPRECIOS_LLAVE(1) = 0
Set ps_llave_precios = PSPRECIOS_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPRECIOS_LLAVE(0) = 0
PSPRECIOS_LLAVE(1) = 0
ps_llave_precios.Requery
' local
'--------
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? AND PRE_FLAG_UNIDAD = ?  "
Set PSPRECIOS_LOCAL = CN.CreateQuery("", pub_cadena)
PSPRECIOS_LOCAL(0) = 0
PSPRECIOS_LOCAL(1) = 0
PSPRECIOS_LOCAL(2) = 0
Set ps_local_precios = PSPRECIOS_LOCAL.OpenResultset(rdOpenKeyset, rdConcurValues)



pub_cadena = "SELECT * FROM BDCHEPEN.DBO.ARTI WHERE BDCHEPEN.DBO.ARTI.ART_CODCIA = '05' AND BDCHEPEN.DBO.ARTI.ART_TIPO = 'Y'"
Set PSPZONA_LLAVE = CN.CreateQuery("", pub_cadena)
Set pszona_llave1 = PSPZONA_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
pszona_llave1.Requery

'MsgBox "PARAR PASO POR PASO"
posab = Val(InputBox("absolute inicial"))
'Stop
' ver si existe un codigo en zona que no existe en PRINCIPAL
'***********************************************************
pszona_llave1.AbsolutePosition = Val(posab)
Do Until pszona_llave1.EOF
 cmdAgregar.Caption = pszona_llave1.AbsolutePosition & " / " & pszona_llave1.RowCount
 DoEvents
 If Trim(pszona_llave1!ART_alterno) = "0" Then GoTo PASAPOR_VER
otravez_ver:
  SQ_OPER = 3
  pu_alterno = pszona_llave1!ART_alterno
  pu_codcia = LK_CODCIA
  LEER_ART_LLAVE
  If art_llave_alt.EOF Then
   'GoTo PASA_ESPERAR
    'MsgBox " No Existe : " & pszona_llave1!art_alterno & " " & pszona_llave1!art_nombre
    'Debug.Print pszona_llave1!art_alterno & " " & pszona_llave1!art_nombre
'    codi = InputBox("Codigo no Existe Reemplzar por :", "no existe", pszona_llave1!art_alterno)
''    pszona_llave1.Delete
'    If codi <> "" Then
     'pszona_llave1.Edit
     'pszona_llave1!art_tipo = "X"
     'pszona_llave1.Update
'     GoTo otravez_ver
    'End If
'PASA_ESPERAR:
    Debug.Print "XXXXX  - " & pszona_llave1!ART_alterno & "   ---   " & pszona_llave1!art_nombre
  Else
fila = 9
    pub_cadena = "UPDATE BDCHEPEN.DBO.FACART SET BDCHEPEN.DBO.FACART.FAR_CODCIA = '99' , BDCHEPEN.DBO.FACART.FAR_CODART = " & art_llave_alt!art_key & " WHERE BDCHEPEN.DBO.FACART.FAR_CODCIA = '05' AND BDCHEPEN.DBO.FACART.FAR_CODART = " & pszona_llave1!art_key & " "
    CN.Execute pub_cadena, rdExecDirect
    pub_cadena = "INSERT  INTO BDATOS.DBO.FACART SELECT * FROM BDCHEPEN.DBO.FACART WHERE  BDCHEPEN.DBO.FACART.FAR_CODCIA  = '99' AND   BDCHEPEN.DBO.FACART.FAR_CODART = " & art_llave_alt!art_key & " "
    CN.Execute pub_cadena, rdExecDirect
'    On Error GoTo dale
    pub_cadena = "UPDATE BDATOS.DBO.FACART SET BDATOS.DBO.FACART.FAR_CODCIA = '05' WHERE  BDATOS.DBO.FACART.FAR_CODCIA = '99' AND  BDATOS.DBO.FACART.FAR_CODART = " & art_llave_alt!art_key & ""
    CN.Execute pub_cadena, rdExecDirect
'    pub_cadena = "DELETE FROM BDATOS.DBO.FACART  WHERE  BDATOS.DBO.FACART.FAR_CODCIA = '99'"
 '   CN.Execute pub_cadena, rdExecDirect
'    If fila = 9 Then Stop
Print ""
   'End If
  
'            SQ_OPER = 2
'            PUB_CODART = art_llave_alt!art_key
'            pu_codcia = LK_CODCIA
'            LEER_PRE_LLAVE
'            Do Until pre_mayor.EOF
'                  PSPRECIOS_LLAVE(0) = pszona_llave1!art_key
'                  PSPRECIOS_LLAVE(1) = pre_mayor!pre_equiv
'                  ps_llave_precios.Requery
'                  If ps_llave_precios.EOF Then
'                     Debug.Print " NO HAY EQUIV  = " & art_llave_alt!art_alterno & "  --- " & art_llave_alt!art_nombre
'                    pszona_llave1.Edit
'                    pszona_llave1!art_tipo = "X"
'                    pszona_llave1.Update
'
'                  End If
'              pre_mayor.MoveNext
'            Loop
     
  End If
PASAPOR_VER:
  pszona_llave1.MoveNext
Loop
MsgBox "Ok Revision Lista", 48






Stop

' PROCEDE A QUITAR TODAS LOS PRE_SECUECIA = 1 AND PRE_EQVUI = 1  ESTARINA DEMAS Y CAUSA PROBLEMA
'***********************************************************
pub_cadena = "SELECT * FROM BDCHEPEN.DBO.PRECIOS WHERE BDCHEPEN.DBO.PRECIOS.PRE_CODCIA = '02' AND BDCHEPEN.DBO.PRECIOS.PRE_SECUENCIA = 1 AND BDCHEPEN.DBO.PRECIOS.PRE_EQUIV = 1 "
Set PSPZONA_FACART = CN.CreateQuery("", pub_cadena)
Set pszona_facart = PSPZONA_FACART.OpenResultset(rdOpenKeyset, rdConcurValues)
pszona_facart.Requery
Do Until pszona_facart.EOF
  cmdAgregar.Caption = pszona_facart.AbsolutePosition & " / " & pszona_facart.RowCount
  pub_cadena = "UPDATE BDCHEPEN.DBO.PRECIOS SET BDCHEPEN.DBO.PRECIOS.PRE_FLAG_UNIDAD = 'A' WHERE BDCHEPEN.DBO.PRECIOS.PRE_CODCIA = '02' AND BDCHEPEN.DBO.PRECIOS.PRE_CODART = " & pszona_facart!PRE_CODART & " AND BDCHEPEN.DBO.PRECIOS.PRE_SECUENCIA = 0"
  CN.Execute pub_cadena, rdExecDirect
  pszona_facart.Edit
  pszona_facart!PRE_CODCIA = "9"
  pszona_facart.Update
pszona_facart.MoveNext
Loop
'***********************************************************
'***********************************************************


' PROCEDE A CAMBIO DE CODIGO EN EL FACART DE LIMA BDATOS
'***********************************************************
'SELECT DISTINCT FAR_CODART , FAR_EQUIV  FROM FACART WHERE FAR_CODART = 214
Stop

pub_cadena = "SELECT DISTINCT BDCHEPEN.DBO.FACART.FAR_CODART, BDCHEPEN.DBO.FACART.FAR_EQUIV FROM BDCHEPEN.DBO.FACART WHERE BDCHEPEN.DBO.FACART.FAR_CODCIA = '" & ZON_CODCIA & "' AND FAR_CODART = ? "
Set PSPZONA_FACART = CN.CreateQuery("", pub_cadena)
PSPZONA_FACART(0) = 0
Set pszona_facart = PSPZONA_FACART.OpenResultset(rdOpenKeyset, rdConcurValues)

pszona_llave1.Requery
Do Until pszona_llave1.EOF
 cmdAgregar.Caption = pszona_llave1.AbsolutePosition & " / " & pszona_llave1.RowCount
 DoEvents
 If Trim(pszona_llave1!ART_alterno) = "0" Then GoTo PASAPOR_PROCESO
 PSPZONA_FACART(0) = pszona_llave1!art_key
 pszona_facart.Requery
 Do Until pszona_facart.EOF
    ' LLAVE DE CODIGO LOCAL O SERVIDOR PINCIPAL '
    SQ_OPER = 3
    pu_alterno = pszona_llave1!ART_alterno
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_llave_alt.EOF Then
      Stop
    End If
    PSPRECIOS_LLAVE(0) = pszona_facart!far_codart  ' En la Zona
    PSPRECIOS_LLAVE(1) = pszona_facart!FAR_equiv
    ps_llave_precios.Requery
    '---------------------------------
    If ps_llave_precios.EOF Then
        Debug.Print "Prodcuto Abservado especial " & pszona_facart!far_codart
        GoTo sigue_facart
    Else
      If ps_llave_precios!pre_FLAG_UNIDAD = "A" Then  ' UNIDAD ACTIVA DE ZONA
SUBE_OTRO:
         PSPRECIOS_LOCAL(0) = LK_CODCIA
         PSPRECIOS_LOCAL(1) = art_llave_alt!art_key
         PSPRECIOS_LOCAL(2) = "A"
         ps_local_precios.Requery
         If ps_local_precios.EOF Then
            Stop
         End If
         pub_cadena = "UPDATE BDCHEPEN.DBO.FACART SET BDCHEPEN.DBO.FACART.FAR_CANTIDAD = (BDCHEPEN.DBO.FACART.FAR_CANTIDAD / BDCHEPEN.DBO.FACART.FAR_EQUIV) * " & ps_local_precios!PRE_EQUIV & "  , BDCHEPEN.DBO.FACART.FAR_EQUIV =  " & ps_local_precios!PRE_EQUIV & " ,  BDCHEPEN.DBO.FACART.FAR_CODCIA = '99', BDCHEPEN.DBO.FACART.FAR_CODART = " & art_llave_alt!art_key & " WHERE BDCHEPEN.DBO.FACART.FAR_CODCIA = '" & ZON_CODCIA & "' AND BDCHEPEN.DBO.FACART.FAR_CODART = " & pszona_facart!far_codart & " AND BDCHEPEN.DBO.FACART.FAR_EQUIV = " & pszona_facart!FAR_equiv
         CN.Execute pub_cadena, rdExecDirect
      Else
         PSPRECIOS_LOCAL(0) = LK_CODCIA
         PSPRECIOS_LOCAL(1) = art_llave_alt!art_key
         PSPRECIOS_LOCAL(2) = " "
         ps_local_precios.Requery
         If ps_local_precios.EOF Then
           GoTo SUBE_OTRO
            Stop
         End If
         pub_cadena = "UPDATE BDCHEPEN.DBO.FACART SET BDCHEPEN.DBO.FACART.FAR_CANTIDAD = (BDCHEPEN.DBO.FACART.FAR_CANTIDAD / BDCHEPEN.DBO.FACART.FAR_EQUIV) * " & ps_local_precios!PRE_EQUIV & "  , BDCHEPEN.DBO.FACART.FAR_EQUIV =  " & ps_local_precios!PRE_EQUIV & " ,  BDCHEPEN.DBO.FACART.FAR_CODCIA = '99', BDCHEPEN.DBO.FACART.FAR_CODART = " & art_llave_alt!art_key & " WHERE BDCHEPEN.DBO.FACART.FAR_CODCIA = '" & ZON_CODCIA & "' AND BDCHEPEN.DBO.FACART.FAR_CODART = " & pszona_facart!far_codart & " AND BDCHEPEN.DBO.FACART.FAR_EQUIV = " & pszona_facart!FAR_equiv
         CN.Execute pub_cadena, rdExecDirect
      End If
    End If

sigue_facart:
   pszona_facart.MoveNext
 Loop
 
 
PASAPOR_PROCESO:
pszona_llave1.MoveNext
Loop
MsgBox "TERMINO", 48

Stop

' ACTUALIZAR SALDOS A LA FECHA EN ZONAS
'***************************************

pub_cadena = "SELECT * FROM BDCHEPEN.DBO.ARTI, BDCHEPEN.DBO.ARTICULO,BDCHEPEN.DBO.PRECIOS  WHERE (BDCHEPEN.DBO.ARTI.ART_KEY = BDCHEPEN.DBO.ARTICULO.ARM_CODART) AND (BDCHEPEN.DBO.ARTI.ART_CODCIA = BDCHEPEN.DBO.ARTICULO.ARM_CODCIA) AND (BDCHEPEN.DBO.ARTICULO.ARM_CODART = BDCHEPEN.DBO.PRECIOS.PRE_CODART) AND (BDCHEPEN.DBO.ARTICULO.ARM_CODCIA = BDCHEPEN.DBO.PRECIOS.PRE_CODCIA) AND (BDCHEPEN.DBO.PRECIOS.PRE_FLAG_UNIDAD = 'A') AND BDCHEPEN.DBO.ARTICULO.ARM_CODCIA = '01'"
Set PSPZONA_FACART = CN.CreateQuery("", pub_cadena)
Set pszona_facart = PSPZONA_FACART.OpenResultset(rdOpenKeyset, rdConcurValues)
pszona_facart.Requery
Do Until pszona_facart.EOF
  cmdAgregar.Caption = pszona_facart.AbsolutePosition & " / " & pszona_facart.RowCount
  DoEvents
  SQ_OPER = 3
  pu_alterno = pszona_facart!ART_alterno
  pu_codcia = LK_CODCIA
  LEER_ART_LLAVE
  If art_llave_alt.EOF Then
    MsgBox " No Existe : " & pszona_llave1!ART_alterno & " " & pszona_llave1!art_nombre
  Else
   SQ_OPER = 2
   pu_codcia = LK_CODCIA
   PUB_CODART = art_llave_alt!art_key
   LEER_PRE_LLAVE
   w_equiv = 0
   Do Until pre_mayor.EOF
      If pre_mayor!pre_FLAG_UNIDAD = "A" Then
         w_equiv = pre_mayor!PRE_EQUIV
         Exit Do
      End If
      pre_mayor.MoveNext
   Loop

   SQ_OPER = 1
   PUB_CODART = art_llave_alt!art_key
   pu_codcia = LK_CODCIA
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      Stop
   Else
   arm_llave.Edit
   W_STOCK = Format(pszona_facart!arm_stock / pszona_facart!PRE_EQUIV, "0.00")
   arm_llave!arm_stock = Format(W_STOCK * w_equiv, "0")
   arm_llave!ARM_INGRESOS = arm_llave!arm_stock
   arm_llave!ARM_SALIDAS = 0
   arm_llave!ARM_COSPRO = pszona_facart!ARM_COSPRO
   arm_llave.Update
   End If
  
  End If
  
pszona_facart.MoveNext
Loop
'***********************************************************
'***********************************************************
MsgBox "termino"

Exit Sub




dale:
fila = 0
Resume Next
End Sub


Public Sub cias_marca()
Dim xcuenta As Integer
lstcias.Clear
If Trim(Nulo_Valors(par_llave!par_art_cias)) <> "" Then
    xcuenta = 1
    For fila = 1 To 30
      pu_codcia = Mid(Trim(par_llave!par_art_cias), xcuenta, 2)
      If Trim(pu_codcia) = "" Then Exit For
      PSPAR_MULTI(0) = pu_codcia
      par_multi.Requery
      lstcias.AddItem Trim(par_multi!par_codcia) & " - " & Trim(par_multi!par_nombre)
      xcuenta = xcuenta + 2
    Next fila
End If
fracias.Visible = True
        
End Sub
Public Sub LEVANTE_ZONA_ORIGINAL()
End Sub

Public Function Pedido_Estimado(wt_codart As Currency, wt_dias As Integer, wt_fecha1 As Date, wt_fecha2 As Date) As Currency
Dim wdias As Integer
Dim wflag As String
Dim ps_estimaped  As rdoResultset
Dim wdias_filtro As Integer
Dim ws_resultado As Currency
Dim sum_canti As Currency
Dim wdias_afecto As Integer
Dim wfactor As Currency
Dim sum_cantidad As Currency
Dim WQ_cANTIDAD As Currency
Dim fecha1
Dim fecha2
Dim wdividir As Integer
Dim wrango As Integer
Dim wcias As String
If LK_CODCIA = "01" Then
 wcias = "IN ( '01','07')"
Else
 wcias = "IN ( '" & LK_CODCIA & "')"
End If

LOCAL_V1 = 0
LOCAL_V2 = 0
LOCAL_V3 = 0


wdias_filtro = wt_fecha2 - wt_fecha1
 
wdias = Val(wt_dias)
wdividir = 0
wrango = wdias_filtro * -1 'Val(wt_dias) * -1
fecha1 = DateAdd("d", wrango, LK_FECHA_DIA)
fecha2 = DateAdd("d", wdias, fecha1)
wflag = "S"
wdividir = 0
sum_cantidad = 0
Do Until wflag = "N"
    If fecha1 > LK_FECHA_DIA Then
       Exit Do
    End If
    pub_cadena = " SELECT SUM(FAR_CANTIDAD) As VENTA_CANT From FACART " & _
       " WHERE FAR_CODCIA " & wcias & " AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E' " & _
       " AND (FAR_FECHA_COMPRA >= '" & Format(fecha1, "dd/mm/yyyy") & "' AND FAR_FECHA_COMPRA <= '" & Format(fecha2, "dd/mm/yyyy") & "' ) " & _
       " AND FAR_ESTADO2 <>'L' AND FAR_CODART = " & wt_codart & " AND far_mortal <> 9 GROUP BY FAR_CODCIA "
    Set ps_estimaped = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
    If Not ps_estimaped.EOF Then
      If Nulo_Valor0(ps_estimaped!venta_cant) <> 0 Then
         wdividir = wdividir + 1
         sum_cantidad = sum_cantidad + Nulo_Valor0(ps_estimaped!venta_cant)
      End If
    End If
    fecha1 = DateAdd("d", 1, fecha2)
    fecha2 = DateAdd("d", wdias, fecha1)
Loop
If sum_cantidad <> 0 Then
   WQ_cANTIDAD = Format((sum_cantidad / wdividir), "000000")
Else
  WQ_cANTIDAD = 0
End If
Pedido_Estimado = WQ_cANTIDAD

End Function



Public Sub add_convenio(WCODART As Currency)
Dim rsconvenio As rdoResultset
Dim WPRECIO As Currency
Dim wunidad As String
pub_cadena = "SELECT  TOP 1 PRE_PRE2 , PRE_UNIDAD FROM PRECIOS WHERE PRE_CODART = " & WCODART & " AND PRE_CODCIA = '01' ORDER BY  PRE_EQUIV  DESC"
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
WPRECIO = 0
wunidad = ""
If Not X.EOF Then
WPRECIO = Val(X!PRE_PRE2)
wunidad = Trim(X!PRE_UNIDAD)
End If


pub_cadena = "select * from tabconvenio where tco_codcia = '01' and tCO_fecha = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' and tCO_codart = " & WCODART
Set rsconvenio = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
If rsconvenio.EOF Then
rsconvenio.AddNew
rsconvenio!tco_codcia = "01"
rsconvenio!tco_fecha = LK_FECHA_DIA
rsconvenio!tco_codart = WCODART
Else
rsconvenio.Edit
End If
rsconvenio!tco_descrip = wunidad
rsconvenio!tco_PRECIO = WPRECIO
rsconvenio!tco_CODUSU = LK_CODUSU
rsconvenio.Update

pub_cadena = "select * from tabconvenio where tco_codcia = '25' and tCO_fecha = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' and tCO_codart = " & WCODART
Set rsconvenio = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
If rsconvenio.EOF Then
rsconvenio.AddNew
rsconvenio!tco_codcia = "25"
rsconvenio!tco_fecha = LK_FECHA_DIA
rsconvenio!tco_codart = WCODART
Else
rsconvenio.Edit
End If
rsconvenio!tco_descrip = wunidad
rsconvenio!tco_PRECIO = WPRECIO
rsconvenio!tco_CODUSU = LK_CODUSU
rsconvenio.Update

pub_cadena = "select * from tabconvenio where tco_codcia = '03' and tCO_fecha = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' and tCO_codart = " & WCODART
Set rsconvenio = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
If rsconvenio.EOF Then
rsconvenio.AddNew
rsconvenio!tco_codcia = "03"
rsconvenio!tco_fecha = LK_FECHA_DIA
rsconvenio!tco_codart = WCODART
Else
rsconvenio.Edit
End If
rsconvenio!tco_descrip = wunidad
rsconvenio!tco_PRECIO = WPRECIO
rsconvenio!tco_CODUSU = LK_CODUSU
rsconvenio.Update

pub_cadena = "select * from tabconvenio where tco_codcia = '07' and tCO_fecha = '" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & "' and tCO_codart = " & WCODART
Set rsconvenio = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
If rsconvenio.EOF Then
rsconvenio.AddNew
rsconvenio!tco_codcia = "07"
rsconvenio!tco_fecha = LK_FECHA_DIA
rsconvenio!tco_codart = WCODART
Else
rsconvenio.Edit
End If
rsconvenio!tco_descrip = wunidad
rsconvenio!tco_PRECIO = WPRECIO
rsconvenio!tco_CODUSU = LK_CODUSU
rsconvenio.Update


End Sub

Sub busca()
Dim WDIRCLI As Integer

Dim VALOR As Currency

Dim WR_SUMA As Currency
Dim WWF As String
Dim ST_ACTUAL As Currency
Dim ps_llave_precios As rdoResultset
Dim WfART_llave As rdoResultset
Dim WPSART_LLAVE As rdoQuery

Dim PSFAR_PEDIDOS As rdoQuery
Dim far_pedidos As rdoResultset

Dim PSPRE_LOTES As rdoQuery
Dim pre_lotes As rdoResultset

Dim wcanti_unid As Currency
Dim WS_FILA As Integer
Dim xl  As Object
Dim E As Integer
Dim WS_SALDO_FINAL As Currency
Stop
Dim ADDPRE As rdoResultset
Dim WCOSTO As Currency
Dim sp_lot As rdoResultset
Dim ws_dif As Currency
Dim cto_fila As Integer

Stop

pub_cadena = "SELECT  * FROM arti WHERe art_codcia='" & LK_CODCIA & "'   "
Set X = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
Do Until X.EOF
  DoEvents
  pub_cadena = "SELECT  * FROM precios WHERE pre_codcia = '" & LK_CODCIA & "' AND pre_codart = " & X!art_key & "  and pre_flag_unidad='A' "
  Set ADDPRE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
  If ADDPRE.EOF Then
    Debug.Print X!art_key
  End If
  X.MoveNext
Loop
MsgBox "proceso terminado"
End Sub
