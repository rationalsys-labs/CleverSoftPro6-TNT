VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCompRp 
   KeyPreview = -1
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compilatorio de Guias RP"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   12495
   Begin Clever.Boton CmdPri 
      Height          =   375
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BTYPE           =   8
      TX = "Listado del compilatorio"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14606046
      BCOLO           =   8612156
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCompRp.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdINfRp 
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      BTYPE           =   8
      TX = "Estadísticas del Diario RP"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14606046
      BCOLO           =   8612156
      FCOL            =   0
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCompRp.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Clever.Boton CmdEliBus 
      Height          =   375
      Left            =   9000
      TabIndex        =   20
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   8
      TX = "Limpiar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8612156
      BCOLO           =   10390910
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmCompRp.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10800
      TabIndex        =   18
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame FraHer 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   Begin Clever.Boton Cmd
      Left=120
      Top=1080
      Width=1935
      Height=375
      Index=5
      TX="Ver historial"
      BTYPE=6
   End
   Begin Clever.Boton Cmd
      Left=120
      Top=1560
      Width=1935
      Height=375
      Index=8
      TX="Actualizar"
      BTYPE=6
   End
   Begin Clever.Boton Cmd
      Left=120
      Top=2040
      Width=1935
      Height=375
      Index=9
      TX="Salir"
      BTYPE=6
   End
      Begin Clever.Boton Cmd 
         Index = 6
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX = "Ver / reimprimir cartilla"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton Cmd 
         Index = 7
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         BTYPE           =   9
         TX = "Reactivar al Diario"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   6120
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22/03/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraInf 
      Height          =   5535
      Left            =   2400
      TabIndex        =   22
      Top             =   480
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid GridRP 
         SelectionMode = 1
         Height          =   5385
         Left            =   15
         TabIndex        =   23
         Top             =   120
         Width           =   9930
         _ExtentX        =   17515
         _ExtentY        =   9499
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   16185077
         BackColorBkg    =   16777215
         GridColor       =   15790320
         GridColorFixed  =   15790320
         GridLinesFixed  =   1
         MergeCells      =   1
         AllowUserResizing=   1
         BorderStyle     =   0
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
   End
   Begin VB.Frame FraDiarRp 
      Height          =   5535
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Visible = 0
      Width           =   9975
   Begin VB.Label LblInfo
      Left=120
      Top=3480
      Width=2055
      Height=255
      Index=8
      Caption="15D con salida"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=3840
      Width=1935
      Height=255
      Index=9
      Caption="48N con salida"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=2040
      Width=1935
      Height=255
      Index=4
      Caption="Muestras 15N"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=2400
      Width=1935
      Height=255
      Index=5
      Caption="Documentos 15D"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=2760
      Width=1935
      Height=255
      Index=6
      Caption="Muestras 48N"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=3120
      Width=1935
      Height=255
      Index=7
      Caption="15N con salida"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=1320
      Width=1935
      Height=255
      Index=2
      Caption="Estado COT"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=1680
      Width=1935
      Height=255
      Index=3
      Caption="Retenidos"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=960
      Width=1935
      Height=255
      Index=1
      Caption="Estado Ventas"
      BackStyle=0

   End
   Begin VB.Label LblInfo
      Left=120
      Top=600
      Width=1935
      Height=255
      Index=0
      Caption="RP con salida"
      BackStyle=0

   End
   Begin VB.Label LblTit
      Left=3360
      Top=240
      Width=6495
      Height=255
      Caption="GUÍAS DEL GRUPO SELECCIONADO"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
      Begin VB.Frame Frame1 
         Height          =   4815
         Left            =   3360
         TabIndex        =   24
         Top             =   480
         Width           =   6495
         Begin MSFlexGridLib.MSFlexGrid GridResumen 
            Height          =   4665
            Left            =   15
            TabIndex        =   25
            Top             =   120
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   8229
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   16185077
            BackColorBkg    =   16777215
            GridColor       =   15790320
            GridColorFixed  =   15790320
            GridLinesFixed  =   1
            MergeCells      =   1
            AllowUserResizing=   1
            BorderStyle     =   0
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
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":0590
         PICN            =   "FrmCompRp.frx":05AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   Begin VB.Label LabelDegrade1
      Left=120
      Top=240
      Width=3135
      Height=255
      Caption="ESTADÍSTICAS DEL DIARIO"
      BackStyle=0
      BackStyle=1
      BackColor=&H0083443B&
      ForeColor=&H00FFFFFF&

   End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   1
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":0FBE
         PICN            =   "FrmCompRp.frx":0FDA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   8
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":19EC
         PICN            =   "FrmCompRp.frx":1A08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   3
         Left            =   2760
         TabIndex        =   9
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":241A
         PICN            =   "FrmCompRp.frx":2436
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   10
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":2E48
         PICN            =   "FrmCompRp.frx":2E64
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   5
         Left            =   2760
         TabIndex        =   11
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":3876
         PICN            =   "FrmCompRp.frx":3892
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   6
         Left            =   2760
         TabIndex        =   12
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":42A4
         PICN            =   "FrmCompRp.frx":42C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   7
         Left            =   2760
         TabIndex        =   13
         Top             =   3120
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":4CD2
         PICN            =   "FrmCompRp.frx":4CEE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   14
         Top             =   3480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":5700
         PICN            =   "FrmCompRp.frx":571C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdEstadistica 
         Height          =   285
         Index           =   9
         Left            =   2760
         TabIndex        =   15
         Top             =   3840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":612E
         PICN            =   "FrmCompRp.frx":614A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Clever.Boton CmdGenARep 
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         Top             =   4920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         BTYPE           =   6
         TX              =   "Actualizar Datos"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8612156
         BCOLO           =   10390910
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmCompRp.frx":6B5C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame3 
         Height          =   375
         Left            =   2280
         TabIndex        =   37
         Top             =   510
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=0
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   2280
         TabIndex        =   39
         Top             =   870
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=1
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame4 
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   1230
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=2
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   1590
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=3
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame6 
         Height          =   375
         Left            =   2280
         TabIndex        =   45
         Top             =   1950
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=4
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame7 
         Height          =   375
         Left            =   2280
         TabIndex        =   47
         Top             =   2310
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=5
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame8 
         Height          =   375
         Left            =   2280
         TabIndex        =   49
         Top             =   2670
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=6
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame9 
         Height          =   375
         Left            =   2280
         TabIndex        =   51
         Top             =   3030
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=7
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame10 
         Height          =   375
         Left            =   2280
         TabIndex        =   53
         Top             =   3390
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=8
      Caption="0"
      BackStyle=0

   End
      End
      Begin VB.Frame Frame11 
         Height          =   375
         Left            =   2280
         TabIndex        =   55
         Top             =   3750
         Width           =   495
   Begin VB.Label Info
      Left=15
      Top=120
      Width=470
      Height=285
      Index=9
      Caption="0"
      BackStyle=0

   End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar"
      Height          =   255
      Left            =   10200
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LblEstado
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=0
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=1
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=2
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=3
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=4
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=10
      Visible=0
   End
   Begin Clever.Boton Cmd
      Left=0
      Top=0
      Width=20
      Height=20
      Index=11
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=0
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=1
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=2
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=3
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=4
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=5
      Visible=0
   End
   Begin VB.TextBox Txt
      Left=0
      Top=0
      Width=20
      Height=20
      Index=6
      Visible=0
   End
   Begin VB.ComboBox CmbCliente
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
      Style=2
   End
   Begin VB.ComboBox CmbMoneda
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
      Style=2
   End
   Begin VB.ComboBox CmbGrupo
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
      Style=2
   End
   Begin VB.CheckBox ChkSalida
      Left=0
      Top=0
      Width=20
      Height=20
      Visible=0
   End
End
Attribute VB_Name = "FrmCompRp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PosicionInicialAplicada As Boolean
Private Const EsCompilatorio As Boolean = True
Private Datos As ADODB.Recordset
Private Editando As Boolean
Private Cargando As Boolean
Private Sub Form_Load()
 On Error GoTo Fallo
 If Not AccesoPermitido(IIf(EsCompilatorio, 10, 9)) Then Unload Me: Exit Sub
 PrepararVista
 CargarOpciones
 Actualizar
 Exit Sub
Fallo:
 MsgBox Err.Description, vbExclamation, App.Title
 Unload Me
End Sub
Private Sub CargarOpciones()
 Dim R As ADODB.Recordset
 Cargando = True
 CmbGrupo.Clear
 CmbCliente.Clear
 Set R = ConsultaOperacion("SELECT Item,Asient,CLIENT FROM dbo.DbClientes WHERE Activo=1 ORDER BY CLIENT")
 Do While Not R.EOF
  CmbCliente.AddItem CStr(R!Asient) & " - " & CStr(R!CLIENT)
  CmbCliente.ItemData(CmbCliente.NewIndex) = CLng(R!Item)
  R.MoveNext
 Loop
 CerrarRecordset R
 CmbMoneda.Clear
 Set R = ConsultaOperacion("SELECT CodEst,Moneda FROM dbo.Int_Moneda WHERE ActEst=1 ORDER BY CodEst")
 Do While Not R.EOF
  CmbMoneda.AddItem CStr(R!Moneda): CmbMoneda.ItemData(CmbMoneda.NewIndex) = CLng(R!CodEst)
  R.MoveNext
 Loop
 CerrarRecordset R
 If Not EsCompilatorio Then
  CargarComboGruposRP CmbGrupo
 Else
  CmbGrupo.AddItem "General": CmbGrupo.ItemData(0) = -1
 End If
 CmbGrupo.ListIndex = 0
 CmbGrupo.Enabled = Not EsCompilatorio
 ChkSalida.Visible = Not EsCompilatorio
 Cargando = False
End Sub
Private Sub Actualizar()
 Dim F As Long, I As Long, Grupo As Long
 On Error GoTo Fallo
 Cargando = True
 CerrarRecordset Datos
 GridRP.Rows = 1
 Grupo = -1
 If CmbGrupo.ListIndex >= 0 Then Grupo = CmbGrupo.ItemData(CmbGrupo.ListIndex)
 Set Datos = ConsultaOperacion("dbo.ConsultarPizarraRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Compilatorio", adBoolean, 0, EsCompilatorio), Array("@Filtro", adVarWChar, 100, Left$(TxtBuscar.Text, 100)), Array("@Grupo", adInteger, 0, Grupo), Array("@SoloSalida", adBoolean, 0, ChkSalida.Value = 1)), True)
 GridRP.FormatString = "Marca|Guía|Cliente|Producto|Gestión|Salida|Total|Fecha"
 If EsCompilatorio Then GridRP.FormatString = "|Guía|Cliente|Producto|Situación|Cartilla|Importe|Fecha de envío"
 GridRP.ColWidth(0) = 600: GridRP.ColWidth(1) = 1400: GridRP.ColWidth(2) = 3200
 GridRP.ColWidth(3) = 850: GridRP.ColWidth(4) = 1300: GridRP.ColWidth(5) = 1300
 GridRP.ColWidth(6) = 1100: GridRP.ColWidth(7) = 1150
 Do While Not Datos.EOF
  GridRP.AddItem ""
  F = GridRP.Rows - 1
  GridRP.TextMatrix(F, 1) = CStr(Datos!NroGui)
  GridRP.TextMatrix(F, 2) = ValorTexto(Datos!Client)
  GridRP.TextMatrix(F, 3) = ValorTexto(Datos!TipGui)
  If EsCompilatorio Then
   GridRP.TextMatrix(F, 4) = ValorTexto(Datos!FINAL)
   GridRP.TextMatrix(F, 5) = ValorTexto(Datos!NroPla)
  Else
   GridRP.TextMatrix(F, 4) = ValorTexto(Datos!EstadoGestion)
   GridRP.TextMatrix(F, 5) = IIf(Val(ValorTexto(Datos!Salida)) = 1, "SI", "NO")
  End If
  GridRP.TextMatrix(F, 6) = Format$(CCur(IIf(IsNull(Datos!MonTot), 0, Datos!MonTot)), "0.00")
  If Not IsNull(Datos!PuGui) Then GridRP.TextMatrix(F, 7) = Format$(Datos!PuGui, "dd/mm/yyyy")
  Datos.MoveNext
 Loop
 Cargando = False
 If GridRP.Rows > 1 Then GridRP.Row = 1
 MostrarFicha
 Habilitar False
 ActualizarVista
 LblEstado.Caption = CStr(GridRP.Rows - 1) & " registros. Doble clic para marcar las guías del lote."
 Exit Sub
Fallo:
 Cargando = False
 CerrarRecordset Datos
 GridRP.Rows = 1
 MostrarFicha
 LblEstado.Caption = "No se pudo actualizar: " & Err.Description
End Sub
Private Function Seleccionar() As Boolean
 If Datos Is Nothing Then Exit Function
 If GridRP.Row < 1 Or GridRP.Row >= GridRP.Rows Or Datos.RecordCount = 0 Then Exit Function
 Datos.MoveFirst
 Datos.Move GridRP.Row - 1
 Seleccionar = True
End Function
Private Sub Elegir(ByVal Combo As ComboBox, ByVal Id As Long)
 Dim I As Long
 Combo.ListIndex = -1
 For I = 0 To Combo.ListCount - 1
  If Combo.ItemData(I) = Id Then Combo.ListIndex = I: Exit Sub
 Next I
End Sub
Private Sub MostrarFicha()
 Dim I As Long, Campos As Variant
 If Cargando Or Editando Then Exit Sub
 For I = 0 To 6: Txt(I).Text = "": Next I
 CmbCliente.ListIndex = -1: CmbMoneda.ListIndex = -1
 If Not Seleccionar Then Exit Sub
 Elegir CmbCliente, Val(ValorTexto(Datos!IdCliente))
 Elegir CmbMoneda, Val(ValorTexto(Datos!Moneda))
 Campos = Array("PesFis", "PesVol", "MonNet", "MonRc", "MonIgv", "MonTot", "Info")
 For I = 0 To 6
  Txt(I).Text = ValorTexto(Datos.Fields(CStr(Campos(I))).Value)
 Next I
End Sub
Private Sub Habilitar(ByVal Edicion As Boolean)
 Dim I As Long, Hay As Boolean
 Editando = Edicion
 Hay = (GridRP.Rows > 1)
 For I = 0 To 9: Cmd(I).Enabled = Not Edicion: Next I
 For I = 0 To 4: Cmd(I).Visible = Not EsCompilatorio: Next I
 Cmd(7).Visible = EsCompilatorio
 Cmd(10).Visible = Edicion: Cmd(11).Visible = Edicion
 GridRP.Enabled = Not Edicion: TxtBuscar.Enabled = Not Edicion
 CmbGrupo.Enabled = Not Edicion And Not EsCompilatorio
 ChkSalida.Enabled = Not Edicion
 CmbCliente.Enabled = Edicion: CmbMoneda.Enabled = Edicion
 For I = 0 To 6: Txt(I).Locked = Not Edicion: Next I
 If Not EsCompilatorio Then MostrarEditor Edicion
End Sub
Private Function FilasXML(ByVal Lote As Boolean) As String
 Dim I As Long, R As String
 If Datos Is Nothing Then Exit Function
 For I = 1 To GridRP.Rows - 1
  If (Lote And GridRP.TextMatrix(I, 0) = "X") Or (Not Lote And I = GridRP.Row) Then
   Datos.MoveFirst: Datos.Move I - 1
   R = R & "<r" & AtributoXML("Item", CStr(Datos!ITEM)) & AtributoXML("Version", VersionXML(Datos.Fields("VersionFila").Value)) & "/>"
  End If
 Next I
 FilasXML = R
End Function
Private Sub Ejecutar(ByVal Accion As String, ByVal Atributos As String, Optional ByVal Lote As Boolean = False)
 Dim XML As String, R As ADODB.Recordset, Confirmado As Boolean, Mensaje As String
 On Error GoTo Fallo
 XML = FilasXML(Lote)
 If Len(XML) = 0 Then MsgBox "Seleccione o marque las guías.", vbInformation, App.Title: Exit Sub
 If MsgBox("¿Confirmar " & Accion & " para las guías seleccionadas?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes Then Exit Sub
 XML = "<operacion" & Atributos & ">" & XML & "</operacion>"
 Set R = ConsultaOperacion("dbo.GuardarPizarraRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Accion", adVarWChar, 20, Accion), Array("@Datos", adLongVarWChar, Len(XML), XML)), True)
 Confirmado = True
 Mensaje = "Operación confirmada."
 If Not IsNull(R!Planilla) Then Mensaje = "Cartilla " & CStr(R!Planilla) & " confirmada. Puede consultarla y reimprimirla desde Compilatorio."
 CerrarRecordset R
 Habilitar False
 Actualizar
 MsgBox Mensaje, vbInformation, App.Title
 Exit Sub
Fallo:
 Mensaje = Err.Description
 CerrarRecordset R
 If Confirmado Then Habilitar False: Mensaje = "Operación confirmada; recargue sin repetirla. " & Mensaje
 MsgBox Mensaje, vbExclamation, App.Title
End Sub
Private Sub Cmd_Click(Index As Integer)
 Dim A As String, I As Long, Estado As String, Comentario As String, R As ADODB.Recordset, Texto As String, Campos As Variant
 On Error GoTo Fallo
 Select Case Index
 Case 8
  CargarOpciones: Actualizar
 Case 9
  Unload Me
 Case 11
  If MsgBox("¿Descartar el borrador?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) = vbYes Then Habilitar False: MostrarFicha
 Case Else
  If Not Seleccionar Then Exit Sub
  Select Case Index
  Case 0
   Habilitar True
  Case 1
   Ejecutar "Salida", AtributoXML("Salida", IIf(Val(ValorTexto(Datos!Salida)) = 1, "0", "1"))
  Case 2
   FrmEventoRP.Preparar
   FrmEventoRP.Show vbModal, MdiPri
   If Not FrmEventoRP.Aceptado Then Unload FrmEventoRP: Exit Sub
   A = AtributoXML("Estado", CStr(FrmEventoRP.CodigoEstado)) & AtributoXML("Comentario", FrmEventoRP.Comentario)
   Unload FrmEventoRP
   Ejecutar "Evento", A
  Case 3
   Ejecutar "Confirmar", "", True
  Case 4
   Ejecutar "Cerrar", "", True
  Case 5
   Set R = ConsultaOperacion("SELECT h.FECHA,h.STATUS,CONVERT(nvarchar(1000),h.INFO) Info FROM dbo.DbsHisRP h WHERE h.HERENCIA=? ORDER BY h.FECHA,h.Item", Array(Array("Asient", adVarWChar, 20, CStr(Datos!Asient))))
   Do While Not R.EOF
    Texto = Texto & ValorTexto(R!FECHA) & " " & ValorTexto(R!STATUS) & " " & ValorTexto(R!Info) & vbCrLf
    R.MoveNext
   Loop
   CerrarRecordset R
   If Len(Texto) = 0 Then Texto = "Sin eventos."
   MostrarHistorialRP Texto
  Case 6
   If Not EsCompilatorio Then AbrirModulo 10: Exit Sub
   MostrarCartillaRP ValorTexto(Datos!NroPla)
  Case 7
   Ejecutar "Reactivar", ""
  Case 10
   If CmbCliente.ListIndex < 0 Or CmbMoneda.ListIndex < 0 Then Err.Raise 5, , "Seleccione cliente y moneda."
   A = AtributoXML("Cliente", CStr(CmbCliente.ItemData(CmbCliente.ListIndex))) & AtributoXML("Moneda", CStr(CmbMoneda.ItemData(CmbMoneda.ListIndex)))
   Campos = Array("PesoFisico", "PesoVolumen", "Neto", "Recargo", "IGV", "Total")
   For I = 0 To 5
    A = A & AtributoXML(CStr(Campos(I)), NumeroXML(Txt(I).Text))
   Next I
   A = A & AtributoXML("Info", Txt(6).Text)
   Ejecutar "Guardar", A
  End Select
 End Select
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub GridRP_RowColChange()
 MostrarFicha
 RefrescarSeleccion
End Sub
Private Sub GridRP_DblClick()
 If GridRP.Row > 0 Then GridRP.TextMatrix(GridRP.Row, 0) = IIf(GridRP.TextMatrix(GridRP.Row, 0) = "X", "", "X")
End Sub
Private Sub GridRP_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then MostrarFicha: ConsultarFicha: KeyCode = 0
 If KeyCode = vbKeySpace Then GridRP_DblClick: KeyCode = 0
End Sub
Private Sub TxtBuscar_Change()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub CmbGrupo_Click()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub ChkSalida_Click()
 If Not Cargando And Not Editando Then Actualizar
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Editando Then Cancel = (MsgBox("¿Descartar el borrador y cerrar?", vbQuestion Or vbYesNo Or vbDefaultButton2, App.Title) <> vbYes)
 If Not Cancel Then CerrarRecordset Datos
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 And Not Editando Then CargarOpciones: Actualizar: KeyCode = 0
 If KeyCode = vbKeyEscape Then
  If Editando Then Cmd_Click 11 Else Unload Me
 End If
End Sub

Private Sub PrepararVista()
 GridRP.Rows = 2
 GridRP.FixedRows = 1
 GridRP.FixedCols = 0
 FraInf.Visible = True
 FraDiarRp.Visible = False
End Sub
Private Sub MostrarEditor(ByVal Edicion As Boolean)
End Sub
Private Sub ConsultarFicha()
 Dim Texto As String, I As Long
 If Not Seleccionar Then Exit Sub
 For I = 0 To Datos.Fields.Count - 1
  If Datos.Fields(I).Name <> "VersionFila" Then Texto = Texto & Datos.Fields(I).Name & ": " & ValorTexto(Datos.Fields(I).Value) & vbCrLf
 Next I
 MostrarHistorialRP Texto
End Sub
Private Sub RefrescarSeleccion()
 On Error GoTo Salir
 If Cargando Then Exit Sub
 If Seleccionar Then StatusBar1.Panels(2).Text = "Guía: " & CStr(Datos!NroGui) & " - " & ValorTexto(Datos!FINAL) & " - Cartilla: " & ValorTexto(Datos!NroPla)
Salir:
End Sub
Private Sub ActualizarVista()
 StatusBar1.Panels(1).Text = "Compilatorio RP"
 StatusBar1.Panels(2).Text = CStr(GridRP.Rows - 1) & " registros confirmados"
 RefrescarSeleccion
 Form_Resize
End Sub
Private Sub CmdPri_Click()
 FraDiarRp.Visible = False
 FraInf.Visible = True
 Actualizar
End Sub
Private Sub CmdINfRp_Click()
 FraInf.Visible = False
 FraDiarRp.Visible = True
 CargarResumenDiario -1
End Sub
Private Sub CmdEliBus_Click()
 TxtBuscar.Text = ""
End Sub
Private Sub CmdGenARep_Click()
 CargarResumenDiario -1
End Sub
Private Sub CmdEstadistica_Click(Index As Integer)
 CargarResumenDiario Index
End Sub
Private Sub CargarResumenDiario(ByVal Categoria As Long)
 Dim R As ADODB.Recordset, P As String, Incluir As Boolean, I As Long, C(0 To 9) As Long
 Dim Salida As Boolean, VTS As Boolean, COT As Boolean, Retenido As Boolean
 On Error GoTo Fallo
 Set R = ConsultaOperacion("dbo.ConsultarPizarraRP", Array(Array("@Actor", adInteger, 0, IdUsuarioActual), Array("@Compilatorio", adBoolean, 0, False), Array("@Filtro", adVarWChar, 100, ""), Array("@Grupo", adInteger, 0, -2)), True)
 GridResumen.Rows = 1
 GridResumen.Cols = 4
 GridResumen.FormatString = "Guía|Cliente|Origen|Producto"
 GridResumen.ColWidth(0) = 1200: GridResumen.ColWidth(1) = 3000
 GridResumen.ColWidth(2) = 1000: GridResumen.ColWidth(3) = 900
 Do While Not R.EOF
  P = ValorTexto(R!TipGui)
  Salida = (Val(ValorTexto(R!Salida)) = 1)
  VTS = (ValorTexto(R!EstadoGestion) = "VTS")
  COT = (ValorTexto(R!EstadoGestion) = "COT")
  Retenido = (Val(ValorTexto(R!Reten)) = 1)
  If Salida Then C(0) = C(0) + 1
  If VTS Then C(1) = C(1) + 1
  If COT Then C(2) = C(2) + 1
  If Retenido Then C(3) = C(3) + 1
  If P = "15N" Then C(4) = C(4) + 1
  If P = "15D" Then C(5) = C(5) + 1
  If P = "48N" Then C(6) = C(6) + 1
  If P = "15N" And Salida Then C(7) = C(7) + 1
  If P = "15D" And Salida Then C(8) = C(8) + 1
  If P = "48N" And Salida Then C(9) = C(9) + 1
  Select Case Categoria
   Case -1: Incluir = True
   Case 0: Incluir = Salida
   Case 1: Incluir = VTS
   Case 2: Incluir = COT
   Case 3: Incluir = Retenido
   Case 4: Incluir = (P = "15N")
   Case 5: Incluir = (P = "15D")
   Case 6: Incluir = (P = "48N")
   Case 7: Incluir = (P = "15N" And Salida)
   Case 8: Incluir = (P = "15D" And Salida)
   Case 9: Incluir = (P = "48N" And Salida)
  End Select
  If Incluir Then GridResumen.AddItem CStr(R!NroGui) & vbTab & CStr(R!Client) & vbTab & CStr(R!Origen) & vbTab & P
  R.MoveNext
 Loop
 For I = 0 To 9: Info(I).Caption = CStr(C(I)): Next I
 CerrarRecordset R
 LblTit.Caption = "GUÍAS DEL DIARIO: " & CStr(GridResumen.Rows - 1)
 Exit Sub
Fallo:
 CerrarRecordset R
 MsgBox Err.Description, vbExclamation, App.Title
End Sub
Private Sub Form_Resize()
 On Error GoTo Salir
 If Me.WindowState = vbMinimized Then Exit Sub
 If ScaleWidth < 12000 Or ScaleHeight < 6000 Then Exit Sub
 FraHer.Height = ScaleHeight - StatusBar1.Height - 150
 FraInf.Width = ScaleWidth - FraInf.Left - 150
 FraInf.Height = ScaleHeight - StatusBar1.Height - FraInf.Top - 150
 GridRP.Width = FraInf.Width - 60
 GridRP.Height = FraInf.Height - 150
 FraDiarRp.Width = FraInf.Width: FraDiarRp.Height = FraInf.Height
 Frame1.Width = FraDiarRp.Width - Frame1.Left - 150
 Frame1.Height = FraDiarRp.Height - Frame1.Top - 150
 GridResumen.Width = Frame1.Width - 60
 GridResumen.Height = Frame1.Height - 150
 LblTit.Width = Frame1.Width
 GridRP.ColWidth(0) = 0
 GridRP.ColWidth(2) = GridRP.Width - 7500
Salir:
End Sub

Private Sub Form_Activate()
    If PosicionInicialAplicada Then Exit Sub
    PrepararPosicionVentana Me
    PosicionInicialAplicada = True
End Sub
