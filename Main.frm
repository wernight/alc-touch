VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALC Touch: Modifications simultanée sur plusieurs fichiers (32 bits)"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRafraichirStop 
      DownPicture     =   "Main.frx":0442
      Height          =   300
      Left            =   -360
      Picture         =   "Main.frx":0A12
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Rafraichit la liste des fichiers qui seront modifiés."
      Top             =   5280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton cmdRafraichirOri 
      DisabledPicture =   "Main.frx":114A
      DownPicture     =   "Main.frx":19C6
      Height          =   300
      Left            =   -360
      Picture         =   "Main.frx":211A
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Rafraichit la liste des fichiers qui seront modifiés."
      Top             =   4800
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4305
      Index           =   3
      Left            =   7200
      ScaleHeight     =   4275
      ScaleWidth      =   7245
      TabIndex        =   49
      Top             =   5160
      Visible         =   0   'False
      Width           =   7275
      Begin VB.CommandButton cmdStop 
         DownPicture     =   "Main.frx":286E
         Height          =   870
         Left            =   6150
         Picture         =   "Main.frx":43DA
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   -105
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CommandButton cmdEffectuer 
         DisabledPicture =   "Main.frx":5F46
         DownPicture     =   "Main.frx":7AB2
         Height          =   870
         Left            =   6150
         Picture         =   "Main.frx":961E
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   810
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picOpération 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2925
         Left            =   -45
         ScaleHeight     =   2925
         ScaleWidth      =   7275
         TabIndex        =   53
         Top             =   1335
         Visible         =   0   'False
         Width           =   7275
         Begin VB.PictureBox picProgress 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   2055
            Picture         =   "Main.frx":B18A
            ScaleHeight     =   195
            ScaleWidth      =   3525
            TabIndex        =   76
            Top             =   615
            Width           =   3525
         End
         Begin MSComctlLib.ListView lwFichiersModif 
            Height          =   1815
            Left            =   210
            TabIndex        =   59
            ToolTipText     =   "Liste des fichiers modifiés (après leurs modification)."
            Top             =   1110
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   3201
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Nom"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Modifié"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Attributs"
               Object.Width           =   1411
            EndProperty
         End
         Begin VB.CheckBox chkAffFichiersModifs 
            Caption         =   "Afficher les fichiers modifiés après les modifications (ci-dessous)"
            Height          =   195
            Left            =   945
            TabIndex        =   58
            ToolTipText     =   "Une fois que les fichier auront été modifiés, affiche ces fichiers dans la liste (ci-dessous)."
            Top             =   900
            Width           =   4815
         End
         Begin VB.TextBox txtProgress 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   2025
            TabIndex        =   57
            Top             =   570
            Width           =   3600
         End
         Begin VB.TextBox txtOpération 
            BackColor       =   &H8000000F&
            DragMode        =   1  'Automatic
            Enabled         =   0   'False
            Height          =   285
            Left            =   2025
            TabIndex        =   55
            Text            =   "(aucune)"
            Top             =   270
            Width           =   3600
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Progression :"
            Height          =   195
            Index           =   5
            Left            =   945
            TabIndex        =   56
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Opération :"
            Height          =   195
            Index           =   4
            Left            =   945
            TabIndex        =   54
            Top             =   300
            Width           =   780
         End
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   360
         Index           =   12
         Left            =   2445
         ScaleHeight     =   360
         ScaleWidth      =   2190
         TabIndex        =   70
         Top             =   0
         Width           =   2190
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   390
         Index           =   11
         Left            =   2370
         ScaleHeight     =   390
         ScaleWidth      =   2250
         TabIndex        =   69
         Top             =   1245
         Width           =   2250
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   1605
         Index           =   10
         Left            =   4380
         ScaleHeight     =   1605
         ScaleWidth      =   450
         TabIndex        =   68
         Top             =   -105
         Width           =   450
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   1605
         Index           =   9
         Left            =   2325
         ScaleHeight     =   1605
         ScaleWidth      =   450
         TabIndex        =   67
         Top             =   -150
         Width           =   450
      End
      Begin VB.CommandButton cmdOK 
         DisabledPicture =   "Main.frx":D5C0
         DownPicture     =   "Main.frx":F12C
         Height          =   1335
         Left            =   2490
         Picture         =   "Main.frx":10C98
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Modifier les fichiers avec les opérations sélectionnées."
         Top             =   150
         Width           =   2190
      End
      Begin VB.Label lblCause 
         AutoSize        =   -1  'True
         Caption         =   "- xxx....."
         Height          =   195
         Left            =   930
         TabIndex        =   52
         Top             =   2100
         Width           =   5400
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Opération impossible :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   51
         Top             =   1650
         Width           =   1890
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4305
      Index           =   2
      Left            =   6960
      ScaleHeight     =   4275
      ScaleWidth      =   7260
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   7290
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   1905
         Index           =   3
         Left            =   4065
         TabIndex        =   43
         Top             =   2010
         Width           =   2700
         Begin VB.CheckBox chkAttrib 
            Caption         =   "Fichier système"
            Height          =   225
            Index           =   3
            Left            =   240
            TabIndex        =   48
            Tag             =   "4"
            ToolTipText     =   "Fichier utilisé par le système."
            Top             =   1455
            Value           =   2  'Grayed
            Width           =   1590
         End
         Begin VB.CheckBox chkAttrib 
            Caption         =   "Fichier caché"
            Height          =   225
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Tag             =   "2"
            ToolTipText     =   "Fichier invisible."
            Top             =   1095
            Value           =   2  'Grayed
            Width           =   1590
         End
         Begin VB.CheckBox chkAttrib 
            Caption         =   "Archive"
            Height          =   225
            Index           =   1
            Left            =   240
            TabIndex        =   46
            Tag             =   "32"
            ToolTipText     =   "Fichier modifié depuis la dernière sauvegarde."
            Top             =   735
            Value           =   2  'Grayed
            Width           =   1590
         End
         Begin VB.CheckBox chkAttrib 
            Caption         =   "Lecture seule"
            Height          =   225
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Tag             =   "1"
            ToolTipText     =   "Accès en écriture interdit."
            Top             =   375
            Value           =   2  'Grayed
            Width           =   1590
         End
         Begin VB.CheckBox chkChangeAttrib 
            Caption         =   "&Attributs"
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
            Left            =   150
            TabIndex        =   44
            ToolTipText     =   "Changer l'attribut des fichiers"
            Top             =   0
            Width           =   1050
         End
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   1590
         Index           =   2
         Left            =   4065
         TabIndex        =   38
         Top             =   195
         Width           =   2700
         Begin VB.TextBox txtRen 
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   42
            Text            =   "*1.???"
            ToolTipText     =   "Renommer les fichiers en les spécification ci-dessous (Voir l'aide)."
            Top             =   1005
            Width           =   2130
         End
         Begin VB.TextBox txtRen 
            Height          =   285
            Index           =   0
            Left            =   240
            TabIndex        =   40
            Text            =   "*1.???"
            ToolTipText     =   "Renomers les fichiers correspondant aux spécifications ci-dessous (Voir l'aide)."
            Top             =   420
            Width           =   2130
         End
         Begin VB.CheckBox chkChangeRen 
            Caption         =   "&Renommer"
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
            Left            =   150
            TabIndex        =   39
            ToolTipText     =   "Renommer les fichiers"
            Top             =   0
            Width           =   1230
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "en"
            Height          =   195
            Index           =   2
            Left            =   1170
            TabIndex        =   41
            Top             =   765
            Width           =   180
         End
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   1680
         Index           =   1
         Left            =   420
         TabIndex        =   33
         Top             =   2235
         Width           =   3240
         Begin VB.OptionButton optCase 
            Caption         =   "Que la première lettre en Majuscule."
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   37
            ToolTipText     =   "La première lettre en majuscule, le reste en minuscule."
            Top             =   1125
            Width           =   2835
         End
         Begin VB.OptionButton optCase 
            Caption         =   "Tout en MAJUSCULE"
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   36
            ToolTipText     =   "Toutes les lettres en majuscules."
            Top             =   705
            Width           =   2835
         End
         Begin VB.OptionButton optCase 
            Caption         =   "Tout en minuscule"
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   35
            ToolTipText     =   "Toutes les lettres en minuscules."
            Top             =   315
            Value           =   -1  'True
            Width           =   2835
         End
         Begin VB.CheckBox chkChangeCase 
            Caption         =   "&Case"
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
            Left            =   150
            TabIndex        =   34
            ToolTipText     =   "Changer la case du nom des fichiers."
            Top             =   0
            Width           =   765
         End
      End
      Begin VB.Frame Frame 
         Appearance      =   0  'Flat
         Caption         =   " "
         ForeColor       =   &H80000008&
         Height          =   1815
         Index           =   0
         Left            =   420
         TabIndex        =   17
         Top             =   210
         Width           =   3240
         Begin VB.CheckBox chkTouch 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   2310
            TabIndex        =   22
            Top             =   375
            Width           =   195
         End
         Begin VB.CheckBox chkTouch 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   1095
            TabIndex        =   21
            Top             =   360
            Width           =   195
         End
         Begin VB.CheckBox chkTouch 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   20
            Top             =   360
            Width           =   195
         End
         Begin VB.PictureBox picTouchOpt 
            BorderStyle     =   0  'None
            Height          =   915
            Left            =   255
            ScaleHeight     =   915
            ScaleWidth      =   2790
            TabIndex        =   75
            Top             =   660
            Width           =   2790
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   3
               Left            =   2340
               MaxLength       =   14
               TabIndex        =   32
               Text            =   "0"
               ToolTipText     =   "Millisecondes"
               Top             =   615
               Width           =   375
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   6
               Left            =   735
               MaxLength       =   14
               TabIndex        =   28
               Text            =   "2000"
               ToolTipText     =   "Année"
               Top             =   615
               Width           =   480
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   5
               Left            =   450
               MaxLength       =   14
               TabIndex        =   27
               Text            =   "1"
               ToolTipText     =   "Mois"
               Top             =   615
               Width           =   270
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   4
               Left            =   165
               MaxLength       =   14
               TabIndex        =   26
               Text            =   "1"
               ToolTipText     =   "Jour"
               Top             =   615
               Width           =   270
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   2
               Left            =   2010
               MaxLength       =   14
               TabIndex        =   31
               Text            =   "0"
               ToolTipText     =   "Secondes"
               Top             =   615
               Width           =   270
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   1
               Left            =   1695
               MaxLength       =   14
               TabIndex        =   30
               Text            =   "0"
               ToolTipText     =   "Minutes"
               Top             =   615
               Width           =   270
            End
            Begin VB.TextBox txtTouch 
               Alignment       =   1  'Right Justify
               Height          =   285
               Index           =   0
               Left            =   1380
               MaxLength       =   14
               TabIndex        =   29
               Text            =   "0"
               ToolTipText     =   "Heure"
               Top             =   615
               Width           =   270
            End
            Begin VB.OptionButton optTouch 
               Caption         =   "Définit par l'utilisateur:"
               Height          =   225
               Index           =   1
               Left            =   15
               TabIndex        =   24
               ToolTipText     =   "Date & Heure définit (ci-dessous)."
               Top             =   300
               Value           =   -1  'True
               Width           =   2355
            End
            Begin VB.OptionButton optTouch 
               Caption         =   "Date et heure courrant"
               Height          =   225
               Index           =   0
               Left            =   15
               TabIndex        =   23
               ToolTipText     =   "Date & Heure du jour au moment de la modification."
               Top             =   0
               Width           =   2355
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               Caption         =   "le                         à       :      :      :"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   25
               Top             =   645
               Width           =   2325
            End
         End
         Begin VB.CheckBox chkChangeTouch 
            Caption         =   "&Date && Heure"
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
            Left            =   150
            TabIndex        =   18
            ToolTipText     =   "Change la date & l'heure de la modification des fichiers."
            Top             =   0
            Width           =   1500
         End
         Begin MSComctlLib.TabStrip TabStripTouch 
            Height          =   1395
            Left            =   90
            TabIndex        =   19
            Top             =   285
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   2461
            MultiRow        =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "    Création"
                  Object.ToolTipText     =   "Date de création du fichier"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "    Modification"
                  Object.ToolTipText     =   "Date de la dernière modification du fichier"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "    Accès"
                  Object.ToolTipText     =   "Date du dernier accès au fichier"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4305
      Index           =   1
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   7275
      TabIndex        =   63
      Top             =   390
      Width           =   7275
      Begin MSComctlLib.ListView lwFichiers 
         Height          =   2280
         Index           =   1
         Left            =   3690
         TabIndex        =   15
         ToolTipText     =   "Liste des fichiers modifiables (cochez ceux que vous voulez modifier)"
         Top             =   1980
         Visible         =   0   'False
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   4022
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nom"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modifié"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Attributs"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   525
         Index           =   8
         Left            =   7110
         ScaleHeight     =   525
         ScaleWidth      =   255
         TabIndex        =   66
         Top             =   1575
         Width           =   255
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   120
         Index           =   6
         Left            =   5775
         ScaleHeight     =   120
         ScaleWidth      =   1545
         TabIndex        =   64
         Top             =   1470
         Width           =   1545
      End
      Begin MSComctlLib.ListView lwFichiers 
         Height          =   2280
         Index           =   0
         Left            =   3450
         TabIndex        =   14
         ToolTipText     =   "Liste des fichiers qui seront modifiés."
         Top             =   1830
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   4022
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nom"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Modifié"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Attributs"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CheckBox chkOnlySel 
         Caption         =   "Fichiers sélectionnés :"
         Height          =   195
         Left            =   3480
         TabIndex        =   11
         ToolTipText     =   "Ne change que les fichiers coché dans la liste (ci-dessous)"
         Top             =   1620
         Width           =   1995
      End
      Begin VB.TextBox txtSpécification 
         Height          =   285
         Left            =   4860
         TabIndex        =   10
         Text            =   "*.*"
         ToolTipText     =   "Si vous voulez changer que certains fichiers, entrer ici les spécifications (du DOS)"
         Top             =   1155
         Width           =   2055
      End
      Begin VB.Frame fraAttribToChange 
         Caption         =   "Attributs des fichiers à sélectionner"
         Height          =   885
         Left            =   3555
         TabIndex        =   4
         ToolTipText     =   "Seul les fichiers ayant les attributs ci-dessous seront modifiés."
         Top             =   150
         Width           =   3360
         Begin VB.CheckBox chkAttribToChange 
            Caption         =   "Système"
            Height          =   195
            Index           =   3
            Left            =   1905
            TabIndex        =   8
            Tag             =   "4"
            ToolTipText     =   "Inclure les fichiers ayant l'attribut système activé."
            Top             =   555
            Value           =   2  'Grayed
            Width           =   1350
         End
         Begin VB.CheckBox chkAttribToChange 
            Caption         =   "Cachés"
            Height          =   195
            Index           =   2
            Left            =   1905
            TabIndex        =   7
            Tag             =   "2"
            ToolTipText     =   "Inclure les fichiers ayant l'attribut invisible activé."
            Top             =   285
            Value           =   2  'Grayed
            Width           =   1350
         End
         Begin VB.CheckBox chkAttribToChange 
            Caption         =   "Archive"
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   6
            Tag             =   "32"
            ToolTipText     =   "Inclure les fichiers ayant l'attribut archive activé."
            Top             =   555
            Value           =   2  'Grayed
            Width           =   1545
         End
         Begin VB.CheckBox chkAttribToChange 
            Caption         =   "Lecture seule"
            Height          =   195
            Index           =   0
            Left            =   165
            TabIndex        =   5
            Tag             =   "1"
            ToolTipText     =   "Inclure les fichiers ayant l'attribut lecture seule activé."
            Top             =   270
            Value           =   2  'Grayed
            Width           =   1545
         End
      End
      Begin VB.CheckBox chkSub 
         Caption         =   "Inclure les sous répertoires"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         ToolTipText     =   "Chercher les fichiers dans le répertoire définit et ses sous-répertoires."
         Top             =   4035
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   3465
         Left            =   90
         TabIndex        =   2
         ToolTipText     =   "Répertoire où se trouve les fichiers à modifier"
         Top             =   510
         Width           =   3255
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         ToolTipText     =   "Lecteur"
         Top             =   150
         Width           =   3255
      End
      Begin VB.PictureBox picCache 
         BorderStyle     =   0  'None
         Height          =   525
         Index           =   7
         Left            =   5820
         ScaleHeight     =   525
         ScaleWidth      =   195
         TabIndex        =   65
         Top             =   1530
         Width           =   195
      End
      Begin VB.CommandButton cmdRafraichir 
         DisabledPicture =   "Main.frx":12804
         DownPicture     =   "Main.frx":13080
         Height          =   480
         Left            =   5925
         Picture         =   "Main.frx":13660
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Rafraichit la liste des fichiers qui seront modifiés."
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label lblNbrFichierSel 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   5475
         TabIndex        =   12
         ToolTipText     =   "Nombre de fichiers sélectionnés ou dans la liste."
         Top             =   1620
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "&Spécification :"
         Height          =   195
         Index           =   0
         Left            =   3645
         TabIndex        =   9
         Top             =   1185
         Width           =   1005
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   8281
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1: Fichiers à modifier"
            Object.ToolTipText     =   "Les fichiers sur lesquels seront effectués les opérations."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2: Opérations à effectuer sur ces fichiers"
            Object.ToolTipText     =   "Opérations à effectuer sur les fichiers définis en (1)."
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3: Effectuer les opérations"
            Object.ToolTipText     =   "Effectuer les opérations défnis en (2) sur les fichiers définis en (1)"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   7380
      TabIndex        =   61
      Top             =   5400
      Width           =   7380
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   60
      ScaleHeight     =   255
      ScaleWidth      =   7380
      TabIndex        =   60
      Top             =   4680
      Width           =   7380
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   4
      Left            =   5130
      ScaleHeight     =   915
      ScaleWidth      =   690
      TabIndex        =   62
      Top             =   4740
      Width           =   690
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   2
      Left            =   3420
      ScaleHeight     =   915
      ScaleWidth      =   690
      TabIndex        =   81
      Top             =   4770
      Width           =   690
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   13
      Left            =   6885
      ScaleHeight     =   915
      ScaleWidth      =   690
      TabIndex        =   83
      Top             =   4785
      Width           =   690
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   3
      Left            =   1665
      ScaleHeight     =   915
      ScaleWidth      =   690
      TabIndex        =   82
      Top             =   4755
      Width           =   690
   End
   Begin VB.PictureBox picCache 
      BorderStyle     =   0  'None
      Height          =   915
      Index           =   5
      Left            =   -75
      ScaleHeight     =   915
      ScaleWidth      =   690
      TabIndex        =   84
      Top             =   4755
      Width           =   690
   End
   Begin VB.CommandButton cmdAbout 
      DownPicture     =   "Main.frx":13DB8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3945
      Picture         =   "Main.frx":14B44
      Style           =   1  'Graphical
      TabIndex        =   78
      ToolTipText     =   "A propos du programme."
      Top             =   4815
      Width           =   1335
   End
   Begin VB.CommandButton cmdQuitter 
      DownPicture     =   "Main.frx":158D0
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5685
      Picture         =   "Main.frx":1665C
      Style           =   1  'Graphical
      TabIndex        =   77
      ToolTipText     =   "Retourner à Windows"
      Top             =   4815
      Width           =   1335
   End
   Begin VB.CommandButton cmdAide 
      DownPicture     =   "Main.frx":173E8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2220
      Picture         =   "Main.frx":18174
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Aide"
      Top             =   4815
      Width           =   1335
   End
   Begin VB.CommandButton cmdExecRapide 
      DownPicture     =   "Main.frx":18F00
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Picture         =   "Main.frx":19C8C
      Style           =   1  'Graphical
      TabIndex        =   80
      ToolTipText     =   "Effectue les changements (accès rapide)."
      Top             =   4815
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const OPEN_EXISTING = 3
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function CreateFileNS Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SystemTime) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SystemTime, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SystemTime)
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim TouchIniTop As Integer, TabStripTouchSelIndex As Integer, TouchOpt(2) As Integer, TouchDateTime(2, 6) As Currency

Private Function PlaySound(ID, Sync As Boolean) As Long
    Dim WaveData() As Byte          ' buffer to hold wave resource
    
    ' Joue un son
    WaveData() = LoadResData(ID, "WAVE")
    If Sync Then
        PlaySound = sndPlaySound(WaveData(0), &H4)
    Else
        PlaySound = sndPlaySound(WaveData(0), &H1 Or &H4)
    End If
End Function

Private Sub chkAffFichiersModifs_Click()
    ' Aff/Cache la liste des fichiers modifs
    lwFichiersModif.Visible = chkAffFichiersModifs.Value
End Sub

Private Sub chkAttrib_Click(Index As Integer)
    Static iValue(0 To 3) As Integer, bIni As Boolean, bBusy As Boolean
    
    ' Initialise
    If bIni = False Then
        For i% = 0 To 3
            iValue(i%) = GetSetting(App.EXEName, "Etat précédent", "Attributs" & i%, chkAttrib(i%))
            chkAttrib(i%) = iValue(i%)
        Next
        bIni = True
        Exit Sub
    End If
    
    ' Permet de mettre en grisé
    If bBusy Then Exit Sub
    bBusy = True
    iValue(Index) = (iValue(Index) + 1) Mod 3
    chkAttrib(Index).Value = iValue(Index)
    bBusy = False

    ' Coche l'opt°
    OptChange 3
End Sub

Private Sub chkAttribToChange_Click(Index As Integer)
    Static iValue(0 To 3) As Integer, bIni As Boolean, bBusy As Boolean
    
    ' Initialise
    If bIni = False Then
        For i% = 0 To 3
            iValue(i%) = GetSetting(App.EXEName, "Etat précédent", "AttribToChange" & i%, chkAttribToChange(i%))
            chkAttribToChange(i%) = iValue(i%)
        Next
        bIni = True
        Exit Sub
    End If
    
    ' Permet de mettre en grisé
    If bBusy Then Exit Sub
    bBusy = True
    iValue(Index) = (iValue(Index) + 1) Mod 3
    chkAttribToChange(Index).Value = iValue(Index)
    bBusy = False
End Sub

Private Sub chkOnlySel_Click()
    ' Affiche la liste correspondante
    If chkOnlySel Then
        lwFichiers(0).Visible = False
    Else
        lwFichiers(1).Visible = False
    End If
    lwFichiers(chkOnlySel).Visible = True

    ' Nbr de fichier
    lwFichiers_Click chkOnlySel
    
    ' Affiche le nbr de fichiers sél.
    lwFichiers_Click chkOnlySel
End Sub

Private Sub chkTouch_Click(Index As Integer)
    Dim bValue As Integer
    
    ' Coche la case générale de l'opt°
    For i = 0 To 2
        If chkTouch(i).Value <> 0 Then bValue = 1
    Next
    chkChangeTouch.Value = bValue
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub cmdAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdAide_Click()
    Dim str As String
    str = App.Path
    If Right(str, 1) <> "\" Then str = str & "\"
    ShellExecute hwnd, "open", str & "Aide.html", "", str, vbNormalFocus
End Sub

Private Sub cmdAide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdExecRapide_Click()
    TabStrip1_Click
    If cmdOK.Enabled Then
        cmdOK_Click
    Else
        MsgBox "Vous ne pouvez pas exécuter les changement," & Chr(10) & "pour plus d'informations veuillez cliquer sur l'onglet " & Chr(10) & """" & Mid(TabStrip1.Tabs(3).Caption, 2) & """.", vbInformation
    End If
End Sub

Private Sub cmdExecRapide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdQuitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdRafraichir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then PlaySound 100, False
End Sub

Private Sub cmdOK_Click()
    Static bStop As Boolean
    Dim f As String, CurSubPath As String, Attrib As Integer, Directory As String
    Dim NbrFichiers As Long, NbrFichiersModif As Long, NbrFichiersTt As Long
    Dim AttribToBeSet As Integer, AttribToBeUnSet As Integer, iFileAttrib As Integer
    Dim hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME, SystemTime As SystemTime, lpLocalFileTime As FILETIME
    Dim strRen(1, 9) As String, nRenPos As Integer, strRenResult As String
    
    On Error GoTo OKErr
    
    ' Début
    If cmdOK.Picture <> cmdStop.Picture Then
        ' Met en attente
        Screen.MousePointer = vbArrowHourglass
        cmdOK.Picture = cmdStop.Picture
        cmdOK.DownPicture = cmdStop.DownPicture
        cmdRafraichir.Enabled = False
        picTab(1).Enabled = False
        picTab(2).Enabled = False
        picProgress.Width = 0
        txtOpération.Enabled = True
        chkAffFichiersModifs.Enabled = False
        lwFichiersModif.ListItems.Clear
        For i% = 0 To 2
            If optCase(i%) Then CaseToBeSet = i%
        Next
        For i% = 0 To 3
            If chkAttrib(i%) = 1 Then AttribToBeSet = AttribToBeSet Or chkAttrib(i%).Tag
            If chkAttrib(i%) = 0 Then AttribToBeUnSet = AttribToBeUnSet Or chkAttrib(i%).Tag
            If chkAttribToChange(i%) <> 0 Then Attrib = Attrib Or chkAttribToChange(i%).Tag
        Next
        Attrib = Attrib Or vbDirectory
        ' OnlySel: Définit la liste des fichiers a changer
        If chkOnlySel Then
            For i% = 1 To lwFichiers(chkOnlySel).ListItems.Count
                If lwFichiers(chkOnlySel).ListItems(i%).Checked Then lstFichiersAModif.AddItem lwFichiers(chkOnlySel).ListItems(i%).Text
            Next
        End If
        bStop = False
    Else
        ' Enlève l'attente
        bStop = True
        picTab(1).Enabled = True
        picTab(2).Enabled = True
        cmdOK.Picture = cmdEffectuer.Picture
        cmdOK.DownPicture = cmdEffectuer.DownPicture
        cmdRafraichir.Enabled = True
        txtOpération.Enabled = False
        txtOpération.Text = "Opérations annulée."
        txtProgress.Enabled = False
        txtProgress.Text = ""
        chkAffFichiersModifs.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Répertoire
    Directory = Dir1.Path
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    ' Calcule la taille total
    If Not chkOnlySel Then
        txtOpération.Text = "Calcule le nombre de fichiers à modifier."
        txtProgress.Enabled = True
        f = Dir(Directory & txtSpécification.Text, Attrib)
        If f = "" Then
            MsgBox "Erreur: Répertoire invalide", vbCritical
            cmdRafraichir_Click
            Exit Sub
        End If
        Do While True
            ' Donne le fichier
            If f = "." Or f = ".." Then GoTo NextDo2
            ' Fin
            If f = "" And CurSubPath = "" Then
                Exit Do
            ElseIf f = "" Then
                ' Se place au rép. d'avant et au fichier suivant
                For i% = Len(CurSubPath) - 1 To 2 Step -1
                    If Mid(CurSubPath, i%, 1) = "\" Then GoTo RepFound2
                Next
                i% = 0
RepFound2:
                PathToFind$ = Mid(CurSubPath, i% + 1, Len(CurSubPath) - i% - 1)
                CurSubPath = Left(CurSubPath, i%)
                f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
                Do
                    f = Dir
                Loop Until f = PathToFind$ And GetAttr(Directory & CurSubPath & f) = vbDirectory
            Else
                ' Répertoire
                If GetAttr(Directory & CurSubPath & f) = vbDirectory Then
                    If chkSub Then
                        ' Se place ds ce rép.
                        CurSubPath = CurSubPath & f & "\"
                        f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
                    End If
                ' Fichier
                Else
                    ' Ajoute le fichier/répertoire au total
                    NbrFichiersTt = NbrFichiersTt + 1
                End If
            End If
NextDo2:
            DoEvents
            ' Donne le fichier (suite)
            If bStop Then Exit Sub
            f = Dir
        Loop
        txtProgress.Text = ""
        txtProgress.Enabled = False
    End If
    
    ' Effectue les opérations
    txtOpération.Text = "Effectu les changements."
    f = Dir(Directory & txtSpécification.Text, Attrib)
    If f = "" Then
        MsgBox "Erreur: Répertoire invalide", vbCritical
        cmdRafraichir_Click
        Exit Sub
    End If
    Do While True
        ' Donne le fichier
        If f = "." Or f = ".." Then GoTo NextDo3
        ' Fin
        If f = "" And CurSubPath = "" Then
            Exit Do
        ElseIf f = "" Then
            ' Se place au rép. d'avant et au fichier suivant
            For i% = Len(CurSubPath) - 1 To 2 Step -1
                If Mid(CurSubPath, i%, 1) = "\" Then GoTo RepFound3
            Next
            i% = 0
RepFound3:
            PathToFind$ = Mid(CurSubPath, i% + 1, Len(CurSubPath) - i% - 1)
            CurSubPath = Left(CurSubPath, i%)
            f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
            Do
                f = Dir
            Loop Until f = PathToFind$ And GetAttr(Directory & CurSubPath & f) = vbDirectory
        Else
            ' Répertoire
            If GetAttr(Directory & CurSubPath & f) = vbDirectory Then
                If chkSub Then
                    ' Se place ds ce rép.
                    CurSubPath = CurSubPath & f & "\"
                    f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
                End If
            ' Fichier
            Else
                ' Attributs voulus ?
                iFileAttrib = GetAttr(Directory & CurSubPath & f)
                For i% = 0 To 3
                    If chkAttribToChange(i%) <> 2 Then _
                        If ((iFileAttrib And chkAttribToChange(i%).Tag) = chkAttribToChange(i%).Tag) _
                            <> (chkAttribToChange(i%) = 1) Then GoTo AprèsChgt
                Next

                ' OnlySel & Fichier non sélectionné -> Rien faire
                If chkOnlySel Then
                    For i% = 0 To lstFichiersAModif.ListCount - 1
                        If lstFichiersAModif.List(i%) = CurSubPath & f Then
                            lstFichiersAModif.RemoveItem i%
                            GoTo EffChgt
                        End If
                    Next
                    GoTo AprèsChgt
                End If
EffChgt:
                                    
                ' Ren
                If chkChangeRen Then
                    ' Cherches les équivalents aux ?x et *x
                    nRenPos = 1
                    For i = 1 To Len(txtRen(0))
                        X = Mid(txtRen(0), i, 1)
                        ' Enregistre ce qui remplace le ?x
                        If X = "?" Then
                            strRen(0, Int(Mid(txtRen(0), i + 1, 1))) = Mid(f, nRenPos, 1)
                            i = i + 1
                            nRenPos = nRenPos + 1
                        ' Enregistre ce qui remplace le *x
                        ElseIf X = "*" Then
                            If Mid(txtRen(0), i + 2, 1) <> "?" Then
                                ' Cherche la fin des caractères
                                If i + 2 > Len(txtRen(0)) Then
                                    j = i + 2
                                    X = Len(f)
                                Else
                                    For j = i + 2 To Len(txtRen(0))
                                        If Mid(txtRen(0), j, 1) = "*" Or Mid(txtRen(0), j, 1) = "?" Then Exit For
                                    Next
                                    X = InStr(nRenPos, f, Mid(txtRen(0), i + 2, j - i - 2), vbTextCompare)
                                    X = X - 1
                                End If
                                If X > 0 Then _
                                    strRen(1, Int(Mid(txtRen(0), i + 1, 1))) = Mid(f, nRenPos, X - nRenPos + 1)
                                i = i + 1
                                nRenPos = X + 1
                            Else
                                ' Cherche la fin des ?
                                For j = i + 2 To Len(txtRen(0))
                                    If Mid(txtRen(0), j, 1) <> "?" Then Exit For
                                Next
                                ' Si les ? correspondent à des caractères
                                If nRenPos + j - i <= Len(f) Then
                                    ' Cherche la fin des caractères
                                    For k = j + 1 To Len(txtRen(0))
                                        If Mid(txtRen(0), k, 1) = "?" Or Mid(txtRen(0), k, 1) = "*" Or nRenPos + k - i > Len(t) Then Exit For
                                    Next
                                    X = InStr(j, f, Mid(txtRen(0), j, k - j), vbTextCompare)
                                    strRen(1, Int(Mid(txtRen(0), i + 1, 1))) = Mid(f, nRenPos, j - i)
                                    i = i + 1
                                    nRenPos = nRenPos + j - i
                                End If
                            End If
                        ' Check le nom
                        ElseIf UCase(X) <> UCase(Mid(f, nRenPos, 1)) Then
                            GoTo FinRen
                        Else
                            nRenPos = nRenPos + 1
                        End If
                        
                        ' Fin trop proche ? -> Eliminer ce fichier
                        If nRenPos > Len(f) + 1 Then GoTo FinRen
                    Next
                    
                    ' Check
                    If nRenPos <> Len(f) + 1 Then GoTo FinRen
                    
                    ' Renomme
                    strRenResult = ""
                    For i = 1 To Len(txtRen(1))
                        X = Mid(txtRen(1), i, 1)
                        ' Remplace les ?x
                        If X = "?" Then
                            strRenResult = strRenResult + strRen(0, Int(Mid(txtRen(1), i + 1, 1)))
                            i = i + 1
                        ' Remplace les *x
                        ElseIf X = "*" Then
                            strRenResult = strRenResult + strRen(1, Int(Mid(txtRen(1), i + 1, 1)))
                            i = i + 1
                        ' Autre char
                        Else
                            strRenResult = strRenResult + X
                        End If
                    Next
                    Name Directory & CurSubPath & f As Directory & CurSubPath & strRenResult
FinRen:
                End If
                
                ' Case
                If chkChangeCase Then
                    Select Case CaseToBeSet
                    Case 0:
                        Name Directory & CurSubPath & f As Directory & CurSubPath & Format(f, "<")
                        f = Format(f, "<")
                    Case 1:
                        Name Directory & CurSubPath & f As Directory & CurSubPath & Format(f, ">")
                        f = Format(f, ">")
                    Case 2:
                        Name Directory & CurSubPath & f As Directory & CurSubPath & Format(Left(f, 1), ">") & Format(Right(f, Len(f) - 1), "<")
                        f = Format(Left(f, 1), ">") & Format(Right(f, Len(f) - 1), "<")
                    End Select
                    DoEvents
                End If
                
                ' Attrib
                If chkChangeAttrib Then
                    SetAttr Directory & CurSubPath & f, AttribToBeSet Or GetAttr(Directory & CurSubPath & f) And (Not AttribToBeUnSet)
                    DoEvents
                End If
                
                ' TOUCH
                If chkChangeTouch Then
                    X = 0
                    If (GetAttr(Directory & CurSubPath & f) And vbReadOnly) = vbReadOnly Then
                        X = MsgBox("Le fichier " & Directory & CurSubPath & f & " est en lecture seule. Etes vous sur de vouloir changer la date et l'heure ?", vbQuestion Or vbYesNoCancel)
                        If X = vbCancel Then
                            cmdOK_Click
                            Exit Sub
                        ElseIf X = vbYes Then
                            SetAttr Directory & CurSubPath & f, GetAttr(Directory & CurSubPath & f) Xor vbReadOnly
                        Else
                            GoTo FinTouch
                        End If
                    End If
                    
                    ' Ouvre un fichier pour avoir son handle
                    hFile = CreateFileNS(Directory & CurSubPath & f, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_COMPRESSED Or FILE_ATTRIBUTE_DIRECTORY Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_TEMPORARY, 0)
                    If X <> 0 Then SetAttr Directory & CurSubPath & f, GetAttr(Directory & CurSubPath & f) Or vbReadOnly
                    If hFile = -1 Then ' Le fichier ne peux pas s'ouvrir
                        If MsgBox("Erreur: Impossible de modifier le fichier." & Chr(10) & "Fichier en cours: " & Directory & CurSubPath & f & Chr(10) & Chr(10) & "Voulez-vous continuer ?", vbQuestion Or vbYesNo) = vbNo Then _
                            cmdOK_Click Else GoTo FinTouch
                    End If
                    ' Récuprère le différentes dates du fichiers
                    GetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
                    
                    ' Change les dates/heure
                    For i = 0 To 2
                        ' Change la date ?
                        If chkTouch(i).Value Then
                            If TouchOpt(i) = 0 Then
                                ' Date en cours
                                GetLocalTime SystemTime
                            Else
                                ' Date définie par l'utilisateur
                                SystemTime.wHour = TouchDateTime(i, 0)
                                SystemTime.wMinute = TouchDateTime(i, 1)
                                SystemTime.wSecond = TouchDateTime(i, 2)
                                SystemTime.wMilliseconds = TouchDateTime(i, 3)
                                SystemTime.wDay = TouchDateTime(i, 4)
                                SystemTime.wMonth = TouchDateTime(i, 5)
                                SystemTime.wYear = TouchDateTime(i, 6)
                            End If
                            
                            ' Change la date en mémoire
                            SystemTimeToFileTime SystemTime, lpLocalFileTime
                            Select Case i
                            Case 0:
                                LocalFileTimeToFileTime lpLocalFileTime, lpCreationTime
                            Case 1:
                                LocalFileTimeToFileTime lpLocalFileTime, lpLastWriteTime
                            Case 2:
                                LocalFileTimeToFileTime lpLocalFileTime, lpLastAccessTime
                            End Select
                        End If
                    Next
                    
                    ' Change la date sur le disque
                    SetFileTime hFile, lpCreationTime, lpLastAccessTime, lpLastWriteTime
                    
                    ' Ferme le fichier
                    CloseHandle hFile
FinTouch:
                End If
            
                ' Ajoute le fichier à la liste
                If chkAffFichiersModifs.Value Then
                    iFileAttrib = GetAttr(Directory & CurSubPath & f)
                    Set itmX = lwFichiersModif.ListItems.Add(, , CurSubPath & f)
                    itmX.SubItems(1) = CStr(FileDateTime(Directory & CurSubPath & f))
                    sFileAttrib = ""
                    If (iFileAttrib And vbReadOnly) = vbReadOnly Then sFileAttrib = sFileAttrib & "R" Else sFileAttrib = sFileAttrib & ". "
                    If (iFileAttrib And vbHidden) = vbHidden Then sFileAttrib = sFileAttrib & "H" Else sFileAttrib = sFileAttrib & ". "
                    If (iFileAttrib And vbSystem) = vbSystem Then sFileAttrib = sFileAttrib & "S" Else sFileAttrib = sFileAttrib & ". "
                    If (iFileAttrib And vbArchive) = vbArchive Then sFileAttrib = sFileAttrib & "A" Else sFileAttrib = sFileAttrib & ". "
                    itmX.SubItems(2) = sFileAttrib
                End If
                
AprèsChgt:
                ' Affiche la progression
                NbrFichiers = NbrFichiers + 1
                picProgress.Width = NbrFichiers * (txtProgress.Width - 4 * Screen.TwipsPerPixelX) / NbrFichiersTt
            End If
        End If
NextDo3:
        DoEvents
        ' Donne le fichier (suite)
        If bStop Then Exit Sub
        f = Dir
    Loop
    
    ' Fin
    cmdOK_Click
    txtOpération.Text = "Opérations terminées sur " & NbrFichiers & " fichiers."
    Exit Sub
OKErr:
    If Err Then If MsgBox("Erreur " & Err & ": " & Error(Err) & "." & Chr(10) & "Fichier en cours: " & Directory & CurSubPath & f & Chr(10) & Chr(10) & "Voulez-vous continuer ?", vbCritica Or vbYesNo) = vbYes Then Resume Next
    cmdOK_Click
End Sub

Private Sub cmdQuitter_Click()
    ' Enregistre les options
    SaveSetting App.EXEName, "Etat précédent", "Lecteur", Drive1.Drive
    SaveSetting App.EXEName, "Etat précédent", "Répertoire", Dir1.Path
    SaveSetting App.EXEName, "Etat précédent", "Spécification", txtSpécification
    SaveSetting App.EXEName, "Etat précédent", "Sous-répertoires", chkSub
    SaveSetting App.EXEName, "Etat précédent", "QueSelListe", chkOnlySel
    For i% = 0 To 2
        SaveSetting App.EXEName, "Etat précédent", "TouchOpt" & i%, TouchOpt(i%)
        SaveSetting App.EXEName, "Etat précédent", "chkTouch" & i%, chkTouch(i%).Value
    Next
    For i% = 0 To 2
        If optCase(i%) Then SaveSetting App.EXEName, "Etat précédent", "Case", i%
    Next
    SaveSetting App.EXEName, "Etat précédent", "RenommerSource", txtRen(0)
    SaveSetting App.EXEName, "Etat précédent", "RenommerDestination", txtRen(1)
    SaveSetting App.EXEName, "Etat précédent", "AfficherFichiersModifiés", chkAffFichiersModifs.Value
    For i% = 0 To 3
        If chkAttrib(i%) Then X = X Or chkAttrib(i%).Tag
    Next
    For i% = 0 To 3
        SaveSetting App.EXEName, "Etat précédent", "Attributs" & i%, chkAttrib(i%)
        SaveSetting App.EXEName, "Etat précédent", "AttribToChange" & i%, chkAttribToChange(i%)
    Next
    
    ' Sort
    End
End Sub

Private Sub cmdRafraichir_Click()
    Static bStop As Boolean
    Dim f As String, CurSubPath As String, Attrib As Integer, Directory As String
    Dim itmX As ListItem, iFileAttrib As Integer, sFileAttrib As String
        
    On Error GoTo SpécificationErr
    
    ' Rafraichit / Stop
    If cmdRafraichir.Picture <> cmdRafraichirStop.Picture Then
        ' Place en attente
        Screen.MousePointer = vbArrowHourglass
        cmdRafraichir.Picture = cmdRafraichirStop.Picture
        cmdRafraichir.DownPicture = cmdRafraichirStop.DownPicture
        chkOnlySel.Enabled = False
        Drive1.Enabled = False
        Dir1.Enabled = False
        chkSub.Enabled = False
        fraAttribToChange.Enabled = False
        txtSpécification.Enabled = False
        TabStrip1.Enabled = False
        lblNbrFichierSel = ""
        ' Définit les var. pour l'opération
        For i% = 0 To 3
            If chkAttribToChange(i%) Then Attrib = Attrib Or chkAttribToChange(i%).Tag
        Next
        Attrib = Attrib Or vbDirectory
        bStop = False
    Else
        bStop = True
        ' Nbr de fichier
        lwFichiers_Click chkOnlySel
        ' Enlève l'attente
        cmdRafraichir.Picture = cmdRafraichirOri.Picture
        cmdRafraichir.DownPicture = cmdRafraichirOri.DownPicture
        Drive1.Enabled = True
        Dir1.Enabled = True
        chkSub.Enabled = True
        fraAttribToChange.Enabled = True
        txtSpécification.Enabled = True
        TabStrip1.Enabled = True
        chkOnlySel.Enabled = True
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    ' Répertoire
    Directory = Dir1.Path
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    
    ' Rafraichit la liste de fichiers
    f = Dir(Directory & txtSpécification.Text, Attrib)
    If f = "" Then
        MsgBox "Erreur: Répertoire invalide", vbCritical
        cmdRafraichir_Click
        Exit Sub
    End If
    lwFichiers(chkOnlySel).ListItems.Clear
    Do While True
        ' Donne le fichier
        If f = "." Or f = ".." Then GoTo NextDo1
        ' Fin
        If f = "" And CurSubPath = "" Then
            Exit Do
        ElseIf f = "" Then
            ' Se place au rép. d'avant et au fichier suivant
            For i% = Len(CurSubPath) - 1 To 2 Step -1
                If Mid(CurSubPath, i%, 1) = "\" Then GoTo RepFound1
            Next
            i% = 0
RepFound1:
            PathToFind$ = Mid(CurSubPath, i% + 1, Len(CurSubPath) - i% - 1)
            CurSubPath = Left(CurSubPath, i%)
            f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
            Do
                f = Dir
            Loop Until f = PathToFind$ And GetAttr(Directory & CurSubPath & f) = vbDirectory
        Else
            ' Répertoire
            If GetAttr(Directory & CurSubPath & f) = vbDirectory Then
                If chkSub Then
                    ' Se place ds ce rép.
                    CurSubPath = CurSubPath & f & "\"
                    f = Dir(Directory & CurSubPath & txtSpécification.Text, Attrib)
                End If
            ' Fichier
            Else
                ' Attributs voulus ?
                iFileAttrib = GetAttr(Directory & CurSubPath & f)
                For i% = 0 To 3
                    If chkAttribToChange(i%) <> 2 Then _
                        If ((iFileAttrib And chkAttribToChange(i%).Tag) = chkAttribToChange(i%).Tag) _
                            <> (chkAttribToChange(i%) = 1) Then GoTo NextDo1
                Next

                ' Ajoute le fichier à la liste
                Set itmX = lwFichiers(chkOnlySel).ListItems.Add(, , CurSubPath & f)
                itmX.SubItems(1) = CStr(FileDateTime(Directory & CurSubPath & f))
                sFileAttrib = ""
                If (iFileAttrib And vbReadOnly) = vbReadOnly Then sFileAttrib = sFileAttrib & "R" Else sFileAttrib = sFileAttrib & ". "
                If (iFileAttrib And vbHidden) = vbHidden Then sFileAttrib = sFileAttrib & "H" Else sFileAttrib = sFileAttrib & ". "
                If (iFileAttrib And vbSystem) = vbSystem Then sFileAttrib = sFileAttrib & "S" Else sFileAttrib = sFileAttrib & ". "
                If (iFileAttrib And vbArchive) = vbArchive Then sFileAttrib = sFileAttrib & "A" Else sFileAttrib = sFileAttrib & ". "
                itmX.SubItems(2) = sFileAttrib
            End If
        End If
NextDo1:
        DoEvents
        ' Donne le fichier (suite)
        If bStop Then Exit Sub
        f = Dir
    Loop
    
    ' Sélectionne
    If chkOnlySel Then
        For i% = 1 To lwFichiers(chkOnlySel).ListItems.Count
            lwFichiers(chkOnlySel).ListItems(i%).Checked = True
        Next
    End If

SpécificationErr:
    If Err Then If MsgBox("Erreur " & Err & ": " & Error(Err) & "." & Chr(10) & "Fichier en cours: " & Directory & CurSubPath & f & Chr(10) & Chr(10) & "Voulez-vous continuer ?", vbQuestion Or vbYesNo) = vbYes Then Resume Next
    ' Fin
    cmdRafraichir_Click
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    On Error Resume Next

    ' Initialise
    TouchIniTop = chkTouch(0).Top
    For i% = 0 To 2
        For j% = 0 To 6
            TouchDateTime(i%, j%) = txtTouch(j%).Text
        Next
    Next
    
    ' Met en place les éléments
    For i% = 2 To 3
        picTab(i%).Move picTab(1).Left, picTab(1).Top
        picTab(i%).BorderStyle = 0
    Next
    lwFichiers(1).Move lwFichiers(0).Left, lwFichiers(0).Top
    picProgress.Width = 0
    
    ' Charge les options
    Drive1.Drive = GetSetting(App.EXEName, "Etat précédent", "Lecteur", Drive1.Drive)
    Dir1.Path = GetSetting(App.EXEName, "Etat précédent", "Répertoire", Dir1.Path)
    txtSpécification = GetSetting(App.EXEName, "Etat précédent", "Spécification", txtSpécification)
    chkSub = GetSetting(App.EXEName, "Etat précédent", "Sous-répertoires", chkSub)
    chkOnlySel = GetSetting(App.EXEName, "Etat précédent", "QueSelListe", chkOnlySel)
    chkOnlySel_Click
    For i = 0 To 2
        chkTouch(i).Value = GetSetting(App.EXEName, "Etat précédent", "chkTouch" & i, 0)
        TouchOpt(i) = GetSetting(App.EXEName, "Etat précédent", "TouchOpt" & i, 0)
    Next
    optCase(GetSetting(App.EXEName, "Etat précédent", "Case", 2)).Value = True
    txtRen(0) = GetSetting(App.EXEName, "Etat précédent", "RenommerSource", txtRen(0))
    txtRen(1) = GetSetting(App.EXEName, "Etat précédent", "RenommerDestination", txtRen(1))
    chkAffFichiersModifs.Value = GetSetting(App.EXEName, "Etat précédent", "AfficherFichiersModifiés", chkAffFichiersModifs.Value)
    chkAttribToChange_Click 0
    chkAttrib_Click 0
    txtTouch_Change 0
    
    ' Se place sur Touch/date de Modification
    TabStripTouch.Tabs(2).Selected = True
    
    ' Décoche les opt°
    chkChangeTouch.Value = 0
    chkChangeCase.Value = 0
    chkChangeRen.Value = 0
    chkChangeAttrib.Value = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdQuitter_Click
End Sub

Private Sub lwFichiers_Click(Index As Integer)
    Dim iNbrFichierSel As Integer
    
    ' Calc. le nbr de fichiers sél.
    If chkOnlySel Then
        For i% = 1 To lwFichiers(chkOnlySel).ListItems.Count
            If lwFichiers(chkOnlySel).ListItems(i%).Checked Then _
                iNbrFichierSel = iNbrFichierSel + 1
        Next
        lblNbrFichierSel = iNbrFichierSel
    Else
        lblNbrFichierSel = lwFichiers(chkOnlySel).ListItems.Count
    End If
End Sub

Private Sub optCase_Click(Index As Integer)
    ' Coche l'opt°
    OptChange 1
End Sub

Private Sub optTouch_Click(Index As Integer)
    ' Affiche les opt° correspondant au Tab
    TouchOpt(TabStripTouch.SelectedItem.Index - 1) = Index
    
    ' Coche l'opération
    chkTouch(TabStripTouch.SelectedItem.Index - 1).Value = 1
    chkTouch_Click 0
End Sub

Private Sub TabStripTouch_BeforeClick(Cancel As Integer)
    ' Affiche les cases à cocher
    For i = 0 To 2
        If i = TabStripTouchSelIndex Then
            ' Place en pos° haute la coche sél°
            chkTouch(i).Top = TouchIniTop - 2 * Screen.TwipsPerPixelY
        Else
            ' Place en pos° basse la coche sél°
            chkTouch(i).Top = TouchIniTop
        End If
    Next
End Sub

Private Sub TabStripTouch_Click()
    Dim chkTouchValue As Integer, chkChangeTouchValue As Integer
    TabStripTouchSelIndex = TabStripTouch.SelectedItem.Index - 1

    ' Affiche les cases à cocher
    For i = 0 To 2
        If i = TabStripTouchSelIndex Then
            ' Place en pos° haute la coche sél°
            chkTouch(i).Top = TouchIniTop - 2 * Screen.TwipsPerPixelY
        Else
            ' Place en pos° basse la coche sél°
            chkTouch(i).Top = TouchIniTop
        End If
    Next
    
    ' Affiche les opt° correspondant au Tab
    chkChangeTouchValue = chkChangeTouch.Value
    chkTouchValue = chkTouch(TabStripTouchSelIndex).Value
    optTouch(TouchOpt(TabStripTouchSelIndex)).Value = True
    For i = 0 To 6
        txtTouch(i).Text = TouchDateTime(TabStripTouchSelIndex, i)
    Next
    chkTouch(TabStripTouchSelIndex).Value = chkTouchValue
    chkChangeTouch.Value = chkChangeTouchValue
End Sub

Private Sub TabStripTouch_KeyPress(KeyAscii As Integer)
    ' Coche avec ESPACE
    If Chr(KeyAscii) = " " Then
        If chkTouch(TabStripTouch.SelectedItem.Index - 1).Value = 1 Then
            chkTouch(TabStripTouch.SelectedItem.Index - 1).Value = 0
        Else
            chkTouch(TabStripTouch.SelectedItem.Index - 1).Value = 1
        End If
    End If
End Sub

Private Sub TabStripTouch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Pour BeforeClick
    For i = 0 To 2
        If X >= TabStripTouch.Tabs(i + 1).Left Then TabStripTouchSelIndex = i
    Next
End Sub

Private Sub txtRen_Change(Index As Integer)
    ' Coche l'opt°
    OptChange 2
End Sub

Private Sub txtTouch_Change(Index As Integer)
    Dim LastDay(1 To 12) As Integer

    LastDay(1) = 31
    If TestDateTime(Val(txtTouch(6)), 6) And (Val(txtTouch(6)) - 1600) Mod 4 = 0 Then
        LastDay(2) = 29
    Else
        LastDay(2) = 28
    End If
    LastDay(3) = 31
    LastDay(4) = 30
    LastDay(5) = 31
    LastDay(6) = 30
    LastDay(7) = 31
    LastDay(8) = 31
    LastDay(9) = 30
    LastDay(10) = 31
    LastDay(11) = 30
    LastDay(12) = 31
    
    ' Vérifit le TOUCH
    txtTouch(Index).BackColor = RGB(255, 255, 255)
    If TestDateTime(Val(txtTouch(Index)), Index) = False Then txtTouch(Index).BackColor = RGB(255, 128, 128)
    If TestDateTime(Val(txtTouch(5)), 5) Then
        If Val(txtTouch(4)) > LastDay(Val(txtTouch(5))) Or TestDateTime(Val(txtTouch(4)), 4) = False Then
            txtTouch(4).BackColor = RGB(255, 128, 128)
        Else
            txtTouch(4).BackColor = RGB(255, 255, 255)
        End If
    End If
    
    ' Affiche les opt° correspondant au Tab
    TouchDateTime(TabStripTouch.SelectedItem.Index - 1, Index) = Val(txtTouch(Index).Text)
    
    ' Coche l'opération
    chkTouch(TabStripTouch.SelectedItem.Index - 1).Value = 1
    chkTouch_Click 0
    optTouch(1).Value = True
End Sub

Private Sub TabStrip1_Click()
    Static iCurIndex As Integer

    ' Initialise
    If iCurIndex = 0 Then iCurIndex = 1

    ' Affiche le tab
    If TabStrip1.SelectedItem.Index <> iCurIndex Then
        picTab(TabStrip1.SelectedItem.Index).Visible = True
        picTab(iCurIndex).Visible = False
        iCurIndex = TabStrip1.SelectedItem.Index
        
        ' Change Touch Tabs
        If iCurIndex = 2 Then TabStripTouch_Click
    End If
    
    ' Teste pour Effectuer
    lblCause = ""
    ' Sélection seule
    If chkOnlySel = 1 And lwFichiers(chkOnlySel).ListItems.Count = 0 Then
        lblCause = lblCause & Chr(10) & "    - Vous avez demandé a effectuer les opérations sur les fichiers cochés, vous devez donc rafraichir la liste des fichiers et cocher des fichiers." & Chr(10)
    ElseIf chkOnlySel = 1 Then
        For i% = 1 To lwFichiers(chkOnlySel).ListItems.Count
            If lwFichiers(chkOnlySel).ListItems(i%).Checked Then GoTo TestOK
        Next
        lblCause = lblCause & Chr(10) & "    - Vous avez demandé a effectuer les opérations sur les fichiers cochés, or vous n'avez rien coché." & Chr(10)
TestOK:
    End If
    ' Pas d'opération
    If (chkChangeTouch + chkChangeCase + chkChangeRen + chkChangeAttrib) = 0 Then _
        lblCause = lblCause & Chr(10) & "    - Vous n'avez pas spécifié d'opération à effectuer. Il faut cocher la case à côté d'une opération pour l'activer." & Chr(10)
    ' Renommer incorrect
    If chkChangeRen.Value Then
        If txtRen(0) = "" Or txtRen(1) = "" Then
            lblCause = lblCause & Chr(10) & "    - Vous avez demandé à renomer des fichiers, mais vous n'avez pas entré de spécifications." & Chr(10)
        Else
            For i = 1 To Len(txtRen(0))
                s = Mid(txtRen(0), i, 1)
                s2 = Mid(txtRen(1), i, 1)
                ' Teste un "?" ou "*" n'a pas de chiffre après (2e zone)
                If (s2 = "?" Or s2 = "*") And (Len(txtRen(1)) = i Or Mid(txtRen(1), i + 1, 1) < "0" Or Mid(txtRen(1), i + 1, 1) > "9") Then _
                    lblCause = lblCause & Chr(10) & "    - Vous avez demandé à renomer des fichiers, et dans la deusième zone un ""*"" ou un ""?"" n'a pas de chiffre après lui. Allez voir l'aide en cas de problème." & Chr(10)
                ' Teste un "?" ou "*" n'a pas de chiffre après (1er zone)
                If (s = "?" Or s = "*") And (Len(txtRen(0)) = i Or Mid(txtRen(0), i + 1, 1) < "0" Or Mid(txtRen(0), i + 1, 1) > "9") Then
                    lblCause = lblCause & Chr(10) & "    - Vous avez demandé à renomer des fichiers, et dans la première zone un ""*"" ou un ""?"" n'a pas de chiffre après lui. Allez voir l'aide en cas de problème." & Chr(10)
                ' Teste  **  *?*  *??*  *...* (1er zone)
                ElseIf s = "*" Then
                    For j% = i + 2 To Len(txtRen(0)) Step 2
                        If Mid(txtRen(0), j%, 1) <> "?" Then Exit For
                    Next
                    If Mid(txtRen(0), j%, 1) = "*" Then _
                        lblCause = lblCause & Chr(10) & "    - Vous avez demandé à renomer des fichiers, et vous avez mis ""**"" ou ""*?*"" dans la première zone. Elvevez une ""*"". Allez voir l'aide en cas de problème." & Chr(10)
                ' Pas de caractères interdits "\ / : * ? " < > |"
                Else
                    If s = "\" Or s = "/" Or s = ":" Or s = """" Or s = "<" Or s = ">" Or s = "|" Or _
                        s2 = "\" Or s2 = "/" Or s2 = ":" Or s2 = """" Or s2 = "<" Or s2 = ">" Or s2 = "|" Then _
                        lblCause = lblCause & Chr(10) & "    - Vous avez demandé à renomer des fichiers, or l'une des zones de texte contient un caractère incorrect ""\ / : "" < > |""." & Chr(10)
                End If
            Next
        End If
    End If
    ' Date incorrecte
    If chkChangeTouch And optTouch(1) Then
        oriTabIndex = TabStripTouch.SelectedItem.Index
        For i% = 0 To 2
            If chkTouch(i%).Value <> 0 Then
                TabStripTouch.Tabs(i% + 1).Selected = True
                For j% = 0 To 6
                    If txtTouch(j%).BackColor = RGB(255, 128, 128) Then
                        lblCause = lblCause & Chr(10) & "    - Vous avez demandé à changer la date et l'heure, or vous avez définit une date ou une heure incorrecte." & Chr(10)
                        i% = 123
                        j% = 123
                        Exit For
                    End If
                Next
            End If
        Next
        TabStripTouch.Tabs(oriTabIndex).Selected = True
    End If
    ' Affiche/cache
    cmdOK.Enabled = (lblCause = "")
    lblCause.Visible = Not cmdOK.Enabled
    picOpération.Visible = cmdOK.Enabled
    lwFichiersModif.Visible = chkAffFichiersModifs.Value
End Sub

Private Sub OptChange(Frame As Integer)
    ' Coche l'opération modif
    Select Case Frame
    Case 0:
        If chkChangeTouch.Enabled Then chkChangeTouch.Value = 1
    Case 1:
        If chkChangeCase.Enabled Then chkChangeCase.Value = 1
    Case 2:
        If chkChangeRen.Enabled Then chkChangeRen.Value = 1
    Case 3:
        If chkChangeAttrib.Enabled Then chkChangeAttrib.Value = 1
    End Select
End Sub

Private Sub txtTouch_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Que des chiffres
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii > 31 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Public Function TestDateTime(ByVal Texte As String, ByVal Index As Integer) As Boolean
    ' Renvoi TRUE si c OK, FALSE si la date/heure est mauvaise
    TestDateTime = True
    Select Case Index
    Case 0:
        If Val(Texte) < 0 Or Val(Texte) > 23 Then TestDateTime = False
    Case 1:
        If Val(Texte) < 0 Or Val(Texte) > 59 Then TestDateTime = False
    Case 2:
        If Val(Texte) < 0 Or Val(Texte) > 59 Then TestDateTime = False
    Case 3:
        If Val(Texte) < 0 Or Val(Texte) > 999 Then TestDateTime = False
    Case 4:
        If Val(Texte) < 1 Or Val(Texte) > 31 Then TestDateTime = False
    Case 5:
        If Val(Texte) < 1 Or Val(Texte) > 12 Then TestDateTime = False
    Case 6:
        If Val(Texte) < 1601 Or Val(Texte) > 2200 Then TestDateTime = False
    End Select
End Function
