VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de MonApplication"
   ClientHeight    =   4005
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6210
   ClipControls    =   0   'False
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   0  'User
   ScaleWidth      =   414
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "About.frx":0442
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   1200
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4725
      TabIndex        =   0
      Top             =   1185
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Infos système..."
      Height          =   345
      Left            =   4740
      TabIndex        =   2
      Top             =   1635
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Auteur :  W. BEROUX - WBC pour ALC."
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1050
      TabIndex        =   9
      Top             =   1785
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"About.frx":074C
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   5670
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Ce logiciel ne devra jamais être vendu."
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   5670
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   120
      Picture         =   "About.frx":0805
      Top             =   120
      Width           =   5910
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   6
      X2              =   400
      Y1              =   179
      Y2              =   179
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description :  Utilitaire pour la modification simultatée sur plusieurs fichiers sous Windows 95/98 ®."
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1050
      TabIndex        =   3
      Top             =   2175
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Titre de l'application"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   5
      Top             =   1140
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   7
      X2              =   400
      Y1              =   180
      Y2              =   180
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   1440
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "ALC Informatique vous offre ce logiciel dans sa version complête et non limitée."
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   255
      TabIndex        =   4
      Top             =   2865
      Width           =   5670
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Options de sécurité des clés de base de registres...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Types racines des clés de base de registres...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Chaîne terminée par un caractère nul Unicode.
Const REG_DWORD = 4                      ' Nombre 32 bits.

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "À propos de " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Essaie d'obtenir le chemin et le nom du programme Infos système dans la base de registre...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Essaie d'obtenir uniquement le chemin du programme Infos système dans la base de registre...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Valide l'existence de la version du fichier 32 bits connu.
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Erreur - Impossible de trouver le fichier...
        Else
            GoTo SysInfoErr
        End If
    ' Erreur - Impossible de trouver l'entrée de la base de registre...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "Les informations système ne sont pas disponibles actuellement", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Compteur de boucle.
    Dim rc As Long                                          ' Code de retour.
    Dim hKey As Long                                        ' Descripteur d'une clé de base de registres ouverte.
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Type de données d'une clé de base de registres.
    Dim tmpVal As String                                    ' Stockage temporaire pour une valeur de clé de base de registres.
    Dim KeyValSize As Long                                  ' Taille de la variable de la clé de base de registres.
    '------------------------------------------------------------
    ' Ouvre la clé de base de registres sous la racine clé {HKEY_LOCAL_MACHINE...}.
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvre la clé de base de registres.
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gère l'erreur...
    
    tmpVal = String$(1024, 0)                             ' Alloue de l'espace pour la variable.
    KeyValSize = 1024                                       ' Définit la taille de la variable.
    
    '------------------------------------------------------------
    ' Extrait la valeur de la clé de base de registres...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtient/Crée la valeur de la clé.
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gère l'erreur.
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ajoute une chaîne terminée par un caractère nul...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Caractère nul trouvé, extrait de la chaîne.
    Else                                                    ' WinNT ne termine pas la chaîne par un caractère nul...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Caractère nul non trouvé, extrait la chaîne uniquement.
    End If
    '------------------------------------------------------------
    ' Détermine le type de valeur de la clé pour la conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Recherche les types de données...
    Case REG_SZ                                             ' Type de données chaîne de la clé de la base de registres.
        KeyVal = tmpVal                                     ' Copie la valeur de la chaîne.
    Case REG_DWORD                                          ' Type de données double mot de la clé de base de registres.
        For i = Len(tmpVal) To 1 Step -1                    ' Convertit chaque bit.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construit la valeur caractère par caractère.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertit le mot double en chaîne.
    End Select
    
    GetKeyValue = True                                      ' Retour avec succès.
    rc = RegCloseKey(hKey)                                  ' Ferme la clé de base de registres
    Exit Function                                           ' Quitte.
    
GetKeyError:      ' Réinitialise après qu'une erreur s'est produite...
    KeyVal = ""                                             ' Affecte une chaîne vide à la valeur de retour.
    GetKeyValue = False                                     ' Retour avec échec.
    rc = RegCloseKey(hKey)                                  ' Ferme la clé de base de registres.
End Function
