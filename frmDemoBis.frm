VERSION 5.00
Begin VB.Form frmDemoBis 
   Caption         =   "UserControl SysTray - Demo (child)"
   ClientHeight    =   4875
   ClientLeft      =   6210
   ClientTop       =   4965
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   5370
   Begin VB.CommandButton cmdClignote2 
      Caption         =   "&Blink icon"
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlternanceIcone2 
      Caption         =   "&Alternate icons"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin DemoSysTrayUserCtrl.ctlSysTrayBalloon ctlSysTrayBalloon2 
      Left            =   2520
      Tag             =   "UserControl n° 1 initialisé."
      Top             =   4440
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin DemoSysTrayUserCtrl.ctlSysTrayBalloon ctlSysTrayBalloon1 
      Left            =   2520
      Tag             =   "UserControl n° 1 initialisé."
      Top             =   4080
      _ExtentX        =   2143
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdMessage2 
      Caption         =   "&Timed message"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdAlternanceIcone1 
      Caption         =   "&Alternate icons"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdMessage1 
      Caption         =   "&Timed message"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdClignote1 
      Caption         =   "&Blink icon"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdStartStop 
      Caption         =   "&Start SysTray"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Empty icon"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "SysTray icon # 2"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   11
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "SysTray icon # 1"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   2520
      Picture         =   "frmDemoBis.frx":0000
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Standard icon #2"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Look at debug window (Ctrl-G) when playing messages"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "UserControl icons only visible at design time"
      Height          =   735
      Left            =   3840
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Alternate icon"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Standard icon #1"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmDemoBis.frx":0CCA
      Top             =   3960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmDemoBis.frx":1994
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuzMonMenu 
      Caption         =   "MenuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAfficherMessage 
         Caption         =   "Afficher le message n°1"
      End
      Begin VB.Menu mnuFaireClignoter 
         Caption         =   "Action sur le clignotement de l'icône n°1"
      End
      Begin VB.Menu mnuzLigne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmDemoBis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdStartStop_Click()

    If cmdStartStop.Caption <> "Stop SysTray" Then
        With ctlSysTrayBalloon1
            ' Icone à afficher dans le SysTray
            Set .IconPicture = Image1
            ' Texte à afficher au survol de la souris (128 caractères unicode maxi)
            .Tooltip = """SysTray et Balloon"" UserControl #1 - Demo"
            ' Démarrage
            .SysTrayAddIcon
        End With
        With ctlSysTrayBalloon2
            ' Icone à afficher dans le SysTray
            Set .IconPicture = Image2
            ' Texte à afficher au survol de la souris (128 caractères unicode maxi)
            .Tooltip = """SysTray et Balloon"" UserControl #2 - Demo"
            ' Démarrage
            .SysTrayAddIcon
        End With
        cmdStartStop.Caption = "Stop SysTray"
    Else
        ' Démontage
        ctlSysTrayBalloon1.SysTrayRemoveIcon
        ctlSysTrayBalloon2.SysTrayRemoveIcon
        cmdStartStop.Caption = "Start SysTray"
    End If
End Sub

Public Sub cmdClignote1_Click()

    With ctlSysTrayBalloon1
        If .BlinkIsRunning Then
            .BlinkStop
            
            cmdAlternanceIcone1.Enabled = True
        Else
            ' Pas d'icone à afficher en alternance
            Set .BlinkIconPicture = Image4
            ' Lance l'animation avec changement tous les 300 mSec
            .BlinkStart (300)
            
            If .BlinkIsRunning Then cmdAlternanceIcone1.Enabled = False
        End If
    End With

End Sub

Private Sub cmdAlternanceIcone1_Click()

    With ctlSysTrayBalloon1
        If .BlinkIsRunning Then
            .BlinkStop
        
            cmdClignote1.Enabled = True
        Else
            ' Deuxième icone à afficher en alternance
            Set .BlinkIconPicture = Image3
            ' Lance l'animation avec changement tous les 300 mSec
            .BlinkStart (300)
        
            If .BlinkIsRunning Then cmdClignote1.Enabled = False
        End If
    End With

End Sub

Private Sub cmdMessage1_Click()
    
    ' Affichage d'un message (Corps : 256 caractères unicode maxi)
    '                        (Titre :  64 caractères unicode maxi)
    ctlSysTrayBalloon1.BalloonTipShow _
        "Planet-Source-Code est vraiment formidable, non ?", _
        "Bonjour (# 1)" & vbCrLf & _
            "This UserControl has the great advantage of being a single object " & _
            "in the project (no module needed)." & vbCrLf & _
            "Start codes used :" & vbCrLf & _
            "- EbartSoft (SubClassing without ""AddressOf"" (vbfrance.com)" & vbCrLf & _
            "- Christopher (Heenix) Lord (planet-source-code.com).", _
        NIIF_INFO, _
        8000

End Sub

Private Sub cmdClignote2_Click()
    
    With ctlSysTrayBalloon2
        If .BlinkIsRunning Then
            .BlinkStop
        
            cmdAlternanceIcone2.Enabled = True
        Else
            ' Pas d'icone à afficher en alternance
            Set .BlinkIconPicture = Image4
            ' Lance l'animation avec changement tous les 400 mSec
            .BlinkStart (400)
            
            If .BlinkIsRunning Then cmdAlternanceIcone2.Enabled = False
        End If
    End With

End Sub

Public Sub cmdAlternanceIcone2_Click()

    With ctlSysTrayBalloon2
        If .BlinkIsRunning Then
            .BlinkStop
        
            cmdClignote2.Enabled = True
        Else
            ' Deuxième icone à afficher en alternance
            Set .BlinkIconPicture = Image3
            ' Lance l'animation avec changement tous les 400 mSec
            .BlinkStart (400)
        
            If .BlinkIsRunning Then cmdClignote2.Enabled = False
        End If
    End With

End Sub

Private Sub cmdMessage2_Click()
    ctlSysTrayBalloon2.BalloonTipShow _
        "Planet-Source-Code est vraiment formidable, non ?", _
        "Bonjour (# 2)" & vbCrLf & _
            "This UserControl has the great advantage of being a single object " & _
            "in the project (no module needed)." & vbCrLf & _
            "Start codes used :" & vbCrLf & _
            "- EbartSoft (SubClassing without ""AddressOf"" (vbfrance.com)" & vbCrLf & _
            "- Christopher (Heenix) Lord (planet-source-code.com).", _
        NIIF_INFO, _
        8000
End Sub


' ##### Menus

Private Sub mnuAfficherMessage_Click()
    Call cmdMessage1_Click
End Sub

Private Sub mnuFaireClignoter_Click()
    Call cmdClignote1_Click
End Sub

Private Sub mnuQuitter_Click()
    ' Pas de précaution à prendre avec le SubClassing du User Control :
    ' Il met fin au SubClassing dans son Terminate
    Unload Me
End Sub


' Evènements générés par le control utilisateur :

'##### Composant SysTray n°2

Private Sub ctlSysTrayBalloon1_BalloonClicked()
    Debug.Print "(1) Balloon Clic"
'    MsgBox "Quelle rapidité !", vbExclamation Or vbOKOnly, App.Title
End Sub

Private Sub ctlSysTrayBalloon1_BalloonTimeOut()
    Debug.Print "(1) Balloon TimeOut"
End Sub

Private Sub ctlSysTrayBalloon1_BalloonShow()
    Debug.Print "(1) Balloon appears"
End Sub

Private Sub ctlSysTrayBalloon1_Balloonclosed()
    Debug.Print "(1) Balloon closed by user"
End Sub

Private Sub ctlSysTrayBalloon1_Click()
    Debug.Print "(1) Click on icon"
End Sub

Private Sub ctlSysTrayBalloon1_DblClick(Button As Integer)
    Debug.Print "(1) DoubleClick on icon, button #" & Button
End Sub

Private Sub ctlSysTrayBalloon1_MouseMove()
'    Debug.Print "(1) Mouse Move"
End Sub

Private Sub ctlSysTrayBalloon1_MouseUp(Button As Integer)
    Debug.Print "(1) ClickUp on icon, button #" & Button
    PopupMenu mnuzMonMenu, vbPopupMenuLeftAlign, , , mnuAfficherMessage
End Sub

Private Sub ctlSysTrayBalloon1_PgmError(Source As String, Code As Long, Description As String)
    MsgBox "(1) Erreur " & CStr(Code) & " - " & Description, vbCritical Or vbOKOnly, _
           "ctlSysTrayBalloon1" & ", procédure """ & Source & """"
End Sub


'##### Composant SysTray n°2

Private Sub ctlSysTrayBalloon2_BalloonClicked()
    Debug.Print "(2) Balloon Clic"
'    MsgBox "Quelle rapidité !", vbExclamation Or vbOKOnly, App.Title
End Sub

Private Sub ctlSysTrayBalloon2_BalloonTimeOut()
    Debug.Print "(2) Balloon TimeOut"
End Sub

Private Sub ctlSysTrayBalloon2_BalloonShow()
    Debug.Print "(2) Balloon appears"
End Sub

Private Sub ctlSysTrayBalloon2_Balloonclosed()
    Debug.Print "(2) Balloon closed by user"
End Sub

Private Sub ctlSysTrayBalloon2_Click()
    Debug.Print "(2) Click on icon"
End Sub

Private Sub ctlSysTrayBalloon2_DblClick(Button As Integer)
    Debug.Print "(2) DoubleClick on icon, button #" & Button
End Sub

Private Sub ctlSysTrayBalloon2_MouseMove()
'    Debug.Print "(2) Mouse Move"
End Sub

Private Sub ctlSysTrayBalloon2_MouseUp(Button As Integer)
    Debug.Print "(2) ClickUp on icon, button #" & Button
    PopupMenu mnuzMonMenu, vbPopupMenuLeftAlign, , , mnuAfficherMessage
End Sub

Private Sub ctlSysTrayBalloon2_PgmError(Source As String, Code As Long, Description As String)
    MsgBox "(2) Erreur " & CStr(Code) & " - " & Description, vbCritical Or vbOKOnly, _
           "ctlSysTrayBalloon1" & ", procédure """ & Source & """"
End Sub

