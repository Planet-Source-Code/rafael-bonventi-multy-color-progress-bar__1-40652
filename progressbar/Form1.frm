VERSION 5.00
Object = "{1FE10612-3B57-11D2-902F-00606732229A}#6.0#0"; "ProgressMeter.ocx"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "Very Simple Progress Bar! 2 Lines of Code"
   ClientHeight    =   2070
   ClientLeft      =   4905
   ClientTop       =   4455
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5895
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Caption         =   "Actions : "
      Height          =   765
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   5805
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   375
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Go"
         Height          =   375
         Left            =   3570
         TabIndex        =   2
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Caption         =   "ProGressBar (Meter) : "
      Height          =   1155
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   5805
      Begin ProgressMeter.Meter met 
         Height          =   525
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   926
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   165
         Left            =   2970
         TabIndex        =   6
         Top             =   750
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   165
         Left            =   2730
         TabIndex        =   5
         Top             =   750
         Width           =   225
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'11/05/2002
'Rafael Bonventi
'rafael.bonventi@systemmarketing.com.br
'icq = 24770989


Private Sub Command1_Click()
    met.Percent = 0
       
End Sub
'met.Percent = met.Percent + Int(100 / ty)
Private Sub Command2_Click()
    Dim intCount As Integer
    Dim intCount2 As Integer
    Dim intAux As Integer
    
    met.Percent = 0
    met.FillColor = &HFF&
    
    intCount2 = 0
    intCount = 0
    
    For intCount = 0 To 10000
        DoEvents
        intAux = intAux + 1
        
        For intCount2 = 0 To 10000
            DoEvents
        Next
        
        If met.Percent = 100 Then
            MsgBox "I hope you all like it, couse i am so tired of the old blue color :-(", vbInformation
            
            Exit For
        Exit Sub
            Else
                met.Percent = met.Percent + 1
                Label1.Caption = Label1.Caption + 1
        End If
        
    Next
    
    
End Sub

Private Sub Form_Load()
    met.FillColor = &H8000000A
End Sub

