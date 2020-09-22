VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmScore 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Score"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   315
      Left            =   1740
      TabIndex        =   11
      Top             =   2340
      Width           =   855
   End
   Begin MSComctlLib.ListView lvwScore 
      Height          =   1935
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "imlBallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Player names"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Hand Total"
         Object.Width           =   5080
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total Score"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Cards Left"
         Object.Width           =   2540
      EndProperty
      Picture         =   "frmScore.frx":0000
   End
   Begin VB.Frame fraBk 
      Height          =   2055
      Left            =   3780
      TabIndex        =   0
      Top             =   2700
      Width           =   3855
      Begin VB.Frame fraTotalScore 
         Caption         =   "Total Score"
         Height          =   1575
         Left            =   2460
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   1020
            TabIndex        =   8
            Top             =   360
            Width           =   90
         End
      End
      Begin VB.Frame fraHandTotal 
         Caption         =   "Hand Total"
         Height          =   1575
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1215
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   9
            Top             =   660
            Width           =   90
         End
         Begin VB.Label lblHandTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Index           =   0
            Left            =   1020
            TabIndex        =   7
            Top             =   420
            Width           =   90
         End
      End
      Begin VB.Label lblPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label lblPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   900
      End
      Begin VB.Label lblPlayerName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   900
      End
   End
   Begin VB.Line linLine 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   4380
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line linLine 
      Index           =   0
      X1              =   60
      X2              =   4380
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "##########"
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
      Left            =   60
      TabIndex        =   12
      Top             =   1980
      Width           =   4335
   End
End
Attribute VB_Name = "frmScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsWinner As Boolean

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i        As Integer
    Dim j        As Integer
    Dim W        As Integer
    Dim Player() As Object
    Dim pName()  As String
    
    ReDim Player(Setting.Opponents) As Object
    ReDim pName(Setting.Opponents) As String
    
    IsWinner = False
    
    With lvwScore
        W = .Width - Screen.TwipsPerPixelX * 6
        
        .ColumnHeaders(1).Width = W * 0.3
        .ColumnHeaders(2).Width = W * 0.23
        .ColumnHeaders(3).Width = W * 0.24
        .ColumnHeaders(4).Width = W * 0.23
    End With
    
    With frmMain
        For i = LBound(Player()) To UBound(Player())
            Select Case i
            Case Is = 0
                Set Player(i) = .crdPlayerOne
            Case Is = 1
                Set Player(i) = .crdPlayerTwo
            Case Is = 2
                Set Player(i) = .crdPlayerThree
            Case Is = 3
                Set Player(i) = .crdPlayerFour
            End Select
        Next i
        
        Select Case Setting.Opponents
        Case Is = 1
            pName(0) = .lblPlayerName(0).Caption
            pName(1) = .lblPlayerName(1).Caption
        Case Is = 2
            pName(0) = .lblPlayerName(0).Caption
            pName(1) = .lblPlayerName(2).Caption
            pName(2) = .lblPlayerName(1).Caption
        Case Is = 3
            pName(0) = .lblPlayerName(0).Caption
            pName(1) = .lblPlayerName(2).Caption
            pName(2) = .lblPlayerName(1).Caption
            pName(3) = .lblPlayerName(3).Caption
        End Select
        
        Dim Total_Score As Integer
        
        For i = LBound(Player()) To UBound(Player())
            Total_Score = Uno.TotalPoints(Player(i))
            lvwScore.ListItems.Add , , pName(i)
            lvwScore.ListItems(i + 1).Bold = True
            lvwScore.ListItems(lvwScore.ListItems.Count).SubItems(1) = _
                Total_Score
            lvwScore.ListItems(lvwScore.ListItems.Count).SubItems(2) = 0
            lvwScore.ListItems(lvwScore.ListItems.Count).SubItems(3) = _
                Uno.TotalCards(Player(i))
        Next i
        
        Dim Lowest_Points As Integer
        Dim Lowest_Card As Integer
        
        Lowest_Points = CInt(lvwScore.ListItems(i).SubItems(1))
        Lowest_Card = CInt(lvwScore.ListItems(i).SubItems(3))
        
        For i = 1 To lvwScore.ListItems.Count
            Lowest_Points = IIf(Lowest_Points < CInt(lvwScore.ListItems(i).SubItems(1)), _
                                Lowest_Points, _
                                CInt(lvwScore.ListItems(i).SubItems(1)))
                                                                
        Next i
        
        For i = 1 To lvwScore.ListItems.Count
            Lowest_Card = IIf(Lowest_Card < CInt(lvwScore.ListItems(i).SubItems(3)), _
                              Lowest_Card, _
                              CInt(lvwScore.ListItems(i).SubItems(3)))
        Next i
        
        For i = 1 To lvwScore.ListItems.Count
            If (Val(lvwScore.ListItems(i).SubItems(1)) = Lowest_Points) And _
               (Val(lvwScore.ListItems(i).SubItems(3)) = Lowest_Card) Then
                lvwScore.ListItems(i).SubItems(2) = Total_Score
                lblMsg.Caption = lvwScore.ListItems(i).Text
                
                lvwScore.ListItems(i).ForeColor = vbRed
                For j = 1 To lvwScore.ListItems(i).ListSubItems.Count
                    lvwScore.ListItems(i).ListSubItems(j).ForeColor = vbRed
                Next j
                
                If i = 1 Then ' player one wins the game
                    IsWinner = True
                    
                    lblMsg.Caption = "You win the game."
                Else
                    lblMsg.Caption = lvwScore.ListItems(i).Text & " wins the game."
                End If
                Exit Sub
            End If
        Next i
        
        Dim Count_Lowest As Integer
        
        For i = 1 To lvwScore.ListItems.Count
            If Val(lvwScore.ListItems(i).SubItems(1)) = Lowest_Points Then
                Count_Lowest = Count_Lowest + 1
            End If
        Next i
        
        If Count_Lowest = 1 Then
            For i = 1 To lvwScore.ListItems.Count
                If Val(lvwScore.ListItems(i).SubItems(1)) = Lowest_Points Then
                    lvwScore.ListItems(i).SubItems(2) = Total_Score
                    'lblMsg.Caption = lvwScore.ListItems(i).Text
                    
                    lvwScore.ListItems(i).ForeColor = vbRed
                    For j = 1 To lvwScore.ListItems(i).ListSubItems.Count
                        lvwScore.ListItems(i).ListSubItems(j).ForeColor = vbRed
                    Next j
                    
                    If i = 1 Then ' player one wins the game
                        IsWinner = True
                        
                        lblMsg.Caption = "You win the game."
                    Else
                        lblMsg.Caption = lvwScore.ListItems(i).Text & " wins the game."
                    End If
                    
                    Exit Sub
                End If
            Next i
        End If
    End With
End Sub

