VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type Position
    X As Integer
    Y As Integer
End Type

Private Type AniInfo
    X1 As Integer
    Y1 As Integer
    X2 As Integer
    Y2 As Integer
End Type

Private Type FallInfo
    X     As Integer
    Y     As Integer
    Speed As Integer
End Type

Private Type SpinInfo
    Angle As Integer
End Type

Dim rcRect              As RECT
Dim rcPic               As RECT

Dim R_Edge              As Integer ' right edge
Dim B_Edge              As Integer ' bottom edge

Dim AniProp()           As AniInfo
Dim FallProp()          As FallInfo
Dim SpinProp()          As SpinInfo
Dim SavePos(1 To 20)    As Position
Dim LinearAni           As New clsLinearAni

Public Reset            As Boolean
Public StopAni          As Boolean

' Bounce
Dim varBounceDistX      As Integer
Dim varBounceDistY      As Integer
Dim varBounceSpeedX     As Integer
Dim varBounceSpeedY     As Integer

' Scatter
Dim varScatterSpeedX    As Integer
Dim varScatterSpeedY    As Integer

' Spin
Dim varSpinDistance     As Integer

' Fall
Dim varFallType         As Integer

Dim varMaxCards         As Integer
Dim varShowTrail        As Boolean

' // Bounce Prop ***
Public Property Get BounceDistX() As Integer
    BounceDistX = varBounceDistX
End Property

Public Property Let BounceDistX(ByVal dx As Integer)
    varBounceDistX = dx
End Property

Public Property Get BounceDistY() As Integer
    BounceDistY = varBounceDistY
End Property

Public Property Let BounceDistY(ByVal dy As Integer)
    varBounceDistY = dy
End Property

Public Property Get BounceSpeedX() As Integer
    BounceSpeedX = varBounceSpeedX
End Property

Public Property Let BounceSpeedX(ByVal SX As Integer)
    varBounceSpeedX = SX
End Property

Public Property Get BounceSpeedY() As Integer
    BounceSpeedY = varBounceSpeedY
End Property

Public Property Let BounceSpeedY(ByVal SY As Integer)
    varBounceSpeedY = SY
End Property
' // Bounce Prop ***

' // Scatter Prop ***
Public Property Get ScatterSpeedX() As Integer
    ScatterSpeedX = varScatterSpeedX
End Property

Public Property Let ScatterSpeedX(ByVal SX As Integer)
    varScatterSpeedX = SX
End Property

Public Property Get ScatterSpeedY() As Integer
    ScatterSpeedY = varScatterSpeedY
End Property

Public Property Let ScatterSpeedY(ByVal SY As Integer)
    varScatterSpeedY = SY
End Property
' // Scatter Prop ***

' // Spin Prop ***
Public Property Get SpinDistance() As Integer
    SpinDistance = varSpinDistance
End Property

Public Property Let SpinDistance(ByVal dx As Integer)
    varSpinDistance = dx
End Property
' // Spin Prop ***

Public Property Get FallType() As Integer
    FallType = varFallType
End Property

Public Property Let FallType(ByVal FallOp As Integer)
    varFallType = FallOp
End Property

Public Property Get MaxCards() As Integer
    MaxCards = varMaxCards
End Property

Public Property Let MaxCards(ByVal N As Integer)
    varMaxCards = N
End Property

Public Property Get ShowTrail() As Boolean
    ShowTrail = varShowTrail
End Property

Public Property Let ShowTrail(ByVal bVal As Boolean)
    varShowTrail = bVal
End Property

Public Sub Bounce(PicBox As PictureBox, Card As Object)
    Dim i    As Integer
    Dim dx   As Integer
    Dim dy   As Integer
    Dim NewX As Integer
    Dim NewY As Integer
    Dim Temp As Integer
    
    On Error GoTo ErrHandler
    
    GetClientRect PicBox.hwnd, rcRect
    R_Edge = rcRect.Right - rcRect.Left
    B_Edge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        ReDim AniProp(1 To MaxCards) As AniInfo
 
        PicBox.Cls
        Call Randomize
        
        For i = LBound(AniProp()) To UBound(AniProp())
            Card(i).Move SavePos(i).X, SavePos(i).Y
            Card(i).Visible = True
        Next i
        
        For i = LBound(AniProp()) To UBound(AniProp())
            AniProp(i).X1 = i * BounceDistX
            AniProp(i).Y1 = i * BounceDistY
            AniProp(i).X2 = BounceSpeedX
            AniProp(i).Y2 = BounceSpeedY
            
            Do While Not LinearAni.Linear(Card(i), SavePos(i).X, SavePos(i).Y, _
                                                   AniProp(i).X1, AniProp(i).Y1)
            Loop
        Next i
        
        Reset = False
    Else
        If Not ShowTrail Then PicBox.Cls
        For i = LBound(AniProp()) To UBound(AniProp())
            If StopAni Then Exit For
            
            GetClientRect Card(i).hwnd, rcPic
            
            Temp = AniProp(i).X1 + AniProp(i).X2
            If Temp + rcPic.Right > R_Edge + AniProp(i).X2 Then
                AniProp(i).X2 = -Abs(AniProp(i).X2)
            ElseIf Temp < 0 Then
                AniProp(i).X2 = Abs(AniProp(i).X2)
            End If
               
            NewX = AniProp(i).X1 + AniProp(i).X2
             
            Temp = AniProp(i).Y1 + AniProp(i).Y2
            If Temp + rcPic.Bottom > B_Edge Then
                AniProp(i).Y2 = -Abs(AniProp(i).Y2)
            ElseIf Temp < 0 Then
                AniProp(i).Y2 = Abs(AniProp(i).Y2)
            End If
             
            NewY = AniProp(i).Y1 + AniProp(i).Y2
              
            AniProp(i).X1 = NewX
            AniProp(i).Y1 = NewY
                
            BitBlt PicBox.hdc, AniProp(i).X1, AniProp(i).Y1, _
                               rcPic.Right, rcPic.Bottom, _
                   Card(i).hdc, 0, 0, vbSrcCopy
            RefreshWindow PicBox.hwnd
            
            SavePos(i).X = AniProp(i).X1
            SavePos(i).Y = AniProp(i).Y1
        Next i
        
        For i = LBound(AniProp()) To UBound(AniProp())
            If Card(i).Visible Then
                Card(i).Visible = False
            End If
        Next i
    End If
    Exit Sub
    
ErrHandler:
End Sub

Public Sub Scatter(PicBox As PictureBox, Card As Object)
    Dim i    As Integer
    Dim dx   As Integer
    Dim dy   As Integer
    Dim NewX As Integer
    Dim NewY As Integer
    Dim Temp As Integer
        
    On Error GoTo ErrHandler
    
    GetClientRect PicBox.hwnd, rcRect
    R_Edge = rcRect.Right - rcRect.Left
    B_Edge = rcRect.Bottom - rcRect.Top
        
    If Reset Then
        ReDim AniProp(1 To MaxCards) As AniInfo
            
        PicBox.Cls
        Call Randomize
            
        For i = LBound(AniProp()) To UBound(AniProp())
            Card(i).Move SavePos(i).X, SavePos(i).Y
            Card(i).Visible = True
        Next i
        
        For i = LBound(AniProp()) To UBound(AniProp())
            AniProp(i).X1 = CInt(Abs(R_Edge - Card(i).Width) * Rnd)
            AniProp(i).Y1 = CInt(Abs(B_Edge - Card(i).Height) * Rnd)
            AniProp(i).X2 = ScatterSpeedX
            AniProp(i).Y2 = ScatterSpeedY
            
            Do While Not LinearAni.Linear(Card(i), SavePos(i).X, SavePos(i).Y, _
                                                   AniProp(i).X1, AniProp(i).Y1)
            Loop
        Next i
            
        Reset = False
    Else
        If Not ShowTrail Then PicBox.Cls
        For i = LBound(AniProp()) To UBound(AniProp())
            If StopAni Then Exit For
            
            GetClientRect Card(i).hwnd, rcPic
            
            Temp = AniProp(i).X1 + AniProp(i).X2
            If Temp + rcPic.Right > R_Edge Then
                AniProp(i).X2 = -Abs(AniProp(i).X2)
            ElseIf Temp < 0 Then
                AniProp(i).X2 = Abs(AniProp(i).X2)
            End If
               
            NewX = AniProp(i).X1 + AniProp(i).X2
             
            Temp = AniProp(i).Y1 + AniProp(i).Y2
            If Temp + rcPic.Bottom > B_Edge Then
                AniProp(i).Y2 = -Abs(AniProp(i).Y2)
            ElseIf Temp < 0 Then
                AniProp(i).Y2 = Abs(AniProp(i).Y2)
            End If
             
            NewY = AniProp(i).Y1 + AniProp(i).Y2
              
            AniProp(i).X1 = NewX
            AniProp(i).Y1 = NewY
                
            BitBlt PicBox.hdc, AniProp(i).X1, AniProp(i).Y1, _
                               rcPic.Right, rcPic.Bottom, _
                   Card(i).hdc, 0, 0, vbSrcCopy
            RefreshWindow PicBox.hwnd
            
            SavePos(i).X = AniProp(i).X1
            SavePos(i).Y = AniProp(i).Y1
        Next i
        
        For i = LBound(AniProp()) To UBound(AniProp())
            If Card(i).Visible Then
                Card(i).Visible = False
            End If
        Next i
    End If
    Exit Sub
    
ErrHandler:
End Sub

Public Sub Spin(PicBox As PictureBox, Card As Object)
    Dim i       As Integer
    Dim cx      As Integer
    Dim cy      As Integer
    Dim RadiusX As Integer
    Dim RadiusY As Integer
    
    On Error GoTo ErrHandler
    
    GetClientRect PicBox.hwnd, rcRect
    R_Edge = rcRect.Right - rcRect.Left
    B_Edge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        ReDim SpinProp(1 To MaxCards) As SpinInfo
        
        PicBox.Cls
        For i = LBound(SpinProp()) To UBound(SpinProp())
            Card(i).Move SavePos(i).X, SavePos(i).Y
            Card(i).Visible = True
        Next i
        
        For i = LBound(SpinProp()) To UBound(SpinProp())
            SpinProp(i).Angle = i * SpinDistance
            
            GetClientRect Card(i).hwnd, rcPic
            
            RadiusX = (R_Edge - rcPic.Right) / 2
            RadiusY = (B_Edge - rcPic.Bottom) / 2
                    
            cx = RadiusX + Cos(Rads(SpinProp(i).Angle)) * RadiusX
            cy = RadiusY + Sin(Rads(SpinProp(i).Angle)) * RadiusY
            
            Do While Not LinearAni.Linear(Card(i), SavePos(i).X, SavePos(i).Y, _
                                                   cx, cy)
            Loop
        Next i
        
        Reset = False
    Else
        If Not ShowTrail Then PicBox.Cls
        For i = LBound(SpinProp()) To UBound(SpinProp())
            If StopAni Then Exit For
            
            GetClientRect Card(i).hwnd, rcPic
            
            RadiusX = (R_Edge - rcPic.Right) / 2
            RadiusY = (B_Edge - rcPic.Bottom) / 2
                    
            cx = RadiusX + Cos(Rads(SpinProp(i).Angle)) * RadiusX
            cy = RadiusY + Sin(Rads(SpinProp(i).Angle)) * RadiusY
                    
            BitBlt PicBox.hdc, cx, cy, _
                               rcPic.Right, rcPic.Bottom, _
                   Card(i).hdc, 0, 0, vbSrcCopy
            RefreshWindow PicBox.hwnd
            
            SpinProp(i).Angle = SpinProp(i).Angle + 2
            SpinProp(i).Angle = SpinProp(i).Angle Mod 360
            
            SavePos(i).X = cx
            SavePos(i).Y = cy
        Next i
        
        For i = LBound(SpinProp()) To UBound(SpinProp())
            If Card(i).Visible Then
                Card(i).Visible = False
            End If
        Next i
    End If
    Exit Sub
    
ErrHandler:
End Sub

Public Sub Fall(PicBox As PictureBox, Card As Object)
    Dim i As Integer
    
    On Error GoTo ErrHandler
    
    GetClientRect PicBox.hwnd, rcRect
    R_Edge = rcRect.Right - rcRect.Left
    B_Edge = rcRect.Bottom - rcRect.Top
    
    If Reset Then
        ReDim FallProp(1 To MaxCards) As FallInfo
        
        PicBox.Cls
        
        For i = LBound(FallProp()) To UBound(FallProp())
            Card(i).Move SavePos(i).X, SavePos(i).Y
            Card(i).Visible = True
        Next i
        
        For i = LBound(FallProp()) To UBound(FallProp())
            GetClientRect Card(i).hwnd, rcPic
                        
            Select Case FallType
            Case Is = 0 ' Down
                FallProp(i).X = Random_Number(0, R_Edge - rcPic.Right)
                FallProp(i).Y = 0
                FallProp(i).Speed = CInt(Rnd * 3) + 1
            Case Is = 1 ' Up
                FallProp(i).X = Random_Number(0, R_Edge - rcPic.Right)
                FallProp(i).Y = B_Edge - rcPic.Bottom
                FallProp(i).Speed = CInt(Rnd * 3) + 1
            Case Is = 2 ' Left
                FallProp(i).X = 0
                FallProp(i).Y = Random_Number(0, B_Edge - rcPic.Bottom)
                FallProp(i).Speed = CInt(Rnd * 3) + 1
            Case Is = 3 ' Right
                FallProp(i).X = R_Edge - rcPic.Right
                FallProp(i).Y = Random_Number(0, B_Edge - rcPic.Bottom)
                FallProp(i).Speed = CInt(Rnd * 3) + 1
            End Select
            
            Do While Not LinearAni.Linear(Card(i), SavePos(i).X, SavePos(i).Y, _
                                                   FallProp(i).X, FallProp(i).Y)
            Loop
        Next i

        Reset = False
    Else
        If Not ShowTrail Then PicBox.Cls

        For i = LBound(FallProp()) To UBound(FallProp())
            GetClientRect Card(i).hwnd, rcPic
            
            Select Case FallType
            Case Is = 0 ' Down
                If FallProp(i).Y > (B_Edge - FallProp(i).Speed) Then
                    FallProp(i).X = Random_Number(0, R_Edge - rcPic.Right)
                    FallProp(i).Y = -rcPic.Bottom
                    FallProp(i).Speed = CInt(Rnd * 3) + 2
                Else
                    FallProp(i).Y = FallProp(i).Y + _
                                    FallProp(i).Speed
                End If
            Case Is = 1 ' Up
                If FallProp(i).Y < -rcPic.Bottom Then
                    FallProp(i).X = Random_Number(0, R_Edge - rcPic.Right)
                    FallProp(i).Y = B_Edge
                    FallProp(i).Speed = CInt(Rnd * 3) + 2
                Else
                    FallProp(i).Y = FallProp(i).Y - _
                                    FallProp(i).Speed
                End If
            Case Is = 2 ' Left
                If FallProp(i).X > (R_Edge - FallProp(i).Speed) Then
                    FallProp(i).X = -rcPic.Right
                    FallProp(i).Y = Random_Number(0, B_Edge - rcPic.Bottom)
                    FallProp(i).Speed = CInt(Rnd * 3) + 2
                Else
                    FallProp(i).X = FallProp(i).X + _
                                    FallProp(i).Speed
                End If
            Case Is = 3 ' Right
                If FallProp(i).X < -rcPic.Right Then
                    FallProp(i).X = R_Edge
                    FallProp(i).Y = Random_Number(0, B_Edge - rcPic.Bottom)
                    FallProp(i).Speed = CInt(Rnd * 3) + 2
                Else
                    FallProp(i).X = FallProp(i).X - _
                                    FallProp(i).Speed
                End If
            End Select
            
            BitBlt PicBox.hdc, FallProp(i).X, FallProp(i).Y, _
                               rcPic.Right, rcPic.Bottom, _
                   Card(i).hdc, 0, 0, vbSrcCopy
            RefreshWindow PicBox.hwnd
        
            SavePos(i).X = FallProp(i).X
            SavePos(i).Y = FallProp(i).Y
        Next i
        
        For i = LBound(FallProp()) To UBound(FallProp())
            If Card(i).Visible Then
                Card(i).Visible = False
            End If
        Next i
    End If
    Exit Sub
    
ErrHandler:
End Sub

Public Sub SaveCardPosition(Card As Object)
    Dim i As Integer
    On Error GoTo ErrHandler
    
    For i = 1 To Card.Count
        SavePos(i).X = Card(i).Left
        SavePos(i).Y = Card(i).Top
    Next i
    Exit Sub
    
ErrHandler:
End Sub

Private Sub Class_Initialize()
    ' Bounce
    varBounceDistX = 5
    varBounceDistY = 5
    varBounceSpeedX = 2
    varBounceSpeedY = 2
    
    ' Scatter
    varScatterSpeedX = 2
    varScatterSpeedY = 2
    
    ' Spin
    varSpinDistance = 5
    
    varFallType = 2
    varShowTrail = False
    
    Reset = True
End Sub

Private Sub Class_Terminate()
    StopAni = True
End Sub
