VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "It's raining letters !"
   ClientHeight    =   8895
   ClientLeft      =   2670
   ClientTop       =   915
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9885
   Begin VB.CommandButton cmdLetItRain 
      Caption         =   "Let it rain !"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   8280
      Width           =   1215
   End
   Begin VB.PictureBox OUT 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   7995
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#################################################
'
' A nice Text effect. Letters falling down
' hitting the ground and being catapulted up
' again.
'
' Author: Over (overkillpage@gmx.net)
'
'#################################################


Option Explicit

Private Sub cmdLetItRain_Click()
    Randomize Timer                         'Init Rnd

    'Declarations
    Dim StartTime(100)                      'Starttime of a up/down movement
    Dim DownMovement(100) As Boolean        'are we doing a up or down movement ???
    Dim MoveDistance As Double              'distance target has moved since the start of the movement
    Dim YPos(100) As Double                 'Holds the y position of a letter
    Dim MovementDone(100) As Boolean        'Is set to true when a up / down movement is completed
    Dim StartHeight(100) As Double               'From which hight will the letter fall down ?
    Dim UpMovementTime(100) As Double            'How long will it the letter take to move up
    Dim PowerLoss(100) As Double                 'losing xx% of power when touching the ground
    Dim Message As String                   'Message you want to display
    Dim Looop As Integer                    'Loop var
    Dim TextColor(100) As ColorConstants    'Color of one letter

    
    'Settings
    
    OUT.ScaleMode = 4
    OUT.FontName = "Courier New"
    
    Message = "Ohh my god ! It's raining letters today !!! Contact me: overkillpage@gmx.net"    'Message you want to display
    
    For Looop = 1 To Len(Message)
    
        PowerLoss(Looop) = 0.2 + ((Rnd * 25) / 100)                  'losing xx% of power when touching the ground
        StartHeight(Looop) = 0
        TextColor(Looop) = RGB(80 + Looop * 2, 80 + Looop * 2, 255)
    
    Next Looop
        
    For Looop = 1 To Len(Message)
        StartTime(Looop) = Timer                       'Setting up startime for a following movement, needed for calculation of position
    Next Looop
    
    Do
        
        OUT.Cls                             'Clear picturebox
        
        'Looping throung the textmessage
        For Looop = 1 To Len(Message)
        
        
            If DownMovement(Looop) = True Then
                
                MoveDistance = (StartHeight(Looop) + (0.5 * 9.81 * ((Timer - StartTime(Looop)) ^ 2))) 'Calculating falling distance
                
                If YPos(Looop) >= OUT.ScaleHeight - 1 Then MovementDone(Looop) = True     'The letter reached the bottom border. The Downmovement is complete
        
            Else
                MoveDistance = (StartHeight(Looop) + (0.5 * 9.81 * (UpMovementTime(Looop) - (Timer - StartTime(Looop))) ^ 2)) 'Calculating falling distance
                
                If YPos(Looop) <= StartHeight(Looop) + 0.1 Then MovementDone(Looop) = True      'The letter reached the max. height. The upmovement is complete
                
            End If
            
            YPos(Looop) = MoveDistance
            
            If YPos(Looop) > OUT.ScaleHeight - 1 Then                                   'If the letter fell out of our picturebox ;) we fix it
                YPos(Looop) = OUT.ScaleHeight - 1                                       'At the bottom position
            End If
            
            OUT.CurrentX = OUT.ScaleWidth / 2 - Int((Len(Message) / 2)) + Looop
            OUT.CurrentY = YPos(Looop)                                                  'Setting the letters y position
            OUT.ForeColor = TextColor(Looop)                                            'Setting the letters color
            OUT.Print Mid(Message, Looop, 1)                                            'Text output
        
        Next Looop
        
        DoEvents
    
        For Looop = 1 To Len(Message)
        
            If MovementDone(Looop) = True Then
                
                If DownMovement(Looop) = True Then     'Switch between up/downmovement
                    DownMovement(Looop) = False
                    StartHeight(Looop) = StartHeight(Looop) + ((OUT.ScaleHeight - StartHeight(Looop)) * PowerLoss(Looop))   'New Startheight, because of speed lost ?!?!
                    UpMovementTime(Looop) = Sqr((OUT.ScaleHeight - StartHeight(Looop)) / (0.5 * 9.81))        'How long will the NEXT upmovement last ???
                Else
                    DownMovement(Looop) = True
                End If
                
                StartTime(Looop) = Timer               'Set the StartTime of a new movement
                MovementDone(Looop) = False
            End If
            
         Next Looop
                
    Loop 'Until StartHeight = OUT.ScaleHeight
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub
