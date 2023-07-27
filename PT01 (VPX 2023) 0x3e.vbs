Option Explicit

'*****************************************************************************************************
' CREDITS
' Table created by 0x3e
' Initial table created by fuzzel, jimmyfingers, jpsalas, toxie & unclewilly (in alphabetical order)
' Flipper primitives by zany
' Ball rolling sound script by jpsalas
' Ball shadow by ninuzzu
' Ball control & ball dropping sound by rothbauerw
' DOF by arngrim
' Positional sound helper functions by djrobx
' Plus a lot of input from the whole community (sorry if we forgot you :/)
'*****************************************************************************************************


'First, try to load the Controller.vbs (DOF), which helps controlling additional hardware like lights, gears, knockers, bells and chimes (to increase realism)
'This table uses DOF via the 'SoundFX' calls that are inserted in some of the PlaySound commands, which will then fire an additional event, instead of just playing a sample/sound effect
On Error Resume Next
ExecuteGlobal GetTextFile("controller.vbs")
If Err Then MsgBox "You need the Controller.vbs file in order to run this table (installed with the VPX package in the scripts folder)"
On Error Goto 0

'If using Visual PinMAME (VPM), place the ROM/game name in the constant below,
'both for VPM, and DOF to load the right DOF config from the Configtool, whether it's a VPM or an Original table

'Const cGameName = ""


Dim EnableRetractPlunger
EnableRetractPlunger = false 'Change to true to enable retracting the plunger at a linear speed; wait for 1 second at the maximum position; move back towards the resting position; nice for button/key plungers

Const UseB2S = True
Const TableName = "PT01 (VPX 2023) 0x3e alpha01"

Const GameBalls = 3
Const EasyMode = false


Dim Credits: Credits = 2

If Table1.ShowDT = false OR UseB2S = True then
  ScoreText1.Visible = false
  ScoreText2.Visible = false
  ScoreText3.Visible = false
  ScoreText4.Visible = false
  ScoreTextBG.Visible = false
  PlayerBallText.Visible = false
End If
If UseB2S Then
  Dim Controller
  Set Controller = CreateObject("B2S.Server")
  Controller.B2SName = TableName
  Controller.Run
End If




Const StartLevel = 0
Const MaxBalls = 5
Dim Score(4)
Dim HighScore(4)
Dim MaxPlayers: MaxPlayers = 4
Dim PlayersPlayingGame: PlayersPlayingGame = 0
Dim CurrentPlayer: CurrentPlayer = 0
Dim CurrentBonus: CurrentBonus = 0
Dim DisplayBonus: DisplayBonus = 0
Dim BonusX: BonusX = 0
Dim BonusXTargets(3)
Dim BonusXthisLevel:BonusXthisLevel = True
Dim LastBonus: LastBonus = 0
Dim CurrentBall: CurrentBall = 0
Dim OutTheGate: OutTheGate = 0
Dim BeforeTheFirstBall: BeforeTheFirstBall = 1
Dim TheLevel: TheLevel = Array(1, 2, 2.4, 2.8, 3, 3.1, 3.15, 3.19, 3.2)
Dim TableLevel: TableLevel = StartLevel
Dim Multiballed
Dim OnMission
Dim MissionDone(8)
Dim MissionHits(3)
Dim PlayerLevel(4)
Dim LockedBalls: LockedBalls = 0
Dim LockedFlippers: LockedFlippers = false
Dim Tilt
Dim TiltSensitivity
Dim Tilted
Const MAX_Tilt = 17


Sub AddPlayer()
  debug.print "AddPlayer"
  PlayersPlayingGame = PlayersPlayingGame + 1
  PlayerLevel(PlayersPlayingGame) = StartLevel
  CreditsPlayerBall()
  UpdateScore(PlayersPlayingGame)
End Sub

Sub ResetForNewGame()
  Dim i
  debug.print "ResetForNewGame"
  BeforeTheFirstBall = 1
  CurrentPlayer = 0
  For i = 0 to MaxPlayers
    Score(i) = 0
  Next
  CurrentBall = 1
  LockedFlippers = false
  ResetForNewPlayerBall()
  SetLevel(StartLevel)
  SetBonusX()
  Tilt = 0
  TiltSensitivity = 6
  Tilted = False
End Sub

Sub SetForEndGame()
  debug.print "setForEndGame"
  AddScore(0)
  PlayersPlayingGame = 0
  CurrentPlayer = 0
  CurrentBall = 0
  OutTheGate = 0
  LockedFlippers = true
  LeftFlipper.RotateToStart
  RightFlipper.RotateToStart
  BeforeTheFirstBall = 1
End Sub

Sub ResetHitables()
  Dim h
  For each h in HitAbles
    If InStr(1, h.name, "Target", 1) = 1 Then
      If h.DropSpeed <> 0.4 Then h.IsDropped = 0
    End If
  Next
  If EasyMode Then
    LeftShield.IsDropped = 0
    RightShield.IsDropped = 0
    Guardian.IsDropped = 0
  Else
    LeftShield.IsDropped = 1
    RightShield.IsDropped = 1
    Guardian.IsDropped = 1
  End If
End Sub

Sub ResetDoneMissions()
  Dim i, m: i = 0
  for each m in MissionDone
   MissionDone(i) = false
   i = i + 1
  Next
End Sub

Sub SetLevel(level)
  Dim h, l
  If level > 6 Then level = 6
  For each h in HitAbles
    h.UserValue = level
  Next
  For each l in HitLights
     SetLightLevel l, level
  Next
  If CurrentBonus < level * 55 Then CurrentBonus = level * 55 : DisplayBonus = CurrentBonus
  If TableLevel > 5 Then TableLevel = 5
End Sub

Sub ResetForNewPlayerBall()
  debug.print "ResetForNewPlayerBall level:" & PlayerLevel(CurrentPlayer)
  AddScore(0)
  CurrentBonus = 0
  CurrentPlayer = CurrentPlayer + 1
  BonusX = 0
  BonusXthisLevel = True
  ResetBonusXTargets()
  If CurrentPlayer > PlayersPlayingGame Then 
    CurrentPlayer = 1
    CurrentBall = CurrentBall + 1
  End If
  TableLevel = StartLevel
  ResetHitables()
  SetLevel(TableLevel)
  BonusTimer.Interval = 75
  SetBonusLights()
  UpdateScores()
  CreditsPlayerBall()
  OutTheGate = 0
  MultiBalled = false
  OnMission = 0
  ResetDoneMissions()
  LockedFlippers = false
  Tilt = 0
  Tilted = False
  For each xx in GI:xx.State = 1: Next
  CreateABall()
End Sub

Sub UpdateScores()
  Dim i, basetext: basetext = "---"

  For i = 1 to 4
    If i <= PlayersPlayingGame Then 
      UpdateScore(i)
    Else 
      SetPlayerDisplay i, mid(basetext,5-i) & i & mid(basetext,i)
    End If
  Next
End Sub

Sub B2SSet(id, in_string)
  if UseB2S = false Then Exit Sub
  if id <= 4 AND TypeName(in_string) <> "String" AND IsNumeric(in_string) Then Controller.B2SSetScore id, in_string: Exit Sub

  Dim char, tmp_string
  Dim i, max_len: max_len = 4
  If id = 5 Then max_len = 6
  in_string = CStr(in_string)
  tmp_string = in_string
  For i = 1 to max_len
    If i - len(in_string) <= 0 Then
      char = left(tmp_string,1)
      tmp_string = right(tmp_string,len(tmp_string)-1)
    Else
      char = " "
    End If
    Controller.B2SSetLED i + (id - 1) * 4,B2SLEDValue(char)
  next
End Sub

Sub EndCurrentPlayerTurn()
'  debug.print "EndCurrentPlayerTurn"
  If CurrentBonus > 0 Then AddScore(1) : CurrentBonus = CurrentBonus - 1

  If CurrentBonus < 21 AND LastBonus < 21 AND LastBonus > 0 Then
    BonusTimer.Interval = Cint(2000/LastBonus)
  ElseIf CurrentBonus > 201  Then
    BonusTimer.Interval = 10
  ElseIf CurrentBonus > 101 Then
    BonusTimer.Interval = 30
  ElseIf CurrentBonus > 40 Then
    BonusTimer.Interval = 50
  ElseIf CurrentBonus > 20 Then
    BonusTimer.Interval = 75
  ElseIf CurrentBonus < 20 Then
    BonusTimer.Interval = 100
  End If

  If BonusX > 0 AND BonusTimer.Interval > 30 Then BonusTimer.Interval = 30
 
  If CurrentBonus <> DisplayBonus Then BonusTimer.enabled = true

 If Tilted AND BonusTimer.Interval < 7000 Then
    BonusTimer.enabled = false
    BonusTimer.Interval = 7000
    Tilted = False
    BonusTimer.enabled = true
  ElseIf CurrentBonus < 1 AND DisplayBonus < 1 AND BonusX < 1 Then
    DisplayBonus = LastBonus
    LastBonus = 0
    SetBonusLights()
    DisplayBonus = LastBonus
    If(CurrentPlayer = PlayersPlayingGame)AND(CurrentBall >= GameBalls) Then
      SetForEndGame()
    Else
      ResetForNewPlayerBall()
    End If
  ElseIf CurrentBonus < 1 AND DisplayBonus < 1 Then
    CurrentBonus = LastBonus
    DisplayBonus = LastBonus
    BonusX = BonusX - 1
    SetBonusX()
    SetBonusLights()
    BonusTimer.enabled = true
  End If
End Sub

Sub UpdateDisplayBonus()
'  debug.print "CurrentBonus: " & CurrentBonus & " DisplayBonus: " & DisplayBonus
  If CurrentBonus <> DisplayBonus Then BonusTimer.enabled = true
End Sub

Sub BonusTimer_Timer()
  If CurrentBonus < DisplayBonus Then DisplayBonus = DisplayBonus - 1
  If CurrentBonus > DisplayBonus Then DisplayBonus = DisplayBonus + 1 : LastBonus = CurrentBonus
  SetBonusLights()
  BonusTimer.enabled = false

  If BIP = 0 Then 
    EndCurrentPlayerTurn()
  Else
    UpdateDisplayBonus()
  End If
End SUb

Sub CreateABall()
  debug.print "CreateABall"
		'Plunger.CreateBall
		BallRelease.CreateBall
		BallRelease.Kick 90, 7
		PlaySound SoundFX("ballrelease",DOFContactors), 0,1,AudioPan(BallRelease),0.25,0,0,1,AudioFade(BallRelease)
		BIP = BIP + 1
End Sub

Sub CreditsPlayerBall()
  Dim out_text:  out_text = "c" & Credits & "P" & CurrentPlayer & "b" & CurrentBall
  PlayerBallText.Text = out_text
  B2SSet 5, out_text
End Sub
CreditsPlayerBall()

Sub SetBonusLights()
  Dim l,t, i, t_height, mod_bonus : i = 0 : t_height = 0 : mod_bonus = false
  Dim triangle_bonus: triangle_bonus = Array(10, 19, 27, 34, 40, 45, 49, 52, 54, 55)
  Dim lights_clip : lights_clip = (DisplayBonus Mod 55 )
  For each t in triangle_bonus
    if lights_clip >= t Then t_height = i
    i = i + 1
  Next
  i = 0
  For each l in BonusLights
    i = i + 1
    mod_bonus = CBool((lights_clip MOD triangle_bonus(t_height)) = i)
'    debug.print "i: " & i & " light: " & l.name  & " bonus: " & lights_clip & " mod: " & mod_bonus & " t_height: " & t_height
    Select Case mod_bonus
      Case false : SetLightLevel l, Fix( DisplayBonus / 55 ) - 1
      Case true : SetLightLevel l, Fix( DisplayBonus / 55 )
    End Select
    if ((lights_clip > 9)AND(10 - t_height <= i)) Then SetLightLevel l, Fix( DisplayBonus / 55 )
  Next
End Sub

Sub ResetBonusXTargets()
  Dim i
  For i = 1 to 3
    BonusXTargets(i) = False
  Next
End Sub

Sub SetBonusX()
  Dim l, i, sets_hit: sets_hit = 0
  For i = 1 to 3
    If BonusXTargets(i) = True Then sets_hit = sets_hit + 1
  Next
  If sets_hit = 3 AND BonusX < TableLevel AND BonusXthisLevel = False Then
   BonusX = BonusX + 1
   BonusXthisLevel = True
   ResetBonusXTargets()
  End If
  i = 0
  For each l in BonusXLights
    If BonusX > i Then 
      SetLightLevel l, TableLevel
    Else
     SetLightLevel l, -2
    End If
    i = i + 1
  Next
End Sub

Sub AddScore(points)
  if Tilted Then Exit Sub
  if OutTheGate = 0 Then Exit Sub
  Score(CurrentPlayer) =  Score(CurrentPlayer) + points
  UpdateScore(CurrentPlayer)
End Sub

Sub AddBonus(points)
  If OutTheGate = 0 Then Exit Sub
  If Tilted Then Exit Sub
  If CurrentBonus + points > (TableLevel + 2) * 55 Then
    CurrentBonus = (TableLevel + 2) * 55
  Else
    CurrentBonus =  CurrentBonus + points
  End If
  UpdateDisplayBonus()
End Sub

Sub UpdateScore(inPlayer)
    Select Case inPlayer
      Case 1: ScoreText1.Text = Round(Score(inPlayer)): B2SSet 1, Round(Score(inPlayer))
      Case 2: ScoreText2.Text = Round(Score(inPlayer)): B2SSet 2, Round(Score(inPlayer))
      Case 3: ScoreText3.Text = Round(Score(inPlayer)): B2SSet 3, Round(Score(inPlayer))
      Case 4: ScoreText4.Text = Round(Score(inPlayer)): B2SSet 4, Round(Score(inPlayer))
    End Select
End Sub

Sub SetPlayerDisplay(inPlayer, in_text)
    Select Case inPlayer
      Case 1: ScoreText1.Text = in_text: B2SSet 1, in_text
      Case 2: ScoreText2.Text = in_text: B2SSet 2, in_text
      Case 3: ScoreText3.Text = in_text: B2SSet 3, in_text
      Case 4: ScoreText4.Text = in_text: B2SSet 4, in_text
    End Select
End Sub

Sub Bumper001TryUpgrade()
  Dim b: Set b = Bumper001
  if b.UserValue < TableLevel + 1 Then
    b.UserValue = TableLevel + 1
    SetLightLevel Bu1li1, b.UserValue
  End If
End Sub

Sub Bumper001_Hit()
  PlaySound "fx_bumper" & RndNbr(4), 0, .25, AudioPan(Bumper001), 0.25, 0, 0, 1, AudioFade(Bumper001)
  AddScore(1*TheLevel(Bumper001.UserValue))
  MissionRun("bumper")
End Sub

Sub Bumper002TryUpgrade()
  Dim b: Set b = Bumper002
  if b.UserValue < TableLevel + 1 Then
    b.UserValue = TableLevel + 1 
    SetLightLevel Bu2li1, b.UserValue
  End If
End Sub

Sub Bumper002_Hit()
  PlaySound "fx_bumper"& RndNbr(4), 0, .25, AudioPan(Bumper002), 0.25, 0, 0, 1, AudioFade(Bumper002)
  AddScore(1*TheLevel(Bumper002.UserValue))
  MissionRun("bumper")
End Sub

Sub Gate001_Hit()
  debug.print "gate001 hit"
  Gate001.TimerEnabled = 1
  OutTheGate = 1
  SetBonusLights()
  BeforeTheFirstBall = 0
  Trigger012.Enabled = 0
  Trigger013.Enabled = 0
  Trigger014.Enabled = 0
End Sub

Sub Gate001_Timer
  debug.print "gate001 timer"
  Gate001.TimerEnabled = 0
  Trigger012.Enabled = 1
  Trigger013.Enabled = 1
  Trigger014.Enabled = 1
End Sub

Sub Gate002_Hit()
  debug.print "gate002 hit"
  AddBonus(1)
  If Trigger012.Enabled = false Then
    Gate002.TimerEnabled = 1
    Bumper002.HasHitEvent = 0
    Bu2li1.State = 0
    Gate004.Collidable = 0
    Gate005.Collidable = 0
    Trigger009.Enabled = 0
    Trigger010.Enabled = 0
    Trigger011.Enabled = 0
  End If

End Sub

Sub Gate002_Timer
  debug.print "gate002 timer"
  Gate002.TimerEnabled = 0
  Gate004.Collidable = 1
  Gate005.Collidable = 1
  Bumper002.HasHitEvent = 1
  Bu2li1.State = 1
  Trigger009.Enabled = 1
  Trigger010.Enabled = 1
  Trigger011.Enabled = 1
End Sub

Sub Gate003_Hit()
  Dim t, i, mission : i = 0 : mission = 0
  AddBonus(1)
  If Trigger012.Enabled = true Then
    Gate003.TimerEnabled = 1
    debug.print "gate003 hit start timer"
    Gate004.Collidable = False
    Gate005.Collidable = False
  End If
  If Gate003.UserValue > TableLevel AND OnMission = 0 Then
    Gate003.UserValue = TableLevel

    mission = Get_mission()

    If MissionDone(mission) Then
      MissionDisplay "mm0" & mission, "cplt"
      MissionTimer.enabled = true
      SetLightLevel Tr1li001, TableLevel
    Else
      OnMission = mission
      MissionRun("start")
      debug.print "mission: " & OnMission
    End If
  End If

  SetLightLevel Tr1li001, Gate003.UserValue
End Sub

Function Get_mission()
  Dim i, t, mission: mission = 0
  For each t in Targets4
    If t.UserValue > TableLevel Then mission = mission + 2 ^ i
    'debug.print "T: " & t.name & " i: " & i & " mission: " & mission
    i = i + 1
  Next
  If mission = 0 Then mission = 8
  Get_mission = mission
End Function

Sub Gate003_Timer()
  debug.print "gate003 timer"
  Gate003.TimerEnabled = 0
  Gate005.TimerEnabled = 1
  Gate004.Collidable = True
End Sub

Sub MissionTimer_timer()
  If OnMission = 0 Then UpdateScores()
  MissionTimer.enabled = false
End Sub

Sub MissionScore(HitThing)
  Dim nl_ball_bonus : nl_ball_bonus = 2*(BIP - LockedBalls)
  AddScore((nl_ball_bonus + LockedBalls)*TheLevel(TableLevel))
End Sub

Sub MissionRun(DoThing)
 If OnMission < 1 Then Exit Sub
 MissionTimer.enabled = true
 Dim spinner : spinner = "A"
 Dim bumper : bumper = "W"
 Dim target : target = "X"
 Dim trigger : trigger = "n"
 Select Case OnMission
   Case 1
      If DoThing = "start" Then
        MissionHits(1) = 3
      ElseIf DoThing = "spinner" Then
        MissionHits(1) = MissionHits(1) - 1
        MissionScore(DoThing)
      End If
      If MissionHits(1) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(5*TheLevel(TableLevel))
        AddBonus(20)
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, spinner & MissionHits(1)
      End If
   Case 2
      If DoThing = "start" Then
        MissionHits(1) = 4
      ElseIf DoThing = "bumper" Then
        MissionHits(1) = MissionHits(1) - 1
        MissionScore(DoThing)
      End If
      If MissionHits(1) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(10*TheLevel(TableLevel))
        AddBonus(15)
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, bumper & MissionHits(1)
      End If
   Case 3
      If DoThing = "start" Then
        MissionHits(1) = 4
      ElseIf DoThing = "spinner" Then
        MissionHits(1) = MissionHits(1) - 1
        MissionScore(DoThing)
      End If
      If MissionHits(1) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(15*TheLevel(TableLevel))
        AddBonus(15)
        If LeftShield.IsDropped = 0 AND RightShield.IsDropped = 0 Then Guardian.IsDropped = 0
        If RightShield.IsDropped = 0 Then LeftShield.IsDropped = 0
        RightShield.IsDropped = 0
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, spinner & MissionHits(1)
      End If
   Case 4
      If DoThing = "start" Then
        MissionHits(1) = 8
        MissionHits(2) = 2
      ElseIf DoThing = "bumper" Then
        If MissionHits(1) > 0 Then
          MissionHits(1) = MissionHits(1) - 1
          MissionScore(DoThing)
        End If
      ElseIf DoThing = "spinner" Then
        If MissionHits(2) > 0 Then
          MissionHits(2) = MissionHits(2) - 1
          MissionScore(DoThing)
        End If
      End If
      If MissionHits(1) < 1 AND MissionHits(2) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(10*TheLevel(TableLevel))
        AddBonus(15)
        If LeftShield.IsDropped = 0 AND RightShield.IsDropped = 0 Then Guardian.IsDropped = 0
        RightShield.IsDropped = 0
        LeftShield.IsDropped = 0
        MissionDone(OnMission) = true
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, bumper & MissionHits(1) & spinner & MissionHits(2)
      End If
   Case 5
      If DoThing = "start" Then
        MissionHits(1) = 4
      ElseIf DoThing = "trigger" Then
        MissionHits(1) = MissionHits(1) - 1
        MissionScore(DoThing)
      End If
      If MissionHits(1) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(10*TheLevel(TableLevel))
        AddBonus(15)
        If LeftShield.IsDropped = 0 AND RightShield.IsDropped = 0 Then Guardian.IsDropped = 0
        If LeftShield.IsDropped = 0 Then RightShield.IsDropped = 0
        LeftShield.IsDropped = 0
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, trigger & MissionHits(1)
      End If
   Case 6
      If DoThing = "start" Then
        MissionHits(1) = 8
        MissionHits(2) = 2
      ElseIf DoThing = "target" Then
        If MissionHits(1) > 0 Then
          MissionHits(1) = MissionHits(1) - 1
          MissionScore(DoThing)
        End If
      ElseIf DoThing = "spinner" Then
        If MissionHits(2) > 0 Then
          MissionHits(2) = MissionHits(2) - 1
          MissionScore(DoThing)
        End If
      End If
      If MissionHits(1) < 1 AND MissionHits(2) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(20*TheLevel(TableLevel))
        AddBonus(20)
        If LeftShield.IsDropped = 0 AND RightShield.IsDropped = 0 Then Guardian.IsDropped = 0
        If RightShield.IsDropped = 0 Then LeftShield.IsDropped = 0
        RightShield.IsDropped = 0
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, target & MissionHits(1) & spinner & MissionHits(2)
      End If
   Case 7
      If DoThing = "start" Then
        MissionHits(1) = 3
        MissionHits(2) = 5
      ElseIf DoThing = "trigger" Then
        If MissionHits(1) > 0 Then 
          MissionHits(1) = MissionHits(1) - 1
          MissionScore(DoThing)
        End If
      ElseIf DoThing = "target" Then
        If MissionHits(2) > 0 Then 
          MissionHits(2) = MissionHits(2) - 1
          MissionScore(DoThing)
        End If
      End If
      If MissionHits(1) < 1 AND MissionHits(2) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(20*TheLevel(TableLevel))
        AddBonus(20)
        RightShield.IsDropped = 0
        LeftShield.IsDropped = 0
        Guardian.IsDropped = 0
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, trigger & MissionHits(1) & target & MissionHits(2)
      End If
   Case 8
      If DoThing = "start" Then
        MissionHits(1) = 4
        MissionHits(2) = 2
      ElseIf DoThing = "bumper" Then
        If MissionHits(1) > 0 Then
          MissionHits(1) = MissionHits(1) - 1
          MissionScore(DoThing)
        End If
      ElseIf DoThing = "spinner" Then
        If MissionHits(2) > 0 Then
          MissionHits(2) = MissionHits(2) - 1
          MissionScore(DoThing)
        End If
      End If
      If MissionHits(1) < 1 AND MissionHits(2) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(10*TheLevel(TableLevel))
        AddBonus(10)
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, bumper & MissionHits(1) & spinner & MissionHits(2)
      End If
   Case Else
      If DoThing = "start" Then
        MissionHits(1) = 8
      ElseIf DoThing = "bumper" Then
        MissionHits(1) = MissionHits(1) - 1
        MissionScore(DoThing)
      End If
      If MissionHits(1) < 1 Then
        SetLightLevel Tr1li001, TableLevel
        AddScore(5*TheLevel(TableLevel))
        AddBonus(10)
        MissionDone(OnMission) = true
        MissionDisplay "mm0" & OnMission, "done"
        OnMission = 0
      Else
        MissionDisplay "mm0" & OnMission, bumper & MissionHits(1)
      End If
  End Select
End Sub

Sub MissionDisplay(one, two)
  Select Case CurrentPlayer < 3
    Case true
      SetPlayerDisplay 3, one
      SetPlayerDisplay 4, two
    Case false
      SetPlayerDisplay 1, one
      SetPlayerDisplay 2, two
  End Select
End Sub

Sub Gate004_Hit()
  Gate004.TimerEnabled = True
  Gate003.Collidable = False
  Gate002.Collidable = False
  If LockedBalls > 0 Then
    AddScore(LockedBalls*5*TheLevel(TableLevel))
    AddBonus(LockedBalls*5)
    Kicker001Kick()
    Kicker002Kick()  
    Kicker003Kick()  
    LockedBalls = 0
    MultiBalled = true
  End If
  AddBonus(2)
End Sub

Sub Gate004_Timer
  debug.print "gate004 timer"
  Gate004.TimerEnabled = False
  Gate003.Collidable = True
  Gate00t2.Enabled = 1
End Sub

Sub Gate00t2_Init()
  Gate00t2.Interval = 181
End Sub

Sub Gate00t2_Timer()
  debug.print "gate00t2 timer"
  Gate00t2.Enabled = 0
  Gate002.Collidable = True
End Sub

Sub Gate005_Hit()
  AddBonus(4)
End Sub

Sub Gate005_Timer()
  debug.print "gate005 timer"
  Gate005.Collidable = True
  Gate005.TimerEnabled = False
End Sub

Sub LeftInLane_Hit()
  AddScore(1*TheLevel(TableLevel))
  AddBonus(1)
  MissionRun("trigger")
End Sub

Sub LeftOutLane_Hit()
  AddScore(1*TheLevel(TableLevel))
  AddBonus(2)
  MissionRun("trigger")
End Sub

Sub LeftSlingShot_Hit()
  AddScore(1*TheLevel(TableLevel))
End Sub

Sub RightInLane_Hit()
  AddScore(1*TheLevel(TableLevel))
  AddBonus(1)
  MissionRun("trigger")
End Sub

Sub RightOutLane_Hit()
  AddScore(1*TheLevel(TableLevel))
  AddBonus(2)
  MissionRun("trigger")
End Sub

Sub RightSlingShot_Hit()
  AddScore(1*TheLevel(TableLevel))
End Sub

Sub TargetSet2CheckUp()
  Dim t, all_down, level: all_down = true:level = 0
  For each t in Targets2
    If t.IsDropped = 0 Then all_down = false
    If level < t.UserValue Then level = t.UserValue
  Next
  If all_down AND ( level = Kicker002.UserValue) Then
    Bumper001TryUpgrade()
    Spinner001TryUpgrade()
    Trigger013.UserValue = level
    Guardian.IsDropped = 0
  End If
  If all_down Then
    BonusXTargets(2) = True
    SetBonusX()
    AddScore(3*TheLevel(level))
    AddBonus(6)
    Kicker002.UserValue  = level
    SetLightLevel Tr1li019, level
    SetLightLevel Tr1li020, level
    SetLightLevel Tr1li011, level
    SetLightLevel Tr1li012, level
    Bumper002TryUpgrade()
    Spinner002TryUpgrade()
    If Trigger014.UserValue = level Then
      Trigger012.UserValue = level
      Trigger013.UserValue = level
    End If
    Trigger014.UserValue = level
    For each t in Targets2
      t.UserValue = level
      t.IsDropped = 0
    Next
  End If
End Sub

Sub Target001_Dropped()
  Dim t: Set  t = Target001
  AddScore(3*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet2CheckUp()
End Sub

Sub Target002_Dropped()
  Dim t: Set  t = Target002
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet2CheckUp()
End Sub

Sub Target003_Dropped()
  Dim t: Set  t = Target003
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet2CheckUp()
End Sub

Sub TargetSet1CheckUp()
  Dim t, all_down, level: all_down = true:level = 0
  For each t in Targets1
    If t.IsDropped = 0 Then all_down = false
    If level < t.UserValue Then level = t.UserValue
  Next
  If all_down AND Kicker003.UserValue = level Then
    LeftShield.IsDropped = 0
  End If
  If all_down Then
    BonusXTargets(1) = True
    SetBonusX()
    AddScore(5*TheLevel(level))
    AddBonus(5)
    Kicker003.UserValue = level
    RightShield.IsDropped = 0
    SetLightLevel Tr1li011, level
    SetLightLevel Tr1li012, level
    SetLightLevel Tr1li013, level
    SetLightLevel Tr1li014, level
    SetLightLevel Tr1li015, level
    SetLightLevel Tr1li018, level
    SetLightLevel Tr1li021, level
    SetLightLevel Tr1li022, level
    Bumper002TryUpgrade()
    Spinner002TryUpgrade()
    Trigger009.UserValue = level
    Trigger010.UserValue = level
    Trigger011.UserValue = level
    Trigger012.UserValue = level
    Trigger013.UserValue = level
    Trigger014.UserValue = level
    For each t in Targets1
      t.UserValue = level
      t.IsDropped = 0
    Next
  End If
End Sub

Sub Target004_Dropped()
  Dim t: Set  t = Target004
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet1CheckUp()
End Sub

Sub Target005_Dropped()
  Dim t: Set  t = Target005
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet1CheckUp()
End Sub

Sub Target006_Dropped()
  Dim t: Set  t = Target006
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  t.UserValue = TableLevel + 1
  TargetSet1CheckUp()
End Sub

Sub Target008_Hit()
  Dim mission
  Dim t: Set  t = Target008
  Dim tl: Set tl = Tr1li025
  AddBonus(2)
  MissionRun("target")

  if OnMission > 0 Then Exit Sub
  Select Case t.UserValue > TableLevel
    Case false : t.UserValue = TableLevel + 1 : SetLightLevel tl, t.UserValue
    Case true : t.UserValue = -1 : SetLightLevel tl, t.UserValue
  End Select

  mission = Get_mission()

  If MissionDone(mission) Then
    SetLightLevel Tr1li001, TableLevel
  Else
    SetLightLevel Tr1li001, TableLevel + 1
  End If
  Gate003.UserValue = TableLevel + 1
End Sub

Sub Target009_Hit()
  Dim mission
  Dim t: Set  t = Target009
  Dim tl: Set tl = Tr1li024
  AddBonus(2)
  MissionRun("target")
  if OnMission > 0 Then Exit Sub
  Select Case t.UserValue > TableLevel
    Case false : t.UserValue = TableLevel + 1 : SetLightLevel tl, t.UserValue
    Case true : t.UserValue = -1 : SetLightLevel tl, t.UserValue
  End Select

  mission = Get_mission()

  If MissionDone(mission) Then
    SetLightLevel Tr1li001, TableLevel
  Else
    SetLightLevel Tr1li001, TableLevel + 1
  End If
  Gate003.UserValue = TableLevel + 1
End Sub

Sub Target010_Hit()
  Dim mission
  Dim t: Set  t = Target010
  Dim tl: Set tl = Tr1li023
  AddBonus(2)
  MissionRun("target")
  if OnMission > 0 Then Exit Sub
  Select Case t.UserValue > TableLevel
    Case false : t.UserValue = TableLevel + 1 : SetLightLevel tl, t.UserValue
    Case true : t.UserValue = -1 : SetLightLevel tl, t.UserValue
  End Select

  mission = Get_mission()

  If MissionDone(mission) Then
    SetLightLevel Tr1li001, TableLevel
  Else
    SetLightLevel Tr1li001, TableLevel + 1
  End If
  Gate003.UserValue = TableLevel + 1
End Sub


Sub TargetSet3CheckUp()
  Dim t, all_down, level: all_down = true:level = 0
  For each t in Targets3
    debug.print "TargetSet3 t:" & t.Name & "isdropped:" & t.IsDropped
    If t.IsDropped = 0 Then all_down = false
    If level < t.UserValue Then level = t.UserValue
  Next
  debug.print "TargetSet3CheckUp " & all_down
  If Target011.IsDropped = 1 AND Target012.IsDropped = 1 AND Trigger007.UserValue < Target011.UserValue Then
    Trigger007.UserValue = level
  End If
  If all_down AND ( level = Trigger007.UserValue ) Then
    Bumper002TryUpgrade()
    Spinner002TryUpgrade()
    LeftShield.IsDropped = 0
    Trigger010.UserValue = level
  End If
  If all_down Then
    BonusXTargets(3) = True
    SetBonusX()
    AddScore(3*TheLevel(level))
    AddBonus(6)
    Bumper001TryUpgrade()
    Spinner001TryUpgrade()
    If Trigger009.UserValue = level Then
      Trigger010.UserValue = level
      Trigger011.UserValue = level
    End If
    Trigger009.UserValue = level
    SetLightLevel Tr1li013, level
    SetLightLevel Tr1li014, level
    For each t in Targets3
      t.UserValue = level
      t.IsDropped = 0
    Next
  End If
End Sub

Sub Target011_Dropped()
  Dim t: Set  t = Target011
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr1li017, t.UserValue
    SetLightLevel Tr1li018, t.UserValue
  End If
  TargetSet3CheckUp()
End Sub

Sub Target012_Dropped()
  Dim t: Set  t = Target012
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr1li015, t.UserValue
    SetLightLevel Tr1li016, t.UserValue
  End If
  TargetSet3CheckUp()
End Sub

Sub Target013_Dropped()
  Dim t: Set  t = Target013
  AddScore(2*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("target")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr1li009, t.UserValue
    SetLightLevel Tr1li010, t.UserValue
  End If
  TargetSet3CheckUp()
End Sub

Sub TriggerSet1CheckUpLevel()
  debug.print "Trigger Set 1 check up level"
  Dim t, i, tmp, level, h: tmp = 0 : i = 0
  Dim  bonus_level : bonus_level = Fix( CurrentBonus / 55 )
  AddScore(3*TheLevel(TableLevel))
  AddBonus(3)
  For each t in TriggerSet1
    i = i + 1
    tmp = tmp + t.UserValue
  Next
  If ((tmp Mod i) = 0) Then
    AddScore(15*TheLevel(TableLevel))
    level = Fix(tmp / i)
    debug.print "Trigger Set 1 Complete"
    debug.print "level:" & level & " bonus_level:" & tmp

    Select Case level
      Case 1 : If bonus_level > 0 Then TableLevel = 1 : AddScore(20*TheLevel(TableLevel))
      Case 2 : If bonus_level > 1 Then TableLevel = 2 : AddScore(20*TheLevel(TableLevel))
      Case 3 : If bonus_level > 2 Then TableLevel = 3 : AddScore(20*TheLevel(TableLevel))
      Case 4 : If bonus_level > 3 Then TableLevel = 4 : AddScore(20*TheLevel(TableLevel))
      Case 5 : If bonus_level > 4 Then TableLevel = 5 : AddScore(20*TheLevel(TableLevel))
      Case 6 : If LockedFlippers = false Then Multiballed = false : ResetDoneMissions()
    End Select

    If level = TableLevel Then
      SetLevel(TableLevel)
      PlayerLevel(CurrentPlayer) = TableLevel
      BonusXthisLevel = False
      ResetBonusXTargets()
      If LockedFlippers = false Then Multiballed = false
    End If
  End If
End Sub

Sub SetLightLevel(in_light, level)
  Select Case level
    Case -2 :in_light.Color = RGB(0, 0, 0) : in_light.ColorFull = RGB(0,0, 0)
    Case -1 :in_light.Color = RGB(20, 20, 30) : in_light.ColorFull = RGB(14,4, 4)
    Case 0 : in_light.Color = RGB(255, 255, 255) : in_light.ColorFull = RGB(44,44, 44)
    Case 1 : in_light.Color = RGB(66, 255, 255) : in_light.ColorFull = RGB(222,222, 222)
    Case 2 : in_light.Color = RGB(70, 255, 115) : in_light.ColorFull = RGB(32,111, 72)
    Case 3 : in_light.Color = RGB(255, 255, 0) : in_light.ColorFull = RGB(111,111, 22)
    Case 4 : in_light.Color = RGB(255, 165, 0) : in_light.ColorFull = RGB(111,22, 0)
    Case Else : in_light.Color = RGB(255, 0, 0) : in_light.ColorFull = RGB(22,0, 0)
  End Select
End Sub

Sub Trigger001_Hit()
  Dim t: Set  t = Trigger001
  AddScore(4*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("trigger")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr1Li1, t.UserValue
    SetLightLevel Tr1Li2, t.UserValue
    Spinner002TryUpgrade()
  End If
  TriggerSet1CheckUpLevel()
End Sub

Sub Trigger002_Hit()
  Dim t: Set  t = Trigger002
  AddScore(4*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("trigger")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr2Li1, t.UserValue
    SetLightLevel Tr2Li2, t.UserValue
    Bumper002TryUpgrade()
  End If
  TriggerSet1CheckUpLevel()
End Sub

Sub Trigger003_Hit()
  Dim t: Set t = Trigger003
  AddScore(4*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("trigger")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr3Li1, t.UserValue
    SetLightLevel Tr3Li2, t.UserValue
    Spinner001TryUpgrade()
  End If
  TriggerSet1CheckUpLevel()
End Sub

Sub Trigger004_Hit()
  Dim t: Set t = Trigger004
  AddScore(4*TheLevel(t.UserValue))
  AddBonus(2)
  MissionRun("trigger")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr4Li1, t.UserValue
    SetLightLevel Tr4Li2, t.UserValue
    Bumper001TryUpgrade()
  End If
  TriggerSet1CheckUpLevel()
End Sub

Sub Trigger007_Hit()
  Dim t: Set t = Trigger007
  AddScore(1*TheLevel(t.UserValue))
  AddBonus(1)
  MissionRun("trigger")
  if t.UserValue < TableLevel + 1 Then
    t.UserValue = TableLevel + 1
    SetLightLevel Tr1li015, t.UserValue
    SetLightLevel Tr1li018, t.UserValue
  Elseif t.UserValue = TableLevel + 1 Then
    Spinner001TryUpgrade()
  End If
End Sub

Sub Trigger009_Hit()
  AddScore(3*TheLevel(Trigger009.UserValue))
  AddBonus(3)

  If LeftShield.IsDropped = 0 AND RightShield.IsDropped = 0 Then Guardian.IsDropped = 0
  If EasyMode AND LeftShield.IsDropped = 0 Then RightShield.IsDropped = 0

  LeftShield.IsDropped = 0
End Sub

Sub Trigger010_Hit()
  AddScore(2*TheLevel(Trigger010.UserValue))
  AddBonus(3)
End Sub

Sub Trigger011_Hit()
  AddScore(1*TheLevel(Trigger011.UserValue))
  AddBonus(3)
  MissionRun("trigger")
End Sub

Sub Trigger012_Hit()
  AddScore(3*TheLevel(Trigger012.UserValue))
  AddBonus(3)
  MissionRun("trigger")
End Sub

Sub Trigger013_Hit()
  AddScore(4*TheLevel(Trigger013.UserValue))
  AddBonus(4)
End Sub

Sub Trigger014_Hit()
  AddScore(5*TheLevel(Trigger014.UserValue))
  AddBonus(5)
  LeftShield.IsDropped = 0
  Guardian.IsDropped = 0
  RightShield.IsDropped = 0
End Sub

Sub Spinner001TryUpgrade()
  Dim s: Set s = Spinner001
  if s.UserValue < TableLevel + 1 Then
    s.UserValue = TableLevel + 1
    SetLightLevel Tr1li003, s.UserValue
    SetLightLevel Tr1li004, s.UserValue
    SetLightLevel Tr1li007, s.UserValue
  End If
End Sub

Sub Spinner001_Timer()
  Spinner001.TimerEnabled = 0
End Sub

Sub Spinner001_Spin()
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner001), 0.25, 0, 0, 1, AudioFade(Spinner001)
  AddScore(1*TheLevel(Spinner001.UserValue))
  If Spinner001.TimerEnabled = 0 Then MissionRun("spinner") : Spinner001.TimerEnabled = True
End Sub

Sub Spinner002TryUpgrade()
  Dim s: Set s = Spinner002
  if s.UserValue < TableLevel + 1 Then
    s.UserValue = TableLevel + 1
    SetLightLevel Tr1li005, s.UserValue
    SetLightLevel Tr1li006, s.UserValue
    SetLightLevel Tr1li008, s.UserValue
  End If
End Sub

Sub Spinner002_Timer()
  Spinner002.TimerEnabled = 0
End Sub

Sub Spinner002_Spin()
  PlaySound "fx_spinner", 0, .25, AudioPan(Spinner001), 0.25, 0, 0, 1, AudioFade(Spinner001)
  AddScore(1*TheLevel(Spinner002.UserValue))
  If Spinner002.TimerEnabled = 0 Then MissionRun("spinner") : Spinner002.TimerEnabled = True
End Sub

Sub Kicker001_Hit()
  Kicker001.TimerEnabled = True
End Sub

Sub Kicker001Kick()
    Kicker001.Kick 69, 14
    PlaySound "fx_bumper"& RndNbr(4), 0, .25, AudioPan(Kicker001), 0.25, 0, 0, 1, AudioFade(Kicker001)
End Sub

Sub Kicker001_Timer
  Dim k: Set k = Kicker001
  AddScore(2*TheLevel(k.UserValue))
  AddBonus(1)
  k.UserValue = TableLevel + 1
  If Trigger007.UserValue > TableLevel AND MultiBalled = false AND BIP <= MaxBalls Then
    AddScore(2*TheLevel(k.UserValue))
    AddBonus(3)
    CreateABall()
    LockedBalls = LockedBalls + 1
    SetLightLevel Tr1li002, TableLevel + 1
  Else
    Kicker001Kick()
  End If
  Kicker001.TimerEnabled = False
End Sub

Sub Kicker002_Hit()
  Kicker002.TimerEnabled = True
End Sub

Sub Kicker002Kick()
  Kicker002Protector.IsDropped = 1
  Kicker002.Kick -49, 20
  PlaySound "fx_bumper"& RndNbr(4), 0, .25, AudioPan(Kicker002), 0.25, 0, 0, 1, AudioFade(Kicker002)
End Sub

Sub Kicker002_Timer
  Dim k: Set k = Kicker002
  AddScore(2*TheLevel(k.UserValue))
  AddBonus(1)
  k.UserValue = TableLevel + 1
  If Trigger014.UserValue > TableLevel AND MultiBalled = false AND BIP <= MaxBalls Then
    AddScore(2*TheLevel(Kicker002.UserValue))
    AddBonus(3)
    Kicker002Protector.IsDropped = 0
    CreateABall()
    LockedBalls = LockedBalls + 1
    SetLightLevel Tr1li002, TableLevel + 1
  Else
    Kicker002Kick()
  End If
  Kicker002.TimerEnabled = False
End Sub

Sub Kicker003_Hit()
  Kicker003.TimerEnabled = True
End Sub

Sub Kicker003Kick()
  Kicker003Protector.IsDropped = 1
  Kicker003.Kick 22, 22
  PlaySound "fx_bumper"& RndNbr(4), 0, .25, AudioPan(Kicker003), 0.25, 0, 0, 1, AudioFade(Kicker003)
End Sub

Sub Kicker003_Timer
  Dim k: Set k = Kicker003
  AddScore(2*TheLevel(k.UserValue))
  AddBonus(1)
  k.UserValue = TableLevel + 1
  If Trigger009.UserValue > TableLevel AND MultiBalled = false AND BIP <= MaxBalls Then
    AddScore(2*TheLevel(Kicker002.UserValue))
    AddBonus(3)
    Kicker003Protector.IsDropped = 0
    CreateABall()
    LockedBalls = LockedBalls + 1
    SetLightLevel Tr1li002, TableLevel + 1
  Else
    Kicker003Kick()
  End If
  Kicker003.TimerEnabled = 0
End Sub

Dim EnableBallControl
EnableBallControl = false 'Change to true to enable manual ball control (or press C in-game) via the arrow keys and B (boost movement) keys

Const BallSize = 25  'Ball radius

Sub Table1_KeyDown(ByVal keycode)
    If LockedFlippers = false Then
        If keycode = LeftTiltKey Then CheckTilt 'only check the tilt during game
        If keycode = RightTiltKey Then CheckTilt
        If keycode = CenterTiltKey Then CheckTilt
    End If
	If keycode = PlungerKey Then
        If EnableRetractPlunger Then
            Plunger.PullBackandRetract
        Else
		    Plunger.PullBack
        End If
		PlaySound "plungerpull",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

	If keycode = LeftFlipperKey AND LockedFlippers = false Then
        LeftFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
		LeftFlipper.RotateToEnd
		PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
	End If

	If keycode = RightFlipperKey AND LockedFlippers = false Then
        RightFlipper.TimerEnabled = True 'This line is only for ninuzzu's flipper shadows!
		RightFlipper.RotateToEnd
		PlaySound SoundFX("fx_flipperup",DOFFlippers), 0, .67, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
	End If

	If keycode = LeftTiltKey Then
		Nudge 90, 2
	End If

	If keycode = RightTiltKey Then
		Nudge 270, 2
	End If

	If keycode = CenterTiltKey Then
		Nudge 0, 2
	End If

    If keycode = AddCreditKey Then
      If Credits < 9 Then Credits = Credits + 1 : 	PlaySound "target", 0, 1, AudioPan(Plunger), 0, 0, 0, 0, AudioFade(Plunger)

      CreditsPlayerBall()
    End If

    If keycode = StartGameKey Then
      If((PlayersPlayingGame < MaxPlayers)AND(BeforeTheFirstBall = 1)) Then
        If(Credits > 0)then
          Credits = Credits - 1
          AddPlayer()
          If PlayersPlayingGame = 1 Then ResetForNewGame()
        End If
      End If
    End If

    ' Manual Ball Control
	If keycode = 46 Then	 				' C Key
		If EnableBallControl = 1 Then
			EnableBallControl = 0
		Else
			EnableBallControl = 1
		End If
	End If
    If EnableBallControl = 1 Then
		If keycode = 48 Then 				' B Key
			If BCboost = 1 Then
				BCboost = BCboostmulti
			Else
				BCboost = 1
			End If
		End If
		If keycode = 203 Then BCleft = 1	' Left Arrow
		If keycode = 200 Then BCup = 1		' Up Arrow
		If keycode = 208 Then BCdown = 1	' Down Arrow
		If keycode = 205 Then BCright = 1	' Right Arrow
	End If
End Sub

Sub Table1_KeyUp(ByVal keycode)
	If keycode = PlungerKey Then
		Plunger.Fire
		PlaySound "plunger",0,1,AudioPan(Plunger),0.25,0,0,1,AudioFade(Plunger)
	End If

	If keycode = LeftFlipperKey AND LockedFlippers = false Then
		LeftFlipper.RotateToStart
		PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(LeftFlipper), 0.05,0,0,1,AudioFade(LeftFlipper)
	End If

	If keycode = RightFlipperKey AND LockedFlippers = false Then
		RightFlipper.RotateToStart
		PlaySound SoundFX("fx_flipperdown",DOFFlippers), 0, 1, AudioPan(RightFlipper), 0.05,0,0,1,AudioFade(RightFlipper)
	End If

    'Manual Ball Control
	If EnableBallControl = 1 Then
		If keycode = 203 Then BCleft = 0	' Left Arrow
		If keycode = 200 Then BCup = 0		' Up Arrow
		If keycode = 208 Then BCdown = 0	' Down Arrow
		If keycode = 205 Then BCright = 0	' Right Arrow
	End If
End Sub

Sub CheckTilt                                    'Called when table is nudged
    Tilt = Tilt + TiltSensitivity                'Add to tilt count
    TiltDecreaseTimer.Enabled = True
    If(Tilt > TiltSensitivity)AND(Tilt < 15)Then 'show a warning
        'show warning
    End if
    If NOT Tilted AND Tilt > MAX_Tilt Then ' TILT the table
      Tilted = True
      CurrentBonus = 0
      DisplayBonus = 1
      LastBonus = 0
      BonusX = 0
      SetBonusX()
      SetBonusLights()
      BonusTimer.enabled = true
        'display Tilt
      For each xx in GI:xx.State = 0: Next
      SetLevel(0)
      Kicker001Kick()
      Kicker002Kick()
      Kicker003Kick()
      LockedBalls = 0
      LockedFlippers = true
      MultiBalled = true
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
    End If
End Sub

Sub TiltDecreaseTimer_Timer
    ' DecreaseTilt
    If Tilt > 0 Then
        Tilt = Tilt - 0.1
    Else
        TiltDecreaseTimer.Enabled = False
    End If
End Sub

Sub Drain_Hit()
	PlaySound "drain",0,1,AudioPan(Drain),0.25,0,0,1,AudioFade(Drain)
	Drain.DestroyBall
    
	BIP = BIP - 1
	If BIP > 0 AND BIP = LockedBalls Then
      Kicker001Kick()
      Kicker002Kick()
      Kicker003Kick()
      LockedBalls = 0
      LockedFlippers = true
      MultiBalled = true
      LeftFlipper.RotateToStart
      RightFlipper.RotateToStart
    End If
	If BIP = 0 Then
        UpdateScores()
        EndCurrentPlayerTurn()
	End If
End Sub


Dim BIP
BIP = 0


'*****GI Lights On
dim xx

For each xx in GI:xx.State = 1: Next

'**********Sling Shot Animations
' Rstep and Lstep  are the variables that increment the animation
'****************
Dim RStep, Lstep

Sub RightSlingShot_Slingshot
    PlaySound SoundFX("right_slingshot",DOFContactors), 0,1, 0.05,0.05 '0,1, AudioPan(RightSlingShot), 0.05,0,0,1,AudioFade(RightSlingShot)
    RSling.Visible = 0
    RSling1.Visible = 1
    sling1.rotx = 20
    RStep = 0
    RightSlingShot.TimerEnabled = 1
	gi1.State = 0:Gi2.State = 0
End Sub

Sub RightSlingShot_Timer
    Select Case RStep
        Case 3:RSLing1.Visible = 0:RSLing2.Visible = 1:sling1.rotx = 10
        Case 4:RSLing2.Visible = 0:RSLing.Visible = 1:sling1.rotx = 0:RightSlingShot.TimerEnabled = 0:gi1.State = 1:Gi2.State = 1
    End Select
    RStep = RStep + 1
End Sub

Sub LeftSlingShot_Slingshot
    PlaySound SoundFX("left_slingshot",DOFContactors), 0,1, -0.05,0.05 '0,1, AudioPan(LeftSlingShot), 0.05,0,0,1,AudioFade(LeftSlingShot)
    LSling.Visible = 0
    LSling1.Visible = 1
    sling2.rotx = 20
    LStep = 0
    LeftSlingShot.TimerEnabled = 1
	gi3.State = 0:Gi4.State = 0
End Sub

Sub LeftSlingShot_Timer
    Select Case LStep
        Case 3:LSLing1.Visible = 0:LSLing2.Visible = 1:sling2.rotx = 10
        Case 4:LSLing2.Visible = 0:LSLing.Visible = 1:sling2.rotx = 0:LeftSlingShot.TimerEnabled = 0:gi3.State = 1:Gi4.State = 1
    End Select
    LStep = LStep + 1
End Sub


'*********************************************************************
'                 Positional Sound Playback Functions
'*********************************************************************

' Play a sound, depending on the X,Y position of the table element (especially cool for surround speaker setups, otherwise stereo panning only)
' parameters (defaults): loopcount (1), volume (1), randompitch (0), pitch (0), useexisting (0), restart (1))
' Note that this will not work (currently) for walls/slingshots as these do not feature a simple, single X,Y position
Sub PlayXYSound(soundname, tableobj, loopcount, volume, randompitch, pitch, useexisting, restart)
	PlaySound soundname, loopcount, volume, AudioPan(tableobj), randompitch, pitch, useexisting, restart, AudioFade(tableobj)
End Sub

' Similar subroutines that are less complicated to use (e.g. simply use standard parameters for the PlaySound call)
Sub PlaySoundAt(soundname, tableobj)
    PlaySound soundname, 1, 1, AudioPan(tableobj), 0,0,0, 1, AudioFade(tableobj)
End Sub

Sub PlaySoundAtBall(soundname)
    PlaySoundAt soundname, ActiveBall
End Sub


'*********************************************************************
'                     Supporting Ball & Sound Functions
'*********************************************************************

Function AudioFade(tableobj) ' Fades between front and back of the table (for surround systems or 2x2 speakers, etc), depending on the Y position on the table. "table1" is the name of the table
	Dim tmp
    tmp = tableobj.y * 2 / table1.height-1
    If tmp > 0 Then
		AudioFade = Csng(tmp ^10)
    Else
        AudioFade = Csng(-((- tmp) ^10) )
    End If
End Function

Function AudioPan(tableobj) ' Calculates the pan for a tableobj based on the X position on the table. "table1" is the name of the table
    Dim tmp
    tmp = tableobj.x * 2 / table1.width-1
    If tmp > 0 Then
        AudioPan = Csng(tmp ^10)
    Else
        AudioPan = Csng(-((- tmp) ^10) )
    End If
End Function

Function Vol(ball) ' Calculates the Volume of the sound based on the ball speed
    Vol = Csng(BallVel(ball) ^2 / 2000)
End Function

Function Pitch(ball) ' Calculates the pitch of the sound based on the ball speed
    Pitch = BallVel(ball) * 20
End Function

Function BallVel(ball) 'Calculates the ball speed
    BallVel = INT(SQR((ball.VelX ^2) + (ball.VelY ^2) ) )
End Function


'*****************************************
'   rothbauerw's Manual Ball Control
'*****************************************

Dim BCup, BCdown, BCleft, BCright
Dim ControlBallInPlay, ControlActiveBall
Dim BCvel, BCyveloffset, BCboostmulti, BCboost

BCboost = 1				'Do Not Change - default setting
BCvel = 4				'Controls the speed of the ball movement
BCyveloffset = -0.01 	'Offsets the force of gravity to keep the ball from drifting vertically on the table, should be negative
BCboostmulti = 3		'Boost multiplier to ball veloctiy (toggled with the B key) 

ControlBallInPlay = false

Sub StartBallControl_Hit()
	Set ControlActiveBall = ActiveBall
	ControlBallInPlay = true
End Sub

Sub StopBallControl_Hit()
	ControlBallInPlay = false
End Sub	

Sub BallControlTimer_Timer()
	If EnableBallControl and ControlBallInPlay then
		If BCright = 1 Then
			ControlActiveBall.velx =  BCvel*BCboost
		ElseIf BCleft = 1 Then
			ControlActiveBall.velx = -BCvel*BCboost
		Else
			ControlActiveBall.velx = 0
		End If

		If BCup = 1 Then
			ControlActiveBall.vely = -BCvel*BCboost
		ElseIf BCdown = 1 Then
			ControlActiveBall.vely =  BCvel*BCboost
		Else
			ControlActiveBall.vely = bcyveloffset
		End If
	End If
End Sub


'********************************************************************
'      JP's VP10 Rolling Sounds (+rothbauerw's Dropping Sounds)
'********************************************************************

Const tnob = 5 ' total number of balls
ReDim rolling(tnob)
InitRolling

Sub InitRolling
    Dim i
    For i = 0 to tnob
        rolling(i) = False
    Next
End Sub

Sub RollingTimer_Timer()
    Dim BOT, b
    BOT = GetBalls

    ' stop the sound of deleted balls
    For b = UBound(BOT) + 1 to tnob
        rolling(b) = False
        StopSound("fx_ballrolling" & b)
    Next

    ' exit the sub if no balls on the table
    If UBound(BOT) = -1 Then Exit Sub

    For b = 0 to UBound(BOT)
        ' play the rolling sound for each ball
        If BallVel(BOT(b) ) > 1 AND BOT(b).z < 30 Then
            rolling(b) = True
            PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        Else
            If rolling(b) = True Then
                StopSound("fx_ballrolling" & b)
                rolling(b) = False
            End If
        End If

        ' play ball drop sounds
        If BOT(b).VelZ < -1 and BOT(b).z < 55 and BOT(b).z > 27 Then 'height adjust for ball drop sounds
            PlaySound "fx_ball_drop" & b, 0, ABS(BOT(b).velz)/17, AudioPan(BOT(b)), 0, Pitch(BOT(b)), 1, 0, AudioFade(BOT(b))
        End If
    Next
End Sub

'**********************
' Ball Collision Sound
'**********************

Sub OnBallBallCollision(ball1, ball2, velocity)
	PlaySound("fx_collide"), 0, Csng(velocity) ^2 / 2000, AudioPan(ball1), 0, Pitch(ball1), 0, 0, AudioFade(ball1)
End Sub


Sub Pins_Hit (idx)
	PlaySound "pinhit_low", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Targets_Hit (idx)
	PlaySound "target", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 0, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Thin_Hit (idx)
	PlaySound "metalhit_thin", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals_Medium_Hit (idx)
	PlaySound "metalhit_medium", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Metals2_Hit (idx)
	PlaySound "metalhit2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Gates_Hit (idx)
	PlaySound "gate4", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
End Sub

Sub Spinner_Spin
	PlaySound "fx_spinner", 0, .25, AudioPan(Spinner), 0.25, 0, 0, 1, AudioFade(Spinner)
End Sub

Sub Bumper_Bump
	PlaySound "fx_bumper", 0, .25, AudioPan(Bumper), 0.25, 0, 0, 1, AudioFade(Bumper)
End Sub

Sub Rubbers_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 20 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 20 then
 		RandomSoundRubber()
 	End If
End Sub

Sub Posts_Hit(idx)
 	dim finalspeed
  	finalspeed=SQR(activeball.velx * activeball.velx + activeball.vely * activeball.vely)
 	If finalspeed > 16 then 
		PlaySound "fx_rubber2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End if
	If finalspeed >= 6 AND finalspeed <= 16 then
 		RandomSoundRubber()
 	End If
End Sub

Sub RandomSoundRubber()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "rubber_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "rubber_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "rubber_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Sub LeftFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RightFlipper_Collide(parm)
 	RandomSoundFlipper()
End Sub

Sub RandomSoundFlipper()
	Select Case Int(Rnd*3)+1
		Case 1 : PlaySound "flip_hit_1", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 2 : PlaySound "flip_hit_2", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
		Case 3 : PlaySound "flip_hit_3", 0, Vol(ActiveBall), AudioPan(ActiveBall), 0, Pitch(ActiveBall), 1, 0, AudioFade(ActiveBall)
	End Select
End Sub

Function RndNbr(n) 'returns a random number between 1 and n
    Randomize timer
    RndNbr = Int((n * Rnd) + 1)
End Function

Function B2SLEDValue(CharPar)           'to be used with dB2S 7-segments-LED
  B2SLEDValue = 0
  Select Case CharPar
 
Case " ": B2SLEDValue = 	&h00
Case "!": B2SLEDValue = 	&h86
Case """": B2SLEDValue = 	&h22
Case "#": B2SLEDValue = 	&h7E
Case "$": B2SLEDValue = 	&h6D
Case "%": B2SLEDValue = 	&hD2
Case "&": B2SLEDValue = 	&h46
Case "'": B2SLEDValue = 	&h20
Case "(": B2SLEDValue = 	&h29
Case ")": B2SLEDValue = 	&h0B
Case "*": B2SLEDValue = 	&h21
Case "+": B2SLEDValue = 	&h70
Case "-": B2SLEDValue = 	&h40
Case ".": B2SLEDValue = 	&h80
Case "/": B2SLEDValue = 	&h52
Case "0": B2SLEDValue = 	&h3F
Case "1": B2SLEDValue = 	&h06
Case "2": B2SLEDValue = 	&h5B
Case "3": B2SLEDValue = 	&h4F
Case "4": B2SLEDValue = 	&h66
Case "5": B2SLEDValue = 	&h6D
Case "6": B2SLEDValue = 	&h7D
Case "7": B2SLEDValue = 	&h07
Case "8": B2SLEDValue = 	&h7F
Case "9": B2SLEDValue = 	&h6F
Case ":": B2SLEDValue = 	&h09
Case ";": B2SLEDValue = 	&h0D
Case "<": B2SLEDValue = 	&h61
Case "=": B2SLEDValue = 	&h48
Case ">": B2SLEDValue = 	&h43
Case "?": B2SLEDValue = 	&hD3
Case "@": B2SLEDValue = 	&h5F
Case "A": B2SLEDValue = 	&h77
Case "B": B2SLEDValue = 	&h7C
Case "C": B2SLEDValue = 	&h39
Case "D": B2SLEDValue = 	&h5E
Case "E": B2SLEDValue = 	&h79
Case "F": B2SLEDValue = 	&h71
Case "G": B2SLEDValue = 	&h3D
Case "H": B2SLEDValue = 	&h76
Case "I": B2SLEDValue = 	&h30
Case "J": B2SLEDValue = 	&h1E
Case "K": B2SLEDValue = 	&h75
Case "L": B2SLEDValue = 	&h38
Case "M": B2SLEDValue = 	&h37
Case "N": B2SLEDValue = 	&h54
Case "O": B2SLEDValue = 	&h3F
Case "P": B2SLEDValue = 	&h73
Case "Q": B2SLEDValue = 	&h6B
Case "R": B2SLEDValue = 	&h33
Case "S": B2SLEDValue = 	&h6D
Case "T": B2SLEDValue = 	&h78
Case "U": B2SLEDValue = 	&h3E
Case "V": B2SLEDValue = 	&h3E
Case "W": B2SLEDValue = 	&h7E
Case "X": B2SLEDValue = 	&h76
Case "Y": B2SLEDValue = 	&h6E
Case "Z": B2SLEDValue = 	&h5B
Case "[": B2SLEDValue = 	&h39
Case "\": B2SLEDValue = 	&h64
Case "]": B2SLEDValue = 	&h0F
Case "^": B2SLEDValue = 	&h23
Case "_": B2SLEDValue = 	&h08
Case "`": B2SLEDValue = 	&h02
Case "a": B2SLEDValue = 	&h77
Case "b": B2SLEDValue = 	&h7C
Case "c": B2SLEDValue = 	&h58
Case "d": B2SLEDValue = 	&h5E
Case "e": B2SLEDValue = 	&h7B
Case "f": B2SLEDValue = 	&h71
Case "g": B2SLEDValue = 	&h6F
Case "h": B2SLEDValue = 	&h74
Case "i": B2SLEDValue = 	&h10
Case "j": B2SLEDValue = 	&h0C
Case "k": B2SLEDValue = 	&h75
Case "l": B2SLEDValue = 	&h30
Case "m": B2SLEDValue = 	&h37
Case "n": B2SLEDValue = 	&h54
Case "o": B2SLEDValue = 	&h5C
Case "p": B2SLEDValue = 	&h73
Case "q": B2SLEDValue = 	&h67
Case "r": B2SLEDValue = 	&h50
Case "s": B2SLEDValue = 	&h6D
Case "t": B2SLEDValue = 	&h78
Case "u": B2SLEDValue = 	&h1C
Case "v": B2SLEDValue = 	&h1C
Case "w": B2SLEDValue = 	&h7E
Case "x": B2SLEDValue = 	&h76
Case "y": B2SLEDValue = 	&h6E
Case "z": B2SLEDValue = 	&h5B
Case "{": B2SLEDValue = 	&h46
Case "|": B2SLEDValue = 	&h30
Case "}": B2SLEDValue = 	&h70
Case "~": B2SLEDValue = 	&h01
  end select
  B2SLEDValue = cint(B2SLEDValue)
End Function