Attribute VB_Name = "modPlayer_Combat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > Player_HighIndex Then Exit Function
    Select Case Vital
    Case HP
        Select Case GetPlayerClass(Index)
        Case 1    ' Warrior
            GetPlayerMaxVital = (GetPlayerStat(Index, Endurance) / 2) * 15 + 150
        Case 2    ' Wizard
            GetPlayerMaxVital = (GetPlayerStat(Index, Endurance) / 2) * 5 + 65
        Case 3    ' Whisperer
            GetPlayerMaxVital = (GetPlayerStat(Index, Endurance) / 2) * 5 + 65
        Case Else    ' Anything else - Warrior by default
            GetPlayerMaxVital = (GetPlayerStat(Index, Endurance) / 2) * 15 + 150
        End Select
    Case MP
        Select Case GetPlayerClass(Index)
        Case 1    ' Warrior
            GetPlayerMaxVital = (GetPlayerStat(Index, Intelligence) / 2) * 5 + 25
        Case 2    ' Wizard
            GetPlayerMaxVital = (GetPlayerStat(Index, Intelligence) / 2) * 30 + 85
        Case 3    ' Whisperer
            GetPlayerMaxVital = (GetPlayerStat(Index, Intelligence) / 2) * 30 + 85
        Case Else    ' Anything else - Warrior by default
            GetPlayerMaxVital = (GetPlayerStat(Index, Intelligence) / 2) * 5 + 25
        End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
    Case HP
        i = (GetPlayerStat(Index, Stats.Willpower) * 0.8) + 6
    Case MP
        i = (GetPlayerStat(Index, Stats.Willpower) / 4) + 12.5
    End Select

    If i < 2 Then i = 2
    GetPlayerVitalRegen = i
End Function

Public Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim weaponNum As Long, BasedAtribute As Byte

    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    If GetPlayerEquipmentNum(Index, Weapon) > 0 Then
        weaponNum = GetPlayerEquipmentNum(Index, Weapon)
        BasedAtribute = Item(weaponNum).AtributeBase

        If Item(weaponNum).Data2_Percent > 0 Then
            If BasedAtribute > 0 Then
                GetPlayerDamage = 2 + (((GetPlayerStat(Index, BasedAtribute) / 2) / 100) * Item(weaponNum).Data2)
            Else
                GetPlayerDamage = 2 + (((GetPlayerStat(Index, Strength) / 2) / 100) * Item(weaponNum).Data2)
            End If
        Else
            GetPlayerDamage = 2 + (GetPlayerStat(Index, Strength) / 2) + Item(weaponNum).Data2
        End If
    Else
        GetPlayerDamage = 2 + (GetPlayerStat(Index, Strength) / 2)
    End If
    
    ' Bonus de conjunto
    If TempPlayer(Index).Bonus.Dano > 0 Then
        GetPlayerDamage = GetPlayerDamage + TempPlayer(Index).Bonus.Dano
    End If
End Function

Public Function GetPlayerDefence(ByVal Index As Long) As Long
    Dim Defence As Long, i As Long, ItemNum As Long, DefencePercent As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    ' base defence
    Defence = 1

    ' add in a player's agility
    GetPlayerDefence = GetPlayerStat(Index, Agility)

    For i = 1 To Equipment.Equipment_Count - 1
        If i <> Equipment.Weapon Then
            ItemNum = GetPlayerEquipmentNum(Index, i)
            If ItemNum > 0 Then
                If Item(ItemNum).Data2 > 0 Then
                    If Item(ItemNum).Data2_Percent > 0 Then
                        If Item(ItemNum).AtributeBase > 0 Then
                            DefencePercent = DefencePercent + ((((GetPlayerStat(Index, Item(ItemNum).AtributeBase) / 100) * Item(ItemNum).Data2)) / 3)
                        Else
                            DefencePercent = DefencePercent + ((((GetPlayerDefence / 100) * Item(ItemNum).Data2)) / 3)
                        End If
                    Else
                        Defence = Defence + Item(ItemNum).Data2
                    End If
                End If
            End If
        End If
    Next i

    ' Bonus de conjunto
    If TempPlayer(Index).Bonus.Defesa > 0 Then
        Defence = Defence + TempPlayer(Index).Bonus.Defesa
    End If

    ' calculate
    GetPlayerDefence = (GetPlayerDefence / 3) + Defence + DefencePercent
End Function

Function GetPlayerSpellDamage(ByVal Index As Long, ByVal SpellNum As Long) As Long
    Dim Damage As Long

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then
        Exit Function
    End If

    ' return damage
    Damage = Spell(SpellNum).Vital
    ' 10% modifier
    If Damage <= 0 Then Damage = 1
    GetPlayerSpellDamage = Rand(Damage - ((Damage / 100) * 10), Damage + ((Damage / 100) * 10))
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerBlock(ByVal Index As Long, ByVal Damage As Long) As Long
    Dim rndNum As Long
    Dim Rate As Single

    CanPlayerBlock = 0

    ' Se não tiver dano pra bloquear
    If Damage <= 1 Then Exit Function

    ' Chance de block se estiver usando um shield
    If GetPlayerEquipmentNum(Index, Shield) <= 0 Then Exit Function

    ' Obtém a chance de block do shield
    Rate = Item(GetPlayerEquipmentNum(Index, Shield)).BlockChance

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanPlayerBlock = (Damage / 2)
    End If

End Function

Public Function CanPlayerCrit(ByVal Attacker As Long, Optional ByVal VictimType As Byte, Optional ByVal VictimID As Long) As Boolean
    Dim rndNum As Long
    Dim Value As Long
    Dim Rate As Single
    Dim NpcNum As Long

    ' Obtém a chance de acerto.
    If VictimType = TARGET_TYPE_PLAYER Then
        Value = GetPlayerRawStat(VictimID, Stats.Strength) - GetPlayerRawStat(Attacker, Stats.Strength)
        Rate = CSng(Value / GetPlayerRawStat(VictimID, Stats.Strength))

    ElseIf VictimType = TARGET_TYPE_NPC Then
        NpcNum = MapNpc(GetPlayerMap(Attacker)).NPC(VictimID).Num
        Value = NPC(NpcNum).Stat(Stats.Strength) - GetPlayerRawStat(Attacker, Stats.Strength)
        Rate = CSng(Value / NPC(NpcNum).Stat(Stats.Strength))
    End If

    ' Inverte os valores para obter a chance de esquiva.
    Rate = 100 - (Rate * 100)

    ' Limita a chance em 50%
    If Rate > 50 Then Rate = 50
    If Rate < 0 Then Rate = 1

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanPlayerCrit = True
    End If

End Function

Public Function CanPlayerDodge(ByVal Victim As Long, Optional ByVal AttackerType As Byte, Optional ByVal AttackerID As Long) As Boolean
    Dim rndNum As Long
    Dim Value As Long
    Dim Rate As Single
    Dim NpcNum As Long

    If TempPlayer(Victim).StunDuration > 0 Then Exit Function

    ' Obtém a chance de acerto do atacante.
    If AttackerType = TARGET_TYPE_PLAYER Then
        Value = GetPlayerRawStat(AttackerID, Stats.Agility) - GetPlayerRawStat(Victim, Stats.Agility)
        Rate = CSng(Value / GetPlayerRawStat(AttackerID, Stats.Agility))

    ElseIf AttackerType = TARGET_TYPE_NPC Then
        NpcNum = MapNpc(GetPlayerMap(Victim)).NPC(AttackerID).Num
        Value = NPC(NpcNum).Stat(Stats.Agility) - GetPlayerRawStat(Victim, Stats.Agility)
        Rate = CSng(Value / NPC(NpcNum).Stat(Stats.Agility))
    End If

    ' Inverte os valores para obter a chance de esquiva.
    Rate = 100 - (Rate * 100)

    ' Limita a chance em 50%
    If Rate > 50 Then Rate = 50
    If Rate < 0 Then Rate = 1

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function
