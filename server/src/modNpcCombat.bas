Attribute VB_Name = "modNpc_Combat"
Option Explicit

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcMaxVital = 0
        Exit Function
    End If

    Select Case Vital
    Case HP
        GetNpcMaxVital = ((Npc(NpcNum).Stat(Endurance) / 2)) * 10
    Case MP
        GetNpcMaxVital = (Npc(NpcNum).Stat(Intelligence) / 2) * 5 + 35
    End Select

End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    'Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then
        GetNpcVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
    Case HP
        i = (Npc(NpcNum).Stat(Stats.Willpower) * 0.8) + 6
    Case MP
        i = (Npc(NpcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select

    GetNpcVitalRegen = i

End Function

Function GetNpcSpellDamage(ByVal NpcNum As Long, ByVal SpellNum As Long) As Long
    Dim damage As Long

    ' Check for subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    ' return damage
    damage = Spell(SpellNum).Vital
    ' 10% modifier
    If damage <= 0 Then damage = 1
    GetNpcSpellDamage = Rand(damage - ((damage / 100) * 10), damage + ((damage / 100) * 10))
End Function

Public Function GetNpcDamage(ByVal NpcNum As Long) As Long
' return the calculation
    GetNpcDamage = 2 + (Npc(NpcNum).Stat(Strength) / 2)
End Function

Public Function GetNpcDefence(ByVal NpcNum As Long) As Long
    Dim Defence As Long

    ' base defence
    Defence = 1

    ' add in a player's agility
    GetNpcDefence = (Defence + (Npc(NpcNum).Stat(Agility) / 3))

End Function

Public Function CanNpcBlock(ByVal NpcNum As Long) As Long
    Dim rndNum As Long
    Dim Rate As Single

    CanNpcBlock = 0

    ' Limita a chance em 50%
    Rate = 50

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanNpcBlock = (Npc(NpcNum).Stat(Agility) / 2)
    End If

End Function

Public Function CanNpcCrit(ByVal Attacker As Long, ByVal VictimID As Long) As Boolean
    Dim rndNum As Long
    Dim Value As Long
    Dim Rate As Single

    ' Obtém a chance de acerto.
    Value = GetPlayerRawStat(VictimID, Stats.Strength) - Npc(Attacker).Stat(Stats.Strength)
    Rate = CSng(Value / GetPlayerRawStat(VictimID, Stats.Strength))

    ' Inverte os valores para obter a chance de esquiva.
    Rate = 100 - (Rate * 100)

    ' Limita a chance em 50%
    If Rate > 50 Then Rate = 50
    If Rate < 0 Then Rate = 1

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanNpcCrit = True
    End If

End Function

Public Function CanNpcDodge(ByVal Victim As Long, ByVal AttackerID As Long) As Boolean
    Dim rndNum As Long
    Dim Value As Long
    Dim Rate As Single

    ' Obtém a chance de acerto do atacante.
    Value = GetPlayerRawStat(AttackerID, Stats.Agility) - Npc(Victim).Stat(Stats.Agility)
    Rate = CSng(Value / GetPlayerRawStat(AttackerID, Stats.Agility))

    ' Inverte os valores para obter a chance de esquiva.
    Rate = 100 - (Rate * 100)

    ' Limita a chance em 50%
    If Rate > 50 Then Rate = 50
    If Rate < 0 Then Rate = 1

    rndNum = Rand(1, 100)
    If rndNum <= Rate Then
        CanNpcDodge = True
    End If
End Function
