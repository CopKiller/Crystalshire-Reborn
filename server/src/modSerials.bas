Attribute VB_Name = "modSvSerials"
Option Explicit

'For Clear functions
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

' serial number
Public Const MAX_SERIAL_NUMBER As Integer = 100

' quant items give for serial
Public Const MAX_SERIAL_ITEMS As Byte = 10

Private Const MAX_SERIAL_LENGTH As Byte = 14

' Serial Database
Public Serial(1 To MAX_SERIAL_NUMBER) As SerialRec

Public Type SerialRec
    ' INIT
    Name As String * NAME_LENGTH
    Serial As String * NAME_LENGTH
    ' CONFIG
    NamePlayer As String * NAME_LENGTH
    GiveOne As Byte
    Blocked As Byte
    BirthDay As Byte
    ' ITEMS
    Item(1 To MAX_SERIAL_ITEMS) As Integer
    ItemValue(1 To MAX_SERIAL_ITEMS) As Long
    ' ADICIONAIS
    VipDays As Integer
    GiveSpell As Integer
    GiveGuildSlot As Byte
    ' MSG
    Msg As String * DESC_LENGTH
End Type

Sub SendSerialWindow(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSerialWindow
    SendDataTo Index, Buffer.ToArray()

    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleEditSerial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSerialEditor
    SendDataTo Index, Buffer.ToArray()
    Buffer.Flush: Set Buffer = Nothing
End Sub

Sub HandleSendSerial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Buffer As clsBuffer
    Dim SerialNumber As String

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    SerialNumber = Buffer.ReadString

    Buffer.Flush: Set Buffer = Nothing

    ClaimSerial Index, SerialNumber
End Sub

Private Sub ClaimSerial(ByVal Index As Long, ByVal SerialNumber As String)
    Dim SerialIndex As Byte
    Dim I As Byte
    Dim GuildSlotsSomados As Byte

    ' Está jogando?
    If Not IsPlaying(Index) Then
        Exit Sub
    End If
    
    If SerialNumber = vbNullString Then
        Exit Sub
    End If

    ' Validar o tamanho do serial.
    If Len(SerialNumber) > MAX_SERIAL_LENGTH Then
        Call AlertMsg(Index, DIALOGUE_SERIAL_INCORRECT, NO, False)
        Exit Sub
    End If

    ' Verificar se o serial existe e obter o número dele.
    SerialIndex = ValidadeSerialExist(SerialNumber)
    If SerialIndex = 0 Then
        Call AlertMsg(Index, DIALOGUE_SERIAL_INCORRECT, NO, False)
        Exit Sub
    End If

    ' Verificar se o serial está blockeado por poder ser reinvindicado apenas uma vez!
    If Serial(SerialIndex).Blocked > 0 Then
        Call AlertMsg(Index, DIALOGUE_SERIAL_INCORRECT, NO, False)
        Exit Sub
    End If
    
    ' Verificar se o jogador possui espaço pra receber tudo do pacote!
    If Not CanGetSerial(Index, SerialIndex) Then Exit Sub

    ' Tudo certo, vamos entregar os brindes!
    For I = 1 To MAX_SERIAL_ITEMS
        If Serial(SerialIndex).Item(I) > 0 Then
            If GiveInvItem(Index, Serial(SerialIndex).Item(I), Serial(SerialIndex).ItemValue(I), 0) Then
                Call PlayerMsg(Index, "Você recebeu na sua bolsa " & Serial(SerialIndex).ItemValue(I) & " " & Trim$(Item(Serial(SerialIndex).Item(I)).Name), BrightGreen)
            End If
        End If
    Next I
    
    ' ADICIONAIS
    If Serial(SerialIndex).VipDays > 0 Then
        Call AddPremiumTime(Index, Serial(SerialIndex).VipDays)
    End If
    
    If Serial(SerialIndex).GiveSpell > 0 Then
        Call GivePlayerSpell(Index, Serial(SerialIndex).GiveSpell)
    End If
    
    If Serial(SerialIndex).GiveGuildSlot > 0 Then
        If Player(Index).Guild_ID > 0 Then
            If Guild(Player(Index).Guild_ID).Capacidade < MAX_GUILD_MEMBERS Then
                GuildSlotsSomados = Guild(Player(Index).Guild_ID).Capacidade + Serial(SerialIndex).GiveGuildSlot
                If GuildSlotsSomados > MAX_GUILD_MEMBERS Then
                    GuildSlotsSomados = MAX_GUILD_MEMBERS
                    Guild(Player(Index).Guild_ID).Capacidade = GuildSlotsSomados
                    ReDim Preserve GuildMembers(Player(Index).Guild_ID).Membro(1 To Guild(Player(Index).Guild_ID).Capacidade)
                    
                    For I = 1 To Guild(Player(Index).Guild_ID).Capacidade
                        If GuildMembers(Player(Index).Guild_ID).Membro(I).Name = vbNullString Then
                            GuildMembers(Player(Index).Guild_ID).Membro(I).MembroDisponivel = True
                        End If
                    Next I
                    
                    SaveGuild Player(Index).Guild_ID
                    GuildCache_Create Player(Index).Guild_ID

                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If Player(I).Guild_ID = Player(Index).Guild_ID Then
                                SendUpdateGuildTo I, Player(Index).Guild_ID
                                Call PlayerMsg(I, "Novo espaço da guild alterado pra " & Guild(Player(Index).Guild_ID).Capacidade & " Membros!", BrightGreen)
                            End If
                        End If
                    Next I
                Else
                    Guild(Player(Index).Guild_ID).Capacidade = Guild(Player(Index).Guild_ID).Capacidade + Serial(SerialIndex).GiveGuildSlot
                    ReDim Preserve GuildMembers(Player(Index).Guild_ID).Membro(1 To Guild(Player(Index).Guild_ID).Capacidade)
                    
                    For I = 1 To Guild(Player(Index).Guild_ID).Capacidade
                        If GuildMembers(Player(Index).Guild_ID).Membro(I).Name = vbNullString Then
                            GuildMembers(Player(Index).Guild_ID).Membro(I).MembroDisponivel = True
                        End If
                    Next I
                    
                    SaveGuild Player(Index).Guild_ID
                    GuildCache_Create Player(Index).Guild_ID

                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If Player(I).Guild_ID = Player(Index).Guild_ID Then
                                SendUpdateGuildTo I, Player(Index).Guild_ID
                                Call PlayerMsg(I, "Novo espaço da guild alterado pra " & Guild(Player(Index).Guild_ID).Capacidade & " Membros!", BrightGreen)
                            End If
                        End If
                    Next I
                End If
            End If
        End If
    End If

    ' Serial pode ser obtido apenas uma vez por jogador?
    If Serial(SerialIndex).GiveOne > 0 Then
        Player(Index).Serial(SerialIndex).Hash = GetSerialHash(SerialNumber)
    ElseIf Serial(SerialIndex).BirthDay > 0 Then
        Player(Index).Serial(SerialIndex).YearBirthday = CInt(Year(Date))
    Else ' Caso não seja de obter 1 vez por jogador então bloqueia o serial após o primeiro jogador utilizar!
        Serial(SerialIndex).Blocked = YES
        SaveSerial SerialIndex
        SerialCache_Create SerialIndex
        SendSerialsAll SerialIndex
    End If
    
    ' Envia a mensagem caso tenha uma.
    If LenB(Trim$(Serial(SerialIndex).Msg)) > 0 Then
        Call PlayerMsg(Index, Trim$(Serial(SerialIndex).Msg), BrightBlue)
    End If
    
    ' Salva o jogador
    SavePlayer Index
    
    Call AddLog(GetPlayerName(Index) & " Usou o serial nº " & SerialIndex & ": " & SerialNumber, PLAYER_LOG)

    Call AlertMsg(Index, DIALOGUE_SERIAL_CLAIMED, NO, False)
End Sub

Private Function CanGetSerial(ByVal Index As Long, ByVal SerialIndex As Byte) As Boolean
    Dim I As Byte

    If Serial(SerialIndex).GiveOne > 0 Then
        If Player(Index).Serial(SerialIndex).Hash = GetSerialHash(Trim$(Serial(SerialIndex).Serial)) Then
            Call PlayerMsg(Index, "Serial Falhou: Você ja utilizou este serial infelizmente.", BrightRed)
            CanGetSerial = False
            Exit Function
        End If
    Else
        Player(Index).Serial(SerialIndex).Hash = 0
    End If

    If Trim$(Serial(SerialIndex).NamePlayer) <> vbNullString Then
        If GetPlayerName(Index) <> Trim$(Serial(SerialIndex).NamePlayer) Then
            Call PlayerMsg(Index, "Serial Falhou: Serial vinculado apenas ao jogador " & Trim$(Serial(SerialIndex).NamePlayer), BrightRed)
            CanGetSerial = False
            Exit Function
        End If
    End If

    ' Verifica se o serial entrega itens e verifica se o jogador tem espaço
    For I = 1 To MAX_SERIAL_ITEMS
        If Serial(SerialIndex).Item(I) > 0 Then
            If FindOpenInvSlot(Index, Serial(SerialIndex).Item(I)) = 0 Then
                Call PlayerMsg(Index, "Serial Falhou: Você precisa de espaço na bolsa pra obter " & Serial(SerialIndex).ItemValue(I) & " " & Trim$(Item(Serial(SerialIndex).Item(I)).Name), BrightRed)
                CanGetSerial = False
                Exit Function
            End If
        End If
    Next I

    ' Verifica se o pacote tem spell e processa se tem espaço, etc...
    If Serial(SerialIndex).GiveSpell > 0 Then
        If HasSpell(Index, Serial(SerialIndex).GiveSpell) Then
            Call PlayerMsg(Index, "Você já possui a spell " & Trim$(Spell(Serial(SerialIndex).GiveSpell).Name) & " mas vai receber o pacote normalmente!", BrightGreen)
        ElseIf FindOpenSpellSlot(Index) = 0 Then
            Call PlayerMsg(Index, "Serial Falhou: Você não tem slot pra receber a spell " & Trim$(Spell(Serial(SerialIndex).GiveSpell).Name) & " libere espaço de spells!", BrightRed)
            CanGetSerial = False
            Exit Function
        End If
    End If

    If Serial(SerialIndex).BirthDay > 0 Then
        If GetPlayerBirthday(Index) <> Date Then
            Call PlayerMsg(Index, "Serial Falhou: Não é dia do seu aniversário!", BrightRed)
            Call SendTimeToBirthday(Index)
            CanGetSerial = False
            Exit Function
        End If

        If Player(Index).Serial(SerialIndex).YearBirthday = Year(Date) Then
            Call PlayerMsg(Index, "Serial Falhou: Você ja utilizou este serial no seu aniversário, volte no próximo!.", BrightRed)
            CanGetSerial = False
            Exit Function
        End If
    End If

    CanGetSerial = True
End Function

Function ValidadeSerialExist(ByVal SerialNumber As String) As Byte
    Dim I As Byte

    ValidadeSerialExist = 0

    For I = 1 To MAX_SERIAL_NUMBER
        If Trim$(Serial(I).Name) <> vbNullString Then
            If Trim$(Serial(I).Serial) = SerialNumber Then
                ValidadeSerialExist = I
                Exit Function
            End If
        End If
    Next I
End Function

Sub HandleSaveSerial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Buffer As clsBuffer
    Dim SerialSize As Long
    Dim SerialData() As Byte
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    n = Buffer.ReadLong    'CLng(Parse(1))

    If n < 0 Or n > MAX_SERIAL_NUMBER Then
        Buffer.Flush: Set Buffer = Nothing: Exit Sub
    End If

    ' Update the item
    SerialSize = LenB(Serial(n))
    ReDim SerialData(SerialSize - 1)
    SerialData = Buffer.ReadBytes(SerialSize)
    CopyMemory ByVal VarPtr(Serial(n)), ByVal VarPtr(SerialData(0)), SerialSize

    Buffer.Flush: Set Buffer = Nothing

    ' Save it
    Call SerialCache_Create(n)
    Call SendSerialsAll(n)
    Call SaveSerial(n)
    Call AddLog(GetPlayerName(Index) & " saved Serial #" & n & ".", ADMIN_LOG)
End Sub

Sub HandleRequestSerial(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)

' Prevent hacking
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    ' Envia os seriais novamente, usado ao cancelar o editor após fazer a limpa.
    Call SendSerial(Index)
End Sub

Sub SaveSerials()
    Dim I As Long

    For I = 1 To MAX_SERIAL_NUMBER
        Call SaveSerial(I)
    Next

End Sub

Sub SaveSerial(ByVal SerialNum As Long)
    Dim FileName As String
    Dim F As Long
    FileName = App.Path & "\data\Serials\serial" & SerialNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
    Put #F, , Serial(SerialNum)
    Close #F
End Sub

Sub SendDataToAdmins(ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex

        If IsPlaying(I) Then
            If GetPlayerAccess(I) >= ADMIN_DEVELOPER Then
                Call SendDataTo(I, Data)
            End If
        End If

    Next I

End Sub

Sub CheckSerials()
    Dim I As Long

    For I = 1 To MAX_SERIAL_NUMBER

        If Not FileExist("\data\serials\serial" & I & ".dat") Then
            Call SaveSerial(I)
        End If
    Next
End Sub

Sub LoadSerials()
    Dim FileName As String
    Dim I As Long
    Dim F As Long

    Call CheckSerials

    For I = 1 To MAX_SERIAL_NUMBER
        FileName = App.Path & "\data\serials\serial" & I & ".dat"
        F = FreeFile
        Open FileName For Binary As #F
        Get #F, , Serial(I)
        Close #F
    Next
End Sub

Sub ClearSerial(ByVal Index As Long)
    Call ZeroMemory(ByVal VarPtr(Serial(Index)), LenB(Serial(Index)))
    Serial(Index).Name = vbNullString
    Serial(Index).Serial = vbNullString
    Serial(Index).Msg = vbNullString
    Serial(Index).NamePlayer = vbNullString
End Sub

Sub ClearSerials()
    Dim I As Long

    For I = 1 To MAX_SERIAL_NUMBER
        Call ClearSerial(I)
    Next

End Sub

Public Function GetSerialHash(ByVal Code As String) As Long
    Dim I As Byte

    For I = 1 To Len(Code)
            GetSerialHash = GetSerialHash + Asc(Mid(Code, I))
    Next I

End Function













