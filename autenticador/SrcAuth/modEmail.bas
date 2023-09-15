Attribute VB_Name = "modEmail"
Option Explicit

Private oSmtp As New EASendMailObjLib.mail

Private Const SMTP_Host As String = "smtp.office365.com"
Private Const SMTP_Port As Integer = 587
Private Const SMTP_User As String = "felipe_157@windowslive.com"
Private Const SMTP_Pass As String = "filipebispocarne"

Public Sub DestroyEmailObject()
    Set oSmtp = Nothing
End Sub

Public Sub CreateEmailObject()
    'Set oSmtp = CreateObject("EASendMailObjLib.Mail")
    'oSmtp.LicenseCode = "TryIt"

    ' Configurações do servidor SMTP do Outlook
    'oSmtp.ServerAddr = SMTP_Host  ' Endereço do servidor SMTP do Outlook
    'oSmtp.ServerPort = SMTP_Port  ' Porta para comunicação com o servidor SMTP
    'oSmtp.SSL_init  ' Detecta automaticamente a necessidade de SSL/TLS

    ' Autenticação do usuário no servidor SMTP
    'oSmtp.Username = SMTP_User  ' Seu endereço de email Outlook
    'oSmtp.Password = SMTP_Pass  ' Senha do seu email
End Sub

Public Sub SendEmail(ByVal Index As Long, ByVal sEmail As String)

    Dim sSenha As String
    
    sSenha = GetPass(sEmail)
    
    If sSenha = vbNullString Then
        SendAlertMsg Index, DIALOGUE_ACCOUNT_EMAILINVALID, MenuCount.menuLogin
        Exit Sub
    End If
    
    oSmtp.LicenseCode = "TryIt"
    
    ' Configuração do remetente
    oSmtp.FromAddr = SMTP_User  ' Seu endereço de email Outlook

    ' Configuração do destinatário
    'oSmtp.ClearRecipients
    oSmtp.AddRecipientEx sEmail, 0  ' Endereço de email do destinatário

    ' Configuração do assunto do email
    oSmtp.Subject = GAME_NAME & " - Esqueceu sua senha?"

    ' Configuração do corpo do email
    oSmtp.BodyText = "Você requisitou sua senha na " & GAME_NAME & "." & vbNewLine & vbNewLine & "Sua senha é: " & vbNewLine & sSenha

    ' Não será necessário enviar anexos
    ' Se desejar enviar um anexo, descomente as linhas abaixo e forneça o caminho do arquivo desejado
    'If oSmtp.AddAttachment("d:\test.txt") <> 0 Then
    '    MsgBox "Failed to add attachment with error:" & oSmtp.GetLastErrDescription()
    'End If
    
    ' Gmail SMTP (Servidor do Gmail)
    oSmtp.ServerAddr = SMTP_Host
    
    ' set direct SSL 465 port,
    oSmtp.ServerPort = SMTP_Port
    
    ' detect SSL/TLS automatically
    oSmtp.SSL_init
    
    ' Autenticação do usuário no servidor SMTP
    oSmtp.Username = SMTP_User  ' Seu endereço de email Outlook
    oSmtp.Password = SMTP_Pass  ' Senha do seu email

    SetStatus "Enviando email..."

    If oSmtp.SendMail() = 0 Then
        SetStatus "Requisição completa, senha foi recuperada com sucesso!"
        SendAlertMsg Index, DIALOGUE_ACCOUNT_EMAILSUCCESS, MenuCount.menuLogin
    Else
        SetStatus "Ocorreu uma falha na requisição, erro: " & oSmtp.GetLastErrDescription()
        SendAlertMsg Index, DIALOGUE_ACCOUNT_EMAILINVALID, MenuCount.menuLogin
    End If
End Sub

Public Function GetPass(ByVal Email As String) As String
    Dim F As Long
    Dim s As String
    Dim g() As String

    F = FreeFile
    Open App.Path & "\emailList.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        g = Split(s, ":")

        If Trim$(LCase(g(0))) = Trim$(LCase$(Email)) Then
            GetPass = Trim$(g(1))
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function
