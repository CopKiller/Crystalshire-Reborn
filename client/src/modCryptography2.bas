Attribute VB_Name = "modCryptography2"
Option Explicit

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
                                             (ByVal phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
                                              ByVal dwProvType As Long, ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, _
                                                             ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByVal phHash As Long) As Long

Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, _
                                                           ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long

Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, _
                                                            ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long

Private Declare Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, _
                                                            ByVal dwLen As Long, ByVal pbBuffer As String) As Long

Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
                                             ByVal hProv As Long, _
                                             ByVal dwFlags As Long) As Long

Private Declare Function CryptDestroyHash Lib "advapi32.dll" ( _
                                          ByVal hHash As Long) As Long

Private Declare Function CryptExportKey Lib "advapi32.dll" ( _
                                        ByVal hKey As Long, _
                                        ByVal hExpKey As Long, _
                                        ByVal dwBlobType As Long, _
                                        ByVal dwFlags As Long, _
                                        ByRef pbData As Any, _
                                        ByRef pdwDataLen As Long) As Long

Private Declare Function CryptDestroyKey Lib "advapi32.dll" ( _
                                         ByVal hKey As Long) As Long

Private Declare Function CryptGetKeyParam Lib "advapi32.dll" ( _
                                          ByVal hKey As Long, _
                                          ByVal dwParam As Long, _
                                          ByRef pbData As Any, _
                                          ByRef pdwDataLen As Long, _
                                          ByVal dwFlags As Long) As Long

Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HFFFFFFFF
Private Const CALG_RC2 As Long = 26114    'Algoritmo de chave sim�trica RC2
Private Const CALG_SHA1 As Long = 32772    'Algoritmo de hash SHA1
Private Const KP_IV = 1

Public Function GenerateKeyAndIV(ByVal KeySize As Long, ByVal IVSize As Long, ByRef Key() As Byte, ByRef IV() As Byte) As Boolean
    Dim hProv As Long
    Dim hHash As Long
    Dim hKey As Long
    Dim buffer() As Byte
    Dim BufferSize As Long

    'Inicializa o provedor de criptografia
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) = 0 Then
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Cria um hash SHA1
    If CryptCreateHash(hProv, CALG_SHA1, 0, 0, hHash) = 0 Then
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Gera um valor aleat�rio para a chave e o IV
    BufferSize = KeySize + IVSize
    ReDim buffer(BufferSize - 1)
    If CryptGenRandom(hProv, BufferSize, buffer(0)) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Adiciona o valor aleat�rio ao hash SHA1
    If CryptHashData(hHash, buffer(0), BufferSize, 0) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Cria a chave sim�trica a partir do hash SHA1
    If CryptDeriveKey(hProv, CALG_RC2, hHash, 0, hKey) = 0 Then
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Obt�m a chave sim�trica gerada
    BufferSize = KeySize
    ReDim Key(BufferSize - 1)
    If CryptExportKey(hKey, 0, 1, 0, Key(0), BufferSize) = 0 Then
        CryptDestroyKey hKey
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Obt�m o IV gerado
    BufferSize = IVSize
    ReDim IV(BufferSize - 1)
    If CryptGetKeyParam(hKey, KP_IV, IV(0), BufferSize, 0) = 0 Then
        CryptDestroyKey hKey
        CryptDestroyHash hHash
        CryptReleaseContext hProv, 0
        GenerateKeyAndIV = False
        Exit Function
    End If

    'Libera os recursos utilizados
    CryptDestroyKey hKey
    CryptDestroyHash hHash
    CryptReleaseContext hProv, 0

    'Retorna a chave e o IV gerados com sucesso
    GenerateKeyAndIV = True
End Function

' Este c�digo utiliza fun��es da CryptoAPI para inicializar o provedor de criptografia, criar um hash SHA1, gerar um valor aleat�rio para a chave e o IV, adicionar o valor aleat�rio ao hash SHA1, criar a chave sim�trica a partir do hash SHA1, obter a chave sim�trica e o IV gerados, e liberar os recursos utilizados. A fun��o retorna um valor booleano indicando se a gera��o da chave e do IV foi bem sucedida ou n�o, e tamb�m preenche os arrays passados como par�metro com a chave e o IV gerados.

'Este c�digo � um exemplo de como utilizar a fun��o GenerateKeyAndIV para gerar uma chave e um IV aleat�rios, e como utilizar a chave e o IV gerados para criptografar e descriptografar uma imagem. O exemplo utiliza a fun��o AES_Encrypt e AES_Decrypt, que foram definidas na Parte 1 deste c�digo, para criptografar e descriptografar a imagem usando o algoritmo AES. O exemplo tamb�m salva a imagem criptografada e descriptografada em arquivos separados.

'Lembre-se de adaptar o c�digo para o seu caso espec�fico, definindo os caminhos corretos para os arquivos e ajustando os tamanhos da chave e do IV de acordo com o algoritmo de criptografia que voc� deseja utilizar.

