Attribute VB_Name = "Basic"
Option Explicit
' funcoes utilizadas por todo o sistema
Public Banco_path As String, Banco_path_CEP As String

Public Banco_senha As String
Public ocnBanco As New ADODB.Connection

Public SomenteSitio As Boolean


'25/07/12
Public UsuarioID As Integer 'codigo do usuario para montar as funcionalidades

'funcao sleep
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function Open_Banco() As Boolean
' procedimento de abertura do banco de dados
On Error GoTo erro
'ocnBanco.Open "dsn=db_wf"
'ocnBanco.Properties("Jet OLEDB:Database Password") = "sistema"
ocnBanco.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Banco_path & ";Jet OLEDB:Database Password=" & Banco_senha

'ocnBanco.Open "Driver={SQL Server};Server=.;Uid=sa;Pwd=;Database=wf"


Open_Banco = True
Exit Function
erro:
Open_Banco = False
MsgBox Err.Description
End Function

Function Monta_SQL(campo As String, tipo As String, DADO As String) As String
Dim x As Integer
Select Case UCase$(tipo)
    Case "NUMERO"
        Monta_SQL = campo & "=" & Trim$(DADO)
    Case "TEXTO"
        If InStr(1, DADO, "*", vbTextCompare) > 0 Then
            Monta_SQL = campo & " like '" & DADO & "'"
            Monta_SQL = Replace(Monta_SQL, "*", "%")
        Else
            Monta_SQL = campo & "='" & Trim$(DADO) & "'"
        End If
    Case "DATA"
        ' INVERTER A DATA PADRAO AMERICANO
        Monta_SQL = campo & "=#" & Format(Trim$(DADO), "mm/dd/yyyy") & "#"
End Select
End Function

Sub Main()
SomenteSitio = False
Banco_senha = "sistema"

If App.Path <> "E:\Work Fire\Sistema\pf" Then
'If App.Path <> "C:\temp\wf\Sistema\pf" Then
    'Em modo executável
    Banco_path = App.Path & "\banco\data.mdb"
    Banco_path_CEP = App.Path & "\banco\cepbr.mdb"
Else
    'Em modo debug
    'Banco_path = "C:\temp\wf\" & "\banco\1bd_wf.mdb" 'para usar em modo debug
    'Banco_path_CEP = "C:\temp\wf\" & "\banco\cepbr.mdb" 'caminho
    
    Banco_path = "E:\Work Fire\" & "\banco\1bd_wf.sku" 'para usar em modo debug
    Banco_path_CEP = "E:\Work Fire\" & "\banco\cepbr.mdb" 'caminho
End If

If Open_Banco = True Then
mdi.Show
Else
MsgBox "Não foi possível abrir o Banco de Dados"
End
End If
End Sub





Public Function ChecaInscrE(pUF As String, pInscr As String) As Boolean
    ChecaInscrE = False
    Dim StrBase              As String
    Dim StrBase2             As String
    Dim StrOrigem            As String
    Dim StrDigito1           As String
    Dim StrDigito2           As String
    Dim intPos               As Integer
    Dim intValor             As Integer
    Dim intSoma              As Integer
    Dim intResto             As Integer
    Dim intNumero            As Integer
    Dim intPeso              As Integer
    Dim intDig               As Integer
    
    StrBase = ""
    StrBase2 = ""
    StrOrigem = ""
    If Trim$(pInscr) = "ISENTO" Then
        ChecaInscrE = True
        Exit Function
    End If
    For intPos = 1 To Len(Trim$(pInscr))
         If InStr(1, "0123456789P", Mid$(pInscr, intPos, 1), vbTextCompare) > 0 Then
             StrOrigem = StrOrigem & Mid$(pInscr, intPos, 1)
         End If
    Next
    Select Case pUF
      Case "AC"    ' Acre
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "01" And Mid$(StrBase, 3, 2) <> "00" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "AL"    ' Alagoas
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "24" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intSoma = intSoma * 10
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto = 10, "0", Str$(intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "AM"    ' Amazonas
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           If intSoma < 11 Then
               StrDigito1 = Right$(Str$(11 - intSoma), 1)
           Else
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           End If
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "AP"    ' Amapa
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intPeso = 0
           intDig = 0
           If Left$(StrBase, 2) = "03" Then
               intNumero = val(Left$(StrBase, 8))
               If intNumero >= 3000001 And _
                  intNumero <= 3017000 Then
                   intPeso = 5
                   intDig = 0
               ElseIf intNumero >= 3017001 And _
                      intNumero <= 3019022 Then
                   intPeso = 9
                   intDig = 1
               ElseIf intNumero >= 3019023 Then
                   intPeso = 0
                   intDig = 0
               End If
               intSoma = intPeso
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               intValor = 11 - intResto
               If intValor = 10 Then
                   intValor = 0
               ElseIf intValor = 11 Then
                   intValor = intDig
               End If
               StrDigito1 = Right$(Str$(intValor), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "BA"    ' Bahia
           StrBase = Left$(Trim$(StrOrigem) & "00000000", 8)
           If InStr(1, "0123458", Left$(StrBase, 1), vbTextCompare) > 0 Then
               intSoma = 0
               For intPos = 1 To 6
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (8 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 10
               StrDigito2 = Right$(IIf(intResto = 0, "0", Str$(10 - intResto)), 1)
               StrBase2 = Left$(StrBase, 6) & StrDigito2
               intSoma = 0
               For intPos = 1 To 7
                    intValor = val(Mid$(StrBase2, intPos, 1))
                    intValor = intValor * (9 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 10
               StrDigito1 = Right$(IIf(intResto = 0, "0", Str$(10 - intResto)), 1)
           Else
               intSoma = 0
               For intPos = 1 To 6
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (8 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito2 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 6) & StrDigito2
               intSoma = 0
               For intPos = 1 To 7
                    intValor = val(Mid$(StrBase2, intPos, 1))
                    intValor = intValor * (9 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           End If
           StrBase2 = Left$(StrBase, 6) & StrDigito1 & StrDigito2
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "CE"    ' Ceara
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           intValor = 11 - intResto
           If intValor > 9 Then
               intValor = 0
           End If
           StrDigito1 = Right$(Str$(intValor), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "DF"    ' DiStr$ito Federal
           StrBase = Left$(Trim$(StrOrigem) & "0000000000000", 13)
           If Left$(StrBase, 3) = "073" Then
               intSoma = 0
               intPeso = 2
               For intPos = 11 To 1 Step -1
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 11) & StrDigito1
               intSoma = 0
               intPeso = 2
               For intPos = 12 To 1 Step -1
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
               Next
               intResto = intSoma Mod 11
               StrDigito2 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 12) & StrDigito2
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "ES"    ' Espirito Santo
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "GO"    ' Goias
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If InStr(1, "10,11,15", Left$(StrBase, 2), vbTextCompare) > 0 Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               If intResto = 0 Then
                   StrDigito1 = "0"
               ElseIf intResto = 1 Then
                   intNumero = val(Left$(StrBase, 8))
                   StrDigito1 = Right$(IIf(intNumero >= 10103105 And intNumero <= 10119997, "1", "0"), 1)
               Else
                   StrDigito1 = Right$(Str$(11 - intResto), 1)
               End If
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "MA"    ' Maranhão
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "12" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "MT"    ' Mato Grosso
           StrBase = Left$(Trim$(StrOrigem) & "0000000000", 10)
           intSoma = 0
           intPeso = 2
           For intPos = 10 To 1 Step -1
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 9 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 10) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "MS"    ' Mato Grosso do Sul
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "28" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "MG"    ' Minas Gerais
           StrBase = Left$(Trim$(StrOrigem) & "0000000000000", 13)
           StrBase2 = Left$(StrBase, 3) & "0" & Mid$(StrBase, 4, 8)
           intNumero = 2
           For intPos = 1 To 12
                intValor = val(Mid$(StrBase2, intPos, 1))
                intNumero = IIf(intNumero = 2, 1, 2)
                intValor = intValor * intNumero
                If intValor > 9 Then
                    StrDigito1 = Format(intValor, "00")
                    intValor = val(Left$(StrDigito1, 1)) + _
                               val(Right$(StrDigito1, 1))
                End If
                intSoma = intSoma + intValor
           Next
           intValor = intSoma
           While Right$(Format(intValor, "000"), 1) <> "0"
               intValor = intValor + 1
           Wend
           StrDigito1 = Right$(Format(intValor - intSoma, "00"), 1)
           StrBase2 = Left$(StrBase, 11) & StrDigito1
           intSoma = 0
           intPeso = 2
           For intPos = 12 To 1 Step -1
                intValor = val(Mid$(StrBase2, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 11 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           StrDigito2 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = StrBase2 & StrDigito2
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "PA"    ' Para
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "15" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "PB"    ' Paraiba
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           intValor = 11 - intResto
           If intValor > 9 Then
               intValor = 0
           End If
           StrDigito1 = Right$(Str$(intValor), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "PE"    ' Pernambuco
           StrBase = Left$(Trim$(StrOrigem) & "00000000000000", 14)
           intSoma = 0
           intPeso = 2
           For intPos = 13 To 1 Step -1
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 9 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           intValor = 11 - intResto
           If intValor > 9 Then
               intValor = intValor - 10
           End If
           StrDigito1 = Right$(Str$(intValor), 1)
           StrBase2 = Left$(StrBase, 13) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "PI"    ' Piaui
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "PR"    ' Parana
           StrBase = Left$(Trim$(StrOrigem) & "0000000000", 10)
           intSoma = 0
           intPeso = 2
           For intPos = 8 To 1 Step -1
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           intSoma = 0
           intPeso = 2
           For intPos = 9 To 1 Step -1
                intValor = val(Mid$(StrBase2, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           StrDigito2 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = StrBase2 & StrDigito2
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "RJ"    ' Rio de Janeiro
           StrBase = Left$(Trim$(StrOrigem) & "00000000", 8)
           intSoma = 0
           intPeso = 2
           For intPos = 7 To 1 Step -1
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * intPeso
                intSoma = intSoma + intValor
                intPeso = intPeso + 1
                If intPeso > 7 Then
                    intPeso = 2
                End If
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 7) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "RN"    ' Rio Grande do Norte
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "20" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intSoma = intSoma * 10
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto > 9, "0", Str$(intResto)), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "RO"    ' Rondonia
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           StrBase2 = Mid$(StrBase, 4, 5)
           intSoma = 0
           For intPos = 1 To 5
                intValor = val(Mid$(StrBase2, intPos, 1))
                intValor = intValor * (7 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           intValor = 11 - intResto
           If intValor > 9 Then
               intValor = intValor - 10
           End If
           StrDigito1 = Right$(Str$(intValor), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "RR"    ' Roraima
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           If Left$(StrBase, 2) = "24" Then
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 9
               StrDigito1 = Right$(Str$(intResto), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "RS"    ' Rio Grande do Sul
           StrBase = Left$(Trim$(StrOrigem) & "0000000000", 10)
           intNumero = val(Left$(StrBase, 3))
           If intNumero > 0 And intNumero < 468 Then
               intSoma = 0
               intPeso = 2
               For intPos = 9 To 1 Step -1
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 9 Then
                        intPeso = 2
                    End If
               Next
               intResto = intSoma Mod 11
               intValor = 11 - intResto
               If intValor > 9 Then
                   intValor = 0
               End If
               StrDigito1 = Right$(Str$(intValor), 1)
               StrBase2 = Left$(StrBase, 9) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
      Case "SC"    ' Santa Catarina
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "SE"    ' Sergipe
           StrBase = Left$(Trim$(StrOrigem) & "000000000", 9)
           intSoma = 0
           For intPos = 1 To 8
                intValor = val(Mid$(StrBase, intPos, 1))
                intValor = intValor * (10 - intPos)
                intSoma = intSoma + intValor
           Next
           intResto = intSoma Mod 11
           intValor = 11 - intResto
           If intValor > 9 Then
               intValor = 0
           End If
           StrDigito1 = Right$(Str$(intValor), 1)
           StrBase2 = Left$(StrBase, 8) & StrDigito1
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "SP"    ' São Paulo
           If Left$(StrOrigem, 1) = "P" Then
               StrBase = Left$(Trim$(StrOrigem) & "0000000000000", 13)
               StrBase2 = Mid$(StrBase, 2, 8)
               intSoma = 0
               intPeso = 1
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso = 2 Then
                        intPeso = 3
                    End If
                    If intPeso = 9 Then
                        intPeso = 10
                    End If
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(Str$(intResto), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1 & Mid$(StrBase, 11, 3)
           Else
               StrBase = Left$(Trim$(StrOrigem) & "000000000000", 12)
               intSoma = 0
               intPeso = 1
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso = 2 Then
                        intPeso = 3
                    End If
                    If intPeso = 9 Then
                        intPeso = 10
                    End If
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(Str$(intResto), 1)
               StrBase2 = Left$(StrBase, 8) & StrDigito1 & Mid$(StrBase, 10, 2)
               intSoma = 0
               intPeso = 2
               For intPos = 11 To 1 Step -1
                    intValor = val(Mid$(StrBase, intPos, 1))
                    intValor = intValor * intPeso
                    intSoma = intSoma + intValor
                    intPeso = intPeso + 1
                    If intPeso > 10 Then
                        intPeso = 2
                    End If
               Next
               intResto = intSoma Mod 11
               StrDigito2 = Right$(Str$(intResto), 1)
               StrBase2 = StrBase2 & StrDigito2
           End If
           If StrBase2 = StrOrigem Then
               ChecaInscrE = True
           End If
      Case "TO"    ' Tocantins
           StrBase = Left$(Trim$(StrOrigem) & "00000000000", 11)
           If InStr(1, "01,02,03,99", Mid$(StrBase, 3, 2), vbTextCompare) > 0 Then
               StrBase2 = Left$(StrBase, 2) & Mid$(StrBase, 5, 6)
               intSoma = 0
               For intPos = 1 To 8
                    intValor = val(Mid$(StrBase2, intPos, 1))
                    intValor = intValor * (10 - intPos)
                    intSoma = intSoma + intValor
               Next
               intResto = intSoma Mod 11
               StrDigito1 = Right$(IIf(intResto < 2, "0", Str$(11 - intResto)), 1)
               StrBase2 = Left$(StrBase, 10) & StrDigito1
               If StrBase2 = StrOrigem Then
                   ChecaInscrE = True
               End If
           End If
    End Select
    
End Function


Function Impressora() As Boolean
On Error GoTo erro
'With pf_cer_rela.ctr
With mdi.ctr
    .Min = 0
    .Max = 999
    .FromPage = 1
    .ToPage = 999
    .CancelError = True
    .PrinterDefault = True
    .Flags = cdlPDNoSelection Or cdlPDPageNums Or cdlPDAllPages Or cdlPDHidePrintToFile
    .ShowPrinter
End With
    On Error GoTo 0
    Screen.MousePointer = vbArrow
    Impressora = True
    Exit Function
erro:
    On Error GoTo 0
    Screen.MousePointer = vbArrow
    MsgBox "Impressão Cancelada!", vbCritical, "Impressora"
    Impressora = False
    Exit Function
End Function






Function CNPJ(RecebeCNPJ As String) As Boolean

Dim Numero(14) As Integer, soma As Integer, resultado1 As Integer, resultado2 As Integer
Dim x As Integer
Dim s As String, ch As String

s = ""
For x = 1 To Len(RecebeCNPJ)
ch = Mid$(RecebeCNPJ, x, 1)
If Asc(ch) >= 48 And Asc(ch) <= 57 Then
s = s & ch
End If
Next
RecebeCNPJ = s

If Len(RecebeCNPJ) <> 14 Then
    MsgBox "Obrigatório 14 dígitos", vbCritical, "function CNPJ"
ElseIf RecebeCNPJ = "00000000000000" Then
    CNPJ = True
    Exit Function
Else

Numero(1) = CInt(Mid$(RecebeCNPJ, 1, 1))
Numero(2) = CInt(Mid$(RecebeCNPJ, 2, 1))
Numero(3) = CInt(Mid$(RecebeCNPJ, 3, 1))
Numero(4) = CInt(Mid$(RecebeCNPJ, 4, 1))
Numero(5) = CInt(Mid$(RecebeCNPJ, 5, 1))
Numero(6) = CInt(Mid$(RecebeCNPJ, 6, 1))
Numero(7) = CInt(Mid$(RecebeCNPJ, 7, 1))
Numero(8) = CInt(Mid$(RecebeCNPJ, 8, 1))
Numero(9) = CInt(Mid$(RecebeCNPJ, 9, 1))
Numero(10) = CInt(Mid$(RecebeCNPJ, 10, 1))
Numero(11) = CInt(Mid$(RecebeCNPJ, 11, 1))
Numero(12) = CInt(Mid$(RecebeCNPJ, 12, 1))
Numero(13) = CInt(Mid$(RecebeCNPJ, 13, 1))
Numero(14) = CInt(Mid$(RecebeCNPJ, 14, 1))

soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2

soma = soma - (11 * (Int(soma / 11)))

If soma = 0 Or soma = 1 Then
resultado1 = 0
Else
resultado1 = 11 - soma
End If
If resultado1 = Numero(13) Then
soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
soma = soma - (11 * (Int(soma / 11)))
If soma = 0 Or soma = 1 Then
resultado2 = 0
Else
resultado2 = 11 - soma
End If
If resultado2 = Numero(14) Then
CNPJ = True
Else
CNPJ = False
End If
Else
CNPJ = False
End If
End If
End Function

Sub BOTDisable()
mdi.Toolbar.Buttons(1).Enabled = False
mdi.Toolbar.Buttons(3).Enabled = False
mdi.Toolbar.Buttons(2).Enabled = False
mdi.Toolbar.Buttons(4).Enabled = False
mdi.Toolbar.Buttons(5).Enabled = False
mdi.Toolbar.Buttons(6).Enabled = False
mdi.Toolbar.Buttons(7).Enabled = False
mdi.Toolbar.Buttons(8).Enabled = False
mdi.Toolbar.Buttons(9).Enabled = False: mdi.Toolbar.Buttons(10).Enabled = False
mdi.Toolbar.Buttons(10).Enabled = False
End Sub

Function Data(vData As Date)
'Transforma a data no padrão americano para gravar no banco de dados
Data = Format(vData, "mm/dd/yyyy")
End Function
