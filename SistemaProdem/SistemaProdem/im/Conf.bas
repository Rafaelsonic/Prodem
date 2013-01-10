Attribute VB_Name = "Module1"
Const Server = "nivaldo"
Const Server_Arq = "nivaldo"
Const Unidade_Sistema = "G:"
Const Unidade_Arq = "P:"
Const Share_Sistema = "sistema"
Const Share_Arq = ""
Sub Main()
On Error GoTo Erro
'Atualiza o horario com a máquina servidora
Shell ("net time \\" & Server & " /set /y")

'Mapeia unidade de rede nescessária para o sistema
Shell ("net use " & Unidade_Sistema & " \\" & Server & "\" & Share_Sistema)

'Mapeia Unidade dos documentos
'Shell ("net use " & Unidade_Arq & " \\" & Server_Arq & "\" & Share_Arq)
Exit Sub
Erro:
MsgBox Err.Description, vbCritical, "Configuração : Informe o ero ocorrido"
End Sub
