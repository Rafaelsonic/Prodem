Attribute VB_Name = "im_Download"
Option Explicit
'Funções para baixar arquivos

Public Function download(arqUrl As String, ArqTemp As String) As Boolean
Dim Arquivo() As Byte
Open ArqTemp For Binary Access Write As #1
    Arquivo() = im_webxml.Inet.OpenURL(arqUrl, icByteArray)
    Put #1, , Arquivo()
Close #1
download = True
Exit Function
trata_erro:
download = False
End Function


