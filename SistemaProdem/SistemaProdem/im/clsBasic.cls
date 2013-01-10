VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Function Monta_Where(campo As String, tipo As String, DADO As String) As String
Dim operador(1) As String
operador(0) = ";"
operador(1) = ":"
'operador() = ">"
'operador() = "<"


Select Case UCase$(tipo)
    Case "NUMERO"
        If InStr(1, DADO, operador(0)) > 0 Then
        
            Dim val As Integer
            Dim fim As Integer
            Dim cont As Integer
            fim = 1
            cont = 1
            Monta_Where = campo & " in ("
            For val = 1 To Len(DADO)
                Do While Mid$(DADO, fim, 1) <> ";"
                    If fim >= Len(DADO) Then
                        fim = Len(DADO) + 1
                        Exit Do
                    End If
                    fim = fim + 1
                Loop
                If cont = 1 Then
                    Monta_Where = Monta_Where & Mid$(DADO, val, fim - val)
                Else
                    Monta_Where = Monta_Where & "," & Mid$(DADO, val, fim - val)
                End If
                val = fim
                fim = fim + 1
                cont = cont + 1
            Next val
            Monta_Where = Monta_Where & ")"
            
        ElseIf InStr(1, DADO, operador(1)) > 0 Then
        
        
        
        Else
            Monta_Where = campo & "=" & Trim$(DADO)
        End If
    Case "TEXTO"
        If InStr(1, DADO, "*", vbTextCompare) > 0 Then
            Monta_Where = campo & " like '" & DADO & "'"
            Monta_Where = Replace(Monta_Where, "*", "%")
        Else
            Monta_Where = campo & "='" & Trim$(DADO) & "'"
        End If
    Case "DATA"
        ' INVERTER A DATA PADRAO AMERICANO
        Monta_Where = campo & "=#" & Format(Trim$(DADO), "mm/dd/yyyy") & "#"
End Select
End Function

Public Sub Menssagem(msg As String, titulo As String, ico As Integer, tipo As Integer)
'tipo 0 - msgbox
'tipo 1 - barra de status
MsgBox msg, ico, titulo
End Sub

