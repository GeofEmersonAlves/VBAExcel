VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmJogoVelha 
   Caption         =   "Jogo da Velha"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   OleObjectBlob   =   "frmJogoVelha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmJogoVelha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Jogo(1 To 3, 1 To 3) As Integer                             'Armazena as jogadas (1:Usuário, -1:Excel, 0:Vazia)
Dim Linha(1 To 3), Coluna(1 To 3), Diagonal(1 To 2) As Integer  'Soma das linhas, colunas e diagonais do jogo
Dim bytNumJogada As Integer                                     'Quantidade de jogadas efetuadas
    


Private Sub cmdIniciar_Click()
    Dim bytContAuxLin, bytContAuxCol As Byte
    Dim strNomeBotao As String
    
    If cmdIniciar.Caption = "Iniciar" Or cmdIniciar.Caption = "Novo jogo" Then
        For bytContAuxLin = 1 To 3 'Limpa a jogada anterior e habilita as jogadas
            For bytContAuxCol = 1 To 3
                strNomeBotao = "lbl" & Trim(Str(bytContAuxLin)) & Trim(Str(bytContAuxCol))
                Jogo(bytContAuxLin, bytContAuxCol) = 0
                frmJogoVelha.Controls(strNomeBotao).Enabled = True
                frmJogoVelha.Controls(strNomeBotao).Caption = ""
                frmJogoVelha.Controls(strNomeBotao).SpecialEffect = 1
            Next
        Next
        cmdIniciar.Caption = "Fechar"
        frmSimbolo.Enabled = False
        bytNumJogada = 0
    Else
        If MsgBox("Tem certeza que deseja encerrar o jogo", vbQuestion + vbYesNo) = vbYes Then
            Unload frmJogoVelha
        End If
    End If
        
        
End Sub

Private Sub lbl11_Click()
    MostraJogada 1, 1, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 1, 1
    End If
End Sub

Private Sub lbl12_Click()
    MostraJogada 1, 2, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 1, 2
    End If
End Sub

Private Sub lbl13_Click()
    MostraJogada 1, 3, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 1, 3
    End If
End Sub

Private Sub lbl21_Click()
    MostraJogada 2, 1, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 2, 1
    End If
End Sub


Private Sub lbl22_Click()
    MostraJogada 2, 2, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 2, 2
    End If
End Sub

Private Sub lbl23_Click()
    MostraJogada 2, 3, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 2, 3
    End If
End Sub

Private Sub lbl31_Click()
    MostraJogada 3, 1, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 3, 1
    End If
End Sub

Private Sub lbl32_Click()
    MostraJogada 3, 2, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 3, 2
    End If
End Sub

Private Sub lbl33_Click()
    MostraJogada 3, 3, IIf(optX.Value, "X", "0"), 1
    If Ganhou Then    'Se o usuário ganhou mostra uma mensagem
        MsgBox "P A R A B É N S! " & Chr(10) & Chr(13) & "Você ganhou", vbInformation + vbOKOnly
        FimDoJogo
    Else 'Efetua uma jogada do Excel
        DoJogadaExcel 3, 3
    End If
End Sub

'Mostra a jogada
Private Sub MostraJogada(ByVal bytLin As Byte, ByVal bytCol As Byte, stJogador As String, intValJogada As Integer)
    Dim stNomeBotao As String
    
    If Jogo(bytLin, bytCol) = 0 Then    'Verifica se esta posição já esta preenchida
        stNomeBotao = "lbl" & Trim(Str(bytLin)) & Trim(Str(bytCol))
        frmJogoVelha.Controls(stNomeBotao).Caption = stJogador
        frmJogoVelha.Controls(stNomeBotao).SpecialEffect = 2
        Jogo(bytLin, bytCol) = intValJogada
        bytNumJogada = bytNumJogada + 1
    End If
    
End Sub

'Retorna True se o usuário ganhou
Private Function Ganhou() As Boolean
    Dim blGanhou As Boolean
    Dim bytCtAux As Byte
     
    blGanhou = False
    
    'Soma as posições jogadas, para verificar se o usuário ganhou, se o Excel ganha nesta jogada ou se o usuário pode ganhar na próxima jogada
    For bytCtAux = 1 To 3
        Linha(bytCtAux) = Jogo(bytCtAux, 1) + Jogo(bytCtAux, 2) + Jogo(bytCtAux, 3)
    Next
    For bytCtAux = 1 To 3
        Coluna(bytCtAux) = Jogo(1, bytCtAux) + Jogo(2, bytCtAux) + Jogo(3, bytCtAux)
    Next
    Diagonal(1) = Jogo(1, 1) + Jogo(2, 2) + Jogo(3, 3)
    Diagonal(2) = Jogo(1, 3) + Jogo(2, 2) + Jogo(3, 1)
    
    'Verifica se o usuario ganhou
    bytCtAux = 1
    Do
        If Linha(bytCtAux) = 3 Then
            blGanhou = True
        ElseIf Coluna(bytCtAux) = 3 Then
            blGanhou = True
        ElseIf bytCtAux < 3 Then
            If Diagonal(bytCtAux) = 3 Then
                blGanhou = True
            End If
        End If
        bytCtAux = bytCtAux + 1
    Loop While Not blGanhou And bytCtAux <= 3
    
    
   Ganhou = blGanhou
End Function


'Efetua a jogada do excel
Private Sub DoJogadaExcel(bytLin As Byte, bytCol As Byte)
    Dim blGanhaProx, blGanhei As Boolean
    Dim stNomeBotao As String
    Dim bytCtL, bytCtC As Byte 'Contadores para linha, coluna e contador auxiliar

    blGanhaProx = False
    blGanhei = False
    
    'Verifica se é possível ganhar nesta jogada e efetua esta jogada
    bytCtAux = 1
    Do
        If Linha(bytCtAux) = -2 Then             'Verifica ganha em alguma linha
            blGanhei = True
            bytCtL = bytCtAux
            bytCtC = 1
            Do While Jogo(bytCtL, bytCtC) <> 0
                bytCtC = bytCtC + 1
            Loop
        ElseIf Coluna(bytCtAux) = -2 Then        'Verifica se ganha em alguma coluna
            blGanhei = True
            bytCtC = bytCtAux
            bytCtL = 1
            Do While Jogo(bytCtL, bytCtC) <> 0
                bytCtL = bytCtL + 1
            Loop
        ElseIf bytCtAux < 3 Then                'Verifica se ganha em alguma diagonal
            If Diagonal(bytCtAux) = -2 Then
                blGanhei = True
                If bytCtAux = 1 Then    'Diagonal 1
                    bytCtL = bytCtAux
                    bytCtC = bytCtAux
                    Do While Jogo(bytCtL, bytCtC) <> 0
                        bytCtL = bytCtL + 1
                        bytCtC = bytCtC + 1
                    Loop
                
                Else                    'Diagonal 2
                    bytCtL = 1
                    bytCtC = 3
                    Do While Jogo(bytCtL, bytCtC) <> 0
                        bytCtL = bytCtL + 1
                        bytCtC = bytCtC - 1
                    Loop
                End If
            End If
        End If
        bytCtAux = bytCtAux + 1
    Loop While Not blGanhei And bytCtAux <= 3

    If blGanhei Then
        MostraJogada bytCtL, bytCtC, IIf(optX.Value, "0", "X"), -1
        MsgBox "HI HI HI! " & Chr(10) & Chr(13) & "HE HE HE!" & Chr(10) & Chr(13) & "Você perdeu", vbInformation + vbOKOnly
        FimDoJogo
    Else    'Se não é possível ganhar nesta jogada, verifica se o usuário ganhará na próxima
        bytCtAux = 1
        Do
            If Linha(bytCtAux) = 2 Then             'Verifica se o usuario ganha em alguma linha
                blGanhaProx = True
                bytCtL = bytCtAux
                bytCtC = 1
                Do While Jogo(bytCtL, bytCtC) <> 0
                    bytCtC = bytCtC + 1
                Loop
            ElseIf Coluna(bytCtAux) = 2 Then        'Verifica se o usuario ganha em alguma coluna
                blGanhaProx = True
                bytCtC = bytCtAux
                bytCtL = 1
                Do While Jogo(bytCtL, bytCtC) <> 0
                    bytCtL = bytCtL + 1
                Loop
            ElseIf bytCtAux < 3 Then                'Verifica se o usuario ganha em alguma diagonal
                If Diagonal(bytCtAux) = 2 Then
                    blGanhaProx = True
                    If bytCtAux = 1 Then    'Diagonal 1
                        bytCtL = bytCtAux
                        bytCtC = bytCtAux
                        Do While Jogo(bytCtL, bytCtC) <> 0
                            bytCtL = bytCtL + 1
                            bytCtC = bytCtC + 1
                        Loop
    
                    Else                    'Diagonal 2
                        bytCtL = 1
                        bytCtC = 3
                        Do While Jogo(bytCtL, bytCtC) <> 0
                            bytCtL = bytCtL + 1
                            bytCtC = bytCtC - 1
                        Loop
                    End If
                End If
            End If
            bytCtAux = bytCtAux + 1
        Loop While Not blGanhaProx And bytCtAux <= 3
           
        If blGanhaProx Then 'Mostra a jogada
            MostraJogada bytCtL, bytCtC, IIf(optX.Value, "0", "X"), -1
            MsgBox "Pensou e ia me pegar é!", vbExclamation + vbOKOnly
        Else
            'Se o usuário nao ganhara na proxima jogada verifica qual é a melhor jogada agora para o Excel
            If bytNumJogada = 1 Then    'Verifica se é a primeira jogada
                If bytLin = 2 And bytCol = 2 Then       'Se a primeira jogada for no centro joga no canto
                    bytCtL = 1
                    bytCtC = 3
                ElseIf (bytLin + bytCol) Mod 2 = 0 Then   'Se a primeira jogada for num cantos joga no canto oposto
                    bytCtL = IIf(bytLin = 1, 3, 1)
                    bytCtC = IIf(bytCol = 1, 3, 1)
                Else                                      'Senão joga num canto longe da jogada
                    If bytLin = 1 Then
                        bytCtL = 3
                        bytCtC = 3
                    ElseIf bytLin = 3 Then
                        bytCtL = 1
                        bytCtC = 1
                    ElseIf bytCol = 1 Then
                        bytCtL = 1
                        bytCtC = 3
                    Else
                        bytCtL = 1
                        bytCtC = 1
                    End If
                End If
            Else        'Se não for a primeira jogada
                'Verifica se tem como fechar o usuário
                If Linha(1) = -1 And Coluna(1) = -1 And Jogo(1, 1) = 0 Then
                    bytCtL = 1
                    bytCtC = 1
                ElseIf (Diagonal(1) = -1 And Jogo(1, 1) = 0 And Jogo(3, 3) = 0) Or (Linha(1) = -1 And Jogo(1, 1) = 0 And Jogo(1, 2) = 0) Then
                    bytCtL = 1
                    bytCtC = 1
                ElseIf Diagonal(1) = -1 And Jogo(1, 1) = 0 And Jogo(3, 1) = 0 Then
                    bytCtL = 2
                    bytCtC = 2
                ElseIf Coluna(3) = -1 And Diagonal(2) = -1 And Jogo(3, 1) = 0 Then
                    bytCtL = 3
                    bytCtC = 1
                ElseIf Coluna(1) = -1 And Jogo(1, 1) = -1 And (Jogo(2, 1) + Jogo(3, 1) = 0) Then
                    bytCtL = 2
                    bytCtC = 1
                ElseIf Coluna(2) = -1 And Jogo(1, 2) = -1 And (Jogo(2, 2) + Jogo(3, 2) = 0) Then
                    bytCtL = 2
                    bytCtC = 2
                ElseIf Coluna(3) = -1 And Jogo(1, 3) = -1 And (Jogo(2, 3) + Jogo(3, 3) = 0) Then
                    bytCtL = 2
                    bytCtC = 3
                ElseIf Jogo(2, 2) = 0 Then 'Verifica se o centro esta vazio
                    bytCtL = 2
                    bytCtC = 2
                    
                Else    'Senao, joga na primeira casa vazia
                    bytCtL = 1
                    bytCtC = 1
                    Do
                        If bytCtL = 3 Then
                            bytCtL = 1
                            bytCtC = bytCtC + 1
                        Else
                            bytCtL = bytCtL + 1
                        End If
                        
                    Loop While Jogo(bytCtL, bytCtC) <> 0 And bytCtC < 3 And bytCtL < 3
                End If
            End If
            
            If bytCtC <= 3 And bytCtL <= 3 And bytNumJogada < 9 Then
                MostraJogada bytCtL, bytCtC, IIf(optX.Value, "0", "X"), -1
            Else
                If bytNumJogada = 9 Then    'Deu velha
                    MsgBox "D E U   V É L H A !!!" & Chr(10) & Chr(13) & "Ninguem ganhou", vbExclamation + vbOKOnly
                    FimDoJogo
                End If
            End If
        End If
    End If

End Sub

Private Sub FimDoJogo()
    Dim bytContAuxLin, bytContAuxCol As Byte
    Dim strNomeBotao As String
    
    For bytContAuxLin = 1 To 3 'Limpa a jogada anterior e habilita as jogadas
        For bytContAuxCol = 1 To 3
            strNomeBotao = "lbl" & Trim(Str(bytContAuxLin)) & Trim(Str(bytContAuxCol))
            frmJogoVelha.Controls(strNomeBotao).Enabled = False
        Next
    Next
    cmdIniciar.Caption = "Novo jogo"
End Sub

