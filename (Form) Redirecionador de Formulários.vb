Private Sub UserForm_Initialize()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If

End Sub

Private Sub CommandButton1_Click()
    'Esconder formulário atual
    Me.Hide
    
    'Mostrar formulário
    ConcRecebimento.Show
End Sub

Private Sub CommandButton2_Click()
    'Esconder formulário atual
    Me.Hide
    
    'Mostrar formulário
    ConcConsumo.Show
End Sub

Private Sub CommandButton3_Click()
    'Esconder formulário atual
    Me.Hide
    
    'Mostrar formulário
    ConcResist.Show
End Sub

Private Sub CommandButton4_Click()
    MsgBox "A funcionalidade ainda não foi implementada!", vbExclamation
End Sub

Private Sub CommandButton5_Click()
    MsgBox "A funcionalidade ainda não foi implementada!", vbExclamation
End Sub

Private Sub CommandButton6_Click()
    MsgBox "A funcionalidade ainda não foi implementada!", vbExclamation
End Sub

Private Sub CommandButton7_Click()
    Application.Visible = True ' torna o Excel visível novamente
    ThisWorkbook.Activate ' ativa a janela da pasta de trabalho atual
    ThisWorkbook.Windows(1).Visible = True ' torna a janela da pasta de trabalho atual visível
    ThisWorkbook.Windows(1).Activate ' ativa a janela da pasta de trabalho atual
    
    ' ativa a edição de planilhas
    Application.EnableEvents = True
    Application.Interactive = True
    Application.ScreenUpdating = True
    
    'Esconder formulário atual
    Me.Hide
End Sub

Private Sub CommandButton8_Click()
    'Esconder formulário atual
    Me.Hide
    
    'Mostrar formulário
    EquipLista.Show
End Sub

Private Sub CommandButton9_Click()
    MsgBox "A funcionalidade ainda não foi implementada!", vbExclamation
End Sub

Private Sub CommandButton10_Click()
    MsgBox "A funcionalidade ainda não foi implementada!", vbExclamation
End Sub

Private Sub CommandButton11_Click()
    'Esconder formulário atual
    Me.Hide
    'Limpar a variável de username
    Username.usernameGlobal = ""
    'Mostrar formulário
    Login.Show
End Sub

Private Sub UserForm_Terminate()
    ' Salvar a pasta de trabalho
    ThisWorkbook.Save
    
    ' Fechar o Excel
    Application.Quit
End Sub

Private Sub Label25_Click()
    ' Abrir o site no navegador padrão
    Shell "cmd /c start https://forms.gle/Lcf3bpwyfFuxoPPs6", vbHide
End Sub
