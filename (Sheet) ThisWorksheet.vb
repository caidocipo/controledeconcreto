Private Sub Workbook_Open()
    'Garante que a planilha não seja mostrada ao usuário ao abrir
    Application.Visible = False ' oculta a janela do Excel
    Login.Show ' mostra o formulário Login
End Sub