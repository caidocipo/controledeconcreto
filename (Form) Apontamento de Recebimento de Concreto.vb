'Apontamento de Recebimento de Concreto

Private Sub UserForm_Activate()

    ' Verificar se a variável global "usernameGlobal" está vazia
    If usernameGlobal = "" Then
        ' Fechar o formulário atual
        Me.Hide
        ' Redirecionar para o formulário de login
        Login.Show
    End If
    
    TextBox5.Value = Format(Date, "dd/mm/yyyy")
    
    Dim ws0 As Worksheet
    Set ws0 = ThisWorkbook.Worksheets("recebConcreto")

    With ws0
        Dim lastRow0 As Long
        lastRow0 = .Cells(.Rows.Count, 1).End(xlUp).Row

        Dim i0 As Long
        Dim lista0 As String

        For i0 = lastRow0 To 2 Step -1 ' Percorrer de forma decrescente
            If Not IsEmpty(.Cells(i0, 1).Value) Then ' Verificar se o valor não está vazio
                lista0 = lista0 & .Cells(i0, 1).Value & vbCrLf
            End If
        Next i0

        ListBox1.Clear ' Limpa a lista atual do ListBox2
        ListBox1.List = Split(lista0, vbCrLf)
    End With
    
    ComboBox3.Clear
    ComboBox3.AddItem "Concreto auto-adensável fck: 15 MPa"
    ComboBox3.AddItem "Concreto auto-adensável fck: 20 MPa"
    ComboBox3.AddItem "Concreto auto-adensável fck: 25 MPa"
    ComboBox3.AddItem "Concreto usinado fck: 10 MPa"
    ComboBox3.AddItem "Concreto usinado fck: 20 MPa"
    ComboBox3.AddItem "Concreto usinado fck: 25 MPa"
    ComboBox3.AddItem "Concreto usinado fck: 30 MPa"
    ComboBox3.AddItem "Concreto manual fck: 10 MPa"
    ComboBox3.AddItem "Concreto manual fck:15 MPa"
    ComboBox3.AddItem "Argamassa usinada"
    ComboBox3.AddItem "Taxa de bomba"
    
    ComboBox4.Clear
    ComboBox4.AddItem "Contrato"
    ComboBox4.AddItem "Pedido"
        
End Sub

Private Sub ListBox1_AfterUpdate()
    ' Verifica se há uma linha selecionada na ListBox1
    If ListBox1.ListIndex <> -1 Then
        ' Obtém o valor da linha selecionada na ListBox1
        Dim selectedValue As String
        selectedValue = ListBox1.Value
        
        ' Define a planilha onde os dados estão armazenados
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets("recebConcreto")
        
        ' Procura o valor na coluna A da planilha e retorna a linha correspondente
        Dim searchRange As Range
        Set searchRange = ws.Columns("A")
        
        Dim foundCell As Range
        Set foundCell = searchRange.Find(What:=selectedValue, LookIn:=xlValues, LookAt:=xlWhole)
        
        ' Verifica se o valor foi encontrado
        If Not foundCell Is Nothing Then
            ' Obtém a linha correspondente ao valor encontrado
            Dim selectedRow As Long
            selectedRow = foundCell.Row
            
            ' Define o valor dos TextBoxes como os valores das colunas correspondentes da linha selecionada
            TextBox10.Value = Format(ws.Cells(selectedRow, "D").Value, "dd/mm/yyyy")
            TextBox11.Value = ws.Cells(selectedRow, "E").Value
            TextBox12.Value = ws.Cells(selectedRow, "F").Value
            TextBox13.Value = ws.Cells(selectedRow, "G").Value
            TextBox14.Value = ws.Cells(selectedRow, "H").Value
        End If
    End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentValue As String
    Dim validKeys As String
    Dim i As Integer
    
    currentValue = TextBox5.Text
    validKeys = "0123456789"
    
    ' Verificar se a tecla pressionada é um número válido
    If InStr(validKeys, Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
    ' Adicionar o caractere "/" após os dois primeiros dígitos e após os próximos dois dígitos
    If Len(currentValue) = 2 Or Len(currentValue) = 5 Then
        TextBox5.Text = currentValue & "/"
    End If
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 48 And KeyAscii <= 57 Then ' permite apenas números de 0 a 9
        If Len(Me.TextBox6.Value) = 2 Then ' adiciona dois pontos após a hora ser digitada
            Me.TextBox6.Value = Me.TextBox6.Value & ":"
            Me.TextBox6.SelStart = Len(Me.TextBox6.Value)
        ElseIf Len(Me.TextBox6.Value) = 5 Then ' limita a entrada em 5 caracteres para evitar que o usuário digite mais do que HH:MM
            Me.TextBox6.Value = Me.TextBox6.Value & ":"
            Me.TextBox6.SelStart = Len(Me.TextBox6.Value)
        End If
    ElseIf KeyAscii = 8 Then ' permite uso da tecla backspace
        If Len(Me.TextBox6.Value) = 3 Then ' remove o caracter ":" se ele existir
            Me.TextBox6.Value = Left(Me.TextBox6.Value, 2)
            Me.TextBox6.SelStart = Len(Me.TextBox6.Value)
        ElseIf Len(Me.TextBox6.Value) = 6 Then ' remove o caracter ":" se ele existir
            Me.TextBox6.Value = Left(Me.TextBox6.Value, 5)
            Me.TextBox6.SelStart = Len(Me.TextBox6.Value)
        End If
    Else
        KeyAscii = 0 ' desabilita a tecla pressionada caso não seja um número ou backspace
    End If
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii >= 48 And KeyAscii <= 57 Then ' permite apenas números de 0 a 9
        If Len(Me.TextBox7.Value) = 2 Then ' adiciona dois pontos após a hora ser digitada
            Me.TextBox7.Value = Me.TextBox7.Value & ":"
            Me.TextBox7.SelStart = Len(Me.TextBox7.Value)
        ElseIf Len(Me.TextBox7.Value) = 5 Then ' limita a entrada em 5 caracteres para evitar que o usuário digite mais do que HH:MM
            Me.TextBox7.Value = Me.TextBox7.Value & ":"
            Me.TextBox7.SelStart = Len(Me.TextBox7.Value)
        End If
    ElseIf KeyAscii = 8 Then ' permite uso da tecla backspace
        If Len(Me.TextBox7.Value) = 3 Then ' remove o caracter ":" se ele existir
            Me.TextBox7.Value = Left(Me.TextBox7.Value, 2)
            Me.TextBox7.SelStart = Len(Me.TextBox7.Value)
        ElseIf Len(Me.TextBox7.Value) = 6 Then ' remove o caracter ":" se ele existir
            Me.TextBox7.Value = Left(Me.TextBox7.Value, 5)
            Me.TextBox7.SelStart = Len(Me.TextBox7.Value)
        End If
    Else
        KeyAscii = 0 ' desabilita a tecla pressionada caso não seja um número ou backspace
    End If
End Sub

Private Sub CommandButton1_Click()
    ' Verifica se todos os campos foram preenchidos corretamente
        If ComboBox3.Value = "" Or ComboBox4.Value = "" Or TextBox1.Value = "" _
        Or TextBox2.Value = "" Or TextBox3.Value = "" Or TextBox4.Value = "" _
        Or TextBox5.Value = "" Or TextBox6.Value = "" Or TextBox7.Value = "" _
        Or TextBox8.Value = "" Or TextBox15.Value = "" Or TextBox16.Value = "" Or TextBox17.Value = "" Then
        MsgBox "Por favor, preencha todos os campos.", vbExclamation, "Campos obrigatórios"
        Exit Sub
    End If
    
    On Error GoTo Error_Handler 'Ativa a manipulação de erros
    
    'Definindo a planilha e a próxima linha disponível
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("recebConcreto")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row 'Detecta a última linha preenchida
    
    'Escrevendo os dados na planilha na ordem especificada
    ws.Cells(lastRow + 1, "B").Value = Application.WorksheetFunction.Max(ws.Range("B:B")) + 1
    ws.Cells(lastRow + 1, "C").Value = "(0184) Cidade Jardim"
    ws.Cells(lastRow + 1, "D").Value = Format(TextBox5.Value, "dd/mm/yyyy")
    ws.Cells(lastRow + 1, "E").Value = TextBox1.Value
    ws.Cells(lastRow + 1, "F").Value = TextBox2.Value
    ws.Cells(lastRow + 1, "G").Value = TextBox3.Value
    ws.Cells(lastRow + 1, "H").Value = TextBox4.Value
    ws.Cells(lastRow + 1, "I").Value = Format(TextBox6.Value, "hh:mm:ss")
    ws.Cells(lastRow + 1, "J").Value = Format(TextBox7.Value, "hh:mm:ss")
    ws.Cells(lastRow + 1, "L").Value = TextBox8.Value
    ws.Cells(lastRow + 1, "M").Value = Now()
    ws.Cells(lastRow + 1, "N").Value = ComboBox3.Value
    ws.Cells(lastRow + 1, "O").Value = TextBox15.Value
    ws.Cells(lastRow + 1, "P").Value = ComboBox4.Value
    ws.Cells(lastRow + 1, "Q").Value = TextBox16.Value
    ws.Cells(lastRow + 1, "R").Value = usernameGlobal
    ws.Cells(lastRow + 1, "T").Value = TextBox17.Value
    
    'Mensagem de sucesso e fechamento do formulário
    MsgBox "Cadastro realizado com sucesso!", vbInformation, "Sucesso"
    Me.Hide
    ConcRecebimento.Show ' Abre o formulário ConcRecebimento
    
    Exit Sub 'Sai do tratamento de erros
    
Error_Handler: 'Tratamento de erros
    'Mensagem de erro e cancelamento do cadastro
    MsgBox "Ocorreu um erro no cadastro, por favor, verifique suas informações.", vbCritical, "Erro"
    Exit Sub 'Sai do tratamento de erros
End Sub

Private Sub CommandButton3_Click()
    Me.Hide 'Fecha o formulário atual (ConcrecebConcreto)
    WelcomeScreen.Show 'Abre o formulário WelcomeScreen
End Sub

Private Sub UserForm_Terminate()
    ' Salvar a pasta de trabalho
    ThisWorkbook.Save
    
    ' Fechar o Excel
    Application.Quit
End Sub

Private Sub Label40_Click()
    ' Abrir o site no navegador padrão
    Shell "cmd /c start https://forms.gle/Lcf3bpwyfFuxoPPs6", vbHide
End Sub