Imports MySql.Data.MySqlClient
Imports Trabalho_Final_de_OBD

Public Class Form1
    Dim conn As New MySqlConnection
    Dim myCommand As New MySqlCommand
    Dim myAdapter As New MySqlDataAdapter
    Dim myData As New DataTable
    Dim SQL As String

    Public banco_host As String = ""
    Public banco_user As String = ""
    Public banco_pass As String = ""
    Public banco_banco As String = ""

    Private connString As String = "host=" & banco_host & ";Port=3306; user id=" & banco_user & "; password=" & banco_pass & "; database=" & banco_banco

    '' AO INICIAR:
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabPage2.Enabled = False
        TabPage3.Enabled = False
        TabPage4.Enabled = False
        TabPage5.Enabled = False

    End Sub

    '' BOTAO CONECTAR:
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        banco_host = TextBox1.Text
        banco_banco = TextBox2.Text
        banco_user = TextBox3.Text
        banco_pass = TextBox4.Text
        connString = "host=" & banco_host & ";Port=3306; user id=" & banco_user & "; password=" & banco_pass & "; database=" & banco_banco

        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand("show tables;", myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()
            reader = myCommand.ExecuteReader()
            If reader.HasRows Then
                MsgBox("Conexão OK!")
                TabPage2.Enabled = True
                TabPage3.Enabled = True
                TabPage4.Enabled = True
                TabPage5.Enabled = True
            Else
                MsgBox("Algum erro na conexão!")
                TabPage2.Enabled = False
                TabPage3.Enabled = False
                TabPage4.Enabled = False
                TabPage5.Enabled = False
            End If
        Catch ex As Exception
            MsgBox("Erro: " & ex.Message)
            TabPage2.Enabled = False
            TabPage3.Enabled = False
            TabPage4.Enabled = False
            TabPage5.Enabled = False
        Finally
            myCon1.Close()
        End Try

    End Sub

    Private Sub lerBanco()
        Dim sqlQuery1 As String = "SELECT DISTINCT p.nome FROM products limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                'reader("product_name")
            Loop
        Catch ex As Exception
            MsgBox("Erro. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    Private Sub addBanco()
        Dim sqlQuery1 As String = "INSERT INTO `paciente` (`nome`) VALUES ('@nome');"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        'myCommand.Parameters.AddWithValue("@login", login)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Erro. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' Entrando na tab Nutricionista / Pacientes / Dieta / Registros / Visualização
    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter, TabPage4.Enter, TabPage5.Enter, TabPage6.Enter
        If TabPage2.Enabled Then
            ListBox1.Items.Clear()
            ListBox2.Items.Clear()
            carregarNutricionistas()
            carregarPacientes()
            carregarProdutos()
            carregarDietas()
        End If
    End Sub

    Private Sub carregarPacientes()
        Dim sqlQuery1 As String = "SELECT * FROM paciente limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            ListBox2.Items.Clear()  '' 
            ListBox5.Items.Clear()  '' 
            ListBox8.Items.Clear()  '' Visualizar Registros
            ListBox9.Items.Clear()  '' Adicionar Registro Paciente
            '' Pacientes:
            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                ListBox2.Items.Add(reader("nome") & " [" & reader("id") & "]")
                ListBox5.Items.Add(reader("nome") & " [" & reader("id") & "]")
                ListBox8.Items.Add(reader("nome") & " [" & reader("id") & "]")
                ListBox9.Items.Add(reader("nome") & " [" & reader("id") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar Pacientes. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub
    Private Sub carregarNutricionistas()
        Dim sqlQuery1 As String = "SELECT * FROM nutricionista limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            ListBox1.Items.Clear()
            ComboBox4.Items.Clear()

            '' Nutricionistas:
            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                ListBox1.Items.Add(reader("nome") & " [" & reader("id") & "]")
                ComboBox4.Items.Add(reader("nome") & " [" & reader("id") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar nutricionistas. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub
    Private Sub carregarProdutos()
        Dim sqlQuery1 As String = "SELECT id,produto,categoria_id from produto limit 2000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            ComboBox6.Items.Clear()
            ListBox6.Items.Clear()

            '' Produtos:
            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                ComboBox6.Items.Add(reader("produto") & " [" & reader("id") & "]")
                ListBox6.Items.Add(reader("produto") & " [" & reader("id") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar nutricionistas. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub
    Private Sub carregarDietas()
        Dim sqlQuery1 As String = "SELECT d.id as 'dietaId', p.nome from dieta as d, paciente as p where d.paciente_id = p.id limit 2000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            ListBox7.Items.Clear()

            '' Dietas:
            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                ListBox7.Items.Add(reader("nome") & " [" & reader("dietaId") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar receitas de dieta. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' BOTAO ADD NUTRICIONISTA
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        addNutricionista(TextBox5.Text)
        ListBox1.Items.Clear()
        carregarNutricionistas()
    End Sub
    Private Sub addNutricionista(nome As String)
        Dim sqlQuery1 As String = "INSERT INTO nutricionista (nome) VALUES (@nome);"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        myCommand.Parameters.AddWithValue("@nome", nome)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
            MsgBox("Adicionado.")
        Catch ex As Exception
            MsgBox("Erro ao adicionar nutricionista. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' BOTAO ADD PACIENTE
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        addPaciente(TextBox6.Text)
        ListBox2.Items.Clear()
        carregarPacientes()
    End Sub
    Private Sub addPaciente(nome As String)
        Dim sqlQuery1 As String = "INSERT INTO paciente (nome) VALUES (@nome);"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        myCommand.Parameters.AddWithValue("@nome", nome)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
            MsgBox("Adicionado.")
        Catch ex As Exception
            MsgBox("Erro ao adicionar paciente. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' Entrando na tab PRODUTOS
    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        If TabPage3.Enabled Then
            '' Itens nutricionais
            ListBox3.Items.Clear()
            ComboBox3.Items.Clear()

            '' Itens do produto:
            itensNutrucionaisProduto = New List(Of ItemNutricional)
            ListBox4.Items.Clear()
            ComboBox1.Text = "g"

            '' Categorias
            ComboBox2.Items.Clear()

            '' Carregar dados:
            carregaritensNutricionais()
            carregarCategorias()
        End If
    End Sub

    Dim itensNutricionais As New List(Of ItemNutricional)
    Private Sub carregaritensNutricionais()
        Dim sqlQuery1 As String = "SELECT * FROM itens limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            itensNutricionais.Clear()

            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                itensNutricionais.Add(New ItemNutricional(reader("id"), reader("descricao")))
                ListBox3.Items.Add(reader("descricao") & "[" & reader("id") & "]")
                ComboBox3.Items.Add(reader("descricao") & "[" & reader("id") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar itens nutricionais. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub
    Dim categorias As New List(Of CategoriaProduto)
    Private Sub carregarCategorias()
        Dim sqlQuery1 As String = "SELECT * FROM categoria limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                categorias.Add(New CategoriaProduto(reader("id"), reader("descricao")))
                ComboBox2.Items.Add(reader("descricao") & " [" & reader("id") & "]")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar categorias. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' BOTAO ADD ITEM NUTRICIONAL:
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        addItemNutricional(TextBox7.Text)
        ListBox3.Items.Clear()
        ComboBox3.Items.Clear()
        carregaritensNutricionais()
    End Sub
    Private Sub addItemNutricional(descricao As String)
        Dim sqlQuery1 As String = "INSERT INTO itens (descricao) VALUES (@descricao);"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        myCommand.Parameters.AddWithValue("@descricao", descricao)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
            MsgBox("Adicionado.")
        Catch ex As Exception
            MsgBox("Erro ao adicionar nutricionista. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub


    '' BOTAO Limpar produto
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        limparCamposProduto()
        limparCamposProdutoItens()
    End Sub
    Private Sub limparCamposProduto()
        TextBox9.Text = ""
        ComboBox2.SelectedIndex = -1
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""

        ListBox4.Items.Clear()

    End Sub
    Private Sub limparCamposProdutoItens()
        'itensNutrucionaisProduto = New List(Of ItemNutricional)
        ComboBox3.SelectedIndex = -1
        TextBox8.Text = "0,0"
        ComboBox1.SelectedIndex = 0
        NumericUpDown1.Value = 0
    End Sub

    '' BOTAO ADD ITEM NUTRICIONAL AO PRODUTO
    Dim itensNutrucionaisProduto As New List(Of ItemNutricional)
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox3.Text = "" Then : MsgBox("Selecione o item nutricional.") : Return : End If
        If TextBox8.Text = "" Then : MsgBox("Informe a quantidade.") : Return : End If
        If ComboBox1.Text = "" Then : MsgBox("Informe ou selecione a medida.") : Return : End If

        Try
            Dim x() As String = ComboBox3.Text.Split("[")
            x(1) = x(1).Replace("]", "")
            itensNutrucionaisProduto.Add(New ItemNutricional(Integer.Parse(x(1).Trim), x(0).Trim, TextBox8.Text.Replace(".", ","), ComboBox1.Text, NumericUpDown1.Value))
            ListBox4.Items.Add(ComboBox3.Text & " - " & TextBox8.Text & " - " & ComboBox1.Text & " - " & NumericUpDown1.Value)
            limparCamposProdutoItens()

        Catch ex As Exception
            MsgBox("Erro ao adicionar item nutricional, há algum erro. " & ex.Message)
        End Try

    End Sub

    '' BOTAO SALVAR PRODUTO
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim categoriaStr As String = ComboBox2.Text
        If categoriaStr.Contains("[") Then
            Dim x() As String = categoriaStr.Split("[")
            categoriaStr = x(0).Trim
        End If

        If ComboBox2.Text = "" Then : MsgBox("Informe uma categoria.") : Return : End If
        If TextBox9.Text = "" Then : MsgBox("Informe uma descrição do produto.") : Return : End If
        If TextBox10.Text = "" Then : MsgBox("Informe o tipo de dose, por exemplo g, colher de sopa, unidade.") : Return : End If
        If TextBox11.Text = "" Then : MsgBox("Informe a quantidade da dose.") : Return : End If
        If TextBox12.Text = "" Then : MsgBox("Informe informação calorica em Kcal.") : Return : End If

        adicionaProduto(categoriaStr, TextBox9.Text,
                        TextBox10.Text, TextBox11.Text,
                        TextBox12.Text, NumericUpDown2.Value, itensNutrucionaisProduto)

    End Sub
    Private Sub adicionaProduto(ByVal categoria As String, ByVal descricao As String,
                                     ByVal doseTipo As String, ByVal doseQuantidade As Double,
                                     ByVal calorias As Integer, ByVal vd As Integer, ByVal itensProduto As List(Of ItemNutricional))

        Dim produtoId As Integer
        Dim tabelaNutricionalId As Integer

        'INSERT INTO categoria (descricao) VALUES ('Laticínios')
        'INSERT INTO produto (produto, categoria_id) VALUES ('Iogurte de Morango', '1')
        'INSERT INTO tabela_nutricional (produto_id, kcalorias, dose_tipo, dose_qnt) VALUES ('1', '203', 'g', '200')
        'INSERT INTO tabela_nutricional_itens (tabela_nutricional_id, item_id, quantidade, medida, porcentagem) VALUES ('1', '3', '6.1', 'g', '11')
        Dim sqlQuery1 As String = "SELECT * FROM categoria LIMIT 1000;"
        Dim sqlQuery2 As String = "INSERT INTO categoria (descricao) VALUES (@descricao); SELECT LAST_INSERT_ID();"
        Dim sqlQuery3 As String = "INSERT INTO produto (produto, categoria_id) VALUES (@produto, @catId); SELECT LAST_INSERT_ID();"
        Dim sqlQuery4 As String = "INSERT INTO tabela_nutricional (produto_id, kcalorias, valor_diario, dose_tipo, dose_qnt) VALUES (@produtoId, @calorias, @vd, @doseTipo, @doseQnt); SELECT LAST_INSERT_ID();"
        Dim sqlQuery5 As String = "INSERT INTO tabela_nutricional_itens (tabela_nutricional_id, item_id, quantidade, medida, porcentagem) VALUES (@tabNutId, @itemId, @qnt, @medida, @vd); SELECT LAST_INSERT_ID();"

        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()
            reader = myCommand.ExecuteReader()

            '' Verifica se ja tem no banco
            Dim idCategoria As Integer = 0
            Do While reader.Read()
                If reader("descricao") = categoria Then
                    idCategoria = reader("id")
                End If
            Loop
            reader.Close()

            '' Se nao tiver no banco adiciona
            If idCategoria = 0 Then
                myCommand.CommandText = sqlQuery2
                myCommand.Parameters.Clear()
                myCommand.Parameters.AddWithValue("@descricao", categoria)
                idCategoria = myCommand.ExecuteScalar
                'MsgBox(idCategoria) 'debug
            End If
            reader.Close()

            '' Adicionar Produto:
            'Dim sqlQuery3 As String = "INSERT INTO produto (produto, categoria_id) VALUES (@produto, @catId); SELECT LAST_INSERT_ID();"
            myCommand.CommandText = sqlQuery3
            myCommand.Parameters.Clear()
            myCommand.Parameters.AddWithValue("@produto", descricao)
            myCommand.Parameters.AddWithValue("@catId", idCategoria)
            produtoId = myCommand.ExecuteScalar
            'MsgBox(produtoId) 'debug,
            reader.Close()

            '' Adicionar Tabela Nutricional
            'Dim sqlQuery4 As String = "INSERT INTO tabela_nutricional (produto_id, kcalorias, dose_tipo, dose_qnt) 
            'VALUES (@produtoId, @calorias, @doseTipo, @doseQnt); SELECT LAST_INSERT_ID();"
            myCommand.CommandText = sqlQuery4
            myCommand.Parameters.Clear()
            myCommand.Parameters.AddWithValue("@produtoId", produtoId)
            myCommand.Parameters.AddWithValue("@calorias", calorias)
            myCommand.Parameters.AddWithValue("@vd", vd)
            myCommand.Parameters.AddWithValue("@doseTipo", doseTipo)
            myCommand.Parameters.AddWithValue("@doseQnt", doseQuantidade)
            tabelaNutricionalId = myCommand.ExecuteScalar
            'MsgBox(tabelaNutricionalId) 'debug
            reader.Close()

            For Each ip In itensProduto
                'Dim sqlQuery5 As String = "INSERT INTO tabela_nutricional_itens (tabela_nutricional_id, item_id, quantidade, medida, porcentagem) 
                'VALUES (@tabNutId, @itemId, @qnt, @medida, @vd);"
                myCommand.CommandText = sqlQuery5
                myCommand.Parameters.Clear()
                myCommand.Parameters.AddWithValue("@tabNutId", tabelaNutricionalId)
                myCommand.Parameters.AddWithValue("@itemId", ip.id)
                myCommand.Parameters.AddWithValue("@qnt", ip.quantidade)
                myCommand.Parameters.AddWithValue("@medida", ip.medida)
                myCommand.Parameters.AddWithValue("@vd", ip.vd)
                myCommand.ExecuteNonQuery()
                'MsgBox(tabelaNutricionalId) 'debug
            Next
            reader.Close()

            MsgBox("Produto adicionado.")
            limparCamposProduto()
            limparCamposProdutoItens()
            itensNutrucionaisProduto = New List(Of ItemNutricional)

        Catch ex As Exception
            MsgBox("Erro ao adicionar Produto. " & ex.Message)
        Finally
            myCon1.Close()
        End Try



    End Sub

    Private Sub ListBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox5.SelectedIndexChanged
        Label18.Text = "Paciente: " & ListBox5.SelectedItem
    End Sub

    '' BOTAO SALVAR DIETA
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If ListBox5.SelectedIndex = -1 Then : MsgBox("Selecione um paciente.") : Return : End If
        If ComboBox4.SelectedIndex = -1 Then : MsgBox("Selecione o nutricionista responsável.") : Return : End If
        If DateTimePicker1.Value = DateTimePicker2.Value Then : MsgBox("Selecione uma data diferente.") : Return : End If

        Try
            Dim x() = ListBox5.SelectedItem.ToString.Split("[")
            Dim pacId As Integer = x(1).Replace("]", "")

            Dim y() = ComboBox4.SelectedItem.ToString.Split("[")
            Dim nutId As Integer = y(1).Replace("]", "")

            addDieta(pacId, nutId, DateTimePicker1.Value, DateTimePicker2.Value)

        Catch ex As Exception
            MsgBox("Erro ao verificar código do paciente e nutricionista.")
        End Try

    End Sub
    Private Sub addDieta(pacienteId As Integer, nutricionistaId As Integer, dataInicio As Date, dataFim As Date)
        'INSERT INTO dieta (nutricionista_id, paciente_id, data_inicio, data_final) VALUES ('1', '1', '2018-11-28', '2018-11-28')
        Dim sqlQuery1 As String = "INSERT INTO dieta (nutricionista_id, paciente_id, data_inicio, data_final) VALUES (@nut, @pac, @dti, @dtf);"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        myCommand.Parameters.AddWithValue("@nut", nutricionistaId)
        myCommand.Parameters.AddWithValue("@pac", pacienteId)
        myCommand.Parameters.AddWithValue("@dti", dataInicio)
        myCommand.Parameters.AddWithValue("@dtf", dataFim)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
            MsgBox("Adicionado.")
        Catch ex As Exception
            MsgBox("Erro ao adicionar nutricionista. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub


    '' Adicionar Registro Paciente:
    Private Sub ListBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox9.SelectedIndexChanged
        Dim x() = ListBox9.SelectedItem.ToString.Split("[")
        Dim pacId As Integer = x(1).Replace("]", "")

        Dim dietaId As Integer = buscarDietaIdPaciente(pacId)

        If dietaId = 0 Then
            MsgBox("O paciente não tem uma receita de dieta prescrita, crie primeiro.")
        Else
            Label25.Text = "Dieta: " & dietaId
            Label25.Tag = dietaId
        End If


    End Sub
    Private Function buscarDietaIdPaciente(pacId As Integer) As Integer
        Dim retorno As Integer = 0

        Dim sqlQuery1 As String = "SELECT id FROM dieta where dieta.paciente_id = " & pacId & " limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Do While reader.Read()
                retorno = reader("id")
            Loop

        Catch ex As Exception
            MsgBox("Erro ao carregar categorias. " & ex.Message)
        Finally
            myCon1.Close()
        End Try

        Return retorno
    End Function
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        '' Botão salvar registro:
        If ComboBox6.SelectedIndex = -1 Then : MsgBox("Selecione um produto.") : Return : End If
        If TextBox14.Text = "" Or TextBox14.Text = "0,0" Then : MsgBox("Informe a dose.") : Return : End If
        If Label25.Text = "Dieta: ~Selecione Paciente~" Then : MsgBox("Selecione um paciente.") : Return : End If

        Try
            Dim x() = ListBox9.SelectedItem.ToString.Split("[")
            Dim pacId As Integer = x(1).Replace("]", "")

            Dim y() = ComboBox6.SelectedItem.ToString.Split("[")
            Dim produtoId As Integer = y(1).Replace("]", "")

            adicionarRegistroDieta(pacId, produtoId, TextBox14.Text, CInt(TextBox15.Text), Label25.Tag)

            ComboBox6.SelectedIndex = -1
            TextBox14.Text = "0,0"
            Label25.Text = "Dieta: ~Selecione Paciente~"
        Catch ex As Exception
            MsgBox("Erro ao verificar código do paciente e do produto.")
        End Try

    End Sub
    Private Sub adicionarRegistroDieta(pacId As Integer, produtoId As Integer, dose As String, kCaloria As Integer, dietaId As Integer)
        'INSERT INTO registro(dieta_id, produto_id, dose, data_hora) VALUES ('1', '1', '47', '2018-11-28 10:44:59')
        Dim sqlQuery1 As String = "INSERT INTO registro (dieta_id, produto_id, dose, kcaloria, data_hora) 
                                VALUES (@did, @pid, @dose, @kc, @dthora);"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)

        myCommand.Parameters.AddWithValue("@did", dietaId)
        myCommand.Parameters.AddWithValue("@pid", produtoId)
        myCommand.Parameters.AddWithValue("@dose", dose.Replace(",", "."))
        myCommand.Parameters.AddWithValue("@kc", kCaloria)
        myCommand.Parameters.AddWithValue("@dthora", DateTime.Now)

        Try
            myCon1.Open()
            myCommand.ExecuteNonQuery()
            MsgBox("Adicionado.")
        Catch ex As Exception
            MsgBox("Erro ao adicionar registro à dieta. " & ex.Message)
        Finally
            myCon1.Close()
        End Try

    End Sub

    '' VISUALIZAR PRODUTOS:
    Private Sub ListBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox6.SelectedIndexChanged
        Dim x() = ListBox6.SelectedItem.ToString.Split("[")
        Dim produtoId As Integer = x(1).Replace("]", "")

        carregarInfoProduto(produtoId)

    End Sub

    Private Sub carregarInfoProduto(produtoId As Integer)
        Dim sqlQuery1 As String = "SELECT p.produto, t.kcalorias, t.valor_diario, i.descricao, ti.quantidade, ti.medida, ti.porcentagem 
                              FROM produto AS p, itens AS i, tabela_nutricional AS t,	tabela_nutricional_itens AS ti
                              WHERE t.produto_id = p.id AND ti.item_id = i.id AND ti.tabela_nutricional_id = t.id AND p.id = " & produtoId & " limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Dim dt = New DataTable()
            dt.Load(reader)
            DataGridView2.AutoGenerateColumns = True
            DataGridView2.DataSource = dt
            DataGridView2.Refresh()

        Catch ex As Exception
            MsgBox("Erro ao carregar informações do produto. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' VISUALIZAR DIETAS:
    Private Sub ListBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox7.SelectedIndexChanged
        Dim x() = ListBox7.SelectedItem.ToString.Split("[")
        Dim pacId As Integer = x(1).Replace("]", "")

        visualizarDietas(pacId)
    End Sub
    Private Sub visualizarDietas(pacId As Integer)
        Dim sqlQuery1 As String = "SELECT DISTINCT pac.nome AS 'Paciente', n.nome AS 'Nutricionista', d.data_inicio AS 'Inicio', d.data_final AS 'Final' 
                        FROM dieta AS d, produto AS p, paciente AS pac, nutricionista AS n 
                        WHERE d.paciente_id = pac.id AND d.nutricionista_id = n.id AND pac.id = " & pacId & " LIMIT 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Dim dt = New DataTable()
            dt.Load(reader)
            DataGridView3.AutoGenerateColumns = True
            DataGridView3.DataSource = dt
            DataGridView3.Refresh()

        Catch ex As Exception
            MsgBox("Erro ao carregar receitas de dieta. " & ex.Message)
        Finally
            myCon1.Close()
        End Try
    End Sub

    '' VISUALIZAR REGISTROS PACIENTE:
    Private Sub ListBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox8.SelectedIndexChanged
        Dim x() = ListBox8.SelectedItem.ToString.Split("[")
        Dim pacId As Integer = x(1).Replace("]", "")

        carregaRegistroDieta(pacId)

    End Sub
    Private Sub carregaRegistroDieta(pacId As Integer)
        'select pac.nome, p.produto, r.dose from dieta as d, produto as p, registro as r, paciente as pac where r.produto_id = p.id and d.paciente_id = pac.id;
        Dim sqlQuery1 As String = "SELECT DISTINCT pac.nome, p.produto, r.dose, r.kcaloria, r.data_hora 
                                FROM dieta AS d, produto AS p, registro AS r, paciente AS pac 
                                WHERE r.produto_id = p.id and d.paciente_id = pac.id AND r.dieta_id = d.id AND pac.id = " & pacId & " limit 1000;"
        Dim myCon1 As New MySqlConnection(connString)
        Dim myCommand As New MySqlCommand(sqlQuery1, myCon1)
        Dim reader As MySqlDataReader

        Try
            myCon1.Open()

            reader = myCommand.ExecuteReader()
            Dim dt = New DataTable()
            dt.Load(reader)
            DataGridView1.AutoGenerateColumns = True
            DataGridView1.DataSource = dt
            DataGridView1.Refresh()

            reader = myCommand.ExecuteReader()
            Dim total As Double
            Do While reader.Read()
                total += reader("kcaloria")
            Loop

            Label28.Text = "Info: Total de " & total & " Kcal"

        Catch ex As Exception
            MsgBox("Erro ao carregar registros. " & ex.Message)
        Finally
            myCon1.Close()
        End Try

    End Sub

End Class
