<% @Language=VBScript %>

<%
    dim lblMensagem
    dim nome, telefone, dataNascimento  

    function validarForm()        
        if (len(nome) = 0 or len(telefone) = 0) then
            lblMensagem = "Os campos NOME e TELEFONE são obrigatórios!"
            validarForm = 1
        else
            validarForm = 0
        end if
    end function

    sub SalvarRegistro()
        dim conn, command
        'Declaração da conexão
        set conn = Server.CreateObject("ADODB.Connection")
        conn.ConnectionString ="Provider=SQLOLEDB.1;Data Source=localhost\SQLEXPRESS;Initial Catalog=Agenda;UID=dbTeste;PWD=SenhaTeste#2;"

        'Monta o comando SQL
        command =   "INSERT INTO Agenda.dbo.tbl_contatos (Nome, Telefone, DataNascimento) " &_
                    "VALUES ('" & nome & "', '" & telefone & "', " & dataNascimento & ")"
        
        on error resume next
            'Executa o insert no banco
            conn.Open()
            conn.Execute command
            conn.close

            if( err.number = 0) then
                lblMensagem = "Contato inserido com sucesso!"
                response.redirect "./Index.asp"
            else
                lblMensagem = "Erro ao inserir contato: " & err.description
            end if
        on error goto 0
    end sub
%>

<%
    'Verifica se é um POST'
    if (request.form <> "") then      
        nome = Trim(Request.Form("nome"))
        telefone = Trim(Request.Form("telefone"))
        dataNascimento = Request.Form("dataNascimento")

        if(len(dataNascimento) = 0) then
            dataNascimento = "NULL"
        else
            dataNascimento = "'" & dataNascimento & "'"
        end if

        'Verifica os campos do form e executa o insert
        if(validarForm = 0) then
            call SalvarRegistro()
        end if

    end if
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link rel="stylesheet" type="text/css" media="screen" href="style.css">
    <script src="main.js"></script>
</head>
<body>
    <h5 style='color: red; font-weight: bold;'>
        <% response.write(lblMensagem) %>
    </h5>
    <h3>
        Lista de Contatos - Novo
    </h3>
    <form action="Inserir.asp" method="POST">
        <table>
            <tbody>
                <tr>
                    <td>Nome</td>
                    <td>
                        <input required maxlength="100"  type="text" name="nome">
                    </td>
                </tr>
                <tr>
                    <td>Telefone</td>
                    <td>
                        <input maxlength="20" required type="text" name="telefone">
                    </td>
                </tr>
                <tr>
                    <td>Data Nascimento</td>
                    <td>
                        <input type="date" name="dataNascimento">
                    </td>
                </tr>
                <tr>
                    <td colspan="100%">
                        <button type="submit">Salvar</button>
                    </td>
                </tr>
            </tbody>
        </table>
    </form>
</body>
</html>