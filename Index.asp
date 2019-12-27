<%@Language=VBScript %>
<% 
    option explicit 
    dim conn, rs, lblMensagem, lblCor
    set conn = Server.CreateObject("ADODB.Connection")
    conn.ConnectionString ="Provider=SQLOLEDB.1;Data Source=localhost\SQLEXPRESS;Initial Catalog=Agenda;UID=dbTeste;PWD=SenhaTeste#2;"    

    sub ExcluirContato()
        dim arrContatos, contato 'necessário declarar essa var CONTATO usada no FOREACH por causa do OPTION EXPLICIT
        arrContatos = request.form("idContatos") 'todo: validar esse parametro
        
        if(len(arrContatos) >= 1) then
            conn.open()
            conn.Execute "DELETE FROM agenda.dbo.tbl_contatos WHERE Id IN(" & arrContatos & ")"            
            conn.close()
            lblMensagem = "Contatos excluídos com sucesso!"
            lblCor = "green"
        else
            lblMensagem = "É necessário informar os contatos que serão excluídos!"
            lblCor = "red"
        end if
    end sub
%>

<%
    if( request.form <> "") then
        call ExcluirContato()
    end if
%>

<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <link href="style.css" rel="stylesheet" type="text/css">
    <script src="main.js"></script>

</head>
<body>
    <% Response.Write("<div style='font-weight: bold; color:" & lblCor & ";'>" & lblMensagem & "</div>") %>

    <h3>
        Lista de Contatos
        <div style="float: right;">
            <a class="link-btn" href= "Inserir.asp">Novo</a>
            <button id="btnExcluir" disabled form="frmContatos" type="submit" onclick="return confirm('Deseja excluír os contatos selecionados?')">Excluir</button>
        </div>
        
    </h3>

    <!--tabela com os contatos -->
    <form action="Index.asp" method="POST" id="frmContatos">
        <table>
            <thead>
                <tr>
                    <th>
                        <input type="checkbox" onclick="toggleCheckboxes(this)" />
                    </th>
                    <th>Nome</th>
                    <th>Telefone</th>
                    <th>Data Nascimento</th>
                </tr>
            </thead>
            <tbody>
                <%
                    on error resume next                        
                        conn.Open()
                        Set rs = conn.Execute("SELECT * FROM tbl_contatos ORDER BY Nome")

                        'Percorre o recordSet de contatos
                        do until rs.EOF

                            dim id, nome, telefone, dataNascimento
                            
                            id = rs("Id")
                            nome = rs("Nome")
                            telefone = rs("Telefone")
                            dataNascimento = rs("DataNascimento") 'FormatDateTime

                            'Tratamento de valores NULOS antes de formatar
                            if(len(dataNascimento) = 10) then
                                dataNascimento = FormatDateTime(dataNascimento)
                            else
                                dataNascimento = "--"
                            end if

                            Response.Write("<tr>")
                            Response.Write("<td style='text-align: center;'><input onclick='checkboxAction(this)' type='checkbox' name='idContatos' value='" & id & "' /></td>")
                            Response.Write("<td>" & nome & "</td>")
                            Response.Write("<td>" & telefone & "</td>")
                            Response.Write("<td style='text-align:center;'>" & dataNascimento & "</td>")
                            Response.Write("</tr>")

                            rs.MoveNext
                        loop

                        rs.close
                        conn.close

                        if Err.Number <> 0 then
                            Response.Write("<p style='color: red;'>" & Err.Description & "</p>")           
                        end if
                    On Error Goto 0
                %>
            </tbody>
        </table>    
    </form>

</body>
</html>