<%@Language=VBScript %>
<% 
     'option explicit
    
    ' dim teste, a
    ' teste = Array("apple","orange","cherries")

    ' 'response.write(join(teste))
    
    ' for each x in teste
    '     'response.write("okokok")
    '     a = 0
    ' next
'fruits is an array 
         dim fruits, fruitnames
         fruits = Array("apple","orange","cherries")

         'iterating using For each loop. 
         For each item in fruits
            response.write("<br />" & item)
         Next


%>