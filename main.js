function toggleCheckboxes(obj){
    let checkboxes = obj.closest("table").querySelector("tbody").querySelectorAll("[type=checkbox]");
    let marcados;

    checkboxes = Array.from(checkboxes);

    checkboxes.map((element) => {
        element.checked = obj.checked;
        console.log('passo 1')
    });

    marcados = obj.closest("table").querySelector("tbody").querySelectorAll("[type=checkbox]:checked");
    toggleExcluir(marcados)
}

function checkboxAction(obj){
    let masterCheckbox = obj.closest("table").querySelector("[type=checkbox]")
    let checkboxes = obj.closest("tbody").querySelectorAll("[type=checkbox]")
    let marcados = obj.closest("tbody").querySelectorAll("[type=checkbox]:checked")

    //ativa o checkbox master de acordo com os checkboxes selecionados
    if(checkboxes.length === marcados.length){
        masterCheckbox.checked = obj.checked
    }
    else{
        masterCheckbox.checked = false;
    }

    toggleExcluir(marcados);
}

//ativa o botão de excluir somente se houver checkbox selecionado
function toggleExcluir(marcados){
    let btnExcluir = document.querySelector("#btnExcluir")
     
    if(marcados.length > 0){
        btnExcluir.disabled = false;
    }
    else{
        btnExcluir.disabled = true;
    }

    console.log('passo 2')
}