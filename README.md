# Exerc-cio-1
nome : Lista 3 - 1
descrição : > -
  Layout do suplemento: três entradas e um botão, para permitir cadastrar o Nome,
  a Nota 1 e a Nota 2 de cada aluno.
  Comportamento do suplemento: a cada clique no botão, o script previamente
  deve verificar se a primeira célula da planilha está vazia e, caso esteja,
  preencher os cabeçalhos nas quatro primeiras colunas da planilha: "Nome",
  "Nota 1", "Nota 2" e "Média Final". Em seguida, o script deve procurar um
  próxima linha livre da planilha e preencher cada coluna ali com as informações
  digitadas nos campos, na ordem correta, por fim limpando os valores dos campos
  e deixando o suplemento pronto para inserir outro aluno (dica: calcule a média
  final de acordo com os valores digitados nos campos antes de tentar preencher
  na planilha).
  Tempo estimado para conclusão: 1 hora e 30 minutos.
anfitrião : EXCEL
api_set : {}
script :
  conteúdo : >
    show.addEventListener ("click", async () => {
      esperar Excel.run (função assíncrona (contexto) {
        let sheet = context.workbook.worksheets.getActiveWorksheet ();
        let table = sheet.tables.getItemOrNullObject ("TabelaAluno");
        table.load ("isNullObject");
        aguarde context.sync ();
        if (table.isNullObject) {
          tabela = planilha.tables.add ("A1: D1", verdadeiro);
          table.name = "TabelaAluno";
          table.getHeaderRowRange (). values ​​= [["Nome", "Nota 1", "Nota 2", "Média"]];
        };
        let nome = txtname.value;
        deixe nota1 = parseFloat (txtnoteone.value);
        deixe nota2 = parseFloat (txtnotetwo.value);
        deixe mídia = (nota1 + nota2) / 2;
        table.rows.add (null, [[nome, nota1, nota2, media]]);
        table.getRange (). format.autofitColumns ();
        table.getRange (). format.autofitRows ();
        return context.sync ();
      }). catch (função (erro) {
        console.log ("Erro:" + erro);
        if (instância de erro de OfficeExtension.Error) {
          console.log ("Informações de depuração:" + JSON.stringify (error.debugInfo));
        }
      });
    });
    // Força somente números nos inputs
    //
    https://www.geeksforgeeks.org/how-to-force-input-field-to-enter-numbers-only-using-javascript/
    function onlyNumberKey (evt) {
      // Somente caracteres ASCII permitidos neste intervalo
      var ASCIICode = evt.which? evt.which: evt.keyCode;
      if (ASCIICode> 31 && (ASCIICode <48 || ASCIICode> 57)) return false;
      return true;
    }
  linguagem : texto datilografado
modelo :
  conteúdo : " <h1> Cadastro de Notas </h1> \ n <div> \ n \ t <input id = \" txtname \ " placeholder = \" Nome do Aluno \ " > </br> \ n \ t < input id = \ " txtnoteone \" placeholder = \ " Nota 1 \" onkeypress = \ " return onlyNumberKey (event) \" maxlength = \ " 2 \" > </br> \ n \ t <input id = \ " txtnotetwo \ "placeholder = \ " Nota 2 \" onkeypress =\ " return onlyNumberKey (event) \" maxlength = \ " 2 \" > \ n </div> \ n \ n \ t <div> \ n \ t \ t <button id = \ " show \" > Adicionar < / botão> \ n \ t </div> "
  linguagem : html
estilo :
  conteúdo : | -
    section.samples {
        margem superior: 20px;
    }
    section.samples .ms-Button, section.setup .ms-Button {
        display: bloco;
        margin-bottom: 5px;
        margem esquerda: 20px;
        largura mínima: 80px;
    }
  idioma : css
bibliotecas : |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @ types / office-js
  office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css
  core-js@2.4.1/client/core.min.js
  @ types / core-js
  jquery@3.1.1
  @ types / jquery @ 3.3.1
