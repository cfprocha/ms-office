## Essas são as soluções presentes na página:

- **Encurtador de links ([lnk-sht](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/lnk-sht.bas)):**
  Esta macro encurtará o link que estiver na célula. Algumas vezes, durante o trabalho, você precisa compilar muitas informações em uma única planilha e isso inclui os links para dados externos. O problema é que esses links não possuem um tamanho padronizado, o que estraga o visual da planilha. A forma de padronizar eles é através do encurtamento, deixando todos com o mesmo tamanho. Essa macro usa o Bitly para encurtar os links, deixando todos do mesmo tamanho. 
- **Crie um temporizador, para executar uma macro ([tmr](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/tmr.bas)):**
  Uma macro será executada após um certo período de tempo. Algumas vezes criamos macros e desejamos que elas sejam executadas várias vezes, após períodos específicos de tempo. Quando você executa "minhaMacro" descrita abaixo, ela definirá um intervalo de tempo de 15 minutos, entre uma execução e outra. Caso deseje fazer ela parar, basta executar a macro "paraTimer".
- **Descaracterizar o Excel ([desc-xl](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/desc-xl.bas)):**
  Esse código oculta as barras de comandos, fórmulas e os cabeçalhos, além de alterar o nome da janela. Ele pode ser usado para, visualmente, descaracterizar o Excel, dando a entender, para um usuário leigo, que está trabalhando em um aplicativo diferente.
- **Alerta de inclusão ou exclusão de linha ([aie](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/aie.bas)):**
  A finalidade desse código é emitir uma mensagem de alerta, toda vez que o usuário inserir, ou excluir uma linha na planilha.
- **Fecha planilha inativa ([fecha-inativa](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/fecha-inativa.bas)):**
  Esse código serve para fechar uma planilha que esteja aberta há mais de 30 minutos, sem receber nenhuma interação. Ele é útil para os casos em que o usuário se afasta do computador e esquece de fechar a planilha ativa. Basta copiar os códigos nos locais indicados e quando abrir a planilha ele já estará em efeito.
- **Valida e-mail ([email-valido](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/email-valido.bas)):**
  Use esse código para validar a sintaxe de um e-mail informado em uma caixa de diálogo, ou constante de uma célula da sua planilha.
- **Personalizar o tempo de salvamento ([salva-period](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/salva-period.bas)):**  
  A finalidade desse código é definir o período de tempo após o qual a planilha será salva. Isso é uma forma de personalizar o autosalvamento, através de programação.
- **Onde estão os arquivos temporários ([onde-temp](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/onde-temp.bas)):**
  Esse código serve para mostrar onde ficam os arquivos temporários do seu Windows (onde está a pasta dos temporários). Como ele pode ter sido movida, assim temos uma identificação rápida dela.


## Como usar os códigos

1. Crie um botão para disparar a macro (cmdStart) e iniciar a funcionalidade dela, na planilha principal
2. Acesse Tools (Ferramentas) --> Macro --> Visual Basic Editor (ou pressione Alt + F11)
3. Na janela do VBE, clique em Insert (Inserir) --> Module (Módulo)
4. Dê dois cliques no Module1 e cole o código da solução.
