## Essas são as soluções presentes na página:

- **Encurtador de links ([lnk-sht](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/lnk-sht.md)):**
  Esta macro encurtará o link que estiver na célula. Algumas vezes, durante o trabalho, você precisa compilar muitas informações em uma única planilha e isso inclui os links para dados externos. O problema é que esses links não possuem um tamanho padronizado, o que estraga o visual da planilha. A forma de padronizar eles é através do encurtamento, deixando todos com o mesmo tamanho. Essa macro usa o Bitly para encurtar os links, deixando todos do mesmo tamanho. 
- **Crie um temporizador, para executar uma macro ([tmr](https://github.com/cfprocha/codigos/blob/main/VBA/Excel/tmr.md)):**
  Uma macro será executada após um certo período de tempo. Algumas vezes criamos macros e desejamos que elas sejam executadas várias vezes, após períodos específicos de tempo. Quando você executa "minhaMacro" descrita abaixo, ela definirá um intervalo de tempo de 15 minutos, entre uma execução e outra. Caso deseje fazer ela parar, basta executar a macro "paraTimer". 

## Como usar os códigos

1. Cria um botão para disparar a macro (cmdStart) e iniciar a funcionalidade dela, na planilha principal
2. Acesse Tools (Ferramentas) --> Macro --> Visual Basic Editor (ou pressione Alt + F11)
3. Na janela do VBE, clique em Insert (Inserir) --> Module (Módulo)
4. Dê dois cliques no Module1 e cole o código da solução.
