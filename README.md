# escala-servico
Escala de serviço militar, usando funções do LibreOffice Calc, sem Visual Basic, para ser compatível com outras suítes office.

1.0. Introdução:
 
    A planilha "escala-servico.ods" tem como objetivo auxiliar o sargento-ajudante (sargenteante) ou sargento-brigada na escala de serviço militar. Ela foi desenvolvida no com funções básicas do LibreOffice Calc, algumas matriciais, montando um algoritmo sem Visual Basic, sendo leve e compatível com qualquer suíte Office.

2.0. Objetivos:

    2.1. Explicar como utilizar a planilha para escalar militares;
    2.2. Explicar como funciona o vínculo de mala-direta do aditamento e pernoite;
    2.3. Explicar como funciona a lógica de programação para os interessados, a fim de coletar ideias para futuras melhorias.

3.0. Desenvolvimento:

    3.1. Utilização da planilha, 1ª aba, "escala", linha 1:
            
        Data In: data de início da célula "C1" SEMPRE começando numa sexto dia da semana(sexta-feira);
        
        D1: quantidade total de militares a serem escalados do segundo ao sexto dia da semana (escala preta, escala A, etc);
        
        E1: quantidade total de militares a serem escalados do sétimo ao primeiro dia da semana (escala vermelha, escala B, etc);
        
        Esc: escala preta e escala vermelha editável, ao lado, para feriados ou alterações da escala feitas pelo comando. Digite "v" no dia que deseja escala vermelha e "p" do dia que deseja escala preta, e as respectivas folgas serão recalculadas;

    3.2. Utilização da planilha, 1ª aba, "escala", linha 2:

        Nr: número da antiguidade do militar, abaixo, apenas para referência e não pode ser alterado;

        P Grad: posto ou graduação do militar, abaixo;
        
        Nome G: nome de guerra do militar, abaixo;

        Sit: situação do militar, abaixo, pronto (Pron) ou não pronto (N Pron);

        Dstn: destino ou motivo do militar, baixado (Bx), etc;

        Sv: quantidade total de serviços que o militar já cumpriu, abaixo. Ao lado direito de Sv, temos a linha do tempo "G3:OJ3", contendo as datas, com condicionais em preto, vermelho e marrom, onde esta ultima representa o sexto dia da semana, que em algumas instituições é meio expediente. A condicional em amarelo representa o dia seguinte, ou seja, onde os militares serão escalados de véspera, conforme o regulamento.

    3.3. Utilização da planilha, 1ª aba, "escala", linha 3:

        Tot "n" Mil: Total de "n" militares;

        abaixo de Sit: Total de militares prontos;

        abaixo de Dstn: Total de militares em destino;

        abaixo da linha do tempo: contagem de militares escalados com o "/".

    3.4. Utilização da planilha, 1ª aba, "previsao", linhas restantes:

        Na matriz "G3:OJ203", os militares podem ser escalados utilizando-se do sinal literal "/". Ao fazer isto, o nome do militar aparecerá na aba "previsao", na ordem da antiguidade preenchida na aba "escala", de cima para baixo, do mais antigo para o mais moderno, respectivamente. Para que isto aconteça, é necessário indicar o primeiro dia que deseja iniciar a previsão da escala, desde que esta data esteja presente na linha do tempo "G3:OJ3". Curiosamente, UNICODE(/) = 47, que é imediatamente inferior a UNICODE(0) = 48, responsável pela brilhante solução de ranqueamento de dados repetidos usando termos literais, sem a qual não seria possível o feito supracitado.

Desenvolvida por: RAFAEL LUNA. Colaboração de: INALDO JUNIOR, EWERTON DUTRA, LEONARDO COSTA, TATIANE PRIMO, KAUAN SILVA, DANIEL OLIVEIRA, WILLIAM COSTA e MAURICIO DIAS.

