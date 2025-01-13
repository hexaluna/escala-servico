# escala-servico
Escala de serviço militar, usando funções do LibreOffice Calc, sem Visual Basic, para ser compatível com outras suítes office.

1.0. Introdução:
 
    A planilha "escala-servico.ods" tem como objetivo auxiliar o sargento-ajudante (sargenteante) ou sargento-brigada na escala de serviço militar. Ela foi desenvolvida no com funções básicas do LibreOffice Calc, algumas matriciais, montando um algoritmo sem Visual Basic, sendo leve e compatível com qualquer suíte Office.

2.0. Objetivos:

    Explicar como utilizar a planilha para escalar militares;
    
    Explicar como funciona o vínculo de mala-direta do aditamento e pernoite.

3.0. Desenvolvimento:

3.1. Utilização da planilha, 1ª aba, "escala", linha 1:
            
    Data In: data de início da célula "C1" SEMPRE começando numa sexto dia da semana(sexta-feira). Após edição, não poderá ser mais alterada. Proteja a aba "escala", por segurança, para impedir edição da célula "C1".
        
    D1: quantidade total de militares a serem escalados do segundo ao sexto dia da semana (escala preta, escala A, etc);
        
    E1: quantidade total de militares a serem escalados do sétimo ao primeiro dia da semana (escala vermelha, escala B, etc);

    Esc: escala preta e escala vermelha editável, ao lado, para feriados ou alterações da escala feitas pelo comando. Digite "v" no dia que deseja escala vermelha e "p" do dia que deseja escala preta, e as respectivas folgas serão recalculadas. Para escala corrida, substitua o "v" por "p" no período determinado.

3.2. Utilização da planilha, 1ª aba, "escala", linha 2:

    Nr: número da antiguidade do militar, abaixo, apenas para referência e não pode ser alterado;

    P Grad: posto ou graduação do militar, abaixo;
        
    Nome G: nome de guerra do militar, abaixo;

    Sit: situação do militar, abaixo, pronto (Pron) ou não pronto (N Pron);

    Dstn: destino ou motivo do militar, baixado (Bx), etc;

    Sv: quantidade total de serviços que o militar já cumpriu, abaixo. Ao lado direito de Sv, temos a linha do tempo "G3:OJ3", contendo as datas, com condicionais em preto, vermelho e marrom, onde esta ultima representa o sexto dia da semana, que em algumas instituições é meio expediente. A condicional em amarelo representa o dia seguinte, ou seja, onde os militares serão escalados de véspera, conforme o regulamento.

3.3. Utilização da planilha, 1ª aba, "escala", linha 3:

    Tot "n" Mil: Total de "n" militares;

    Abaixo de Sit: Total de militares prontos;

    Abaixo de Dstn: Total de militares em destino;

    Abaixo da linha do tempo: contagem de militares escalados com o "/".

3.4. Utilização da planilha, 1ª aba, "escala", linhas restantes:

    Na matriz "G3:OJ203", os militares podem ser escalados utilizando-se do sinal literal "/". Ao fazer isto, o nome do militar aparecerá na aba "previsao", na ordem da antiguidade preenchida na aba "escala", de cima para baixo, do mais antigo para o mais moderno, respectivamente. Para que isto aconteça, é necessário indicar o primeiro dia que deseja iniciar a previsão da escala, desde que esta data esteja presente na linha do tempo "G3:OJ3". Se a aba "escala" foi protegida, conforme orientações do item 3.1, não existirão problemas gráficos ao escalar militares utilizando o sinal "/". 
    Curiosamente, "=UNICODE(/) = 47", que é imediatamente inferior a UNICODE(0) = 48, responsável pela brilhante solução de ranqueamento de dados repetidos usando termos literais, sem a qual não seria possível o feito supracitado.

3.5. Utilização da planilha, 2ª aba, "previsao":

    Na célula "B2", pode ser editado o início da previsão da escala de serviço, num total de 15 (quinze) dias, inclusive, para até 50 (cinquenta) militares escalados numa guarnição de serviço (Gu Sv);

    Abaixo de "Data", podem ser indicados os postos da Gu Sv a serem escalados, em ordem de antiguidade.

3.6. Utilização da planilha, 3ª aba, "alteracao":

    A aba "alteracao" é uma cópia da aba "previsao", mas pode ser editada, para fins de permutas, substituições ou trocas de serviço, previstas ou emergenciais.

3.5. Vínculo de mala-direta, 4ª aba, "transpor":

    A aba "transpor" é uma transposição da aba "previsao" e "alteracao, para servir de base de dados de mala-direta, a fim de vincular seus campos nos arquivos "aditamento.odt" e "pernoite.odt", proporcionando documentos preenchidos automaticamente, não isentando a responsabilidade de conferência destes pelos auxiliares e chefes imediatos para sua apreciação, assinatura, publicação e leitura.

4.0. Conclusão:

    Para sugestões, envie um e-mail para: hexaluna@outlook.com

Desenvolvida por: Rafael LUNA. 
Colaboração de: INALDO Junior, Danilo DERRE, Bill LESKO, Gilberto SCHIAVINATTO, Ewerton DUTRA, LEONARDO COSTA, Tatiane PRIMO, Kauan SILVA, Daniel OLIVEIRA, WILLIAM Costa e Mauricio DIAS.
