# EncontraDistanciasVBA
Encontra distância entre dois pontos a partir do Excel

(Update em set/2018: O Google mudou a chamada da API. Agora é necessário ter abrir uma conta e colocar a chave na chamada, de modo que o código abaixo não funciona mais. Não devo mexer no código por não eu ter demanda para isto no momento)

Esta é uma ferramenta muito útil para Logística.

Você insere uma série de origens e destinos. Pode ser endereço, CEP, cidade, qualquer coisa entendível pelo GMaps. A macro consulta o Google maps para cada linha, e traz a distância e o tempo.

Demora um pouco, mas é melhor do que fazer no braço.

Os endereços devem ser entendíveis no GMaps. Se não encontrar, vai voltar um zero.

E o GMaps pode errar também. Se procurar a rua Pernambuco em São Paulo, ele pode retornar a rua São Paulo, em Pernambuco. Portanto, tem que tomar cuidado e não confiar cegamente na informação.
 
Há um limite de consultas de 2500 consultas diárias por dia no Google Maps, com esta API grátis. 



Ferramentas em Excel VBA
https://ferramentasexcelvba.wordpress.com/


Ideias técnicas com um pouco de filosofia
https://ideiasesquecidas.com/
