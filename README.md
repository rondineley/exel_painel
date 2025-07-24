Exel Painel

Acabei de desenvolver um aplicativo desktop usando PyQt5 que permite listar, classificar e abrir arquivos Excel de forma rápida e intuitiva. 

<img width="799" height="628" alt="image" src="https://github.com/user-attachments/assets/a4c115eb-86c1-4ac1-be49-68a8d4a07466" />

Galeria de Arquivos Excel em PyQt5
Este projeto é uma aplicação gráfica desenvolvida com PyQt5 para visualizar e gerenciar arquivos do tipo Excel (.xls, .xlsx) presentes em um diretório e suas subpastas. A interface foi projetada para ser simples, intuitiva e com um tema escuro estilo AMOLED.

Funcionalidades
Seleção de Diretório:
O usuário pode selecionar um diretório e a aplicação irá listar todos os arquivos do tipo Excel encontrados nele, incluindo os arquivos dentro de subpastas.

Exibição da Lista:
A lista de arquivos é exibida de maneira vertical, mostrando o nome do arquivo seguido de seu caminho completo.

Ordenação de Arquivos:
Os arquivos podem ser ordenados de duas maneiras:

Por Nome: Ordena os arquivos em ordem alfabética.

Por Data: Ordena os arquivos de acordo com a data de modificação.

Exportação para CSV:
A aplicação permite exportar a lista de arquivos, com nome e caminho, para um arquivo CSV.

Abrir Arquivos:
Ao clicar duas vezes em um arquivo na lista, a aplicação tenta abrir o arquivo Excel correspondente usando o aplicativo padrão do sistema.

Tecnologias Usadas
Python 3.x

PyQt5 para a interface gráfica

os e csv para manipulação de arquivos e exportação de dados

Exemplo de Uso
Selecione o diretório onde seus arquivos Excel estão localizados.

A lista será automaticamente atualizada com todos os arquivos encontrados, incluindo os de subpastas.

Você pode classificar os arquivos por nome ou por data de modificação.

Clique duas vezes em qualquer arquivo para abri-lo no aplicativo associado ao seu sistema.

Exporte a lista de arquivos para um arquivo CSV clicando no botão "Exportar Lista".

Instalação
Clone o repositório:

bash
Copiar
Editar
git clone https://github.com/seu-usuario/galeria-arquivos-excel.git
cd galeria-arquivos-excel
Instale as dependências:

bash
Copiar
Editar
pip install PyQt5
Execute a aplicação:

bash
Copiar
Editar
python galeria_exel.py
Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests ou abrir issues com sugestões de melhorias.

#Python #PyQt5 #DesenvolvimentoDeSoftware #Automação #Excel
