# CSM EXCEL

Modelo simples de preenchimento por meio de checkbox e optionButton no excel.

# SCENES

- [x] Criação do modelo em excel para facilitar o copy/paste de um arquivo csv.
- [x] Flexibilidade no preenchimento.
- [x] Validação dos campos.
- [x] Regra de negócio separado em arquivos .vbs para facilitar manutenção.
- [x] Modelo pontual criado para atendimento ativo.
- [x] Funcionalidade baseada no office 2011.
- [x] Listar
    - [x] Contratos preenchidos
    - [x] Contratos pendentes para preenchimento
    - [x] Todos contratos
- [x] Performance das Macros 
    - [x] Leves
    - [x] Dinamicas
    - [x] Intuitivas
- [x] Template com dados para teste.

## FUNCTIONS

- `wkb_open` // popula as formulas na worksheet base

- `front_initialize` // Popula o combobox
- `front_checkBox1` // função para o click no checkbox
- `front_comboBox1` // função para o click no combobox
- `front_clearCheckBoxStatic` // limpeza dos checkboxes para novo preenchimento
- `front_execUpdateData` // atualiza as linhas da base
- `front_popupateCheckBox` // atualiza os checkboxes para consulta
- `front_populateCheckBoxUpdated` // atualiza os checkboxes para consulta de contratos já preenchidos
- `front_handleHack` // função para identificar a linha correta para o update

- `mod_clearCheckBoxes` // limpeza dos checkbox para novo preenchimento
- `mod_fullScreen` // opção de tela cheia para melhor visualização do preenchimento
- `mod_wideScreen` // tela default
- `mod_handleFormule` // pega a formula com a linguagem DAX
- `mod_handleIncludeForm` // preenche a coluna A na sheet "BASE" para controle usado no preenchimento
