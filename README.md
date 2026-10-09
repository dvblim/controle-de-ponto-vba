# Controle de Ponto em VBA

Painel de controle de ponto e banco de horas para Excel, feito em VBA com **UserForm** e **módulos de classe**.
Você lança os horários do dia, vê a saída sugerida e acompanha o saldo do banco por mês, tudo num painel escuro dentro do próprio Excel.

## O que o painel faz

- **Lançar ponto:** data, tipo do dia, entrada, saída para o almoço, retorno e saída.
  O botão **Agora** marca a hora atual no primeiro horário vazio do dia e já salva.
- **Saída sugerida:** mostra o horário em que você fecha a jornada e a faixa em que o banco não muda
  (ex.: *Saída sugerida hoje: 18:02 · sem afetar o banco entre 17:52 e 18:12*).
- **Meses que entram na conta:** clique nos meses para escolher o que entra nos indicadores.
- **Indicadores:** saldo dos meses marcados, horas positivas e negativas, total trabalhado contra o previsto e banco total.
- **Saldo por mês:** gráfico de barras com o saldo de cada mês marcado.
- **Feriados:** feriados nacionais calculados automaticamente (inclusive Carnaval, Sexta-feira Santa e Corpus Christi)
  e feriados locais cadastrados na aba Config. Mostra também os próximos feriados.
- **Lista de dias:** todos os dias dos meses marcados. Dias úteis sem registro aparecem como **Sem registro**.
  Clique num dia para editá-lo.

## Regra do banco de horas

| Item | Regra |
|---|---|
| Jornada | 8:48 trabalhadas por dia |
| Almoço | não conta como trabalhado e não tem horário nem duração fixos |
| Trabalhado | (saída almoço − entrada) + (saída − retorno) |
| Tolerância | diferença de até 10 min no dia, para mais ou para menos, vale 0 |
| Passou da tolerância | entra a diferença inteira: os 10 min + o que passou deles |
| Feriado, Folga, Atestado | 0 no banco |
| Falta | desconta a jornada inteira (−8:48) |

Exemplos: saiu 8 min antes, banco 0. Saiu 11 min antes, banco −0:11. Ficou 15 min a mais, banco +0:15.

> A jornada, a tolerância, o almoço padrão (usado só na sugestão de saída) e o saldo inicial do banco
> ficam na aba **Config** e podem ser alterados.
> A regra de tolerância acima é a de um acordo de empresa. A CLT (art. 58, §1º) prevê até 5 minutos
> por marcação, com limite de 10 minutos diários. Ajuste os valores conforme a regra que vale para você.

## Instalação

Requisitos: Excel para Windows (2010 ou mais recente).

1. Baixe o repositório (**Code → Download ZIP**) e extraia.
2. Abra `Controle_de_Ponto_modelo.xlsx` e salve como **Pasta de Trabalho Habilitada para Macro (.xlsm)**.
3. Libere o acesso ao projeto VBA:
   **Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade →
   Configurações de Macro →** marque *Confiar no acesso ao modelo de objeto do projeto do VBA*.
4. Pressione **Alt+F11**, vá em **Arquivo → Importar arquivo** e escolha `src/modInstalador.bas`.
5. Pressione **Alt+F8**, rode **InstalarPainel** e selecione a pasta `src`.
6. Salve. O botão **Abrir painel** aparece na aba Registros.

Para atualizar para uma versão nova, substitua a pasta `src` e rode **InstalarPainel** de novo.
O instalador remove os módulos antigos e importa os novos.

Já tem uma planilha de ponto própria? O painel funciona com qualquer arquivo que tenha as abas
**Registros** (com a tabela `tbRegistros`) e **Config** (com os nomes `Jornada`, `TolMin`,
`AlmocoPadrao` e `SaldoInicialMin`) no mesmo formato da planilha modelo.

## Como usar

- Digite os horários só com números: `0814` vira `08:14`. **Enter** salva o dia.
- **Agora** é o jeito mais rápido de bater o ponto: abre o dia de hoje e preenche o próximo horário.
- **Excluir** apaga o registro da data que está no campo Data.
- Tipos de dia: **Normal**, **Feriado**, **Folga**, **Atestado** e **Falta**.
- Cadastre feriados municipais ou estaduais na tabela de feriados da aba **Config**.

## Estrutura do projeto

```
Controle_de_Ponto_modelo.xlsx   planilha modelo (tabela vazia, fórmulas e configurações)
src/
  modInstalador.bas      instalador: cria o formulário, importa os módulos e coloca o botão
  modPainel.bas          abre o painel, cores, formatação de horas e criação de controles
  frmPainel_codigo.txt   código do formulário frmPainel (o instalador cria o formulário)
  clsRegra.cls           regra do banco: jornada, tolerância e saída sugerida
  clsDia.cls             um dia de registro e seus cálculos
  clsRegistros.cls       leitura e gravação da tabela tbRegistros
  clsFeriados.cls        feriados nacionais (com Páscoa e datas móveis) e feriados locais
  clsBotao.cls           eventos dos botões
  clsMesCard.cls         eventos dos cartões de mês
  clsCampo.cls           eventos dos campos (máscara de hora e data, Enter para salvar)
  clsLista.cls           eventos da lista de dias
```

### Como o código está organizado

Todos os controles do painel são criados por código quando o formulário abre.
Como controles criados assim não têm procedimentos de evento no formulário, cada tipo de
controle ganha uma classe com `WithEvents`, que repassa o evento para um método público do formulário:

| Classe | Evento | Chama no formulário |
|---|---|---|
| `clsBotao` | clique | `Acao "salvar"`, `Acao "agora"`... |
| `clsMesCard` | clique no cartão | `AlternarMes` |
| `clsCampo` | digitação, Enter, mudança | `DataMudou`, `CamposMudaram`, `Acao "salvar"` |
| `clsLista` | clique na lista | `DiaSelecionado` |

A regra de negócio fica isolada em `clsRegra`. O formulário só exibe os dados, e a gravação
na planilha fica em `clsRegistros`. O painel grava apenas as colunas de entrada
(Data, Tipo, os quatro horários e Observação), então as fórmulas da planilha continuam válidas.

### Sobre a codificação dos arquivos

Os arquivos de `src` estão em **Windows-1252 com quebras de linha CRLF**, o formato que o editor do VBA
importa sem estragar os acentos. O `.gitattributes` impede que o Git converta esses arquivos.
Se você editar o código fora do Excel, mantenha essa codificação.

## Licença

Distribuído sob a licença MIT. O texto oficial está em [LICENSE](LICENSE) e há uma tradução em [LICENCA.md](LICENCA.md).
