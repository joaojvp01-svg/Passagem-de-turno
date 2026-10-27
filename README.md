# Passagem de Turno · Estab. 40

Site da passagem de turno: lançamento do turno, relatórios do WMS, cadastro da equipe e as 3 páginas de status para enviar no WhatsApp.

Feito com [Vite](https://vitejs.dev) (HTML, CSS e JavaScript, sem framework). Não tem servidor nem banco de dados: tudo roda no navegador.

## Como publicar (GitHub + Vercel)

1. **GitHub:** crie um repositório **privado** (ex.: `passagem-turno`) e envie todos os arquivos desta pasta.
   - Pelo site: *Add file › Upload files*, arraste o conteúdo da pasta (não a pasta `node_modules`) e confirme.
   - Ou pelo terminal, dentro da pasta:
     ```bash
     git init
     git add .
     git commit -m "Passagem de turno"
     git branch -M main
     git remote add origin https://github.com/SEU_USUARIO/passagem-turno.git
     git push -u origin main
     ```
2. **Vercel:** *Add New › Project*, escolha o repositório e clique em **Deploy**. O `vercel.json` já informa tudo:
   - Framework: Vite
   - Build command: `npm run build`
   - Output directory: `dist`
3. **Banco compartilhado (Upstash Redis)** — sem ele, cada navegador guarda só os próprios dados:
   - No projeto da Vercel: *Storage › Connect Database* (ou *Marketplace › Upstash › Redis*). Pode ser um banco novo ou o mesmo já usado em outro projeto: as chaves deste site começam com `passagem:`.
   - Ao conectar, a Vercel cria sozinha `KV_REST_API_URL` e `KV_REST_API_TOKEN` (ou `UPSTASH_REDIS_REST_URL` e `UPSTASH_REDIS_REST_TOKEN`). Os dois pares funcionam.
   - Opcional: em *Settings › Environment Variables*, crie `CODIGO_ACESSO` com uma senha simples. O site pede essa senha uma vez em cada aparelho.
   - Depois de conectar, faça um *Redeploy*. No rodapé do site deve aparecer **"Salvo para toda a equipe"**.
4. A cada `git push` na branch `main`, a Vercel publica a nova versão sozinha.

## Ver sem instalar nada

`npm run build:arquivo` gera `dist-arquivo/index.html`: um único arquivo com tudo dentro, que abre com clique duplo em qualquer navegador. Serve para conhecer o site ou usar sem a Vercel. No Painel, o botão **Carregar exemplo** preenche tudo com dados fictícios.

## Rodar no computador (opcional)

Precisa do Node.js 18 ou mais novo.

```bash
npm install
npm run dev      # abre em http://localhost:5173
npm run build    # gera a pasta dist (o mesmo que a Vercel faz)
npm run build:arquivo  # gera dist-arquivo/index.html, arquivo único
npm run preview  # testa a pasta dist
```

## Como usar no turno

1. No topo, escolha o turno e confira a data.
2. **Equipe** (uma vez, ou quando o quadro mudar): importe a planilha com Nome, Função e Turno. O botão *Baixar modelo* gera o arquivo no formato certo.
3. **Relatórios:** suba a Movimentação de Estoque, o Relatório de Separação e a Conferência, cada um no seu campo.
4. **Lançamento:** registre as ausências, preencha os campos manuais e as observações.
5. **Painel:** confira os pontos de atenção e toque em *Enviar todas as páginas* (inclui a imagem do Planejamento VD, quando houver planejamento salvo). No celular abre o compartilhamento (WhatsApp); no computador, baixa as imagens.

## Relatórios: três envios separados

- **Movimentação de Estoque** (WMS, uma linha por movimentação): *Separação de picking* vira picking, *Onda* vira onda e *Semi-onda* vira semi-onda. Esses três pesos são os do VD (página 1) e os pesos de cada pessoa na página 3. Reposição, armazenagem e transferências aparecem só na mensagem de importação.
- **Relatório de Separação** (Separação Diária: USUARIO, NOME, NO_PEDIDOS, NO_ITENS, PESO_PICKING, PESO_ONDA, DATA): dá pedidos e itens por pessoa. Sem a movimentação, os pesos dele são usados.
- **Conferência** (Conferência Diária: USUARIO, NOME, NO_PEDIDOS, ITENS, PESO_TOTAL, DATA): dá toda a conferência.

Os dois relatórios diários costumam trazer o mês inteiro. O site usa a aba com a coluna DATA e pega só o dia escolhido no topo; se esse dia não estiver no arquivo, usa o mais recente e mostra um seletor para trocar o dia. Cada campo aceita só o seu relatório e avisa se o arquivo for do outro.

A página 3 lista **todos** os colaboradores, em três grupos:
- **Auxiliares** (separação de picking): função com "Auxiliar" no cadastro; sem cadastro, quem separou mais picking que onda.
- **Operadores** (onda e semi-onda): função com "Operador"/"Empilhadeira"; sem cadastro, quem fez mais onda e semi-onda.
- **Conferentes**, do relatório de Conferência.

O nome de cada login vem da coluna *Usuário WMS* da aba Equipe ou dos relatórios de Separação e Conferência (USUARIO + NOME). Login sem nome aparece em maiúsculas.

## Planejamento VD (módulo separado)

Aba própria, com código em `src/planejamento.js`, que reproduz o **RESUMO COMPLETO** da planilha *Acompanhamento de Produção*:

- **Planejamento de carregamento**: NFs por veículo (colunas Sequência, Nota Fiscal, Pedido, Razão Social, Cidade, Peso…). A data de entrega vem do título "dd/mm/aaaa - … ENTREGAS …".
- **Faturamento** (opcional, mesmo modelo do consolidado): dá o PEDIDO+SUB e o peso bruto de cada NF. Sem ele, o site usa o pedido e o peso do próprio planejamento.
- **Pedidos Query**: status de cada pedido no WMS (colunas PEDIDO, SUB e DESCRIÇÃO_STATUS). Pedido do planejamento que não está na Query conta como **Conferido**, igual à planilha.

Pode subir cada arquivo separado ou a planilha `.xlsm` inteira; o site reconhece cada aba pelo cabeçalho. Os 10 últimos planejamentos ficam salvos, e o **dia anterior** é recalculado com a Query mais recente (como a aba RESUMO DIA ANTERIOR). Cada Query enviada registra uma linha no histórico.

Grupos: **Pendente** = Em Curso de Separação + Planejamento concluído + Reservado · **Em Curso** = Em Curso de Conferência + Em transição + Separado · **Finalizado** = Conferido.

Como a Pedidos Query ainda não tem modelo fechado, a leitura procura as colunas pelo nome (`lerQuery` em `src/planejamento.js`). Se o arquivo vier com outros nomes, é ali que se ajusta.

## Onde ficam os dados

- **No site da Vercel com o banco ligado:** lançamentos, relatórios, ausências, equipe e planejamento ficam no Upstash (`api/dados.js`) e aparecem em qualquer computador ou celular. Cada dia é gravado por turno, para dois turnos lançando ao mesmo tempo não apagarem um ao outro.
- O navegador guarda uma cópia de trabalho: sem internet o site continua funcionando e envia tudo quando a conexão volta. Os outros aparelhos buscam as novidades ao abrir o site, ao voltar para a aba e a cada minuto.
- O rodapé mostra a situação: *Salvo para toda a equipe*, *Sem conexão* ou *Salvo só neste navegador* (arquivo HTML, artefato ou site sem banco).
- O arquivo de faturamento não sobe para o banco (é grande); ele já fica aplicado no planejamento salvo.
- Nenhum nome de colaborador fica no código. Por isso o repositório não guarda dados da operação (o `.gitignore` também bloqueia `.xlsx` e `.csv`).

## Regras de cálculo

| Campo | Como é calculado |
|---|---|
| Total separado | picking + onda do relatório de separação |
| Total pendente (VD) | total programado − total separado |
| Carros pendentes (VD e VET) | carros programados − carros expedidos |
| SKU's separados | soma de NO_ITENS do relatório |
| Absenteísmo (%) | ausências (sem contar presença parcial, férias e afastamento) ÷ colaboradores do turno fora de férias/afastamento |
| Dias de ausência | `dia atual // total de dias`, contado a partir do início lançado. Ausências de mais de um dia ficam em `passagemTurno.afast` e aparecem sozinhas em todos os dias do período, no turno em que foram lançadas. *Encerrar* fecha o período no dia anterior. |

## Estrutura

```
index.html          telas (Painel, Lançamento, Relatórios, Equipe, Apoio)
src/main.js         lógica: cálculos, leitura de planilhas, geração das imagens
src/nuvem.js        sincronização com o banco compartilhado
api/dados.js        função da Vercel que lê e grava no Upstash Redis
src/style.css       visual
src/assets/         logo e fotos dos pets
public/             favicon e robots.txt (site fora do Google)
vercel.json         configuração da Vercel
```

## Observações técnicas

- `xlsx` está na versão 0.18.5, a última publicada no npm. O `npm audit` aponta um alerta conhecido dessa versão, ligado à leitura de arquivos maliciosos; aqui só são lidas planilhas da própria operação. Para atualizar, troque no `package.json` por `"xlsx": "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz"`.
- As fontes (Fira Sans, Fira Sans Condensed e Sora) vêm junto no build. O site não depende do Google Fonts.
- O site não aparece no Google (`robots.txt` e cabeçalho `noindex`), mas qualquer pessoa com o link consegue abrir. Para exigir login, use a proteção por senha da Vercel (plano pago) ou coloque o site atrás do login da empresa.
