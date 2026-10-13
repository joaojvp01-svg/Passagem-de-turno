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
   - Não precisa de variáveis de ambiente.
3. A cada `git push` na branch `main`, a Vercel publica a nova versão sozinha.

## Rodar no computador (opcional)

Precisa do Node.js 18 ou mais novo.

```bash
npm install
npm run dev      # abre em http://localhost:5173
npm run build    # gera a pasta dist (o mesmo que a Vercel faz)
npm run preview  # testa a pasta dist
```

## Como usar no turno

1. No topo, escolha o turno e confira a data.
2. **Equipe** (uma vez, ou quando o quadro mudar): importe a planilha com Nome, Função e Turno. O botão *Baixar modelo* gera o arquivo no formato certo.
3. **Relatórios:** suba o relatório do WMS (separação e, se houver, conferência).
4. **Lançamento:** registre as ausências, preencha os campos manuais e as observações.
5. **Painel:** confira os pontos de atenção e toque em *Enviar as 3 páginas*. No celular abre o compartilhamento (WhatsApp); no computador, baixa as imagens.

## Onde ficam os dados

- Lançamentos, relatórios e ausências ficam salvos **no navegador de quem usa**, separados por data e turno. O cadastro da equipe também.
- Outro computador ou celular não enxerga esses dados, e limpar os dados do navegador apaga o histórico.
- Nenhum nome de colaborador fica no código. Por isso o repositório não guarda dados da operação (o `.gitignore` também bloqueia `.xlsx` e `.csv`).

Para todos usarem os mesmos dados, o próximo passo é ligar o site a um banco (ex.: Supabase ou Vercel Postgres).

## Regras de cálculo

| Campo | Como é calculado |
|---|---|
| Total separado | picking + onda do relatório de separação |
| Total pendente (VD) | total programado − total separado |
| Carros pendentes (VD e VET) | carros programados − carros expedidos |
| SKU's separados | soma de NO_ITENS do relatório |
| Absenteísmo (%) | ausências (sem contar presença parcial) ÷ colaboradores do turno |

## Estrutura

```
index.html          telas (Painel, Lançamento, Relatórios, Equipe, Apoio)
src/main.js         lógica: cálculos, leitura de planilhas, geração das imagens
src/style.css       visual
src/assets/         logo e fotos dos pets
public/             favicon e robots.txt (site fora do Google)
vercel.json         configuração da Vercel
```

## Observações técnicas

- `xlsx` está na versão 0.18.5, a última publicada no npm. O `npm audit` aponta um alerta conhecido dessa versão, ligado à leitura de arquivos maliciosos; aqui só são lidas planilhas da própria operação. Para atualizar, troque no `package.json` por `"xlsx": "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz"`.
- As fontes (Fira Sans, Fira Sans Condensed e Sora) vêm junto no build. O site não depende do Google Fonts.
- O site não aparece no Google (`robots.txt` e cabeçalho `noindex`), mas qualquer pessoa com o link consegue abrir. Para exigir login, use a proteção por senha da Vercel (plano pago) ou coloque o site atrás do login da empresa.
