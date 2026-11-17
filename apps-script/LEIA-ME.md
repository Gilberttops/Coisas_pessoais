# App Saúde – como instalar (uns 10 minutos, uma vez só)

1. Crie uma **planilha Google nova** (ex.: "Saúde").
2. Na planilha: **Extensões > Apps Script**.
3. Apague o conteúdo do `Código.gs` e cole o conteúdo de `Code.gs` desta pasta.
4. Clique em **+ > HTML**, dê o nome **Index** (exatamente assim) e cole o conteúdo de `Index.html`.
   Depois **+ > Script**, nome **Strava**, e cole o conteúdo de `Strava.gs`.
5. Salve. No seletor de função, escolha **configurarPlanilha** e clique em **Executar**. Autorize quando o Google pedir. Pode aparecer um erro de "getUi" se você rodar pelo editor; se acontecer, recarregue a planilha e use o menu **Saúde > Criar/conferir abas**.
6. Recarregue a planilha e use **Saúde > Definir senha de acesso** (mínimo 8 caracteres).
7. No Apps Script: **Implantar > Nova implantação > App da Web**. Executar como **Eu**, acesso **Qualquer pessoa**. Copie a URL que termina em `/exec`.
8. No celular, abra essa URL no Chrome, entre com a senha e use **⋮ > Adicionar à tela inicial**. Faça o mesmo no PC, se quiser.

Os dados ficam na planilha, então celular e PC mostram sempre a mesma coisa.

**Atualizar o app depois:** cole o código novo e use **Implantar > Gerenciar implantações > editar > Nova versão**. A URL não muda.
**Perdeu o celular:** menu **Saúde > Encerrar todas as sessões**.
**Não compartilhe a URL com a senha.** Se a URL vazar, troque a senha.

## Rotina
- **Balança Beurer BF 195:** depois de pesar, a balança mostra em sequência gordura %, água %, músculo %, massa óssea e kcal. Lance tudo na aba Peso. Pese de manhã, em jejum e descalço, para os números serem comparáveis.
- **Relógio:** os treinos chegam sozinhos pelo Strava (abaixo). Sono e passos do dia: lançar no Diário ou reimportar o histórico de tempos em tempos.

## Histórico do Mi Fitness (uma vez)
1. Gere `mifit_atividades.csv` e `mifit_diario.csv` com `python converter_mifitness.py "<pasta da cópia de dados do Mi Fitness>"` (esses arquivos têm dados pessoais: nunca suba para o GitHub).
2. Na planilha: **Arquivo > Importar > Fazer upload** de cada arquivo, com **Inserir nova(s) página(s)** e **Converter texto em números, datas e fórmulas = Não**.
3. Menu **Saúde > Importar histórico do Mi Fitness**. Depois pode apagar as duas abas `mifit_…`.
4. Para atualizar depois (sono e passos novos): peça uma nova cópia de dados no Mi Fitness, rode `python converter_mifitness.py "<pasta nova>"` e repita. O que já existe não duplica.

## Strava (treinos automáticos)
1. No **Mi Fitness**: Perfil > Contas de terceiros (ou Configurações > Aplicativos conectados) > **Strava** > conectar.
2. Em **strava.com/settings/api** (pelo computador): crie um app. Nome: "Saúde pessoal", site: qualquer um, **Authorization Callback Domain: `script.google.com`**.
3. Na planilha, menu **Saúde > Strava > 1. Configurar chaves**: cole o Client ID e o Client Secret dessa página.
4. **Saúde > Strava > 2. Conectar conta**: clique no link, autorize (deixe marcada a leitura de atividades) e feche a aba.
5. **Saúde > Strava > 3. Ativar sincronização automática**. Pronto: de hora em hora. Na aba Ajustes do app tem o botão **Buscar treinos agora**.
A primeira busca olha os últimos 30 dias. Treino que já veio do histórico não entra de novo.
