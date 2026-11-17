# App Saúde – como instalar (uns 10 minutos, uma vez só)

1. Crie uma **planilha Google nova** (ex.: "Saúde").
2. Na planilha: **Extensões > Apps Script**.
3. Apague o conteúdo do `Código.gs` e cole o conteúdo de `Code.gs` desta pasta.
4. Clique em **+ > HTML**, dê o nome **Index** (exatamente assim) e cole o conteúdo de `Index.html`.
   Depois **+ > Script**, nome **Strava**, e cole o conteúdo de `Strava.gs`.
5. Salve e recarregue a planilha. Use o menu **Saúde > Criar administrador (primeiro uso)**: seu nome, o usuário para entrar (ex.: `gilberto`) e a senha (mínimo 8 caracteres). Autorize quando o Google pedir. Isso também cria as abas.
6. No Apps Script: **Implantar > Nova implantação > App da Web**. Executar como **Eu**, acesso **Qualquer pessoa**. Copie a URL que termina em `/exec`.
7. No celular, abra o site (GitHub Pages) ou a própria URL `/exec`, entre com usuário e senha e use **⋮ > Adicionar à tela inicial**. Faça o mesmo no PC, se quiser.

Os dados ficam na planilha, então celular e PC mostram sempre a mesma coisa.

## Outros usuários (ex.: esposa)
- No app, **Ajustes > Usuários** (só aparece para administrador): nome, usuário, senha inicial. Passe a senha pessoalmente; ela troca em **Ajustes > Trocar minha senha**.
- Cada usuário vê só os próprios dados (peso, treinos, diário, remédio, meta e altura). Na planilha, todos ficam nas mesmas abas, com a coluna **Usuário** (U1, U2…).
- **Nova senha** redefine a senha de alguém que esqueceu. **Desativar** bloqueia o acesso sem apagar os dados.
- Esqueceu a sua senha de administrador: menu **Saúde > Redefinir senha de um usuário**.

## Remédio (tirzepatida ou semaglutida)
- Aba **Remédio**: data, remédio, dose em mg (botões com as doses comerciais, ou digite) e local da aplicação. A tela já sugere o remédio e a dose da última aplicação.
- O Painel mostra a dose atual, a última aplicação e a próxima (7 dias depois). No gráfico de peso, traços verticais marcam cada mudança de dose.
- É só registro. Dose e troca de dose são decisão do médico.

**Atualizar o app depois:** cole o código novo e use **Implantar > Gerenciar implantações > editar > Nova versão**. A URL não muda.
**Perdeu o celular:** troque a sua senha no app (derruba as suas sessões) ou use o menu **Saúde > Encerrar todas as sessões** (derruba de todos).
**Não compartilhe a URL com a senha.**

## Rotina
- **Balança Beurer BF 195:** depois de pesar, a balança mostra em sequência gordura %, água %, músculo %, massa óssea e kcal. Lance tudo na aba Peso. Pese de manhã, em jejum e descalço, para os números serem comparáveis.
- **Relógio:** os treinos chegam sozinhos pelo Strava (abaixo). Sono e passos do dia: lançar no Diário ou reimportar o histórico de tempos em tempos.

## Histórico do Mi Fitness (uma vez, por pessoa)
1. Gere `mifit_atividades.csv` e `mifit_diario.csv` com `python converter_mifitness.py "<pasta da cópia de dados do Mi Fitness>"` (esses arquivos têm dados pessoais: nunca suba para o GitHub).
2. Na planilha: **Arquivo > Importar > Fazer upload** de cada arquivo, com **Inserir nova(s) página(s)** e **Converter texto em números, datas e fórmulas = Não**.
3. Menu **Saúde > Importar histórico do Mi Fitness** e informe de qual usuário são os dados. Depois apague as duas abas `mifit_…`.
4. Para atualizar depois (sono e passos novos): peça uma nova cópia de dados no Mi Fitness, gere de novo e repita. O que já existe não duplica.

## Strava (treinos automáticos, cada pessoa com o seu)
1. No **Mi Fitness**: Perfil > Contas de terceiros > **Strava** > conectar.
2. Em **strava.com/settings/api** (pelo computador, logado na conta Strava da pessoa): crie um app. Nome: "Saúde pessoal", site: qualquer um, **Authorization Callback Domain: `script.google.com`**. (App novo no Strava aceita um atleta só, por isso cada pessoa cria o seu.)
3. No app Saúde, **Ajustes > Relógio (Strava)**: cole o Client ID e o Client Secret e salve.
4. Toque em **Conectar ao Strava**, autorize (deixe marcada a leitura de atividades) e volte ao app.
5. Uma vez só, o administrador liga a busca automática: menu **Saúde > Ativar sincronização do Strava**. Ela passa de hora em hora por todos os usuários conectados. No app há o botão **Buscar treinos agora**.
A primeira busca olha os últimos 30 dias. Treino que já veio do histórico não entra de novo.
