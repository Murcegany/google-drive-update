# Gmail to Google Drive Integration

Este projeto é um script em C# que faz a integração entre o Gmail e o Google Drive. Ele busca e-mails não lidos, faz o upload de anexos PDF para uma pasta específica no Google Drive, e em seguida, marca esses e-mails como lidos e os exclui. Sendo um projeto de exemplo, sendo necessário substituir por campos verdadeiros.

### Como Funciona

Processamento de E-mails

- O programa pesquisa por e-mails não lidos com anexos PDF de um remetente específico.
- Ele ignora anexos que contenham certas palavras-chave no nome do arquivo (como "Boleto").
- Os arquivos PDF são enviados para uma pasta específica no Google Drive.
- Após o upload, o e-mail é marcado como lido e excluído.

Evitando Duplicações

- Antes de fazer o upload, o programa verifica se o arquivo já existe no Google Drive. Se o arquivo já estiver lá, ele não será carregado novamente.

Persistência de Processamento

- O programa salva o ID do último e-mail processado em um arquivo (lastProcessedMessageId.txt) para garantir que e-mails anteriores não sejam reprocessados.

Modificando o Filtro de E-mails

- Você pode ajustar o filtro de e-mails alterando a string de consulta na linha:
   ```bash
    string query = "from:email_do_remetente@example.com has:attachment is:unread";
   ```

Mudando a Pasta de Destino no Google Drive
- Para alterar a pasta onde os arquivos são salvos no Google Drive, modifique a variável Parents no código:

   ```bash
    Parents = new List<string> { "sua_pasta_id" };
   ```


## Requisitos

- .NET Core SDK 3.1 ou superior
- Conta Google com acesso ao Gmail e Google Drive
- Credenciais de API da Google Cloud Platform (GCP)

## Configuração

### 1. Criar Projeto na Google Cloud Platform

1. Acesse o [Google Cloud Console](https://console.cloud.google.com/).
2. Crie um novo projeto.
3. Ative as APIs do Gmail e do Google Drive para o seu projeto.
   - **Gmail API**: Vá para "APIs & Services" > "Library" e procure por "Gmail API". Clique em "Enable".
   - **Google Drive API**: Vá para "APIs & Services" > "Library" e procure por "Google Drive API". Clique em "Enable".
4. Configure uma tela de consentimento OAuth 2.0 para seu projeto.
   - Vá para "APIs & Services" > "OAuth consent screen" e siga os passos para configurar.
5. Crie credenciais para o OAuth 2.0:
   - Vá para "APIs & Services" > "Credentials".
   - Clique em "Create Credentials" > "OAuth 2.0 Client IDs".
   - Escolha "Desktop app" como tipo de aplicação.
   - Faça o download do arquivo `credentials.json` e salve-o na pasta raiz do projeto.

### 2. Clonar o Repositório e Configurar

1. Clone este repositório para sua máquina local:

   ```bash
   git clone https://github.com/seu-usuario/gmail-drive-integration.git
   cd gmail-drive-integration
   ```

2. Coloque o arquivo credentials.json baixado anteriormente na pasta raiz do projeto.

3. (Opcional) Modifique as variáveis email_do_remetente@example.com e sua_pasta_id no código conforme necessário.

### 3. Compilar e Executar

1. Abra o terminal na pasta do projeto e execute:

   ```bash
    dotnet build
    dotnet run
   ```

2. Ao rodar pela primeira vez, o programa abrirá uma janela do navegador para você autorizar o acesso às suas contas do Gmail e Google Drive. Conceda as permissões necessárias.

