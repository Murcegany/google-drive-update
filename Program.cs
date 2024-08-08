using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Gmail.v1.Data;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

class Program
{
    // Defina os escopos necessários para acessar e modificar e-mails e arquivos no Google Drive
    static string[] Scopes = { GmailService.Scope.GmailModify, DriveService.Scope.Drive, DriveService.Scope.DriveFile };
    static string ApplicationName = "Gmail to Google Drive Integration";
    static string lastProcessedMessageIdFilePath = "lastProcessedMessageId.txt";

    /// <summary>
    /// Ponto de entrada principal para o aplicativo.
    /// Inicializa os serviços do Gmail e Google Drive e inicia o processamento de e-mails em um loop contínuo.
    /// </summary>
    /// <param name="args">Argumentos da linha de comando.</param>
    /// <returns>Tarefa representando a operação assíncrona.</returns>
    static async Task Main(string[] args)
    {
        // Obtém as credenciais do usuário
        UserCredential credential = await GetCredentialsAsync();

        if (credential == null)
        {
            Console.WriteLine("Não foi possível obter as credenciais.");
            return;
        }

        try
        {
            // Inicializa os serviços do Gmail e Google Drive
            var gmailService = InitializeGmailService(credential);
            var driveService = InitializeDriveService(credential);

            Console.WriteLine("Serviços do Gmail e Google Drive inicializados.");

            // Loop contínuo para processar e-mails a cada 5 minutos
            while (true)
            {
                await ProcessEmails(gmailService, driveService);
                await Task.Delay(TimeSpan.FromMinutes(5));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao inicializar serviços: {ex.Message}");
        }
    }

    /// <summary>
    /// Carrega as credenciais do usuário a partir do arquivo de configuração.
    /// </summary>
    /// <returns>Credenciais do usuário ou null se falhar.</returns>
    static async Task<UserCredential> GetCredentialsAsync()
    {
        try
        {
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true));
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao obter credenciais: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Inicializa o serviço do Gmail.
    /// </summary>
    /// <param name="credential">Credenciais do usuário.</param>
    /// <returns>Instância do GmailService.</returns>
    static GmailService InitializeGmailService(UserCredential credential)
    {
        return new GmailService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });
    }

    /// <summary>
    /// Inicializa o serviço do Google Drive.
    /// </summary>
    /// <param name="credential">Credenciais do usuário.</param>
    /// <returns>Instância do DriveService.</returns>
    static DriveService InitializeDriveService(UserCredential credential)
    {
        return new DriveService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = ApplicationName,
        });
    }

    /// <summary>
    /// Recupera o ID do último e-mail processado armazenado em um arquivo.
    /// </summary>
    /// <returns>ID do último e-mail processado ou null se o arquivo não existir.</returns>
    static async Task<string> GetLastProcessedMessageIdAsync()
    {
        if (File.Exists(lastProcessedMessageIdFilePath))
        {
            return await File.ReadAllTextAsync(lastProcessedMessageIdFilePath);
        }
        return null;
    }

    /// <summary>
    /// Armazena o ID do último e-mail processado em um arquivo.
    /// </summary>
    /// <param name="messageId">ID do e-mail a ser armazenado.</param>
    /// <returns>Tarefa representando a operação assíncrona.</returns>
    static async Task SetLastProcessedMessageIdAsync(string messageId)
    {
        try
        {
            await File.WriteAllTextAsync(lastProcessedMessageIdFilePath, messageId);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao salvar o ID do último e-mail processado: {ex.Message}");
        }
    }

    /// <summary>
    /// Processa os e-mails não lidos do Gmail, faz o upload de anexos PDF para o Google Drive e marca os e-mails como lidos.
    /// </summary>
    /// <param name="gmailService">Serviço do Gmail.</param>
    /// <param name="driveService">Serviço do Google Drive.</param>
    /// <returns>Tarefa representando a operação assíncrona.</returns>
    static async Task ProcessEmails(GmailService gmailService, DriveService driveService)
    {
        Console.WriteLine("Iniciando o processamento de e-mails...");

        // Recupera o ID da última mensagem processada
        string lastProcessedMessageId = await GetLastProcessedMessageIdAsync();
        string query = "from:email_do_remetente@example.com has:attachment is:unread"; // Substitua por seu filtro de e-mail
        if (!string.IsNullOrEmpty(lastProcessedMessageId))
        {
            query += $" after:{DateTime.UtcNow.AddDays(-1).ToString("yyyy/MM/dd")}";
        }

        try
        {
            var request = gmailService.Users.Messages.List("me");
            request.Q = query;
            var response = await request.ExecuteAsync();

            if (response.Messages != null && response.Messages.Count > 0)
            {
                Console.WriteLine($"{response.Messages.Count} mensagens encontradas.");
                foreach (var messageItem in response.Messages)
                {
                    try
                    {
                        var message = await gmailService.Users.Messages.Get("me", messageItem.Id).ExecuteAsync();

                        if (message.Payload.Parts != null)
                        {
                            foreach (var part in message.Payload.Parts)
                            {
                                if (!string.IsNullOrEmpty(part.Filename) && part.Filename.EndsWith(".pdf") && part.Body != null && part.Body.AttachmentId != null)
                                {
                                    if (part.Filename.Contains("IgnoreKeyword"))
                                    {
                                        Console.WriteLine($"Anexo ignorado por conter 'IgnoreKeyword': {part.Filename}");
                                        continue;
                                    }

                                    if (await FileExistsInDrive(driveService, part.Filename))
                                    {
                                        Console.WriteLine($"Arquivo {part.Filename} já existe no Google Drive. Ignorando upload.");
                                        continue;
                                    }

                                    Console.WriteLine($"Processando anexo: {part.Filename}");
                                    var attachment = await gmailService.Users.Messages.Attachments.Get("me", message.Id, part.Body.AttachmentId).ExecuteAsync();
                                    byte[] data = Convert.FromBase64String(attachment.Data.Replace('-', '+').Replace('_', '/'));

                                    var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                                    {
                                        Name = part.Filename,
                                        Parents = new List<string> { "sua_pasta_id" } // Substitua pelo ID da sua pasta no Google Drive
                                    };

                                    using (var stream = new MemoryStream(data))
                                    {
                                        var uploadRequest = driveService.Files.Create(fileMetadata, stream, part.MimeType);
                                        uploadRequest.Fields = "id";
                                        var file = await uploadRequest.UploadAsync();

                                        if (file.Status == Google.Apis.Upload.UploadStatus.Completed)
                                        {
                                            Console.WriteLine($"Upload bem-sucedido para {part.Filename}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Falha no upload para {part.Filename}");
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"Anexo inválido ou não encontrado na mensagem {message.Id}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine($"Nenhum anexo encontrado na mensagem {message.Id}");
                        }

                        // Armazena o ID da mensagem processada
                        await SetLastProcessedMessageIdAsync(messageItem.Id);

                        // Marca a mensagem como lida e a exclui
                        var mods = new ModifyMessageRequest { RemoveLabelIds = new List<string> { "UNREAD" } };
                        await gmailService.Users.Messages.Modify(mods, "me", message.Id).ExecuteAsync();
                        Console.WriteLine($"Mensagem {message.Id} marcada como lida.");

                        try
                        {
                            await gmailService.Users.Messages.Delete("me", message.Id).ExecuteAsync();
                            Console.WriteLine($"Mensagem {message.Id} excluída.");
                        }
                        catch (Exception deleteEx)
                        {
                            Console.WriteLine($"Erro ao excluir a mensagem {message.Id}: {deleteEx.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Erro ao processar a mensagem {messageItem.Id}: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Nenhuma mensagem encontrada.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao processar e-mails: {ex.Message}");
        }
    }

    /// <summary>
    /// Verifica se um arquivo com o mesmo nome já existe no Google Drive.
    /// </summary>
    /// <param name="driveService">Serviço do Google Drive.</param>
    /// <param name="fileName">Nome do arquivo.</param>
    /// <returns>True se o arquivo existir, False caso contrário.</returns>
    static async Task<bool> FileExistsInDrive(DriveService driveService, string fileName)
    {
        try
        {
            var request = driveService.Files.List();
            request.Q = $"name = '{fileName.Replace("'", "\\'")}' and trashed = false";
            request.Fields = "files(id, name)";
            var result = await request.ExecuteAsync();
            return result.Files.Count > 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Erro ao verificar existência de arquivo no Drive: {ex.Message}");
            return false;
        }
    }
}
