
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System;
using OpenPop.Mime;
using System.Collections.Generic;
using System.Net.Mail;

namespace ReadingMails
{
    class Program
    {
        //Mapear todos os niveis 
        private static string GetBody(IEnumerable<MessagePart> parts)
        {
            foreach(var item in parts)
            {
                if(item.MessageParts != null && item.MessageParts.Any())
                {
                    return GetBody(item.MessageParts);
                }
                else
                {
                    if (item.IsText)
                    {

                        Encoding valueEncoding = Encoding.GetEncoding(item.BodyEncoding.BodyName.ToUpper());
                        Encoding utfEncoding = Encoding.UTF8;
                        
                        //  Obtém os bytes da string UTF 
                        byte[] bytesUtf = Encoding.Convert(valueEncoding, utfEncoding, item.Body);

                        // Obtém a string ISO a partir do array de bytes convertido
                        var texto = utfEncoding.GetString(bytesUtf);
                        //**********

                        return texto.Replace("\r", "<br>").Replace("\n", " ");
                    }
                }
            }
            return "";
        }
static void Main(string[] args)
        {
            var sustenido = (char)35;

            //Filtrar o assunto com esta Regex
            Regex regex = new Regex(@"\" + sustenido + @"\d+");

            const string host = "email-ssl.com.br";
            const int port = 995;
            const bool useSsl = true;

            const string username = "milton.silva@teste.com.br";
            const string password = "Password";

            using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                client.Connect(host, port, useSsl);

                client.Authenticate(username, password);

                int messageCount = client.GetMessageCount();

                var messages = new List<Message>(messageCount);

                var message = client.GetMessage(messageCount);
                var match = regex.IsMatch(message.Headers.Subject);
                int pendencia_id = 0;
                //Percorrer todas as mensagens
                for (int i = messageCount; i > 0; i--)
                {
                    message = client.GetMessage(i);
                    match = regex.IsMatch(message.Headers.Subject);
                    if(match)
                    {
                        var Teste = regex.Match(message.Headers.Subject);
                        pendencia_id = int.Parse(Teste.Value.Replace(sustenido.ToString(), string.Empty));
                    }
                    else
                    {
                        pendencia_id = 0;

                    }

                    Console.WriteLine(pendencia_id);
                    Console.WriteLine(message.Headers.Subject);

                    messages.Add(message);
                    message.ToMailMessage();

                    System.Net.NetworkCredential testCreds = new System.Net.NetworkCredential("Login", "Password", "Empresa.LOCAL");

                    System.Net.CredentialCache testCache = new System.Net.CredentialCache();

                    testCache.Add(new Uri("\\\\10.0.1.3"), "Basic", testCreds);


                    //caminho do folder
                    string folder = @"\\10.0.1.3\Comum\Milton.Silva'";

                    //Criar folder  
                    System.IO.Directory.CreateDirectory(folder);
                    //salvar arquivo
                    message.Save(new System.IO.FileInfo(folder+DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss")+".eml"));

                    //client.DeleteMessage(messageCount);

                    MailMessage mailMessage = new MailMessage();
                    //Endereço que irá aparecer no e-mail do usuário 
                    mailMessage.From = new MailAddress("milton.silva@teste.com.br");
                    //destinatarios do e-mail, para incluir mais de um basta separar por ponto e virgula  

                    string to = "netoxp70@gmail.com";
                    mailMessage.To.Add(to);

                    //Assunto: o chamado “#98946: Balanca com conversor TCP” requer sua atenção.
                    // string assunto = "enviado por meio do simple farm";
                    mailMessage.Subject = "O chamado " + message.Headers.Subject + " requer sua atenção";
                    mailMessage.IsBodyHtml = true;
                    //conteudo do corpo do e-mail 
                    string mensagem = "<h1>Prezado</h1>" +
                                  "<p>Nosso processo automatizado recebeu este e-mail e " +
                                  "anexou no chamado e registrou que o mesmo foi encaminhado para sua atuação.</p>" +
                                  "<p>Favor ajustar o status imediatamente(visto que o cliente monitora pelo Portal)" +
                                  "e atuar no atendimento.</p> <p>Grato.</p><p>Processo Automatizado de Recepção de E-mail</p><br>";

                    var parts = message.MessagePart.MessageParts;

                    var bodyText = GetBody(parts);
 
                    mailMessage.Body = mensagem + bodyText;

                    //Salvando Imagem no folder
                    foreach (MessagePart emailAttachment in message.FindAllAttachments())
                    {
                        //-Definir variáveis
                        string OriginalFileName = emailAttachment.FileName;
                        string ContentID = emailAttachment.ContentId; // Se isso estiver definido, o anexo será inline.
                        string ContentType = emailAttachment.ContentType.MediaType; // tipo de anexo pdf, jpg, png, etc.

                        //escreve o anexo no disco
                        System.IO.File.WriteAllBytes(folder + OriginalFileName, emailAttachment.Body); // sobrescreve MessagePart.Body com anexo 
                        //salvando folder no disco
                        mailMessage.Attachments.Add(new Attachment(folder + OriginalFileName));
                    }
                    
                    mailMessage.Priority = MailPriority.High;

                    //smtp do e-mail que irá enviar 
                    SmtpClient smtpClient = new SmtpClient("email-ssl.com.br");
                    smtpClient.EnableSsl = false;
                    //credenciais da conta que utilizará para enviar o e-mail 
                    smtpClient.Credentials = new System.Net.NetworkCredential("milton.silva@teste.com.br", "Password");
                    //Mensagem em copia 
                    MailAddress copy = new MailAddress("copia@teste.com.br");
                    mailMessage.CC.Add(copy);
                    smtpClient.Send(mailMessage);
                }
                //Após ler todos, apagar 
                //client.DeleteAllMessages();

                // exclui uma mensagem através de seu ID .. o ID é na verdadea posição no web-mail da mensagem
                // para excluir o client precisa executar a operação de QUIT para realmente excluir (para ser atomico). Em outras palavras, precisa ser 'disposado'//Disconnect
                //client.DeleteMessage(messageCount); 
            }

            Console.ReadLine();
        }        
    }
}
