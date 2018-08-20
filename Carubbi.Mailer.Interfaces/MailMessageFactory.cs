using System.Net.Mail;
using System.Text;

namespace Carubbi.Mailer.Interfaces
{
    /// <summary>
    /// Fabrica responsável por criar objetos MailMessage
    /// </summary>
    public class MailMessageFactory
    {
        /// <summary>
        /// Cria uma MailMessage configurada para o padrão Brasileiro
        /// </summary>
        /// <param name="subject">Assunto</param>
        /// <param name="body">Corpo da mensagem</param>
        /// <returns>Objeto Criado</returns>
        public static MailMessage CreateBrazilianHTMLMessage(string subject, string body)
        { 
            return new MailMessage {
                Subject = subject,
                Body = body,
                BodyEncoding = Encoding.GetEncoding("ISO-8859-1"),
                SubjectEncoding = Encoding.GetEncoding("ISO-8859-1"), 
                IsBodyHtml = true
            };
        }
    }
}
