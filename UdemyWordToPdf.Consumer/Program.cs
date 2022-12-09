using Newtonsoft.Json;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Spire.Doc;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace UdemyWordToPdf.Consumer
{
    class Program
    {
        public static bool EmailSend(string email, MemoryStream ms, string fileName)
        {
            try
            {
                ms.Position = 0;

                ContentType ct = new ContentType(MediaTypeNames.Application.Pdf);

                Attachment attachment = new Attachment(ms, ct);
                attachment.ContentDisposition.FileName = $"{fileName}.pdf";

                MailMessage mailMessage = new MailMessage();

                SmtpClient smtpClient = new SmtpClient();

                mailMessage.From = new MailAddress("muratcan_dongel@hotmail.com");
                mailMessage.To.Add(email);
                mailMessage.Subject = "Pdf Dosyası Oluşturma | bıdıbıdı.com";
                mailMessage.Body = "Pdf dosyanız ektedir.";
                mailMessage.IsBodyHtml = true;
                mailMessage.Attachments.Add(attachment);

                smtpClient.Host = "mail.hotmail.com";
                smtpClient.Port = 587;
                smtpClient.Credentials = new NetworkCredential("muratcan_dongel@hotmail.com", "mcDL5957..");
                smtpClient.Send(mailMessage);

                Console.WriteLine($"Sonuç: {email} adresine gönderilmiştir.");

                ms.Close();
                ms.Dispose();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Mail gönderim sırasında bir hata meydana geldi. {ex.InnerException}");
                return false;
            }
        }

        static void Main(string[] args)
        {
            bool result = false;
            var factory = new ConnectionFactory();

            factory.Uri = new Uri("amqps://zxuhuhnz:uwT3kF1kowQ8859E7O4WgNOon-iqc9VJ@tiger.rmq.cloudamqp.com/zxuhuhnz");

            using (var connection = factory.CreateConnection())
            {
                using var channel = connection.CreateModel();

                channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);
                channel.QueueBind("File", "convert-exchange", "WordToPdf");
                channel.BasicQos(0, 1, false);

                var consumer = new EventingBasicConsumer(channel);
                channel.BasicConsume("File", false, consumer);

                consumer.Received += (model, ea) =>
                {
                    try
                    {
                        Console.WriteLine("Kuyruktan bir mesaj alındı ve işleniyor");

                        Document document = new Document();

                        string message = Encoding.UTF8.GetString(ea.Body.ToArray());

                        MessageWordToPdf messageWordToPdf = JsonConvert.DeserializeObject<MessageWordToPdf>(message);

                        document.LoadFromStream(new MemoryStream(messageWordToPdf.WordByte), FileFormat.Docx2013);

                        using (MemoryStream ms = new MemoryStream())
                        {
                            document.SaveToStream(ms, FileFormat.PDF);
                            result = EmailSend(messageWordToPdf.Email, ms, messageWordToPdf.FileName);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Hata meydana geldi." + ex.Message);
                    }

                    if (result)
                    {
                        Console.WriteLine("Kuyruktan Mesaj başarıyla işlendi...");
                        channel.BasicAck(ea.DeliveryTag, false);
                    }
                };

                Console.WriteLine("Çıkmak için tıklayınız");
                Console.ReadLine();
            }
        }
    }
}
