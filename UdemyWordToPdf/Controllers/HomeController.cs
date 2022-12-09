﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UdemyWordToPdf.Models;

namespace UdemyWordToPdf.Controllers
{
    public class HomeController : Controller
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult WordToPdfPage()
        {
            return View();
        }

        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            var factory = new ConnectionFactory();

            factory.Uri = new Uri(_configuration["ConnectionStrings:RabbitMQCloudString"]);

            using (var connection = factory.CreateConnection())
            {
                using var channel = connection.CreateModel();
                channel.ExchangeDeclare(exchange: "convert-exchange", type: ExchangeType.Direct, durable: true, autoDelete: false, null);
                channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, null);
                channel.QueueBind(queue: "File", exchange: "convert-exchange", routingKey: "WordToPdf");

                MessageWordToPdf messageWordToPdf = new MessageWordToPdf();

                using (MemoryStream ms = new MemoryStream())
                {
                    wordToPdf.WordFile.CopyTo(ms);
                    messageWordToPdf.WordByte = ms.ToArray();
                }

                messageWordToPdf.Email = wordToPdf.Email;
                messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);

                byte[] byteMessage = Encoding.UTF8.GetBytes(serializeMessage);

                var properties = channel.CreateBasicProperties();
                properties.Persistent = true;

                channel.BasicPublish(exchange: "convert-exchange", routingKey: "WordToPdf", basicProperties: properties, body: byteMessage);

                ViewBag.result = "Word dosyanız pdf dosyasına dönüştürüldükten sonra size email olarak gönderilecektir.";

                return View();
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
