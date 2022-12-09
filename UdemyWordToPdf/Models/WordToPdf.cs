using Microsoft.AspNetCore.Http;

namespace UdemyWordToPdf.Models
{
    public class WordToPdf
    {
        public string Email { get; set; }
        public IFormFile WordFile { get; set; }
    }
}
