using Microsoft.AspNetCore.Http;
using System.Collections.Generic;

namespace MipsTester.Models
{
    public class FileUpload
    {
        public IFormFile FormFile { get; set; }
        public List<IFormFile> FormFiles { get; set; }
    }
}
