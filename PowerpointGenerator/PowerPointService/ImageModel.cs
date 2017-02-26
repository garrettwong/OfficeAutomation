using System.IO;
using System.Reflection;

namespace PowerPointService
{
    public class ImageModel
    {
        public string Category { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImagePath { get; set; }
    }
}
