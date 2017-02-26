using PowerPointService;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Http;

namespace PowerpointGenerator.Controllers
{
    public class PowerPointGeneratorController : ApiController
    {
        [HttpGet]
        public void GeneratePowerPoint(string presentationName)
        {
            var powerPointGeneratorService = new PowerPointGeneratorService();

            powerPointGeneratorService.GenerateFourUp(
                presentationName,

                new ImageModel[] {
                    new ImageModel()
                    {
                        Category = "Family",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    },
                    new ImageModel()
                    {
                        Category = "Family",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    },
                    new ImageModel()
                    {
                        Category = "Family",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    },
                    new ImageModel()
                    {
                        Category = "Work",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    },
                    new ImageModel()
                    {
                        Category = "Work",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    },
                    new ImageModel()
                    {
                        Category = "Work",
                        ImageBytes = new byte[] { },
                        ImagePath = ""
                    }
                });
        }

    }
}
