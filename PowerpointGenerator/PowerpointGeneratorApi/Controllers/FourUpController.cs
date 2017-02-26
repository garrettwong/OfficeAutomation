using PowerPointService;
using System.Web.Http;
    
namespace PowerpointGeneratorApi.Controllers
{
    public class FourUpController : ApiController
    {
        private readonly PowerPointGeneratorService _powerPointGeneratorService = new PowerPointGeneratorService();

        [Route("api/FourUp/Generate")]
        public IHttpActionResult Generate(ImageModel[] imageModels)
        {
            var name = "Hello";

            // take in a form upload of images and categories for those images, generate powerpoint
            _powerPointGeneratorService.GenerateFourUp(name, imageModels);

            return Ok();
        }
    }

}
