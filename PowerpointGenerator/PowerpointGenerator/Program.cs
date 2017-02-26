using Microsoft.Owin.Hosting;
using PowerpointGenerator;
using System;
using System.Net.Http;

namespace PowerPointService
{
    class Program
    {
        /// <summary>
        /// Owin Self Host Application
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            string baseAddress = "http://localhost:9000/";

            // Start OWIN host
            using (WebApp.Start<Startup>(url: baseAddress))
            {
                // Create HttpCient and make a request to api/values 
                HttpClient client = new HttpClient();

                //GeneratePowerPoint(args);

                //var response = client.GetAsync(baseAddress + "api/values").Result;

                //Console.WriteLine(response);
                //Console.WriteLine(response.Content.ReadAsStringAsync().Result);
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Console Application
        /// </summary>
        /// <param name="args"></param>
        static void Main2(string[] args) { 
            string presentationName = "";

            // PowerPointGenerator.exe --presentationName "presentationName" --image
            if (args.Length > 2)
            {
                presentationName = args[1];
            }

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
