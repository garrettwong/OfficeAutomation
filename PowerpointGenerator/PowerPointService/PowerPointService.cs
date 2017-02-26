using ImpersonationUtil;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PowerPointService
{

    /// <summary>
    /// CS Automated Power Point - https://code.msdn.microsoft.com/office/CSAutomatePowerPoint-b312d416
    /// Create PowerPoint Programmatically - http://www.free-power-point-templates.com/articles/create-powerpoint-ppt-programmatically-using-c/
    /// 
    /// </summary>
    public class PowerPointGeneratorService
    {
        public void Generate()
        {
            var ppt = new Microsoft.Office.Interop.PowerPoint.Application();
            ppt.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            var pptPresentations = ppt.Presentations;
            var pptPresentation = pptPresentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

            var slides = pptPresentation.Slides;
            var slide = slides
                .AddSlide(1, customLayout);

            var shapes = slide.Shapes;
            var shape = shapes[1];
            var textFrame = shape.TextFrame;
            var textRange = textFrame.TextRange;
            textRange.Text = "All-In-One Code Framework";

            var shape2 = shapes[2];
            var textFrame2 = shape2.TextFrame;
            var textRange2 = textFrame2.TextRange;
            textRange2.Text = "Aloha";
            textRange2.Text += "FPPT.com";
            textRange2.Font.Name = "Arial";
            textRange2.Font.Size = 32;

            textRange2 = slide.Shapes[2].TextFrame.TextRange;
            // \n denotes a bullet
            textRange2.Text = "Content goes here\nYou can add text\nItem 3";

            slide.Shapes.AddPicture(GetImagePath(@"C:\Users\wong_ga\Pictures"),
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue,
                shape.Left, shape.Top + 150.0f, shape.Width + 150.0f, shape.Height + 150.0f);
            slide.NotesPage.Shapes[3].TextFrame.TextRange.Text = "m";



            string fileName = Path.GetDirectoryName(
            Assembly.GetExecutingAssembly().Location) + "\\Sample1.pptx";
            pptPresentation.SaveAs(fileName,
                Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                Microsoft.Office.Core.MsoTriState.msoTriStateMixed);

            // close and quit
            //pptPresentation.Close();
            //ppt.Quit();
        }

        public string GetImagePath(string path)
        {
            var rand = new Random();
            var files = Directory.GetFiles(path, "*");
            return files[rand.Next(files.Length)];
            //return @"C:\Users\wong_ga\Pictures\Greenbear.jpg";
        }

        public void GenerateFourUp(string name, ImageModel[] imageModels)
        {
            using (Impersonation impersonator = new Impersonation("sc", "wong_ga", AppConstants.Password))
            {
                var ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                //ppt.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                //ppt.Visible = Microsoft.Office.Core.MsoTriState.msoTriStateMixed;

                var pptPresentations = ppt.Presentations;
                Presentation pptPresentation = null;

                try
                {
                    pptPresentation = pptPresentations.Open(Directory.GetCurrentDirectory() + "\\App_Data\\FourUpTemplate - Copy.pptx",
                        MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                    Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout =
                    pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];

                    var slides = pptPresentation.Slides;
                    var slide = slides[1];
                    var shapes = slide.Shapes;

                    var shape = shapes[3];
                    var textFrame2 = shape.TextFrame;
                    var textRange2 = textFrame2.TextRange;
                    textRange2.Text = $"{name}\n{DateTime.Now.ToString()}";

                    // four up page
                    var imageHolders = new List<ImagePlaceholder>();
                    slide = slides[2];
                    var initialCount = slide.Shapes.Count;
                    for (int i = 1; i <= initialCount; i++)
                    {
                        var fourUpImageShape = slide.Shapes[i];
                        var reflectionFormat = fourUpImageShape.Reflection;

                        // check type
                        if (fourUpImageShape.Type == MsoShapeType.msoPicture)
                        {
                            // add greenbear pic to four up
                            var imagePlaceholder = new ImagePlaceholder();
                            imagePlaceholder.ImagePath = GetImagePath(@"C:\Users\wong_ga\Pictures");
                            imagePlaceholder.Left = fourUpImageShape.Left;
                            imagePlaceholder.Top = fourUpImageShape.Top;
                            imagePlaceholder.Width = fourUpImageShape.Width;
                            imagePlaceholder.Height = fourUpImageShape.Height;
                            imageHolders.Add(imagePlaceholder);
                            //var pic = slide.Shapes.AddPicture(GetImagePath(),
                            //    Microsoft.Office.Core.MsoTriState.msoFalse,
                            //    Microsoft.Office.Core.MsoTriState.msoTrue,
                            //    fourUpImageShape.Left,
                            //    fourUpImageShape.Top,
                            //    fourUpImageShape.Width,
                            //    fourUpImageShape.Height);

                            //// reflection
                            //try
                            //{
                            //    pic.Reflection.Blur = reflectionFormat.Blur;
                            //    pic.Reflection.Transparency = reflectionFormat.Transparency;
                            //}
                            //catch { }

                            //pic.ZOrder(MsoZOrderCmd.msoSendToBack);
                        }
                    }
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var fourUpImageShape = slide.Shapes[i];
                        if (fourUpImageShape.Type == MsoShapeType.msoPicture)
                        {
                            fourUpImageShape.Delete();
                            i--;
                        }
                    }
                    // add images
                    foreach (var imageHolder in imageHolders)
                    {
                        var pic = slide.Shapes.AddPicture(imageHolder.ImagePath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            imageHolder.Left,
                            imageHolder.Top,
                            imageHolder.Width,
                            imageHolder.Height);
                        pic.ZOrder(MsoZOrderCmd.msoSendToBack);
                    }

                    string fileName = Path.GetDirectoryName(
                        Assembly.GetExecutingAssembly().Location) + $"\\Test_{DateTime.Now.Millisecond}.pptx";
                    pptPresentation.SaveAs(fileName,
                        Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                        Microsoft.Office.Core.MsoTriState.msoTriStateMixed);

                    // group by category
                    var imageModelList = imageModels.ToList();

                    var categoryList = from s in imageModelList
                                       group s by s.Category into g
                                       select g.ToArray();

                    foreach (var categoryImageModelArray in categoryList)
                    {
                        foreach (var imageModel in categoryImageModelArray)
                        {
                            var imageBytes = imageModel.ImageBytes;

                        }
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine("error" + ex.Message);
                }
                finally
                {
                    pptPresentation.Close();
                    ppt.Quit();

                    if (pptPresentation != null)
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pptPresentation);
                    }
                    if (ppt != null)
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ppt);
                    }
                    pptPresentation = null;
                    ppt = null;

                    GC.Collect();
                }
            }
        }
    }

    internal class AppConstants
    {
        public static string Password = "Helloworld123";
    }
}