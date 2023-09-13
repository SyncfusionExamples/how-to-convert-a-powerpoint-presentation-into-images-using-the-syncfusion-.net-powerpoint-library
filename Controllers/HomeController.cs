using Microsoft.AspNetCore.Mvc;
using PPTtoImage.Models;
using System.Diagnostics;
//using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;
using System.IO.Compression;
using Microsoft.AspNetCore.Mvc.TagHelpers;

namespace PPTtoImage.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult ConvertToImage()
        {
            using(FileStream fileStreamInput = new FileStream(Path.GetFullPath("Data/Template.pptx"), FileMode.Open, FileAccess.Read))
            {
                using(IPresentation pptx = Presentation.Open(fileStreamInput))
                {
                    pptx.PresentationRenderer = new PresentationRenderer();
                    Stream imageStream = pptx.Slides[0].ConvertToImage(ExportImageFormat.Jpeg);
                    imageStream.Position = 0;
                    return File(imageStream, "image/jpeg", "Slide1.jpeg");

                    //Stream[] images = pptx.RenderAsImages(ExportImageFormat.Jpeg);
                    //using(MemoryStream ms = new MemoryStream())
                    //{
                    //    using(var zip = new ZipArchive(ms, ZipArchiveMode.Create, true))
                    //    {
                    //        int i = 1;
                    //        foreach(Stream imageStream in images)
                    //        {
                    //            imageStream.Position = 0;
                    //            var image = zip.CreateEntry("Slide_" + i + ".jpeg");
                    //            using(var entryStream = image.Open())
                    //            {
                    //                imageStream.CopyTo(entryStream);
                    //            }
                    //            i++;
                    //        }
                    //    }
                    //    return File(ms.ToArray(), "application/zip", "PPTtoImage.zip");
                    //}
                }
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