using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using TalentifyPPTConversion.Extensions;
using System.IO;
using System.Web;

namespace TalentifyPPTConversion.Controllers
{
    [RoutePrefix("api")]
    public class FileUploadController : ApiController
    {

        public static string pptUploadFolder = "C:/ppt/";

        [Route("upload")]
        [HttpPost]
        public async Task<HttpResponseMessage> UploadFile(HttpRequestMessage request)
        {
            if (!request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(HttpStatusCode.UnsupportedMediaType);
            }

            var data = await Request.Content.ParseMultipartAsync();
            var lessonId="";
            var mediaURL= "http://cdn.talentify.in:9999/lessonXMLs/";

            if (data.Fields.ContainsKey("lessonId"))
            {
                lessonId = data.Fields["lessonId"].Value+"";
            }

            if (data.Fields.ContainsKey("mediaURL"))
            {
                mediaURL = data.Fields["mediaURL"].Value + "";
            }

            var fileName = "abc.pptx";
            if (data.Files.ContainsKey("file"))
            {
                var file = data.Files["file"].File;
                fileName = data.Files["file"].Filename;
                var path = pptUploadFolder + fileName;
                File.WriteAllBytes(path, file);
                var outputPath = pptUploadFolder + lessonId;
                System.IO.Directory.CreateDirectory(outputPath);


                PptConverter pptConverter = new PptConverter(path, outputPath, lessonId,mediaURL);
                pptConverter.convertFile();

            }

            if (data.Fields.ContainsKey("description"))
            {
                var description = data.Fields["description"].Value;
            }

            var respons = "Done... uploaded at " + pptUploadFolder + fileName;
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(respons)
            };
        }
    }
}
