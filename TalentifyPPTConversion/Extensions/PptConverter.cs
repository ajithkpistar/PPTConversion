using System;
using Microsoft.Office.Core;
using System.IO.Compression;
using System.IO;
using Ionic.Zip;

namespace TalentifyPPTConversion.Extensions
{
    public class PptConverter
    {
        String fileName;
        String outputPath;
        String outputFileName;
        String mediaURL;
        String outPath;
        String zipPath;

        public PptConverter(String fileName, String outputPath, String outputFileName,String mediaURL)
        {
            this.fileName = fileName;
            this.outputPath = outputPath;
            this.outputFileName = outputFileName;
            this.mediaURL = mediaURL;
        }

        public void convertFile()
        {
            var app = new Microsoft.Office.Interop.PowerPoint.Application();
            var pres = app.Presentations;
            var file = pres.Open(@fileName, MsoTriState.msoTrue, MsoTriState.msoTrue, MsoTriState.msoFalse);

            // file.Export(@"C:\ppt\presentation1.jpg", "PNG");

            var slideCount = file.Slides.Count;

            //file.SaveAs();

            var xmlPath = outputPath + "/" + outputFileName+"/"+ outputFileName + ".xml".Replace("/", "\\");
            zipPath = "C:/ppt/" + outputFileName+".zip".Replace("/", "\\");

            outPath = (outputPath +"/" + outputFileName + ".png").Replace("/","\\");
            file.SaveCopyAs(outPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPNG, MsoTriState.msoTriStateMixed);

            convertXML(slideCount, xmlPath, outputFileName);
            //createZip();

            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            using (ZipFile zipfile = new ZipFile())
            {
                
                zipfile.AddDirectory(outputPath);
                zipfile.Save(zipPath);
            }


            // ZipFile.CreateFromDirectory(outputPath + "/" + outputFileName + "/", zipPath,CompressionLevel.Fastest,true);
        }



    public void convertXML(int slideCount,String xmlPath,String outputFileName) {
  
            String lessonXML= "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<lesson description=\"NA\" h1=\"13_Lookup Functions\" lesson_type=\"PRESENTATION\">";

            for (int i=1; i<=slideCount;i++) {
                lessonXML += "\n<slide background =\"#ffffff\" background_transition=\"slide\" fragmentCount=\"0\" image_bg=\""+mediaURL+""+ outputFileName + "/"+ outputFileName + "/Slide"+i+ ".PNG\" template=\"NO_CONTENT\" transition=\"slide\">\n" +
                    "<slide_audio> none </slide_audio>\n" +
                    "<id> "+i+ " </id>\n" +
                    "<img fragment_duration =\"0\" url=\"http://images.all-free-download.com/images/graphicthumb/blank_note_document_4179.jpg\"/>\n" +
                    "<ul merged_audio =\"none\">\n" +
                    "<mergedAudioDuration> 0 </mergedAudioDuration>\n" +
                    "</ul>\n" +
                    "<order_id> "+i+ " </order_id>\n" +
                    "<p fragment_duration =\"0\"/>\n" +
                    "<duration> 5000 </duration>\n" +
                    "<student_notes> Not Available </student_notes>\n" +
                    "<teacher_notes> Not Available </teacher_notes>\n" +
                    "<h1 fragment_duration =\"0\">EMPTY_TITLE</h1>\n" +
                    "<h2 fragment_duration =\"0\"/>\n" +
                    "</slide>\n";
            }
            lessonXML += "</lesson>";
            System.IO.File.WriteAllText(xmlPath, lessonXML);
        }

    }
}
