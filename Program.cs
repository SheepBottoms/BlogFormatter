using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace BlogFormatter
{
    class Program
    {
        static string imgTag = "[![picture]({{{{ site.images }}}}/{0}/image{1}.jpg)]({{{{ site.images }}}}/{0}/image{1}.jpg)";
        static string dateTag;
        static string imageLocation;
        static string imageExample;
        static string header = @"---
layout: post
title:  ""{0}""
crawlertitle: ""{0}""
summary: """"
date: {1}
categories: posts
tags: """"
author: huffaker
group: ""{3}""
bg: ""{2}/{4}.jpg""
---

";

        /// <summary>
        /// Converts a word doc into a blog post for a Jekyll site. Text is in plain with tags included for
        /// images. The images are compressed/resized and placed into the assets folder for the given post.
        /// Can be run multiple times for the same post to re-process.
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            var fileName = args[0];
            var tripName = args[1];
            var postName = args[2];
            var blogLocation = args[3];
            dateTag = args[4];

            var imageDir = Directory.CreateDirectory(blogLocation + @"\assets\images\" + dateTag);
            imageLocation = imageDir.FullName;

            // Clear out any files if they exists (incase we are re-publishing)
            foreach (FileInfo file in imageDir.GetFiles())
            {
                file.Delete();
            }

            var fileContent = ReadWordDocument(fileName);

            var postHeader = string.Format(header, postName, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss +0000"), dateTag, tripName, imageExample);
            var postFileName = blogLocation + @"\_posts\" + DateTime.Now.ToString("yyyy-MM-dd") + "-" + postName.Replace(' ', '-').ToLower() + ".markdown";

            // Save the blog post
            File.WriteAllLines(postFileName, new string[] { postHeader, fileContent });
        }

        /// <summary> 
        ///  Read Word Document 
        /// </summary> 
        /// <returns>Plain Text in document </returns> 
        public static string ReadWordDocument(string filepath)
        {
            StringBuilder sb = new StringBuilder();
            using (Stream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = WordprocessingDocument.Open(stream, false))
            {
                OpenXmlElement element = package.MainDocumentPart.Document.Body;
                if (element == null)
                {
                    return string.Empty;
                }

                sb.Append(GetPlainText(element, package.MainDocumentPart));
                return sb.ToString();
            }
        }

        /// <summary> 
        ///  Read Plain Text in all XmlElements of word document 
        /// </summary> 
        /// <param name="element">XmlElement in document</param> 
        /// <returns>Plain Text in XmlElement</returns> 
        public static string GetPlainText(OpenXmlElement element, MainDocumentPart wDoc)
        {
            StringBuilder PlainTextInWord = new StringBuilder();
            foreach (OpenXmlElement section in element.Elements())
            {
                switch (section.LocalName)
                {
                    // Text 
                    case "t":
                        PlainTextInWord.Append(section.InnerText);
                        break;


                    case "cr":                          // Carriage return 
                    case "br":                          // Page break 
                        PlainTextInWord.Append(Environment.NewLine);
                        break;


                    // Tab 
                    case "tab":
                        PlainTextInWord.Append("\t");
                        break;


                    // Paragraph 
                    case "p":
                        PlainTextInWord.Append(GetPlainText(section, wDoc));
                        PlainTextInWord.AppendLine(Environment.NewLine);
                        break;

                    // Run - contains drawing elements
                    case "r":
                        PlainTextInWord.Append(GetPlainText(section, wDoc));
                        break;

                    // Drawing element
                    case "drawing":
                        // Extract the image
                        PlainTextInWord.Append(GetPlainText(section, wDoc));
                        var imageFirst = ((Drawing)section).Inline.Graphic.GraphicData.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.Picture>().FirstOrDefault();
                        var blip = imageFirst.BlipFill.Blip.Embed.Value;
                        ImagePart img = (ImagePart)wDoc.Document.MainDocumentPart.GetPartById(blip);
                        var imageName = string.Format("image{0}", blip);
                        ImageUtilities.FormatImage(img.GetStream(), imageLocation, imageName);
                        imageExample = imageExample == null ? imageName : new Random().Next(4) == 3 ? imageName : imageExample;

                        // Add the image tag
                        PlainTextInWord.Append(string.Format(imgTag, dateTag, blip));
                        break;

                    default:
                        PlainTextInWord.Append(GetPlainText(section, wDoc));
                        break;
                }
            }


            return PlainTextInWord.ToString();
        }
    }
}
