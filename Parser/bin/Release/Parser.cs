using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


public class Parser
{
    //-----------DocBook---------------

    public void createXML(string XmlFilePath, string str)
    {
        try
        {
            if (!File.Exists(XmlFilePath))
            {
                XmlTextWriter textWritter = new XmlTextWriter(XmlFilePath, Encoding.UTF8);

                textWritter.WriteStartDocument();
                string txt = "http://docbook.org/ns/docbook";
                textWritter.WriteStartElement("book " , txt);
               
                
                textWritter.WriteEndElement();

                textWritter.Close();
            }
            else
            {
                //Console.WriteLine("File haven't been created. Xml already exist!");

            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
        addData(XmlFilePath, str);
    }

    private void addData(string XmlFilePath, string str)
    {
        //XmlFilePath += ".xml";
        XmlDocument document = new XmlDocument();

        document.Load(XmlFilePath);
        XmlNode element = document.CreateElement("info");
        document.DocumentElement.AppendChild(element);


        XmlNode title = document.CreateElement("title");
        title.InnerText = XmlFilePath;
        element.AppendChild(title);

        XmlNode chapter = document.CreateElement("chapter");
        document.DocumentElement.AppendChild(chapter); // указываем родителя

        XmlNode para = document.CreateElement("para");
        para.InnerText = str.ToString();
        chapter.AppendChild(para);




        document.Save(XmlFilePath);

    }

    //-------------Docm------------------

    public void WriteDocmDocument(string DocmFilePath, string str)
    {
        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(DocmFilePath, true);  // Open a WordprocessingDocument for editing using the DocmFilePath.
        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;  // Assign a reference to the existing document body.
        Paragraph para = body.AppendChild(new Paragraph());     // Add new text.
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(str.ToString()));
        wordprocessingDocument.Close(); // Close the handle explicitly.
    }

    public string ReadDocmDocument(string DocmFilePath)
    {


        StringBuilder sb = new StringBuilder();
        WordprocessingDocument package = WordprocessingDocument.Open(DocmFilePath, true); // Open a WordprocessingDocument for editing using the DocmFilePath.
        OpenXmlElement element = package.MainDocumentPart.Document.Body;
        if (element == null)
        {
            return string.Empty;
        }
        sb.Append(GetText(element));

        package.Close();
        return sb.ToString();


    }

    private string GetText(OpenXmlElement element)
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
                    PlainTextInWord.Append(GetText(section));
                    PlainTextInWord.AppendLine(Environment.NewLine);
                    break;


                default:
                    PlainTextInWord.Append(GetText(section));
                    break;
            }
        }


        return PlainTextInWord.ToString();
    }
}

