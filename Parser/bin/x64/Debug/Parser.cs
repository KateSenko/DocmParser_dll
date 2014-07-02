using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Design;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
//using DocumentFormat.OpenXml.Drawing.Wordprocessing;
//using DocumentFormat.OpenXml.Drawing;



public class Parser
{
    //-----------DocBook---------------
    public Parser() { }

    //[DllImport("DocumentFormat.OpenXml")]
    //public extern class OpenXmlElement;

    public String calc(String a, String b)
    {
        return "Ok";
    }

    private static void createXML(string XmlFilePath, string str)
    {
        try
        {
            if (!File.Exists(XmlFilePath))
            {
                XmlTextWriter textWritter = new XmlTextWriter(XmlFilePath, Encoding.UTF8);

                textWritter.WriteStartDocument();
                string txt = "http://docbook.org/ns/docbook";
                textWritter.WriteStartElement("book ", txt);


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
        //   addData(XmlFilePath, str);
    }

    // основная функция конвертации

    public void convert(String FilePath, String DocmFileName)
    {
        // FilePath = @"d:\1\";
        // DocmFileName = "130349";
        string LogFilePath = System.IO.Path.Combine(FilePath.ToString(), ("ErrorLog" + ".txt")); //file for containing error docx files


        string DocmFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocmFileName.ToString() + ".docx"));
        string XmlFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocmFileName.ToString() + ".xml"));

        try
        {
            createXML(XmlFilePath, "");
            XmlDocument document = new XmlDocument();
            document.Load(XmlFilePath);
            XmlNode element = document.CreateElement("info");
            document.DocumentElement.AppendChild(element);


            XmlNode title = document.CreateElement("title");
            title.InnerText = DocmFileName;
            element.AppendChild(title);

            XmlNode chapter = document.CreateElement("chapter");
            document.DocumentElement.AppendChild(chapter);


            using (WordprocessingDocument doc = WordprocessingDocument.Open(DocmFilePath, true)) //можно использовать Stream
            {
                var body = doc.MainDocumentPart.Document.Body;
             //   var imageParts = doc.MainDocumentPart.ImageParts;

                foreach (var text in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                {

                    XmlNode para = document.CreateElement("para");
                    para.InnerText = text.Text;
                    chapter.AppendChild(para);
                }
                //foreach (var elements in body.Descendants())
                //{
                //    if (elements.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Text")
                //    {
                //        XmlNode para = document.CreateElement("para");
                //        para.InnerText = elements.InnerText;
                //        chapter.AppendChild(para);

                //        Console.WriteLine("Text!");
                //        Console.ReadKey();

                //    }

                //    var e = doc.MainDocumentPart.Parts.GetEnumerator();
                //    int picNum = 0;

                //    while (e.MoveNext())
                //    {
                //        picNum++;
                //        if (e.Current.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Drawing")
                //        {
                //            Stream stream = e.Current.OpenXmlPart.GetStream();

                //            long length = stream.Length;
                //            byte[] byteStream = new byte[length];
                //            stream.Read(byteStream, 0, (int)length);

                //            FileStream fstream = new FileStream(path + picNum + ".jpg", FileMode.OpenOrCreate);
                //            fstream.Write(byteStream, 0, (int)length);
                //            fstream.Close();
                //        }
                //        if (e.Current.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Text")
                //        {
                //            XmlNode para = document.CreateElement("para");
                //            para.InnerText = e.Current.InnerText;
                //            chapter.AppendChild(para);
                //        }
                //    }
                    //if (elements.ToString() == "DocumentFormat.OpenXml.Wordprocessing.Drawing")
                    //{
                    //    ImagePart imagePart = elements.Current;
                    //}

                    //foreach (var prt in body.Descendants())
                    //{
                    //    if (prt.GetType().ToString().Equals(DocumentFormat.OpenXml.Wordprocessing.Text))
                    //        ;
                    //}
               // }
                
            }
            document.Save(XmlFilePath);
        }
        catch (XmlException ex)
        {
            Console.WriteLine(ex.Message);
            System.IO.StreamWriter ErrorLog = new System.IO.StreamWriter(LogFilePath, true);
            ErrorLog.WriteLine(DocmFileName);
            ErrorLog.Close();

        }

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
        DocumentFormat.OpenXml.Wordprocessing.Paragraph para = body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph());     // Add new text.
        DocumentFormat.OpenXml.Wordprocessing.Run run = para.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Run());
        run.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Text(str.ToString()));
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

    //public void addData(string XmlFilePath, StringBuilder str)
    //{
    //    XmlDocument document = new XmlDocument();

    //    document.Load(XmlFilePath);
    //    XmlNode element = document.CreateElement("info");
    //    document.DocumentElement.AppendChild(element);


    //    XmlNode title = document.CreateElement("title");
    //    title.InnerText = XmlFilePath;
    //    element.AppendChild(title);

    //    XmlNode chapter = document.CreateElement("chapter");
    //    document.DocumentElement.AppendChild(chapter); // указываем родителя

    //    XmlNode para = document.CreateElement("para");
    //    para.InnerText = str.ToString();
    //    chapter.AppendChild(para);




    //    document.Save(XmlFilePath);

    //    // Console.WriteLine("Data have been added to xml!");

    //    // Console.ReadKey();
    //    // Console.WriteLine(XmlToJSON(document));


    //}

    public void ParseDocxDocument(String FilePath, String DocmFileName)
    {
        //  StringBuilder result = new StringBuilder();
        string LogFilePath = System.IO.Path.Combine(FilePath.ToString(), ("ErrorLog" + ".txt")); //file for containing error docx files


        string DocmFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocmFileName.ToString() + ".docx"));
        string XmlFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocmFileName.ToString() + ".xml"));

        try
        {
            createXML(XmlFilePath, "");
            XmlDocument document = new XmlDocument();
            document.Load(XmlFilePath);
            XmlNode element = document.CreateElement("info");
            document.DocumentElement.AppendChild(element);


            XmlNode title = document.CreateElement("title");
            title.InnerText = DocmFileName;
            element.AppendChild(title);

            XmlNode chapter = document.CreateElement("chapter");
            document.DocumentElement.AppendChild(chapter);

            WordprocessingDocument wordProcessingDoc = WordprocessingDocument.Open(DocmFilePath, true);
            List<ImagePart> imgPart = wordProcessingDoc.MainDocumentPart.ImageParts.ToList();
            List<string> tableCellContent = new List<string>();
            IEnumerable<Paragraph> paragraphElement = wordProcessingDoc.MainDocumentPart.Document.Descendants<Paragraph>();
            int imgCounter = 0;

            foreach (OpenXmlElement section in wordProcessingDoc.MainDocumentPart.Document.Body.Elements<OpenXmlElement>())
            {
                if (section.GetType().Name == "Paragraph")
                {
                    Paragraph par = (Paragraph)section;
                    //Add new paragraph tag
                    //result.Append("<div style=\"width:100%; text-align:");

                    ////Append anchor style
                    //if (par.ParagraphProperties != null && par.ParagraphProperties.Justification != null)
                    //    switch (par.ParagraphProperties.Justification.Val.Value)
                    //    {
                    //        case JustificationValues.Left:
                    //            result.Append("left;");
                    //            break;
                    //        case JustificationValues.Center:
                    //            result.Append("center;");
                    //            break;
                    //        case JustificationValues.Both:
                    //            result.Append("justify;");
                    //            break;
                    //        case JustificationValues.Right:
                    //        default:
                    //            result.Append("right;");
                    //            break;
                    //    }
                    //else
                    //    result.Append("left;");

                    ////Append text decoration style
                    //if (par.ParagraphProperties != null && par.ParagraphProperties.ParagraphMarkRunProperties != null && par.ParagraphProperties.ParagraphMarkRunProperties.HasChildren)
                    //    foreach (OpenXmlElement chield in par.ParagraphProperties.ParagraphMarkRunProperties.ChildElements)
                    //    {
                    //        switch (chield.GetType().Name)
                    //        {
                    //            case "Bold":
                    //                result.Append("font-weight:bold;");
                    //                break;
                    //            case "Underline":
                    //                result.Append("text-decoration:underline;");
                    //                break;
                    //            case "Italic":
                    //                result.Append("font-style:italic;");
                    //                break;
                    //            case "FontSize":
                    //                result.Append("font-size:" + ((FontSize)chield).Val.Value + "px;");
                    //                break;
                    //            default: break;
                    //        }
                    //    }

                    //result.Append("\">");

                    //Add image tag
                    IEnumerable<Run> runs = par.Descendants<Run>();
                    foreach (Run run in runs)
                    {
                        if (run.HasChildren)
                        {
                            
                            foreach (OpenXmlElement chield in run.ChildElements.Where(o => o.GetType().Name == "Drawing"))   //добавление картинок
                            {
                               // <imagedata fileref="image.png" width="6in" depth="5.5in" scale="300"/>
                                Console.WriteLine("picture!!");
                                XmlNode imagedata = document.CreateElement("imagedata");
                                chapter.AppendChild(imagedata);
                                XmlAttribute attribute = document.CreateAttribute("fileref");
                                Image img = System.Drawing.Image.FromStream(imgPart[imgCounter].GetStream());
                                DirectoryInfo di = Directory.CreateDirectory(DocmFileName);
                                String imgSavePath = FilePath+"\\" + DocmFileName;
                                img.Save(FilePath  + "\\" +imgCounter + ".jpeg");
                                attribute.Value = string.Format(" src=\"data:image/jpeg;base64\" />");  
                                imagedata.Attributes.Append(attribute);
                                imgCounter++;
                            }
                            foreach (OpenXmlElement table in run.ChildElements.Where(o => o.GetType().Name == "Tbl"))     //обработка таблицы далее
                            {
                                Console.WriteLine("Table!");
                                XmlNode para = document.CreateElement("para");
                                para.InnerText = "HERE IS TABLE";
                                chapter.AppendChild(para);
                            }
                        }
                    }
                    
                    //Append inner text
                    IEnumerable<Text> textElement = par.Descendants<Text>();
                    //if (par.Descendants<Text>().Count() == 0)                   //если встречается пустая строка
                    //{
                    //    XmlNode para = document.CreateElement("para");
                    //    para.InnerText = "";
                    //    chapter.AppendChild(para);
                    //}
                    

                    foreach (Text t in textElement.Where(o => !tableCellContent.Contains(o.Text.Trim())))   //добавление текста
                    {
                        Console.WriteLine("Text!");
                        XmlNode para = document.CreateElement("para");
                        para.InnerText = t.Text;
                        chapter.AppendChild(para);
                     }


                    //result.Append("</div>");
                   //result.Append(Environment.NewLine);


                }
                else if (section.GetType().Name == "Drawing")
                {
                    Console.WriteLine("Picture section!");
                    Picture pic = (Picture)section;
                    XmlNode imagedata = document.CreateElement("imagedata");
                    chapter.AppendChild(imagedata);
                    XmlAttribute attribute = document.CreateAttribute("fileref");
                    Image img = System.Drawing.Image.FromStream(imgPart[imgCounter].GetStream());
                    img.Save(System.IO.Path.GetTempPath() + "\\" + imgCounter + ".jpeg");
                    //attribute.Value = string.Format("<img style=\"{1}\" src=\"data:image/jpeg;base64,{0}\" />", GetBase64Image(imgPart[imgCounter].GetStream()),
                    //                            ((DocumentFormat.OpenXml.Vml.Shape)chield.ChildElements.Where(o => o.GetType().Name == "Shape").FirstOrDefault()).Style);  
                    imagedata.Attributes.Append(attribute);
                    imgCounter++;
                }
                else if (section.GetType().Name == "Tbl")
                {
                    Console.WriteLine("Table 2!");
                    //result.Append("<table>");
                    //Table tab = (Table)section;
                    //foreach (TableRow row in tab.Descendants<TableRow>())
                    //{
                    //    result.Append("<tr>");
                    //    foreach (TableCell cell in row.Descendants<TableCell>())
                    //    {
                    //        result.Append("<td>");
                    //        result.Append(cell.InnerText);
                    //        tableCellContent.Add(cell.InnerText.Trim());
                    //        result.Append("</td>");
                    //    }
                    //    result.Append("</tr>");
                    //}
                    //result.Append("</table>");
                }
                }


            wordProcessingDoc.Close();

           // return result.ToString();


            document.Save(XmlFilePath);
        }
        catch (XmlException ex)
        {
            Console.WriteLine(ex.Message);
            System.IO.StreamWriter ErrorLog = new System.IO.StreamWriter(LogFilePath, true);
            ErrorLog.WriteLine(DocmFileName);
            ErrorLog.Close();

        }
    }

    private static string GetBase64Image(Stream inputData)
    {
        byte[] data = new byte[inputData.Length];
        inputData.Read(data, 0, data.Length);
        return Convert.ToBase64String(data);
    }
}

