﻿using System;
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

public class Parser
{
    public Parser() {}
    
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
    }

    public void ParseDocxDocument(String FilePath, String DocxFileName)
    {
        string DocxFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocxFileName.ToString() + ".docx"));
        string XmlFilePath = System.IO.Path.Combine(FilePath.ToString(), (DocxFileName.ToString() + ".xml"));

        int TagsCount = 0;
        int TotalCount = 10;

        createXML(XmlFilePath, "");
        XmlDocument document = new XmlDocument();
        document.Load(XmlFilePath);
        XmlNode element = document.CreateElement("info");
        document.DocumentElement.AppendChild(element);

        XmlNode title = document.CreateElement("title");
        title.InnerText = DocxFileName;
        element.AppendChild(title);

        XmlNode chapter = document.CreateElement("chapter");
        document.DocumentElement.AppendChild(chapter);

        WordprocessingDocument wordProcessingDoc = WordprocessingDocument.Open(DocxFilePath, true);
        List<ImagePart> imgPart = wordProcessingDoc.MainDocumentPart.ImageParts.ToList();
        imgPart.Reverse();
        List<string> tableCellContent = new List<string>();
        IEnumerable<Paragraph> paragraphElement = wordProcessingDoc.MainDocumentPart.Document.Descendants<Paragraph>();
        int imgCounter = 0;

        foreach (OpenXmlElement section in wordProcessingDoc.MainDocumentPart.Document.Body.Elements<OpenXmlElement>())
        {
            if (section.GetType().Name == "Paragraph")
            {
                Paragraph par = (Paragraph)section;

                string pathString = System.IO.Path.Combine(FilePath, "img");
                DirectoryInfo di = System.IO.Directory.CreateDirectory(pathString);

                IEnumerable<Run> runs = par.Descendants<Run>();
                foreach (Run run in runs)
                {
                    if (run.HasChildren)
                    {
                        foreach (OpenXmlElement chield in run.ChildElements.Where(o => o.GetType().Name == "Drawing"))   //обработка картинок
                        {
                            // <imagedata fileref="image.png" width="6in" depth="5.5in" scale="300"/>
                            Console.WriteLine("Picture!");
                            XmlNode imagedata = document.CreateElement("imagedata");
                            chapter.AppendChild(imagedata);
                            XmlAttribute attribute = document.CreateAttribute("fileref");
                            Image img = System.Drawing.Image.FromStream(imgPart[imgCounter].GetStream());
                            string imgSavePath = pathString + @"\" + imgCounter + ".jpeg";
                            img.Save(imgSavePath);
                            attribute.Value = string.Format(imgSavePath + " />");
                            imagedata.Attributes.Append(attribute);
                            imgCounter++;
                            
                            if (TagsCount < TotalCount)
                                TagsCount++;
                            else
                            {
                                chapter = document.CreateElement("chapter");
                                document.DocumentElement.AppendChild(chapter);
                                TagsCount = 0;
                            }
                        }
                        foreach (OpenXmlElement list in run.ChildElements.Where(o => o.GetType().Name == "NumeredProperty"))
                        {
                            Console.WriteLine("List!");

                        }
                    }
                }

                IEnumerable<Text> textElement = par.Descendants<Text>();

                foreach (Text t in textElement.Where(o => !tableCellContent.Contains(o.Text)))   //добавление текста
                {
                    XmlNode para = document.CreateElement("para");
                    para.InnerText = t.Text;
                    chapter.AppendChild(para);
                    if (TagsCount < TotalCount)
                        TagsCount++;
                    else
                    {
                        chapter = document.CreateElement("chapter");
                        document.DocumentElement.AppendChild(chapter);
                        TagsCount = 0;
                    }
                }
            }
            else if (section.GetType().Name == "Table")
            {
                Table tab = (Table)section;
                IEnumerable<TableRow> tblrow = tab.Descendants<TableRow>();
                Console.WriteLine(tblrow.Count().ToString());

                IEnumerable<TableGrid> tblGrid = tab.Descendants<TableGrid>();
                Console.WriteLine("Table 2!");
                XmlNode table = document.CreateElement("table");
                chapter.AppendChild(table);
                XmlAttribute frame = document.CreateAttribute("frame");     
                frame.Value = "all";                        //      !!!    добавить варианты фрейма      !!!
                table.Attributes.Append(frame);

                XmlNode tgroup = document.CreateElement("tgroup");
                table.AppendChild(tgroup);
                XmlAttribute colspec = document.CreateAttribute("colspec");
                tgroup.Attributes.Append(colspec);

                XmlNode tbody = document.CreateElement("tbody");
                table.AppendChild(tbody);

                foreach (TableRow row in tab.Descendants<TableRow>())
                {
                    XmlNode trow = document.CreateElement("row");
                    if (TagsCount < TotalCount)
                    {
                        TagsCount++;
                        tbody.AppendChild(trow);
                    }
                    else
                    {
                        TagsCount = 0;
                        chapter = document.CreateElement("chapter");
                        document.DocumentElement.AppendChild(chapter);
                        TagsCount = 0;
                        table = document.CreateElement("table");
                        chapter.AppendChild(table);
                        frame = document.CreateAttribute("frame"); 
                        frame.Value = "all";                        //      !!!   изменить верхинй и нижний фреймы для средних частей таблицы     !!!
                        table.Attributes.Append(frame);

                        tgroup = document.CreateElement("tgroup");
                        colspec = document.CreateAttribute("colspec");
                        tgroup.Attributes.Append(colspec);

                        tbody = document.CreateElement("tbody");
                        table.AppendChild(tbody);
                    }
                    foreach (TableCell cell in row.Descendants<TableCell>())
                    {
                        XmlNode entry = document.CreateElement("entry");
                        Console.WriteLine(cell.TableCellProperties.TableCellWidth.Width.InnerText);
                        entry.InnerText = cell.InnerText;
                        trow.AppendChild(entry);
                    }
                }
            }
        }
        wordProcessingDoc.Close();
        document.Save(XmlFilePath);
    }
}

