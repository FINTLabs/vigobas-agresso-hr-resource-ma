using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Net;
using System.IO;

namespace TestXdocument
{
    class Program
    {

        static void Main(string[] args)
        {
            String infile = "c:\\data\\agresso_stamdata.xml";
            String _AgressoGetAllPersons = File.ReadAllText(infile);

            // Parser alle elementer i /ExportInfo/Resources/Resource
            foreach (XElement person in XDocument.Parse(_AgressoGetAllPersons).Element("ExportInfo").Element("Resources").Elements("Resource"))
            {
                try
                {
                    foreach (XElement position in person.Element("Employments").Elements("Employment"))
                    {
                        Console.Write("");
                    }
                }
                catch { }
            }
        }
    }
}
