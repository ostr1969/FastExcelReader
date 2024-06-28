using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.IO;
using System.Xml;
using Prism.Commands;
using System.Xml.Linq;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace ExcelReader
{
    public class FastExcelReader
    {
        string Filename, ExtractPath,TempDir;
        Dictionary<string, XDocument> sheetdocs;
        Dictionary<int, string> Strings;
        // Dictionary<string, string> sheetnames;
        public List<string> ColumnNames;
        public  FastExcelReader(string filename)
        {
            Filename = filename;
            sheetdocs = new Dictionary<string, XDocument>();
            sheetnames = new Dictionary<string, string>();
            Strings = new Dictionary<int, string>();
            ColumnNames = GetAllExcelColumnAddresses();
            ExtractDoc();
            GetSheetnames();
            GetStrings();
        }
        public void LoadSheet(string sheet)
        { if (!sheetdocs.ContainsKey(sheet))
            { var sheettempname = sheetnames[sheet];
                XDocument doc = XDocument.Load($"{TempDir}\\xl\\worksheets\\{sheettempname}.xml");
                sheetdocs[sheet] = doc;
            }
        }
        public string GetCell(string sheetname,string addr1)
        {
            if (!sheetdocs.ContainsKey(sheetname)) LoadSheet(sheetname);
            if (!Regex.IsMatch(addr1, "^[A-Z]{1,3}[1-9][0-9]*$"))
                throw new Exception("Wrong adress pattern");
            addr1 = addr1.ToUpper();
            XNamespace dnamespace= sheetdocs[sheetname].Root.GetDefaultNamespace();
            var cell = sheetdocs[sheetname].Descendants(dnamespace+ "c").
                FirstOrDefault(o => o.Attribute("r").Value == addr1);
            if (cell == null)
                return null;
            
            var val= cell.Element(dnamespace+"v");
            if (!(val == null))
                return val.Value;
            else
                return null;
                
        }

        public List<double> GetDoubleRange(string sheetname, string rangeToGet)
        {
            List<double> res = new List<double>();
            if (!sheetdocs.ContainsKey(sheetname)) LoadSheet(sheetname);
            if (!Regex.IsMatch(rangeToGet, "^[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*$"))
                throw new Exception("Wrong range pattern");
            XNamespace dnamespace = sheetdocs[sheetname].Root.GetDefaultNamespace();
            var startcol = Regex.Match(rangeToGet.Split(':')[0], "[A-Z]+").Value;
            var endcol = Regex.Match(rangeToGet.Split(':')[1], "[A-Z]+").Value;
            var startrow = int.Parse(Regex.Match(rangeToGet.Split(':')[0], "[0-9]+").Value);
            var endrow = int.Parse(Regex.Match(rangeToGet.Split(':')[1], "[0-9]+").Value);
            var startcolindex = ColumnNames.IndexOf(startcol);
            var endcolindex = ColumnNames.IndexOf(endcol);
            if (startcolindex == -1 || endcolindex == -1)
                throw new Exception("Wrong range address");
            for (var row=startrow;row<endrow+1; row++)
            {
                    for (var col= startcolindex; col< startcolindex+1;col++)
                {
                    var celladdr = ColumnNames[col] + row.ToString();
                    var rowel = sheetdocs[sheetname].Descendants(dnamespace+"c").
                         FirstOrDefault(o => o.Attribute("r").Value == celladdr);
                    
                    if (rowel.Attribute("t")?.Value == "s")
                        return null;
                    if (!(rowel == null))
                        if (!(rowel.Descendants(dnamespace + "v").First() == null))
                        {
                            double v = 0;
                            double.TryParse(rowel.Descendants(dnamespace + "v").First().Value, out v);
                            res.Add(v);
                        }
                        else
                            res.Add(0);
                    else
                        res.Add(0);
                            
                }

                    

            }
            return null;
        }

        public void GetStrings()
        {
            XDocument doc = XDocument.Load($"{TempDir}\\xl\\sharedStrings.xml");
            XNamespace dnamespace = doc.Root.GetDefaultNamespace();
            var els = doc.Descendants(dnamespace+"t");
            int i = 0;
            foreach (var el in els)
                Strings[++i] = el.Value;

        }
      public  Dictionary<string, string> sheetnames;
        public void GetSheetnames()
        {
            XDocument wbdoc = XDocument.Load($"{TempDir}\\xl\\workbook.xml");
            
            var sheets = wbdoc.Root.Elements().FirstOrDefault(o=>o.Name.LocalName=="sheets");
            foreach (var sht in sheets.Elements())
            {
                var xlname=sht.Attribute(wbdoc.Root.GetNamespaceOfPrefix("r") + "id").Value;
                var sheetname = sht.Attribute("name").Value;
                sheetnames[sheetname] = xlname.Replace("rId", "sheet");
            }

        }
        public void ExtractDoc()
        {
            TempDir = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Filename));
            Directory.CreateDirectory(TempDir);
            using (ZipArchive archive = ZipFile.OpenRead(Filename))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    entry.ExtractToFile(Path.Combine(TempDir, entry.FullName), true);
                }
            }



        }
        static List<string> GetAllExcelColumnAddresses()
        {
            List<string> columns = new List<string>();

            // Single letter columns (A to Z)
            for (char first = 'A'; first <= 'Z'; first++)
            {
                columns.Add(first.ToString());
            }

            // Double letter columns (AA to ZZ)
            for (char first = 'A'; first <= 'Z'; first++)
            {
                for (char second = 'A'; second <= 'Z'; second++)
                {
                    columns.Add(first.ToString() + second);
                }
            }

            // Triple letter columns (AAA to ZZZ)
            for (char first = 'A'; first <= 'Z'; first++)
            {
                for (char second = 'A'; second <= 'Z'; second++)
                {
                    for (char third = 'A'; third <= 'Z'; third++)
                    {
                        columns.Add(first.ToString() + second + third);
                    }
                }
            }

            return columns;
        }
    }
}
