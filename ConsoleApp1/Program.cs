// See https://aka.ms/new-console-template for more information
using System.Collections.Generic;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;
using Microsoft.VisualBasic;
using iText.Svg.Renderers.Impl;
using System.Runtime.Intrinsics.Arm;
using System.Security.Cryptography.X509Certificates;

public static class PdfHelper
{
    public static IEnumerable<string> ExtractText(string filename)
    {
        using(var r = new PdfReader(filename)) 
        using(var doc = new PdfDocument(r))
        {
            for (int i = 1; i<doc.GetNumberOfPages(); i++)
            {
                ITextExtractionStrategy strategy = new LocationTextExtractionStrategy();
                string text = PdfTextExtractor.GetTextFromPage(doc.GetPage(i), strategy);
                yield return text;
            }
        }
    }
}


namespace pdfText
{
    public static class VR
    {
        public static void vrData(string filename)
        {
            var lines = PdfHelper.ExtractText(filename).ToList();
            Regex rxSensitive = new Regex("This document has been created automatically and is valid without signature", RegexOptions.Compiled);
            string buyPartNumber = "release";
            string datePart1 = "Backlog Immediate requirement";
            string datePart2 = "(30) Quantity";
            string datePart3 = "(66)";
            string datePart4 = "T";
            string datePart5 = "FZ";
            string qtyPart1 = "(30) Quantity (M)";
            string qtyPart2 = "(63) Cum. received qty.";
            string pageText = "Page";
            string pageTwo = "Supplier’s reference number";
            string commaText = ",";
            string dotText = ".";
            Regex rxM = new Regex("M", RegexOptions.Compiled);

            foreach (var line in lines)
            {
                List<Part> parts = new List<Part>();
                MatchCollection matches = rxSensitive.Matches(line);
                if (matches.Count > 0)
                    continue;
                int page = line.IndexOf(pageText) + pageText.Length + 1;
                var pageCount = Convert.ToInt32(line[page].ToString());
                //Console.WriteLine(pageCount);


                if (pageCount == 1)
                {
                    //搜尋件號
                    int partNumberFirst = line.IndexOf(buyPartNumber) - 1;
                    var partNumber = line[partNumberFirst].ToString();
                    for (int i = partNumberFirst - 15; i < partNumberFirst; i++)
                        partNumber += line[i];

                    //Console.WriteLine(partNumber.Trim());

                    //第一頁搜尋日期
                    int dateFirst = line.IndexOf(datePart1) + datePart1.Length + 1;
                    int dateSecond = line.IndexOf(datePart2) - 1;
                    var dateLines = line[dateFirst].ToString();
                    for (int i = dateFirst + 1; i < dateSecond; i++) 
                        dateLines += line[i];

                    //第一頁搜尋數量
                    int qtyFirst = line.IndexOf(qtyPart1) + qtyPart1.Length + 1;
                    int qtySecond = line.IndexOf(qtyPart2) - 1;
                    var qtyLines = line[qtyFirst].ToString();
                    for (int i = qtyFirst + 1; i < qtySecond; i++) 
                        qtyLines += line[i];

                    int dateLength = dateLines.Length;
                    int qtyLength = qtyLines.Length;
                    while (dateLength > 0 && qtyLength > 0)
                    {
                        int i = 0;
                        string date;
                        string year;
                        string qtyText;
                        decimal qty;
                        string pattern = "dd.MM.yyyy";
                        DateTime parseDate;
                        bool result = int.TryParse(dateLines[0].ToString() , out i);
                        int commaFirst = qtyLines.IndexOf(commaText);
                        if (result)
                        {
                            date = dateLines.Substring(0, 10);
                            DateTime.TryParseExact(date, pattern, null, DateTimeStyles.None, out parseDate);
                            qtyText = qtyLines.Substring(0, commaFirst).Trim();
                            int dotFirst = qtyLines.IndexOf(dotText);
                            if (dotFirst != -1)
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                                qty *= 1000;
                            }
                            else
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                            }
                            parts.Add(new Part() { PartNumber = partNumber , PartDate = parseDate , PartQty = qty});
                            dateLength -= 11;
                            qtyLength -= commaFirst + 3;
                            if (dateLength >= 0) dateLines = dateLines.Substring(11).Trim();
                            else continue;
                            if (qtyLength >= 0) qtyLines = qtyLines.Substring(commaFirst + 3).Trim();
                            else continue;
                        }
                        else if (dateLines[0].ToString() == "W")
                        {
                            date = dateLines.Substring(2, 2);
                            year = dateLines.Substring(5, 4);
                            int weekNumber = Convert.ToInt32(date);
                            int yearNumber = Convert.ToInt32(year);
                            DateTime firstDateOfYear = new DateTime(yearNumber, 1, 1);
                            DateTime dayInWeek = firstDateOfYear.AddDays((weekNumber - 1)*7);
                            DateTime firstDayInWeek = dayInWeek.Date;
                            qtyText = qtyLines.Substring(0, commaFirst).Trim();
                            int dotFirst = qtyLines.IndexOf(dotText);
                            if (dotFirst != -1)
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                                qty *= 1000;
                            }
                            else
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                            }
                            parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDayInWeek, PartQty = qty});
                            dateLength -= 10;
                            qtyLength -= commaFirst + 3;
                            if (dateLength >= 0) dateLines = dateLines.Substring(10).Trim();
                            else continue;
                            if (qtyLength >= 0) qtyLines = qtyLines.Substring(commaFirst + 3).Trim();
                            else continue;
                        }
                        else if (dateLines[0].ToString() == "M")
                        {
                            date = dateLines.Substring(2, 2);
                            year = dateLines.Substring(5, 4);
                            int monthNumber = Convert.ToInt32(date);
                            int yearNumber = Convert.ToInt32(year);
                            DateTime firstDateOfMonth = new DateTime(yearNumber, monthNumber, 1);
                            qtyText = qtyLines.Substring(0, commaFirst).Trim();
                            int dotFirst = qtyLines.IndexOf(dotText);
                            if (dotFirst != -1)
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                                qty *= 1000;
                            }
                            else
                            {
                                bool qtyResult = decimal.TryParse(qtyText, out qty);
                            }
                            parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDateOfMonth, PartQty = qty});
                            dateLength -= 10;
                            qtyLength -= commaFirst + 3;
                            if (dateLength >= 0) dateLines = dateLines.Substring(10).Trim();
                            else continue;
                            if (qtyLength >= 0) qtyLines = qtyLines.Substring(commaFirst + 3).Trim();
                            else continue;
                        }

                    }
                    /*
                    foreach (Part apart in parts)
                        Console.WriteLine(apart);
                    */
                    //第一頁搜尋表格
                    int formFirst = line.IndexOf(datePart3);
                    int formSecond = line.IndexOf(datePart4, formFirst) + 2;
                    int formThird = line.IndexOf(datePart5, formSecond);
                    int formLastPart1 = line.IndexOf(datePart4, formThird) + 2;
                    int formLastPart2 = line.IndexOf(datePart5, formLastPart1);

                    var formLinesPart1 = line[formSecond].ToString();
                    for (int i = formSecond + 1; i < formThird; i++)
                        formLinesPart1 += line[i];
                  
                    var formLinesPart2 = line[formLastPart1].ToString();
                    for (int i = formLastPart1 + 1; i < formLastPart2; i++)
                        formLinesPart2 += line[i];

                    VRHelp(formLinesPart1, partNumber);
                    VRHelp(formLinesPart2, partNumber);                 

                }

                else
                {
                    //搜尋件號
                    int partNumberFirst = line.IndexOf(buyPartNumber) - 1;
                    var partNumber = line[partNumberFirst].ToString();
                    for (int i = partNumberFirst - 15; i < partNumberFirst; i++)
                        partNumber += line[i];

                    //1
                    int formFirst = line.IndexOf(pageTwo);
                    int formSecond = line.IndexOf(datePart4, formFirst) + 2;
                    int formThird = line.IndexOf(datePart5, formSecond);
                    //2
                    int formLastPart1 = line.IndexOf(datePart4, formThird) + 3;
                    int formLastPart2 = line.IndexOf(datePart5, formLastPart1);
                    //3
                    int formLastPart3 = line.IndexOf(datePart4, formLastPart2) + 3;
                    int formLastPart4 = line.IndexOf(datePart5, formLastPart3);
                    //4
                    int formLastPart5 = line.IndexOf(datePart4, formLastPart4) + 3;
                    int formLastPart6 = line.IndexOf(datePart5, formLastPart5);
                    //5
                    int formLastPart7 = line.IndexOf(datePart4, formLastPart6) + 3;
                    int formLastPart8 = line.IndexOf(datePart5, formLastPart7);
                    //6
                    int formLastPart9 = line.IndexOf(datePart4, formLastPart8) + 3;
                    int formLastPart10 = line.IndexOf(datePart5, formLastPart9);
                    //7
                    int formLastPart11 = line.IndexOf(datePart4, formLastPart10) + 3;
                    int formLastPart12 = line.IndexOf(datePart5, formLastPart11);

                    var formLinesPart1 = line[formSecond].ToString();
                    for (int i = formSecond + 1; i < formThird; i++)                   
                        formLinesPart1 += line[i];
                     
                    var formLinesPart2 = line[formLastPart1].ToString();
                    for (int i = formLastPart1 + 1; i < formLastPart2; i++)                   
                        formLinesPart2 += line[i];
                    
                    var formLinesPart3 = line[formLastPart3].ToString();
                    for (int i = formLastPart3 + 1; i < formLastPart4; i++)
                        formLinesPart3 += line[i];
                    
                    var formLinesPart4 = line[formLastPart5].ToString();
                    for (int i = formLastPart5 + 1; i < formLastPart6; i++)                   
                        formLinesPart4 += line[i];
                    
                    var formLinesPart5 = line[formLastPart7].ToString();
                    for (int i = formLastPart7 + 1; i < formLastPart8; i++)                    
                        formLinesPart5 += line[i];
                    
                    var formLinesPart6 = line[formLastPart9].ToString();
                    for (int i = formLastPart9 + 1; i < formLastPart10; i++)                    
                        formLinesPart6 += line[i];
                    
                    var formLinesPart7 = line[formLastPart11].ToString();
                    for (int i = formLastPart11 + 1; i < formLastPart12; i++)                    
                        formLinesPart7 += line[i];

                    VRHelp(formLinesPart1, partNumber);
                    VRHelp(formLinesPart2, partNumber);
                    VRHelp(formLinesPart3, partNumber);
                    VRHelp(formLinesPart4, partNumber);
                    VRHelp(formLinesPart5, partNumber);
                    VRHelp(formLinesPart6, partNumber);
                    VRHelp(formLinesPart7, partNumber);


                    
                }              

            }
            
        }
        public class Part
        {
            public string PartNumber { get; set; }
            public DateTime PartDate { get; set; }
            public decimal PartQty { get; set; }
            public override string ToString()
            {
                return "PN: " + PartNumber + " Date: " + PartDate + " Qty: " + PartQty;
            }

        }
        public static void VRHelp(string formLinesPart1, string partNumber)
        {
            string commaText = ",";
            string dotText = ".";
            Regex rxM = new Regex("M", RegexOptions.Compiled);
            List<Part> parts = new List<Part>();
            MatchCollection mMatches = rxM.Matches(formLinesPart1);
            if (mMatches.Count == 1)
            {
                int formMiddle = formLinesPart1.IndexOf("M");
                var dateFormPart1 = formLinesPart1.Substring(0, formMiddle);
                var qtyFormPart1 = formLinesPart1.Substring(formMiddle + 1);
                int dateFormLength = dateFormPart1.Length;
                int qtyFormLength = qtyFormPart1.Length;
                while (dateFormLength > 0 && qtyFormLength > 0)
                {
                    int i = 0;
                    string date;
                    string year;
                    string qtyText;
                    decimal qty;
                    string pattern = "dd.MM.yyyy";
                    DateTime parseDate;
                    bool result = int.TryParse(dateFormPart1[0].ToString(), out i);
                    int commaFirst = qtyFormPart1.IndexOf(commaText);
                    if (result)
                    {
                        date = dateFormPart1.Substring(0, 10);
                        DateTime.TryParseExact(date, pattern, null, DateTimeStyles.None, out parseDate);
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyText.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = parseDate, PartQty = qty });
                        dateFormLength -= 11;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(11).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }

                    else if (dateFormPart1[0].ToString() == "W")
                    {
                        date = dateFormPart1.Substring(2, 2);
                        year = dateFormPart1.Substring(5, 4);
                        int weekNumber = Convert.ToInt32(date);
                        int yearNumber = Convert.ToInt32(year);
                        DateTime firstDateOfYear = new DateTime(yearNumber, 1, 1);
                        DateTime dayInWeek = firstDateOfYear.AddDays((weekNumber - 1) * 7);
                        DateTime firstDayInWeek = dayInWeek.Date;
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyText.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDayInWeek, PartQty = qty });
                        dateFormLength -= 10;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(10).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }
                    else if (dateFormPart1[0].ToString() == "M")
                    {
                        date = dateFormPart1.Substring(2, 2);
                        year = dateFormPart1.Substring(5, 4);
                        int monthNumber = Convert.ToInt32(date);
                        int yearNumber = Convert.ToInt32(year);
                        DateTime firstDateOfMonth = new DateTime(yearNumber, monthNumber, 1);
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyText.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDateOfMonth, PartQty = qty });
                        dateFormLength -= 10;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(10).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }

                }
            }
            else
            {
                int formMiddle = formLinesPart1.LastIndexOf("M");
                var dateFormPart1 = formLinesPart1.Substring(0, formMiddle);
                var qtyFormPart1 = formLinesPart1.Substring(formMiddle + 1);
                int dateFormLength = dateFormPart1.Length;
                int qtyFormLength = qtyFormPart1.Length;
                while (dateFormLength > 0 && qtyFormLength > 0)
                {
                    int i = 0;
                    string date;
                    string year;
                    string qtyText;
                    decimal qty;
                    string pattern = "dd.MM.yyyy";
                    DateTime parseDate;
                    bool result = int.TryParse(dateFormPart1[0].ToString(), out i);
                    int commaFirst = qtyFormPart1.IndexOf(commaText);
                    if (result)
                    {
                        date = dateFormPart1.Substring(0, 10);
                        DateTime.TryParseExact(date, pattern, null, DateTimeStyles.None, out parseDate);
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyFormPart1.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = parseDate, PartQty = qty });
                        dateFormLength -= 11;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(11).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }

                    else if (dateFormPart1[0].ToString() == "W")
                    {
                        date = dateFormPart1.Substring(2, 2);
                        year = dateFormPart1.Substring(5, 4);
                        int weekNumber = Convert.ToInt32(date);
                        int yearNumber = Convert.ToInt32(year);
                        DateTime firstDateOfYear = new DateTime(yearNumber, 1, 1);
                        DateTime dayInWeek = firstDateOfYear.AddDays((weekNumber - 1) * 7);
                        DateTime firstDayInWeek = dayInWeek.Date;
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyFormPart1.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDayInWeek, PartQty = qty });
                        dateFormLength -= 10;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(10).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }
                    else if (dateFormPart1[0].ToString() == "M")
                    {
                        date = dateFormPart1.Substring(2, 2);
                        year = dateFormPart1.Substring(5, 4);
                        int monthNumber = Convert.ToInt32(date);
                        int yearNumber = Convert.ToInt32(year);
                        DateTime firstDateOfMonth = new DateTime(yearNumber, monthNumber, 1);
                        qtyText = qtyFormPart1.Substring(0, commaFirst).Trim();
                        int dotFirst = qtyFormPart1.IndexOf(dotText);
                        if (dotFirst != -1)
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                            qty *= 1000;
                        }
                        else
                        {
                            bool qtyResult = decimal.TryParse(qtyText, out qty);
                        }
                        parts.Add(new Part() { PartNumber = partNumber, PartDate = firstDateOfMonth, PartQty = qty });
                        dateFormLength -= 10;
                        qtyFormLength -= commaFirst + 3;
                        if (dateFormLength > 0) dateFormPart1 = dateFormPart1.Substring(10).Trim();
                        else continue;
                        if (qtyFormLength > 0) qtyFormPart1 = qtyFormPart1.Substring(commaFirst + 3).Trim();
                        else continue;
                    }
                }
            }
            /*
            foreach (Part aPart in parts)
            {
                Console.WriteLine(aPart);
            }*/
            
        }
    }
}

/*
public static class Program
{
    public static void Main()
    {
        pdfText.VR.vrData("Lenawee.pdf");
        Console.Write("Press <Enter> to exit... ");
        while (Console.ReadKey().Key != ConsoleKey.Enter) { }
    }
}*/
