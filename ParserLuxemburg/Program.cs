using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Net;
using NLog;

namespace ParserLuxemburg
{
    class Program
    {
        private static Logger logger = LogManager.GetCurrentClassLogger();
        [STAThread]
        static void Main(string[] args)
        {
            logger.Info("НАЧАЛО РАБОТЫ ПАРСЕРА");
            List<string> ListName = new List<string>();
            List<string> Group = new List<string>();
            List<string> Prog = new List<string>();
            List<string> ISIN = new List<string>();
            List<string> Name = new List<string>();
            List<string> Type = new List<string>();
            List<string> Curenncy = new List<string>();
            List<string> ClosingDate = new List<string>();
            List<string> ClosingPrice = new List<string>();
            string[] url = new string[3];
            string dt = DateTime.Today.ToString("yyyy" + "MM" + "dd");
            url[0] = "https://www.bourse.lu/api/official-list/bdl-market/" + dt;
            url[1] = "https://www.bourse.lu/api/official-list/euro-mtf/" + dt;
            url[2] = "https://www.bourse.lu/api/official-list/lux-sol/" + dt;
            WebClient web = new WebClient();
            string path = args[0];     
            for (int k = 0; k < 3; k++)
            {
                logger.Info("Начало загрузки " + (k + 1) + " таблицы");
                try
                {
                    web.DownloadFile(url[k], path + "\\" + dt + ".pdf");
                    PdfReader pdfReader = new PdfReader(path + "\\" + dt + ".pdf");
                    StringBuilder stringBuilder = new StringBuilder();
                    for (int i = 1; i < pdfReader.NumberOfPages; i++) 
                    {
                        try
                        {
                            int ListNameCount = 0, GroupCount = 0, ProgCount = 0;
                            ITextExtractionStrategy textExtractionStrategy = new LocationTextExtractionStrategy();
                            string str = PdfTextExtractor.GetTextFromPage(pdfReader, i, textExtractionStrategy);
                            var text2 = pdfReader.GetPageContent(i);
                            var mas = str.Split('\n');
                            string line;
                            stringBuilder.Append(PdfTextExtractor.GetTextFromPage(pdfReader, i));
                            for (int j = 0; j < mas.Length - 2; j++) 
                            {
                                var mas1 = mas[j].Split(' ');
                                if (mas1[mas1.Length - 1] == "i")
                                {
                                    line = "";
                                    for (int l = 1; l < mas1.Length - 6; l++)
                                    {
                                        line += mas1[l] + " ";
                                    }
                                    ISIN.Add(mas1[0]);
                                    Name.Add(line.Replace(",", "."));
                                    Type.Add(mas1[mas1.Length - 5].Replace(",", "."));
                                    Curenncy.Add(mas1[mas1.Length - 4].Replace(",", "."));
                                    ClosingDate.Add(mas1[mas1.Length - 3].Replace(",", "."));
                                    ClosingPrice.Add(mas1[mas1.Length - 2].Replace(",", ".") + " " + mas1[mas1.Length - 1].Replace(",", "."));
                                    if (ProgCount == 0)
                                    {
                                        Prog.Add(" ");
                                    }
                                    if (GroupCount == 0)
                                    {
                                        Group.Add(" ");
                                    }
                                    if (ListNameCount == 0)
                                    {
                                        ListName.Add(" ");
                                    }
                                    ProgCount = 0;
                                    GroupCount = 0;
                                    ListNameCount = 0;
                                }
                                else if (mas1[0] == "Prog.:")
                                {
                                    Prog.Add(mas[j].Replace(",", "."));
                                    ProgCount = 1;
                                }
                                else if (j + 1 < mas.Length)
                                {
                                    if (mas[j + 1] == "ISIN Security Type Ccy Last closing price" | mas[j + 1] == "ISIN Security Type Ccy Last closing price Day volume")
                                    {
                                        ListName.Add(mas[j].Replace(",", "."));
                                        j++;
                                    }
                                    else if (j != 0 & mas[j].Trim() != "")
                                    {
                                        Group.Add(mas[j].Replace(",", "."));
                                        GroupCount = 1;
                                    }
                                }
                            }
                            logger.Info("Загрузка страницы " + i + " завершена");
                        }
                        catch (Exception ex)
                        {
                            logger.Debug("Страница " + i + " не загружена");
                            logger.Debug(ex);
                        }
                    }
                    pdfReader.Close();
                    File.Delete(path + "\\" + dt + ".pdf");
                    logger.Info("Загрузка " + (k + 1) + " таблицы завершилась");
                }
                catch (Exception ex)
                {
                    logger.Debug("Файл с сайта" + url[k] + " не был скачен");
                    logger.Debug(ex);
                }                
            }
            logger.Info("Начало загрузки данных в файл");
            try
            {
                using (StreamWriter streamWriter = new StreamWriter(args[0] + "\\" + DateTime.Today.ToString("dd MM yyyy").Replace(" ", "") + ".csv"))
                {
                    streamWriter.WriteLine("sep=,");
                    streamWriter.WriteLine("Sec. list type, Group, Prog., ISIN, Sec. name, Sec. type, Currency, Last closing date, Last closing type");
                    for (int i = 0; i < ISIN.Count; i++)
                    {
                        streamWriter.WriteLine(ListName[i] + ", " + Group[i] + ", " + Prog[i] + ", " + ISIN[i] + ", " + Name[i] + ", " + Type[i] + ", " + Curenncy[i] + ", " + ClosingDate[i] + ", " + ClosingPrice[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Info(ex);
                logger.Debug(ex);
            }           
            logger.Info("ОКОНЧАНИЕ РАБОТЫ ПАРСЕРА");
        }
    }
}
