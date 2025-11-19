using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Nastavení licence pro EPPlus (nekomerční použití)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("Excel Processor - Zpracování Heureka a Sklad Excel souborů");
            Console.WriteLine("===========================================================");

            try
            {
                // Cesty k vstupním souborům
                string heurekaFile = GetInputFilePath("Zadejte cestu k Heureka Excel souboru (nebo stiskněte Enter pro 'heureka.xlsx'): ", "heureka.xlsx");
                string skladFile = GetInputFilePath("Zadejte cestu k Sklad Excel souboru (nebo stiskněte Enter pro 'sklad.xlsx'): ", "sklad.xlsx");
                
                // Cesta k výstupnímu souboru
                string outputFile = GetOutputFilePath("Zadejte cestu k výstupnímu Excel souboru (nebo stiskněte Enter pro 'vysledek.xlsx'): ", "vysledek.xlsx");

                // Zpracování souborů
                ProcessExcelFiles(heurekaFile, skladFile, outputFile);

                Console.WriteLine($"\nÚspěšně dokončeno! Výstupní soubor: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nChyba: {ex.Message}");
                Console.WriteLine($"Detail: {ex.StackTrace}");
            }
        }

        static string GetInputFilePath(string prompt, string defaultValue)
        {
            Console.Write(prompt);
            string? input = Console.ReadLine()?.Trim();
            
            if (string.IsNullOrEmpty(input))
            {
                return defaultValue;
            }
            
            if (!File.Exists(input))
            {
                throw new FileNotFoundException($"Soubor nenalezen: {input}");
            }
            
            return input;
        }

        static string GetOutputFilePath(string prompt, string defaultValue)
        {
            Console.Write(prompt);
            string? input = Console.ReadLine()?.Trim();
            
            return string.IsNullOrEmpty(input) ? defaultValue : input;
        }

        static void ProcessExcelFiles(string heurekaPath, string skladPath, string outputPath)
        {
            Console.WriteLine($"\nNačítání Heureka souboru: {heurekaPath}");
            var heurekaData = ReadHeurekaFile(heurekaPath);
            
            Console.WriteLine($"Načítání Sklad souboru: {skladPath}");
            var skladData = ReadSkladFile(skladPath);

            Console.WriteLine($"Zpracování dat a výpočet rabatů...");
            var processedData = ProcessData(heurekaData, skladData);

            Console.WriteLine($"Zapisování výsledků do: {outputPath}");
            WriteExcelFile(outputPath, processedData);
        }

        // Načtení Heureka Excel souboru
        // Sloupec L - prodejní cena (s DPH)
        // Sloupec I - heureka odkaz
        // Sloupec F - EAN (pro párování)
        // Sloupce P-AH - vzestupně seřazené hodnoty nejnižších cen (P=16, Q=17, ..., AH=34)
        static List<HeurekaRecord> ReadHeurekaFile(string filePath)
        {
            var data = new List<HeurekaRecord>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                if (worksheet.Dimension == null)
                {
                    Console.WriteLine($"Varování: List '{worksheet.Name}' je prázdný.");
                    return data;
                }

                int rowCount = worksheet.Dimension.Rows;

                // Načtení dat z řádků (předpokládáme, že první řádek jsou hlavičky)
                for (int row = 2; row <= rowCount; row++)
                {
                    var eanCell = worksheet.Cells[row, 6];        // Sloupec F - EAN
                    var prodejniCenaCell = worksheet.Cells[row, 12]; // Sloupec L
                    var heurekaOdkazCell = worksheet.Cells[row, 9];  // Sloupec I

                    // Kontrola, zda řádek obsahuje data
                    if (prodejniCenaCell.Value == null && heurekaOdkazCell.Value == null && eanCell.Value == null)
                        continue;

                    // Načtení hodnot ze sloupců P-AH (sloupce 16-34) - nejnižší ceny vzestupně
                    // Načítáme všechny hodnoty včetně prázdných, aby pozice odpovídaly pořadí sloupců
                    var nejnizsiCeny = new List<double?>();
                    for (int col = 16; col <= 34; col++) // Sloupce P až AH (19 sloupců)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        double cena = GetNumericValue(cellValue);
                        nejnizsiCeny.Add(cena > 0 ? cena : (double?)null);
                    }

                    var record = new HeurekaRecord
                    {
                        EAN = eanCell.Value?.ToString()?.Trim() ?? string.Empty,
                        ProdejniCena = GetNumericValue(prodejniCenaCell.Value),
                        HeurekaOdkaz = heurekaOdkazCell.Value?.ToString() ?? string.Empty,
                        NejnizsiCeny = nejnizsiCeny,
                        RowNumber = row
                    };

                    if (record.ProdejniCena > 0 || !string.IsNullOrWhiteSpace(record.HeurekaOdkaz) || !string.IsNullOrWhiteSpace(record.EAN))
                    {
                        data.Add(record);
                    }
                }
            }

            Console.WriteLine($"  Načteno {data.Count} záznamů z Heureka");
            return data;
        }

        // Načtení Sklad Excel souboru
        // Sloupec A - Kód
        // Sloupec B - EAN
        // Sloupec C - Název Produktu
        // Sloupec F - Sklad
        // Sloupec G - Nákupní cena (bez DPH)
        // Sloupec J - Rule
        static List<SkladRecord> ReadSkladFile(string filePath)
        {
            var data = new List<SkladRecord>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                
                if (worksheet.Dimension == null)
                {
                    Console.WriteLine($"Varování: List '{worksheet.Name}' je prázdný.");
                    return data;
                }

                int rowCount = worksheet.Dimension.Rows;

                // Načtení dat z řádků (předpokládáme, že první řádek jsou hlavičky)
                for (int row = 2; row <= rowCount; row++)
                {
                    var kodCell = worksheet.Cells[row, 1];      // Sloupec A
                    var eanCell = worksheet.Cells[row, 2];     // Sloupec B
                    var nazevCell = worksheet.Cells[row, 3];   // Sloupec C
                    var skladCell = worksheet.Cells[row, 6];   // Sloupec F
                    var nakupniCenaCell = worksheet.Cells[row, 7]; // Sloupec G
                    var ruleCell = worksheet.Cells[row, 10];   // Sloupec J

                    // Kontrola, zda řádek obsahuje data
                    if (kodCell.Value == null && eanCell.Value == null)
                        continue;

                    var record = new SkladRecord
                    {
                        Kod = kodCell.Value?.ToString() ?? string.Empty,
                        EAN = eanCell.Value?.ToString() ?? string.Empty,
                        Nazev = nazevCell.Value?.ToString() ?? string.Empty,
                        Sklad = GetNumericValue(skladCell.Value),
                        NakupniCena = GetNumericValue(nakupniCenaCell.Value),
                        Rule = ruleCell.Value?.ToString()?.Trim() ?? string.Empty,
                        RowNumber = row
                    };

                    if (!string.IsNullOrWhiteSpace(record.Kod) || !string.IsNullOrWhiteSpace(record.EAN))
                    {
                        data.Add(record);
                    }
                }
            }

            Console.WriteLine($"  Načteno {data.Count} záznamů ze Skladu");
            return data;
        }

        static double GetNumericValue(object? value)
        {
            if (value == null)
                return 0;

            if (value is double d)
                return d;

            if (value is decimal dec)
                return (double)dec;

            if (value is int i)
                return i;

            if (double.TryParse(value.ToString(), out double result))
                return result;

            return 0;
        }

        // Zpracování dat podle pravidel
        static List<ResultRecord> ProcessData(List<HeurekaRecord> heurekaData, List<SkladRecord> skladData)
        {
            var results = new List<ResultRecord>();

            // Vytvoření slovníku pro rychlé vyhledávání podle EAN
            var heurekaByEAN = new Dictionary<string, HeurekaRecord>();
            foreach (var heureka in heurekaData)
            {
                if (!string.IsNullOrWhiteSpace(heureka.EAN))
                {
                    string eanKey = heureka.EAN.Trim();
                    // Pokud už existuje záznam se stejným EAN, použijeme první nalezený
                    if (!heurekaByEAN.ContainsKey(eanKey))
                    {
                        heurekaByEAN[eanKey] = heureka;
                    }
                }
            }

            int matchedCount = 0;
            int unmatchedCount = 0;

            foreach (var sklad in skladData)
            {
                // Najít odpovídající záznam z Heureka podle EAN
                HeurekaRecord? heureka = null;
                
                if (!string.IsNullOrWhiteSpace(sklad.EAN))
                {
                    string eanKey = sklad.EAN.Trim();
                    if (heurekaByEAN.TryGetValue(eanKey, out var foundHeureka))
                    {
                        heureka = foundHeureka;
                        matchedCount++;
                    }
                }

                // Pokud nenajdeme podle EAN, přeskočíme tento záznam
                if (heureka == null)
                {
                    unmatchedCount++;
                    Console.WriteLine($"  Varování: Nenalezen odpovídající záznam v Heureka pro EAN: {sklad.EAN} (Kód: {sklad.Kod})");
                    continue;
                }

                // Výpočet DPH podle názvu produktu
                double dph = CalculateDPH(sklad.Nazev);

                // Výpočet prodejní ceny bez DPH z původní ceny s DPH
                double prodejniCenaBezDPH = heureka.ProdejniCena / (1 + dph / 100);

                // Výpočet rabatu podle pravidla
                double rabat = CalculateRabat(sklad.Rule, prodejniCenaBezDPH, sklad.NakupniCena, sklad.Sklad);

                // Výpočet nové prodejní ceny s DPH
                double novaProdejniCenaBezDPH = prodejniCenaBezDPH * (1 - rabat / 100);
                double novaProdejniCenaSDPH = novaProdejniCenaBezDPH * (1 + dph / 100);

                // Výpočet pozice na heurece podle sloupců P-AH
                // Sloupce P-AH obsahují vzestupně seřazené nejnižší ceny
                // Pozice = pořadí sloupce, kde se nachází prodejní cena (P=1, Q=2, ..., AH=pořadí)
                int poziceNaHeurece = CalculatePoziceNaHeurece(heureka.ProdejniCena, heureka.NejnizsiCeny);

                var result = new ResultRecord
                {
                    Kod = sklad.Kod,
                    Nazev = sklad.Nazev,
                    EAN = sklad.EAN,
                    Sklad = sklad.Sklad,
                    Nakupka = sklad.NakupniCena,
                    NovaProdejniCenaSDPH = Math.Round(novaProdejniCenaSDPH, 2),
                    NovaProdejniCenaBezDPH = Math.Round(novaProdejniCenaBezDPH, 2),
                    Rabat = Math.Round(rabat, 2),
                    PoziceNaHeurece = poziceNaHeurece,
                    OdkazNaHeureku = heureka.HeurekaOdkaz
                };

                results.Add(result);
            }

            Console.WriteLine($"  Spárováno podle EAN: {matchedCount} záznamů");
            if (unmatchedCount > 0)
            {
                Console.WriteLine($"  Varování: {unmatchedCount} záznamů nebylo spárováno (chybí EAN nebo odpovídající záznam v Heureka)");
            }

            return results;
        }

        // Výpočet pozice na heurece podle sloupců P-AH
        // Sloupce P-AH obsahují vzestupně seřazené nejnižší ceny
        // Najde pozici, kde se nachází prodejní cena (P=1, Q=2, ..., AH=19)
        static int CalculatePoziceNaHeurece(double prodejniCena, List<double?> nejnizsiCeny)
        {
            if (nejnizsiCeny == null || nejnizsiCeny.Count == 0)
            {
                return 0; // Pokud nejsou data, vrátíme 0
            }

            // Procházíme sloupce P-AH (indexy 0-18 v seznamu odpovídají sloupcům P-AH)
            // Sloupec P = index 0 = pozice 1
            // Sloupec Q = index 1 = pozice 2
            // ...
            // Sloupec AH = index 18 = pozice 19
            for (int i = 0; i < nejnizsiCeny.Count; i++)
            {
                var cena = nejnizsiCeny[i];
                if (cena.HasValue)
                {
                    // Porovnání s tolerancí pro zaokrouhlovací chyby (0.01)
                    if (Math.Abs(cena.Value - prodejniCena) < 0.01)
                    {
                        // Pozice = pořadí sloupce (P=1, Q=2, ..., AH=19)
                        return i + 1;
                    }
                }
            }

            // Pokud není nalezena přesná shoda, najdeme nejbližší vyšší hodnotu
            // (prodejní cena je mezi dvěma hodnotami nebo vyšší než všechny)
            for (int i = 0; i < nejnizsiCeny.Count; i++)
            {
                var cena = nejnizsiCeny[i];
                if (cena.HasValue && prodejniCena < cena.Value)
                {
                    return i + 1;
                }
            }

            // Pokud je prodejní cena vyšší než všechny hodnoty, vrátíme poslední pozici + 1
            return nejnizsiCeny.Count + 1;
        }

        // Výpočet DPH podle názvu produktu
        // 0% DPH pro produkty obsahující "kniha, knihy, knížka, knížky"
        // 21% DPH pro ostatní
        static double CalculateDPH(string nazev)
        {
            if (string.IsNullOrWhiteSpace(nazev))
                return 21;

            string nazevLower = nazev.ToLower();
            string[] knihaKeywords = { "kniha", "knihy", "knížka", "knížky" };

            foreach (var keyword in knihaKeywords)
            {
                if (nazevLower.Contains(keyword))
                {
                    return 0;
                }
            }

            return 21;
        }

        // Výpočet rabatu podle pravidla
        // Výpočet Rabatu: (prodejní - nákupní) / prodejní (bez DPH) %
        static double CalculateRabat(string rule, double prodejniCenaBezDPH, double nakupniCena, double sklad)
        {
            if (string.IsNullOrWhiteSpace(rule))
            {
                // Výchozí výpočet: (prodejní - nákupní) / prodejní (bez DPH) %
                if (prodejniCenaBezDPH > 0)
                {
                    return ((prodejniCenaBezDPH - nakupniCena) / prodejniCenaBezDPH) * 100;
                }
                return 0;
            }

            rule = rule.Trim().ToLower();

            switch (rule)
            {
                case "rule 1":
                case "1":
                    // Co nejnižší cena na heurece (rabat 10%)
                    return 10;

                case "rule 2":
                case "2":
                    // Rabat minimálně 10%, kde je více než 3 kusů tak klidně rabat 0 a nejnižší cenu
                    if (sklad > 3)
                    {
                        // Nejnižší možná cena = nákupní cena + minimální marže (např. 1%)
                        double minCena = nakupniCena * 1.01;
                        if (minCena < prodejniCenaBezDPH)
                        {
                            // Vypočítat rabat potřebný pro dosažení minimální ceny
                            double calculatedRabat = ((prodejniCenaBezDPH - minCena) / prodejniCenaBezDPH) * 100;
                            return Math.Max(0, calculatedRabat); // Minimálně 0%
                        }
                        // Pokud minimální cena je vyšší než prodejní, použijeme 0% rabat
                        return 0;
                    }
                    // Pokud je sklad <= 3, minimálně 10% rabat
                    return 10;

                case "rule 3":
                case "3":
                    // Vypočítat rabat 20%
                    return 20;

                case "rule 4":
                case "4":
                    // Cena nákupky + pět procent
                    // Cílová cena = nákupní cena * 1.05
                    double cilovaCena = nakupniCena * 1.05;
                    if (cilovaCena < prodejniCenaBezDPH && prodejniCenaBezDPH > 0)
                    {
                        // Vypočítat rabat potřebný pro dosažení cílové ceny
                        return ((prodejniCenaBezDPH - cilovaCena) / prodejniCenaBezDPH) * 100;
                    }
                    // Pokud cílová cena je vyšší než prodejní, použijeme 0% rabat
                    return 0;

                default:
                    // Výchozí výpočet: (prodejní - nákupní) / prodejní (bez DPH) %
                    if (prodejniCenaBezDPH > 0)
                    {
                        return ((prodejniCenaBezDPH - nakupniCena) / prodejniCenaBezDPH) * 100;
                    }
                    return 0;
            }
        }

        // Zápis výsledků do Excel souboru
        static void WriteExcelFile(string filePath, List<ResultRecord> data)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Výsledek");

                if (data.Count == 0)
                {
                    worksheet.Cells[1, 1].Value = "Žádná data k zobrazení";
                    package.SaveAs(new FileInfo(filePath));
                    return;
                }

                // Zápis hlaviček
                string[] headers = {
                    "Kód", "Název", "EAN", "Sklad", "Nákupka", 
                    "Nová prodejní cena s DPH", "Nová prodejní cena bez DPH", 
                    "Rabat", "Pozice na heurece", "Odkaz na heureku"
                };

                for (int col = 0; col < headers.Length; col++)
                {
                    worksheet.Cells[1, col + 1].Value = headers[col];
                    worksheet.Cells[1, col + 1].Style.Font.Bold = true;
                }

                // Zápis dat
                for (int row = 0; row < data.Count; row++)
                {
                    var record = data[row];
                    int col = 1;

                    worksheet.Cells[row + 2, col++].Value = record.Kod;
                    worksheet.Cells[row + 2, col++].Value = record.Nazev;
                    worksheet.Cells[row + 2, col++].Value = record.EAN;
                    worksheet.Cells[row + 2, col++].Value = record.Sklad;
                    worksheet.Cells[row + 2, col++].Value = record.Nakupka;
                    worksheet.Cells[row + 2, col++].Value = record.NovaProdejniCenaSDPH;
                    worksheet.Cells[row + 2, col++].Value = record.NovaProdejniCenaBezDPH;
                    worksheet.Cells[row + 2, col++].Value = record.Rabat;
                    worksheet.Cells[row + 2, col++].Value = record.PoziceNaHeurece;
                    worksheet.Cells[row + 2, col++].Value = record.OdkazNaHeureku;
                }

                // Automatické přizpůsobení šířky sloupců
                worksheet.Cells.AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }

            Console.WriteLine($"  Zapsáno {data.Count} řádků dat");
        }
    }

    // Datové třídy pro záznamy
    class HeurekaRecord
    {
        public string EAN { get; set; } = string.Empty; // Sloupec F - pro párování
        public double ProdejniCena { get; set; } // Sloupec L - s DPH
        public string HeurekaOdkaz { get; set; } = string.Empty; // Sloupec I
        public List<double?> NejnizsiCeny { get; set; } = new List<double?>(); // Sloupce P-AH - vzestupně seřazené nejnižší ceny
        public int RowNumber { get; set; }
    }

    class SkladRecord
    {
        public string Kod { get; set; } = string.Empty; // Sloupec A
        public string EAN { get; set; } = string.Empty; // Sloupec B
        public string Nazev { get; set; } = string.Empty; // Sloupec C
        public double Sklad { get; set; } // Sloupec F
        public double NakupniCena { get; set; } // Sloupec G - bez DPH
        public string Rule { get; set; } = string.Empty; // Sloupec J
        public int RowNumber { get; set; }
    }

    class ResultRecord
    {
        public string Kod { get; set; } = string.Empty;
        public string Nazev { get; set; } = string.Empty;
        public string EAN { get; set; } = string.Empty;
        public double Sklad { get; set; }
        public double Nakupka { get; set; }
        public double NovaProdejniCenaSDPH { get; set; }
        public double NovaProdejniCenaBezDPH { get; set; }
        public double Rabat { get; set; }
        public int PoziceNaHeurece { get; set; }
        public string OdkazNaHeureku { get; set; } = string.Empty;
    }
}
