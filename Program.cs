using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel; 
using Newtonsoft.Json;
using System.IO; 

namespace ExcelToApiUploader
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string configExcelPath = "konfig.xlsx";
            string spartsFilesDirectory = ""; 
            string aspiratorExcelPath = "asprator1.xlsx";
            string apiUrl = "https://your-api-url.com/endpoint";

            bool saveJsonOutputToFile = true; 
            string outputDirectory = "JsonOutputs"; 

            try
            {
                Console.WriteLine($"\n--- '{configExcelPath}' okunuyor ---");
                var configDataList = ReadExcelFile(configExcelPath, new Dictionary<string, string>
                {
                    { "A", "ProductCode" },
                    { "B", "Name" },
                    { "E", "KonfigEslesmeKodu" }
                });
                Console.WriteLine($"'{configExcelPath}' dosyasından toplam {configDataList.Count} satır okundu.");

                Console.WriteLine($"\n--- '{aspiratorExcelPath}' okunuyor ---");
                var aspiratorDataList = ReadExcelFile(aspiratorExcelPath, new Dictionary<string, string>
                {
                    { "SERİ NO", "AspiratorAColumn" }, 
                    { "MODEL", "AspiratorEslesmeKodu" } 
                });
                Console.WriteLine($"'{aspiratorExcelPath}' dosyasından toplam {aspiratorDataList.Count} satır okundu.");
      
                using var httpClient = new HttpClient();

                int productCount = 0;
                foreach (var configItem in configDataList)
                {
                    productCount++;
                    Console.WriteLine($"\n--- Ürün {productCount} işleniyor ---");

                    var mainProductObject = new Dictionary<string, object>();
                    
                    mainProductObject["ProductCode"] = configItem.GetValueOrDefault("ProductCode", ""); 
                    mainProductObject["FirmID"] = "13"; 
                    mainProductObject["Name"] = configItem.GetValueOrDefault("Name", ""); 
                    mainProductObject["ShortCode"] = "B01"; 
                    mainProductObject["ProductTypeName"] = "ASPIRATOR"; 
                    mainProductObject["ProductModelName"] = configItem.GetValueOrDefault("Name", ""); 

                    var spartsProductsList = new List<Dictionary<string, string>>();
                    string spartsProductCode = configItem.GetValueOrDefault("ProductCode", "");

                    if (!string.IsNullOrEmpty(spartsProductCode))
                    {
                        string spartsExcelFileName = $"{spartsProductCode}.xlsx";
                        string spartsExcelPath = System.IO.Path.Combine(spartsFilesDirectory, spartsExcelFileName);

                        try
                        {
                            Console.WriteLine($"\n--- '{spartsExcelFileName}' okunuyor ---");
                            var spartsData = ReadExcelFile(spartsExcelPath, new Dictionary<string, string>
                            {
                                { "A", "ManufactureCode" },  
                                { "F", "ProductTypeName" },       
                                { "D", "ProductCode" },    
                                { "G", "ShortCode" },    
                                { "E", "Name" }      
                            });
                            Console.WriteLine($"'{spartsExcelFileName}' dosyasından toplam {spartsData.Count} satır okundu.");

                            foreach (var spartsItem in spartsData)
                            {
                                var currentSpartsProduct = new Dictionary<string, string>
                                {
                                    { "ManufactureCode", spartsItem.GetValueOrDefault("ManufactureCode", "") },
                                    { "ProductTypeName", spartsItem.GetValueOrDefault("ProductTypeName", "ASPIRATOR") }, 
                                    { "ProductCode", spartsItem.GetValueOrDefault("ProductCode", "") },
                                    { "ShortCode", spartsItem.GetValueOrDefault("ShortCode", "B01") }, 
                                    { "Name", spartsItem.GetValueOrDefault("Name", "") }
                                };


                                spartsProductsList.Add(currentSpartsProduct);
                            }
                        }
                        catch (System.IO.FileNotFoundException)
                        {
                            Console.WriteLine($"Uyarı: '{spartsExcelFileName}' dosyası bulunamadı. Bu ürün için SpartsProducts eklenmeyecek.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"HATA: '{spartsExcelFileName}' okunurken genel bir hata oluştu: {ex.Message}");
                        }
                    }
                    mainProductObject["SpartsProducts"] = spartsProductsList;

                    var productSerialsList = new List<Dictionary<string, string>>();
                    string configEslesmeKodu = configItem.GetValueOrDefault("KonfigEslesmeKodu", "");

                    if (!string.IsNullOrEmpty(configEslesmeKodu))
                    {
                        var matchedAspiratorSerials = aspiratorDataList
                            .Where(a => a.ContainsKey("AspiratorEslesmeKodu") && a.GetValueOrDefault("AspiratorEslesmeKodu", "") == configEslesmeKodu)
                            .ToList();

                        foreach (var serialItem in matchedAspiratorSerials)
                        {
                            var currentProductSerial = new Dictionary<string, string>
                            {
                                { "CustomFollowCode", serialItem.GetValueOrDefault("AspiratorAColumn", "") },
                                { "Name", serialItem.GetValueOrDefault("AspiratorAColumn", "") }
                            };

                            // currentProductSerial["Status"] = "Aktif"; 
                            // currentProductSerial["Location"] = "Depo A"; 

                            productSerialsList.Add(currentProductSerial);
                        }
                    }
                    mainProductObject["ProductSerials"] = productSerialsList;

                    string json = JsonConvert.SerializeObject(mainProductObject, Formatting.Indented);
                    Console.WriteLine("\n--- API'ye gönderilecek JSON (Kısa Özet) ---");
                    Console.WriteLine(json.Length > 500 ? json.Substring(0, 500) + "...\n(Tam çıktı için dosyayı kontrol edin)" : json);
                    Console.WriteLine("-------------------------------\n");

                    if (saveJsonOutputToFile)
                    {
                        string fileName = $"{aspiratorExcelPath}_{productCount}_{configItem.GetValueOrDefault("ProductCode", "Unknown")}.json";
                        SaveJsonToFile(json, fileName, outputDirectory);
                    }
                }

                Console.WriteLine("Tüm veriler başarıyla hazırlandı.");
                if (saveJsonOutputToFile)
                {
                    Console.WriteLine($"Tüm JSON çıktıları '{Path.GetFullPath(outputDirectory)}' klasörüne kaydedildi.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Genel hata oluştu: {ex.Message}");
            }
        }

        public static List<Dictionary<string, string>> ReadExcelFile(string path, Dictionary<string, string> columnMapping)
        {
            var result = new List<Dictionary<string, string>>();

            if (!System.IO.File.Exists(path))
            {
                throw new System.IO.FileNotFoundException($"'{path}' dosyası bulunamadı.");
            }

            using var workbook = new XLWorkbook(path);
            
            foreach (var worksheet in workbook.Worksheets) 
            {
                Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] Çalışma Sayfası '{worksheet.Name}' işleniyor.");
                var range = worksheet.RangeUsed(); 
                if (range == null || range.RowCount() == 0 || range.ColumnCount() == 0)
                {
                    Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] Sayfa '{worksheet.Name}' boş veya kullanılmış hücre içermiyor. Atlanıyor.");
                    continue; 
                }

                var rows = range.RowsUsed(); 
                var firstRowCells = rows.First().CellsUsed();
                var headers = new List<string>();
                foreach (var cell in firstRowCells)
                {
                    headers.Add(cell.GetString().Trim()); 
                }

                bool useColumnIndexMapping = headers.All(string.IsNullOrWhiteSpace) || !firstRowCells.Any(); 

                Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] Çalışma Sayfası '{worksheet.Name}': useColumnIndexMapping = {useColumnIndexMapping}");
                if (!useColumnIndexMapping) {
                     Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] Algılanan Başlıklar: {string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)))}");
                } else {
                     Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] İlk satırda başlık algılanmadı (tümü boş veya boşluk). Sütun indeksleri kullanılacak.");
                }

                IEnumerable<IXLRangeRow> dataRowsToProcess; 
                if (useColumnIndexMapping)
                {
                    dataRowsToProcess = rows;
                    Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] İlk satırdan itibaren veri olarak okunuyor.");
                }
                else
                {
                    dataRowsToProcess = rows.Skip(1); 
                    Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] İlk satır (başlıklar) atlanıyor.");
                }

                foreach (var row in dataRowsToProcess) 
                {
                    try 
                    {
                        var rowData = new Dictionary<string, string>(); 
                        
                        Console.WriteLine($"    [{System.IO.Path.GetFileName(path)}] Sayfa '{worksheet.Name}' - Satır {row.RowNumber()} işleniyor.");
                        
                        for (int colIndex = 1; colIndex <= range.LastColumnUsed().ColumnNumber(); colIndex++)
                        {
                            var cell = row.Cell(colIndex); 
                            if (cell == null || cell.IsEmpty()) continue; 

                            int cellColumnIndex = colIndex - 1; 
                            string cellValue = cell.GetString();

                            string keyToMap = null; 
                            string debugColumnIdentifier = "";

                            if (useColumnIndexMapping)
                            {
                                debugColumnIdentifier = ((char)('A' + cellColumnIndex)).ToString();
                                if (columnMapping.ContainsKey(debugColumnIdentifier))
                                {
                                    keyToMap = columnMapping[debugColumnIdentifier];
                                }
                            }
                            else
                            {
                                if (cellColumnIndex < headers.Count) 
                                {
                                    debugColumnIdentifier = headers[cellColumnIndex];
                                    if (columnMapping.ContainsKey(debugColumnIdentifier))
                                    {
                                        keyToMap = columnMapping[debugColumnIdentifier];
                                    }
                                }
                            }

                            Console.WriteLine($"      Hücre: {cell.Address.ColumnLetter}{cell.Address.RowNumber} (Değer: '{cellValue}', ColIndex: {cellColumnIndex}, Tanımlayıcı: '{debugColumnIdentifier}', Eşleşen Anahtar: '{keyToMap ?? "BULUNAMADI"}')");

                            if (keyToMap != null)
                            {
                                if (rowData.ContainsKey(keyToMap))
                                {
                                    Console.WriteLine($"        UYARI: '{System.IO.Path.GetFileName(path)}' - Sayfa '{worksheet.Name}' - Satır {row.RowNumber()}: '{keyToMap}' anahtarı zaten mevcut. Eski değer: '{rowData[keyToMap]}', Yeni değer: '{cellValue}'. Eski değerin üzerine yazılıyor.");
                                    rowData[keyToMap] = cellValue; 
                                }
                                else
                                {
                                    rowData.Add(keyToMap, cellValue); 
                                }
                                Console.WriteLine($"        '{cell.Address.ColumnLetter}{cell.Address.RowNumber}' ('{cellValue}') -> '{keyToMap}' olarak eşlendi.");
                            } else {
                                Console.WriteLine($"        '{cell.Address.ColumnLetter}{cell.Address.RowNumber}' ('{cellValue}') -> Eşleşen anahtar bulunamadı (mapping yok).");
                            }
                        }

                        if (rowData.Count > 0)
                        {
                            result.Add(rowData);
                            Console.WriteLine($"    [{System.IO.Path.GetFileName(path)}] Sayfa '{worksheet.Name}' - Satır {row.RowNumber()} eklendi. Toplam eşleşen öğe: {rowData.Count}.");
                        } else {
                            Console.WriteLine($"    [{System.IO.Path.GetFileName(path)}] Sayfa '{worksheet.Name}' - Satır {row.RowNumber()} atlandı (eşleşen öğe yok).");
                        }
                    }
                    catch (Exception rowEx)
                    {
                        Console.WriteLine($"    HATA: '{System.IO.Path.GetFileName(path)}' - Sayfa '{worksheet.Name}' - Satır {row.RowNumber()} işlenirken hata oluştu: {rowEx.Message}");
                    }
                }
                Console.WriteLine($"  [{System.IO.Path.GetFileName(path)}] Çalışma Sayfası '{worksheet.Name}' tamamlandı. Bu sayfadan eklenen satır sayısı: {result.Count - (result.Where(d => d.ContainsKey("WorksheetName") && d["WorksheetName"] == worksheet.Name).Count())}.");
            }
            Console.WriteLine($"[{System.IO.Path.GetFileName(path)}] Tüm sayfalar işlendi. Toplam okunan satır: {result.Count}");
            return result; 
        }

        public static async Task PostToApi(HttpClient client, string url, string jsonData)
        {
            try
            {
                var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
                var response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine("Veri başarıyla gönderildi.");
                }
                else
                {
                    Console.WriteLine($"Gönderim başarısız. Status: {response.StatusCode}");
                    string errorContent = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Hata mesajı: {errorContent}");
                }
            }
            catch (HttpRequestException httpEx)
            {
                Console.WriteLine($"API'ye bağlanırken hata oluştu: {httpEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Post işlemi sırasında beklenmedik hata: {ex.Message}");
            }
        }

        public static void SaveJsonToFile(string jsonData, string fileName, string directory)
        {
            try
            {
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                string fullPath = Path.Combine(directory, fileName);
                File.WriteAllText(fullPath, jsonData);
                Console.WriteLine($"JSON çıktısı başarıyla kaydedildi: {Path.GetFullPath(fullPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"JSON çıktısı dosyaya kaydedilirken hata oluştu: {ex.Message}");
            }
        }
    }

    public static class DictionaryExtensions
    {
        public static TValue GetValueOrDefault<TKey, TValue>(this IDictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
        {
            if (dictionary.TryGetValue(key, out TValue value))
            {
                return value;
            }
            return defaultValue;
        }
    }
}
