using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Newtonsoft.Json;

namespace ExcelToApiUploader
{
    class Program
    {
        static async Task Main(string[] args)
        {
            string excelPath = "asprator.xlsx"; 
            string apiUrl = "https://your-api-url.com/endpoint";


            var columnMapping = new Dictionary<string, string>
            {
                { "SERİ NO", "code" },
                { "ÜRÜN", "model" },
                { "MARKA", "errorCode" },
                { "MODEL", "description" }
            };

            try
            {
                var dataList = ReadExcelFile(excelPath, columnMapping);

                using var httpClient = new HttpClient();

                foreach (var item in dataList)
                {
                    string json = JsonConvert.SerializeObject(item, Formatting.Indented);
                    Console.WriteLine(json); 

                    //await PostToApi(httpClient, apiUrl, json);
                }

                Console.WriteLine("Tüm veriler başarıyla gönderildi.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Hata oluştu: {ex.Message}");
            }
        }

        public static List<Dictionary<string, string>> ReadExcelFile(string path, Dictionary<string, string> columnMapping)
        {
            var result = new List<Dictionary<string, string>>();

            using var workbook = new XLWorkbook(path);

            foreach (var worksheet in workbook.Worksheets)
            {
                var range = worksheet.RangeUsed();
                if (range == null) continue;

                var rows = range.RowsUsed();
                var headers = new List<string>();

                foreach (var cell in rows.First().Cells())
                {
                    headers.Add(cell.GetString());
                }

                foreach (var row in rows.Skip(1))
                {
                    var rowData = new Dictionary<string, string>();

                    for (int i = 0; i < headers.Count; i++)
                    {
                        string excelHeader = headers[i];
                        if (columnMapping.ContainsKey(excelHeader))
                        {
                            string apiField = columnMapping[excelHeader];
                            string cellValue = row.Cell(i + 1).GetString();
                            rowData[apiField] = cellValue;
                        }
                    }

                    // Eğer rowData boş değilse (eşleşen sütun varsa) listeye ekle
                    if (rowData.Count > 0)
                    {
                        result.Add(rowData);
                    }
                }
            }

            return result;
        }


        public static async Task PostToApi(HttpClient client, string url, string jsonData)
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
    }
}
