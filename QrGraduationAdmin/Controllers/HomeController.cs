using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using QrGraduationAdmin.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Threading.Tasks;
using QRCoder; // Добавляем это пространство имен
using System.Drawing;

using static System.Net.Mime.MediaTypeNames;
using Microsoft.AspNetCore.Hosting;


namespace QrGraduationAdmin.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _webHostEnvironment;



        public HomeController(ILogger<HomeController> logger,IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;

        }
        public async Task<IActionResult> Index()
        {
            List<Employee> userList = new List<Employee>();
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Employees"))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    userList = JsonConvert.DeserializeObject<List<Employee>>(apiResponse);
                }
            }
            return View(userList);
        }
        public ViewResult GetEmployees() => View();

        [HttpPost]
        public async Task<IActionResult> GetEmployees(int id)
        {
            Employee employee = new Employee();
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Employees/" + id))
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        string apiResponse = await response.Content.ReadAsStringAsync();
                        employee = JsonConvert.DeserializeObject<Employee>(apiResponse);
                    }
                    else
                        ViewBag.StatusCode = response.StatusCode;
                }
            }
            return View(employee);
        }

        public ViewResult AddEmployee() => View();

        [HttpPost]
        public async Task<IActionResult> AddEmployee(Employee employee)
        {
            Employee recivedEmployee = new Employee();
            using (var httpClient = new HttpClient())
            {
                StringContent content = new StringContent(JsonConvert.SerializeObject(employee), Encoding.UTF8, "application/json");

                using (var response = await httpClient.PostAsync("https://localhost:7062/api/Employees/", content))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    recivedEmployee = JsonConvert.DeserializeObject<Employee>(apiResponse);
                }
            }
            return View(recivedEmployee);
        }

        public async Task<IActionResult> UpdateEmployee(int id)
        {
            Employee employee = new Employee();
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Employees/" + id))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    employee = JsonConvert.DeserializeObject<Employee>(apiResponse);
                }
            }
            return View(employee);
        }

        [HttpPost]
        public async Task<IActionResult> UpdateEmployee(Employee employee, int id)
        {
            Employee receivedEmployee = new Employee();
            using (var httpClient = new HttpClient())
            {
                StringContent content = new StringContent(JsonConvert.SerializeObject(employee), Encoding.UTF8, "application/json");

                using (var response = await httpClient.PutAsync("https://localhost:7062/api/Employees/" + id, content))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    ViewBag.Result = "Success";
                    receivedEmployee = JsonConvert.DeserializeObject<Employee>(apiResponse);
                }
            }
            return View(receivedEmployee);
        }

        [HttpPost]
        public async Task<IActionResult> DeleteEmployees(int IdEmployee)
        {
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.DeleteAsync("https://localhost:7062/api/Employees/" + IdEmployee))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                }
            }

            return RedirectToAction("Index");
        }


        public async Task<IActionResult> EmployeeHistory(int id)
        {
            List<History> historyList = new List<History>();
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync($"https://localhost:7062/api/Histories/Employee/{id}"))
                {
                    if (response.StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        string apiResponse = await response.Content.ReadAsStringAsync();
                        historyList = JsonConvert.DeserializeObject<List<History>>(apiResponse);
                    }
                    else
                    {
                        ViewBag.StatusCode = response.StatusCode;
                    }
                }
            }
            return View(historyList);
        }
        [HttpPost]
        public async Task<IActionResult> SearchHistoriesBySecondName(string secondName, string startDate, string endDate)
        {
            List<History> historyList = new List<History>();

            using (var httpClient = new HttpClient())
            {
                // Формируем URI запроса с параметрами фильтрации
                string requestUri = $"https://localhost:7062/api/Histories/BySecondName/{secondName}?startDate={startDate}&endDate={endDate}";

                // Отправляем запрос на сервер
                using (var response = await httpClient.GetAsync(requestUri))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        // Если запрос успешен, получаем отфильтрованные данные
                        string apiResponse = await response.Content.ReadAsStringAsync();
                        historyList = JsonConvert.DeserializeObject<List<History>>(apiResponse);
                    }
                    else
                    {
                        // Если произошла ошибка, устанавливаем сообщение об ошибке
                        ViewBag.ErrorMessage = "Истории не найдены или произошла ошибка.";
                    }
                }
            }

            return View(historyList);
        }


        public IActionResult ExportToExcel(List<History> historyList)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("History");

                // Заголовки столбцов
                worksheet.Cells[1, 1].Value = "Дата начала";
                worksheet.Cells[1, 2].Value = "Дата окончания";

                // Данные
                for (int i = 0; i < historyList.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = historyList[i].DateStartHistory;
                    worksheet.Cells[i + 2, 2].Value = historyList[i].DateFinishHistory;
                }

                // Сохранение файла
                var stream = new MemoryStream();
                package.SaveAs(stream);

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "History.xlsx");
            }
        }


        public IActionResult ExportToWord()
        {
            var historyList = GetHistoryData(); // Получаем данные истории, замените на ваш метод получения данных

            // Используйте класс Application из пространства имен Microsoft.Office.Interop.Word
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Add();

            var para = doc.Paragraphs.Add();
            para.Range.Text = "History";

            para.Range.InsertParagraphAfter();

            foreach (var history in historyList)
            {
                para.Range.Text = $"Пришел: {history.DateStartHistory}, Ушел: {history.DateFinishHistory}";
                para.Range.InsertParagraphAfter();
            }

            var stream = new MemoryStream();
            doc.SaveAs2(stream);
            doc.Close();
            wordApp.Quit();

            return File(stream.ToArray(), "application/msword", "History.docx");
        }
        private List<History> GetHistoryData()
        {
            // Замените этот метод на ваш метод получения данных истории
            return new List<History>(); // Пример возврата пустого списка, замените на вашу логику
        }

        public IActionResult SearchHistory()
        {
            return View("SearchHistoriesBySecondName");
        }

        public async Task<IActionResult> GetQRCode()
        {
            string qrText = ""; // Здесь сохраните полученный текст QR-кода

            // Получение текста QR-кода из API
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Qrs/1"))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        var qrRecord = JsonConvert.DeserializeObject<Qr>(await response.Content.ReadAsStringAsync());
                        qrText = qrRecord.TextQr;
                    }
                    else
                    {
                        // Обработка ошибки
                        return StatusCode((int)response.StatusCode, "Не удалось получить текст QR-кода из API");
                    }
                }
            }

            // Создание QR-кода на основе полученного текста
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20); // Размер QR-кода

            // Сохранение QR-кода или его передача в представление
            // Пример передачи в представление:
            using (MemoryStream ms = new MemoryStream())
            {
                qrCodeImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return File(ms.ToArray(), "image/png"); // Возвращаем изображение в формате PNG
            }
        }
        public IActionResult ShowQRCode()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ChangeQRText(string newText)
        {
            // Получение первой записи о QR-коде
            Qr qrRecord;
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Qrs/1"))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        qrRecord = JsonConvert.DeserializeObject<Qr>(await response.Content.ReadAsStringAsync());
                    }
                    else
                    {
                        // Обработка ошибки
                        return StatusCode((int)response.StatusCode, "Не удалось получить запись о QR-коде из API");
                    }
                }
            }

            // Изменение значения TextQr
            qrRecord.TextQr = newText;

            // Обновление записи в базе данных
            using (var httpClient = new HttpClient())
            {
                StringContent content = new StringContent(JsonConvert.SerializeObject(qrRecord), Encoding.UTF8, "application/json");

                using (var response = await httpClient.PutAsync("https://localhost:7062/api/Qrs/1", content))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        // В случае успеха перенаправляем пользователя на страницу с QR-кодом
                        return RedirectToAction("ShowQRCode");
                    }
                    else
                    {
                        // Обработка ошибки при обновлении записи
                        return StatusCode((int)response.StatusCode, "Не удалось обновить запись о QR-коде в API");
                    }
                }
            }
        }
        public async Task<IActionResult> SaveQRCode()
        {
            // Создаем относительный URL-адрес для вызова метода GetQRCode
            string qrCodeUrl = Url.Action("GetQRCode", "Home", null, Request.Scheme);

            // Загружаем изображение QR-кода
            using (var httpClient = new HttpClient())
            {
                var response = await httpClient.GetAsync(qrCodeUrl);
                if (response.IsSuccessStatusCode)
                {
                    // Считываем содержимое ответа в байтовый массив
                    byte[] qrCodeBytes = await response.Content.ReadAsByteArrayAsync();

                    // Генерируем уникальное имя для файла
                    var fileName = Guid.NewGuid().ToString() + ".png";

                    // Сохраняем файл на сервере
                    var filePath = Path.Combine(_webHostEnvironment.WebRootPath, "QRImages", fileName);
                    System.IO.File.WriteAllBytes(filePath, qrCodeBytes);

                    // Возвращаем путь к сохраненному файлу
                    ViewBag.QRCodeImagePath = "/QRImages/" + fileName;
                    return View();
                }
                else
                {
                    // Обработка ошибки при получении изображения QR-кода
                    return StatusCode((int)response.StatusCode, "Не удалось получить изображение QR-кода для сохранения");
                }
            }
        }




        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
