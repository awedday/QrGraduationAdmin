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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.InkML;


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
        [HttpPost]
        public async Task<IActionResult> Login(string login, string password)
        {
            // Здесь необходимо реализовать аутентификацию пользователя
            // Примерный код для аутентификации пользователя
            var user = await GetUserByLoginAndPassword(login, password);

            if (user == null)
            {
                return NotFound("Пользователь не найден");
            }

            // Проверяем, является ли пользователь администратором по логину
            if (user.MailEmployee == "benzenkoAP@itimportant.ru" && user.PasswordEmployee == password)
            {
                HttpContext.Session.SetString("UserId", user.IdEmployee.ToString());

                // В данном месте вы можете создать cookie или токен для аутентифицированного пользователя
                return RedirectToAction("Index");
            }
            else
            {
                return Unauthorized("Доступ запрещен. Только администраторы могут войти.");
            }
        }
        public IActionResult Logout()
        {
            // Удалите данные из сессии при выходе
            HttpContext.Session.Remove("UserId");

            return RedirectToAction("Login");
        }

        private async Task<Employee> GetUserByLoginAndPassword(string login, string password)
        {
            Employee user = null;
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync($"https://localhost:7062/api/Employees/login?login={login}&password={password}"))
                {
                    if (response.IsSuccessStatusCode)
                    {
                        string apiResponse = await response.Content.ReadAsStringAsync();
                        user = JsonConvert.DeserializeObject<Employee>(apiResponse);
                    }
                }
            }
            return user;
        }

    

        // Для разметки представления авторизации вы можете использовать форму входа:
        public IActionResult Login()
        {
            return View();
        }
    

    [HttpPost]
        public async Task<IActionResult> SearchHistoriesBySecondName(string secondName)
        {
            List<History> historyList = new List<History>();

            using (var httpClient = new HttpClient())
            {
                // Формируем URI запроса с параметрами фильтрации
                string requestUri = $"https://localhost:7062/api/Histories/BySecondName/{secondName}";

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

        public async Task<List<Employee>> AllUsersForExcel()
        {
            List<Employee> userListAll = new List<Employee>();
            using (var httpClient = new HttpClient())
            {
                using (var response = await httpClient.GetAsync("https://localhost:7062/api/Employees"))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    userListAll = JsonConvert.DeserializeObject<List<Employee>>(apiResponse);
                }
            }
            return userListAll;
        }

        public async Task<IActionResult> ExportEmployeesToExcel()
        {
            List<Employee> employees = new List<Employee>();

            // Здесь получите список всех работников из базы данных или другого источника
            // Например:
            employees = await AllUsersForExcel();

            // Создаем пакет Excel
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Employees");

                // Устанавливаем заголовки столбцов
                worksheet.Cells[1, 1].Value = "IdEmployee";
                worksheet.Cells[1, 2].Value = "FirstNameEmployee";
                worksheet.Cells[1, 3].Value = "SecondNameEmployee";
                worksheet.Cells[1, 4].Value = "MiddleNameEmployee";
                worksheet.Cells[1, 5].Value = "MailEmployee";
                worksheet.Cells[1, 6].Value = "PhoneEmployee";

                // Заполняем данные
                int row = 2;
                foreach (var employee in employees)
                {
                    worksheet.Cells[row, 1].Value = employee.IdEmployee;
                    worksheet.Cells[row, 2].Value = employee.FirstNameEmployee;
                    worksheet.Cells[row, 3].Value = employee.SecondNameEmployee;
                    worksheet.Cells[row, 4].Value = employee.MiddleNameEmployee;
                    worksheet.Cells[row, 5].Value = employee.MailEmployee;
                    worksheet.Cells[row, 6].Value = employee.PhoneEmployee;
                    row++;
                }

                // Сохраняем файл
                var stream = new MemoryStream();
                package.SaveAs(stream);

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Employees.xlsx");
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
