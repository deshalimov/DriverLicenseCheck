using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DriverLicenseCheck
{
    public partial class InquisitorForm : Form
    {
        private static readonly JsonElement secrets = JsonDocument.Parse(
            File.ReadAllText("secrets.json")
            ).RootElement;

        private static readonly string _login =
            secrets.GetProperty("_login").GetString();

        private static readonly string _password =
            secrets.GetProperty("_password").GetString();

        private static readonly string enviroment =
            secrets.GetProperty("enviroment").GetString();




        private static readonly HttpClient _httpClient = new HttpClient();
        // Функция отправки отчета в Spectrum
        public static async Task<(int, string)> CreateReportTestDEPRECATED(string driver_license, string driver_license_date, bool force)
        {
            string url = $"https://b2b-api.spectrumdata.ru/b2b/api/v1/user/reports/{enviroment}/_make";

            // Создаем объект запроса
            var requestData = new
            {
                queryType = "MULTIPART",
                query = " ",
                data = new
                {
                    driver_license,
                    driver_license_date
                },
                options = new
                {
                    FORCE = force
                }
            };

            // Сериализация объекта запроса в JSON
            string jsonRequest = JsonSerializer.Serialize(requestData);
            StringContent content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

            // Установка заголовков
            _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 999999999));

            // отправка запроса
            var response = await _httpClient.PostAsync(url, content);

            if (response.StatusCode.ToString() == "OK")
            {
                string responseBody = await response.Content.ReadAsStringAsync();
                // Десериализация JSON-ответа
                JsonDocument jsonResponse = JsonDocument.Parse(responseBody);

                // Получаем значение uid из data
                string uid = jsonResponse.RootElement
                    .GetProperty("data")[0]
                    .GetProperty("uid")
                    .GetString();

                return ((int)response.StatusCode, uid);
            }
            else
            {
                string errorMessage;
                try
                {
                    string responseError = await response.Content.ReadAsStringAsync();
                    JsonDocument jsonResponseError = JsonDocument.Parse(responseError);

                    errorMessage = jsonResponseError.RootElement
                        .GetProperty("event")
                        .GetProperty("name")
                        .GetString();
                }
                catch (Exception)
                {
                    errorMessage = "Ошибка получения данных";
                }
                return ((int)response.StatusCode, errorMessage);
            }
        }

        /*
        // Функция отправки отчета в Spectrum
        public static async Task<HttpResponseMessage> CreateReportTest(string driver_license, string driver_license_date, bool force)
        {
            string url = $"https://b2b-api.spectrumdata.ru/b2b/api/v1/user/reports/{enviroment}/_make";

            // Создаем объект запроса
            var requestData = new
            {
                queryType = "MULTIPART",
                query = " ",
                data = new
                {
                    driver_license,
                    driver_license_date
                },
                options = new
                {
                    FORCE = force
                }
            };

            // Сериализация объекта запроса в JSON
            var jsonRequest = JsonSerializer.Serialize(requestData);
            var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

            // Установка заголовков
            _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 999999999));

            // отправка запроса
            return await _httpClient.PostAsync(url, content);
        }*/


        // Функция генерация токена
        public static string GenerateToken(int age = 60 * 60 * 24)
        {
            string user = _login;
            string password = _password;
            int timestamp = (int)DateTimeOffset.UtcNow.ToUnixTimeSeconds();

            string passwordHash = Convert.ToBase64String(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(password)));

            string hashWithSalt = $"{timestamp}:{age}:{passwordHash}";
            string saltedHashB64 = Convert.ToBase64String(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(hashWithSalt)));

            string token = $"{user}:{timestamp}:{age}:{saltedHashB64}";
            string tokenB64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(token));

            return tokenB64;
        }

    }
}