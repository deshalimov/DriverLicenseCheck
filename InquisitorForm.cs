using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DriverLicenseCheck
{
    public partial class InquisitorForm : Form
    {
        public InquisitorForm()
        {
            InitializeComponent();
        }

        // таймаут получения данных
        private const int timeout = 90000;

        // Кнопка проверить файл по введенным данным
        private async void SendRequestDataFromGIBDDButton(object sender, EventArgs e)
        {
            if (driverLicense.Text.Length == 0)
                ErrorDriverLicense.Text = "Введите номер ВУ";
            else
                ErrorDriverLicense.Text = "";

            if (issuedDate.Text.Length == 0)
                ErrorReceivingDate.Text = "Введите дату выдачи ВУ";
            else
                ErrorReceivingDate.Text = "";

            if (driverLicense.Text.Length > 0 && issuedDate.Text.Length > 0)
            {
                // очистить поля вывода предыдущей информации
                ClearAllOutput();

                // Создать отчет
                (int code, string message) res = createReportTest(driverLicense.Text, issuedDate.Text, false);

                if (res.code == 200)
                {
                    // Создаем новый поток на получение времени
                    Thread timerThread = new Thread(TimerFunction);
                    timerThread.Start();

                    // Получить данные из ГИБДД
                    AfterSendRequestDataFromGIBDDButton();
                }
                else
                {
                    MessageBox.Show("Код " + res.code + ". " + res.message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Функция проверить файл по введенным данным
        private async void AfterSendRequestDataFromGIBDDButton()
        {
            await Task.Delay(timeout); // ждем указанный таймаут
            // Создать отчет
            (int code, string message) response = createReportTest(driverLicense.Text, issuedDate.Text, false);

            // Если создание отчета прошло успешно
            if (response.code == 200)
            {
                // Получить информацию из отчета
                Dictionary<string, string> resp_report = getInformationTEST(response.message);

                // Вывести информацию из отчета на экран
                textBoxSeriesAndNumber.Invoke(new Action(() => textBoxSeriesAndNumber.Text = resp_report["SeriesAndNumber"]));
                textBoxBirthday.Invoke(new Action(() => textBoxBirthday.Text = resp_report["Birthday"]));
                textBoxReceivingDate.Invoke(new Action(() => textBoxReceivingDate.Text = resp_report["IssuedDate"]));
                textBoxEndDate.Invoke(new Action(() => textBoxEndDate.Text = resp_report["EndDate"]));
                textBoxCategories.Invoke(new Action(() => textBoxCategories.Text = resp_report["Category"]));

                textBoxGibddDataFoundComment.Invoke(new Action(() => textBoxGibddDataFoundComment.Text = resp_report["Comment"]));
                if (resp_report["Comment"].Contains("лишение") || resp_report["Comment"].Contains("не действителен"))
                {
                    textBoxGibddDataFoundComment.ForeColor = Color.Red;
                }
                else
                {
                    textBoxGibddDataFoundComment.ForeColor = Color.Black;
                }


                if (textBoxCategories.Text.Contains("CE") || textBoxCategories.Text.Contains("СЕ"))
                {
                    isActiveCategories.Invoke(new Action(() => isActiveCategories.Text = "Открыта категория СE"));
                    isActiveCategories.ForeColor = Color.Green;
                }
                else
                {
                    isActiveCategories.Invoke(new Action(() => isActiveCategories.Text = "Нет открытой категории СE"));
                    isActiveCategories.ForeColor = Color.Red;
                }

                stateDescription1.Invoke(new Action(() => stateDescription1.Text = resp_report["stateDescription1"]));
                comment1.Invoke(new Action(() => comment1.Text = resp_report["comment1"]));
                limitation1.Invoke(new Action(() => limitation1.Text = resp_report["limitation1"]));
                date1.Invoke(new Action(() => date1.Text = resp_report["date1"]));

                stateDescription2.Invoke(new Action(() => stateDescription2.Text = resp_report["stateDescription2"]));
                comment2.Invoke(new Action(() => comment2.Text = resp_report["comment2"]));
                limitation2.Invoke(new Action(() => limitation2.Text = resp_report["limitation2"]));
                date2.Invoke(new Action(() => date2.Text = resp_report["date2"]));
            }
            else
            {
                MessageBox.Show("Ошибка. Код ошибки: " + response.code + ". " + response.message, "Ошибка создания отчета", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Функция отправки отчета в Spectrum
        public static (int, string) createReportTest(string driver_license, string driver_license_date, bool force)
        {
            string url = "https://b2b-api.spectrumdata.ru/b2b/api/v1/user/reports/report_dl_main@smartseeds/_make";
            HttpClient client = new HttpClient();

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
            client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 999999999));

            // отправка запроса
            var response = client.PostAsync(url, content);

            if (response.Result.StatusCode.ToString() == "OK")
            {
                var responseBody = response.Result.Content.ReadAsStringAsync().Result;
                // Десериализация JSON-ответа
                var jsonResponse = JsonDocument.Parse(responseBody);

                // Получаем значение uid из data
                var uid = jsonResponse.RootElement
                    .GetProperty("data")[0]
                    .GetProperty("uid")
                    .GetString();

                return ((int)response.Result.StatusCode, uid);
            }
            else
            {
                string errorMessage;
                try
                {
                    var responseError = response.Result.Content.ReadAsStringAsync().Result;
                    var jsonResponseError = JsonDocument.Parse(responseError);

                    errorMessage = jsonResponseError.RootElement
                        .GetProperty("event")
                        .GetProperty("name")
                        .GetString();
                }
                catch (Exception)
                {
                    errorMessage = "Ошибка получения данных";
                }
                return ((int)response.Result.StatusCode, errorMessage);
            }
        }

        // Функция получения отчета из Spectrum
        public Dictionary<string, string> getInformationTEST(string uid)
        {
            // Создать словарь данных
            Dictionary<string, string> data_from_GIBDD = new Dictionary<string, string>()
            {
                {"SeriesAndNumber", ""},
                {"Birthday", ""},
                {"IssuedDate", ""},
                {"EndDate", ""},
                {"Category", ""},
                {"Comment", ""},
                {"CategoryCE", ""},
                {"stateDescription1", ""},
                {"comment1", ""},
                {"limitation1", ""},
                {"date1", ""},
                {"stateDescription2", ""},
                {"comment2", ""},
                {"limitation2", ""},
                {"date2", ""}
            };

            // Сгенерить URL адрес
            string url = $"https://b2b-api.spectrumdata.ru/b2b/api/v1/user/reports/{uid}?_content=true&_detailed=false";

            // Создать HTTP клиента
            HttpClient client = new HttpClient();

            // Установка заголовков
            client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 999999999));

            // Отправляем GET запрос
            var response = client.GetAsync(url);

            if (response.Result.StatusCode.ToString() == "OK")
            {
                var responseBody = response.Result.Content.ReadAsStringAsync().Result;

                // Десериализация JSON-ответа
                var jsonResponse = JsonDocument.Parse(responseBody);

                // Если отчет по ВУ нового образца cp_driver_license_v2
                try
                {
                    var data_driver = jsonResponse.RootElement
                        .GetProperty("data")[0]
                        .GetProperty("content")
                        .GetProperty("check_person")
                        .GetProperty("cp_driver_license_v2")
                        .GetProperty("driverLicense");

                    // Серия и номер ВУ
                    data_from_GIBDD["SeriesAndNumber"] = data_driver.GetProperty("series").ToString()
                        + data_driver.GetProperty("number").ToString();

                    // Дата рождения водителя
                    data_from_GIBDD["Birthday"] = FormatDate(data_driver.GetProperty("birthday").ToString());

                    // Дата выдачи ВУ
                    data_from_GIBDD["IssuedDate"] = FormatDate(data_driver.GetProperty("receivingDate").ToString());

                    // Дата окончания ВУ
                    data_from_GIBDD["EndDate"] = FormatDate(data_driver.GetProperty("endDate").ToString());

                    // Открытые категории ВУ
                    string categoriesJSON = data_driver.GetProperty("categories").ToString();
                    string[] categories = JsonSerializer.Deserialize<string[]>(categoriesJSON);
                    string res = "";
                    foreach (var item in categories)
                    {
                        res += item + " ";
                    }
                    data_from_GIBDD["Category"] = res;

                    // Наличие категории CE
                    if (categories.Contains("CE") || categories.Contains("СЕ"))
                    {
                        data_from_GIBDD["CategoryCE"] = "Категория CE открыта";
                    }
                    else
                    {
                        data_from_GIBDD["CategoryCE"] = "Категория CE не открыта";
                    }

                    // Комментарий ГИБДД
                    if (data_driver.GetProperty("gibddDataFoundComment").ToString() == "")
                    {
                        data_from_GIBDD["Comment"] = data_driver.GetProperty("deprivationOfManagementRightsComment").ToString();
                    }
                    else
                    {
                        data_from_GIBDD["Comment"] = data_driver.GetProperty("gibddDataFoundComment").ToString();
                    }

                    // Получить данные по лишениям
                    try
                    {
                        var decisions = data_driver.GetProperty("decisions");

                        if (decisions.GetArrayLength() != 0)
                        {
                            try
                            {
                                data_from_GIBDD["stateDescription1"] = decisions[0].GetProperty("stateDescription").ToString();
                                data_from_GIBDD["comment1"] = decisions[0].GetProperty("comment").ToString();
                                data_from_GIBDD["limitation1"] = decisions[0].GetProperty("limitation").ToString();
                                data_from_GIBDD["date1"] = FormatDate(decisions[0].GetProperty("date").ToString());
                            }
                            catch { }

                            try
                            {
                                data_from_GIBDD["stateDescription2"] = decisions[1].GetProperty("stateDescription").ToString();
                                data_from_GIBDD["comment2"] = decisions[1].GetProperty("comment").ToString();
                                data_from_GIBDD["limitation2"] = decisions[1].GetProperty("limitation").ToString();
                                data_from_GIBDD["date2"] = FormatDate(decisions[1].GetProperty("date").ToString());
                            }
                            catch { }
                        }
                        else
                        {
                            data_from_GIBDD["stateDescription1"] = "Нет данных";
                            data_from_GIBDD["comment1"] = "Нет данных";
                            data_from_GIBDD["limitation1"] = "Нет данных";
                            data_from_GIBDD["date1"] = "Нет данных";
                            data_from_GIBDD["stateDescription2"] = "Нет данных";
                            data_from_GIBDD["comment2"] = "Нет данных";
                            data_from_GIBDD["limitation2"] = "Нет данных";
                            data_from_GIBDD["date2"] = "Нет данных";
                        }
                    }
                    catch { }

                    return data_from_GIBDD;
                }
                // Если при получения отчета нового образца возникли ошибки
                catch
                {
                    // Если отчет по ВУ старого образца
                    try
                    {
                        var data_driver = jsonResponse.RootElement
                            .GetProperty("data")[0]
                            .GetProperty("content")
                            .GetProperty("check_person")
                            .GetProperty("piter_driver_license")
                            .GetProperty("category_info")
                            .GetProperty("category_info_list")
                            .ToString().Replace('\'', '\"');

                        var categoriesArray = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(data_driver);

                        // Извлечение значений "category"
                        List<string> categoryValues = new List<string>();
                        foreach (var item in categoriesArray)
                        {
                            // Проверяем, есть ли ключ "category" и добавляем его значение
                            if (item.TryGetValue("category", out var categoryValue))
                            {
                                categoryValues.Add(categoryValue.ToString());
                            }
                        }

                        // Преобразование в массив string[]
                        string[] categoryNames = categoryValues.ToArray();

                        string res_cat = "";
                        // Вывод значений категорий
                        foreach (var category in categoryNames)
                        {
                            res_cat += category + " ";
                        }
                        // если 
                        if (res_cat == "")
                        {
                            var status_spec = jsonResponse.RootElement
                            .GetProperty("data")[0]
                            .GetProperty("content")
                            .GetProperty("check_person")
                            .GetProperty("piter_driver_license")
                            .GetProperty("status")
                            .ToString();
                            data_from_GIBDD["Category"] = status_spec;
                        }
                        else
                        {
                            data_from_GIBDD["Category"] = res_cat;
                        }

                        // Наличие категории CE
                        if (categoryNames.Contains("CE") || categoryNames.Contains("СЕ"))
                        {
                            data_from_GIBDD["CategoryCE"] = "Категория CE открыта";
                        }
                        else
                        {
                            data_from_GIBDD["CategoryCE"] = "Категория CE не открыта";
                        }


                        var data_driver_lp = jsonResponse.RootElement
                            .GetProperty("data")[0]
                            .GetProperty("query")
                            .GetProperty("data");

                        // Серия и номер ВУ
                        data_from_GIBDD["SeriesAndNumber"] = data_driver_lp.GetProperty("driver_license").ToString();

                        // Дата выдачи ВУ
                        data_from_GIBDD["IssuedDate"] = data_driver_lp.GetProperty("driver_license_date").ToString();

                        data_from_GIBDD["Birthday"] = "Нет данных";
                        data_from_GIBDD["EndDate"] = "Нет данных";
                        data_from_GIBDD["Comment"] = "Нет данных";

                        // Получить данные по лишениям
                        var decisions = jsonResponse.RootElement
                        .GetProperty("data")[0]
                        .GetProperty("content")
                        .GetProperty("check_person")
                        .GetProperty("cp_driver_license_v2")
                        .GetProperty("driverLicense")
                        .GetProperty("decisions");

                        try
                        {
                            if (decisions.GetArrayLength() != 0)
                            {
                                try
                                {
                                    data_from_GIBDD["stateDescription1"] = decisions[0].GetProperty("stateDescription").ToString();
                                    data_from_GIBDD["comment1"] = decisions[0].GetProperty("comment").ToString();
                                    data_from_GIBDD["limitation1"] = decisions[0].GetProperty("limitation").ToString();
                                    data_from_GIBDD["date1"] = FormatDate(decisions[0].GetProperty("date").ToString());
                                }
                                catch { }

                                try
                                {
                                    data_from_GIBDD["stateDescription2"] = decisions[1].GetProperty("stateDescription").ToString();
                                    data_from_GIBDD["comment2"] = decisions[1].GetProperty("comment").ToString();
                                    data_from_GIBDD["limitation2"] = decisions[1].GetProperty("limitation").ToString();
                                    data_from_GIBDD["date2"] = FormatDate(decisions[1].GetProperty("date").ToString());
                                }
                                catch { }
                            }
                            else
                            {
                                data_from_GIBDD["stateDescription1"] = "Нет данных";
                                data_from_GIBDD["comment1"] = "Нет данных";
                                data_from_GIBDD["limitation1"] = "Нет данных";
                                data_from_GIBDD["date1"] = "Нет данных";
                                data_from_GIBDD["stateDescription2"] = "Нет данных";
                                data_from_GIBDD["comment2"] = "Нет данных";
                                data_from_GIBDD["limitation2"] = "Нет данных";
                                data_from_GIBDD["date2"] = "Нет данных";
                            }
                        }
                        catch
                        {
                            data_from_GIBDD["stateDescription1"] = "Данные не получены";
                            data_from_GIBDD["comment1"] = "Данные не получены";
                            data_from_GIBDD["limitation1"] = "Данные не получены";
                            data_from_GIBDD["date1"] = "Данные не получены";
                            data_from_GIBDD["stateDescription2"] = "Данные не получены";
                            data_from_GIBDD["comment2"] = "Данные не получены";
                            data_from_GIBDD["limitation2"] = "Данные не получены";
                            data_from_GIBDD["date2"] = "Данные не получены";
                        }

                        return data_from_GIBDD;
                    }
                    // Если по отчету старого образца возникла ошибка
                    catch
                    {
                        data_from_GIBDD["Category"] = "Данные не получены";
                        data_from_GIBDD["SeriesAndNumber"] = "Данные не получены";
                        data_from_GIBDD["IssuedDate"] = "Данные не получены";
                        data_from_GIBDD["Birthday"] = "Данные не получены";
                        data_from_GIBDD["EndDate"] = "Данные не получены";
                        data_from_GIBDD["Comment"] = "Данные не получены";
                        data_from_GIBDD["CategoryCE"] = "Данные не получены";

                        data_from_GIBDD["stateDescription1"] = "Данные не получены";
                        data_from_GIBDD["comment1"] = "Данные не получены";
                        data_from_GIBDD["limitation1"] = "Данные не получены";
                        data_from_GIBDD["date1"] = "Данные не получены";
                        data_from_GIBDD["stateDescription2"] = "Данные не получены";
                        data_from_GIBDD["comment2"] = "Данные не получены";
                        data_from_GIBDD["limitation2"] = "Данные не получены";
                        data_from_GIBDD["date2"] = "Данные не получены";

                        return data_from_GIBDD;
                    }
                }
            }
            else
            {
                data_from_GIBDD["Category"] = "Данные не получены";
                data_from_GIBDD["SeriesAndNumber"] = "Данные не получены";
                data_from_GIBDD["IssuedDate"] = "Данные не получены";
                data_from_GIBDD["Birthday"] = "Данные не получены";
                data_from_GIBDD["EndDate"] = "Данные не получены";
                data_from_GIBDD["Comment"] = "Данные не получены";
                data_from_GIBDD["CategoryCE"] = "Данные не получены";

                return data_from_GIBDD;
            }
        }

        //Кнопка "Проверить файл"
        private async void buttonCheckFile(object sender, EventArgs e)
        {
            textBoxOutputInfo.Text = "";
            labelErrorPath.Text = "";
            string fileName = textBoxPath.Text;
            if (fileName != "")
            {
                labelErrorPath.Text = "";

                // прочитать данные из файла
                try
                {
                    string[] file = File.ReadAllLines(fileName, Encoding.UTF8);

                    textBoxOutputInfo.Text += "Ожидание получения данных составляет " + (Convert.ToDouble(timeout) / 60000) + " минуты. Пожалуйста, не препятствуйте работе программы. По окончанию получения данных вы получите уведомление." + Environment.NewLine + Environment.NewLine;

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    // Открываем файл Excel
                    FileInfo existingFile = new FileInfo(fileName);

                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        // Получаем первый рабочий лист
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                        // Определяем количество строк и столбцов
                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        for (int i = 1; i < rowCount; i++)
                        {
                            string licence_plate = worksheet.Cells[i + 1, 11].Text.Replace(" ", "");
                            string issue_date_licence_plate = worksheet.Cells[i + 1, 12].Text;

                            textBoxOutputInfo.Text += "Получение данных по ВУ " + licence_plate + Environment.NewLine;

                            // Создать отчет
                            createReportTest(licence_plate, issue_date_licence_plate, false);
                        }
                        // Создаем новый поток на получение времени
                        Thread timerThread = new Thread(TimerFunction);
                        timerThread.Start();

                        StartAfterDelay();
                    }
                }
                catch (Exception x)
                {
                    MessageBox.Show(x.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                labelErrorPath.Text = "Выберите файл";
            }
        }

        async void StartAfterDelay()
        {
            // Добавляем отступ в блоке вывода
            textBoxOutputInfo.Invoke(new Action(() => textBoxOutputInfo.Text += Environment.NewLine));

            // ждем указанный таймаут
            await Task.Delay(timeout);

            // открываем файл Excel
            string fileName = textBoxPath.Text;
            string[] data = File.ReadAllLines(fileName, Encoding.UTF8);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Открываем файл Excel
            FileInfo existingFile = new FileInfo(fileName);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // Получаем первый рабочий лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Новые заголовки
                string[] headers = {
                "Серия и номер ВУ",
                "Дата рождения водителя",
                "Дата выдачи ВУ",
                "Дата окончания действия ВУ",
                "Категории ВУ",
                "Комментарий ГИБДД",
                "Наличие категорий CE",
                "Информация о лишении",
                "Комментарий о лишении",
                "Срок лишения права управления (мес)",
                "Дата постановления лишения",
                "Информация о лишении",
                "Комментарий о лишении",
                "Срок лишения права управления (мес)",
                "Дата постановления лишения"
                };

                // Определяем количество строк и столбцов
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Добавить заголовки в файл
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, colCount + i + 1].Value = headers[i];
                }

                for (int i = 1; i < rowCount; i++)
                {
                    // Получить номер ВУ и дату выдачи
                    string licence_plate = worksheet.Cells[i + 1, 11].Text;
                    string issue_date_licence_plate = worksheet.Cells[i + 1, 12].Text;

                    textBoxOutputInfo.Invoke(new Action(() => textBoxOutputInfo.Text += "Данные получены по ВУ " + licence_plate + Environment.NewLine));

                    // Создать отчет
                    (int code, string message) response = createReportTest(licence_plate, issue_date_licence_plate, false);
                    if (response.code == 200)
                    {
                        // Получить данные из отчета
                        Dictionary<string, string> responceFromGIBDD = getInformationTEST(response.message);

                        for (int j = 0; j < responceFromGIBDD.Count; j++)
                        {
                            worksheet.Cells[i + 1, j + 1 + colCount].Value = responceFromGIBDD.ElementAt(j).Value;
                        }
                    }
                    else
                    {
                        for (int k = 0; k < headers.Length; k++)
                        {
                            worksheet.Cells[i + 1, k + 1 + colCount].Value = response.message;
                        }
                    }
                }

                // Сохраняем изменения в файл
                package.Save();

                // Красим строки в красный цвет если есть нарушение
                ChangeRowColorBySearchData(fileName);

                MessageBox.Show("Файл обновлен", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        // Покрасить в цвет
        static void ChangeRowColorBySearchData(string excelFilePath)
        {
            string[] searchData = { "лишение", "не действителен", "Категория CE не открыта" };
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Открываем файл Excel
            FileInfo existingFile = new FileInfo(excelFilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                // Получаем первый рабочий лист
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                // Получаем количество строк и столбцов
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Перебираем все строки в диапазоне
                for (int row = 1; row <= rowCount; row++)
                {
                    // Перебираем все столбцы в текущей строке
                    for (int col = 1; col <= colCount; col++)
                    {
                        // Получаем значение ячейки
                        var cellValue = worksheet.Cells[row, col].Text;

                        // Проверяем, содержится ли значение из массива в ячейке
                        if (searchData.Any(data => cellValue.IndexOf(data, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            // Если да, окрашиваем строку в красный цвет
                            using (var range = worksheet.Cells[row, 1, row, colCount])
                            {
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                            }

                            // Выходим из внутреннего цикла, поскольку строки уже окрашены
                            break;
                        }
                    }
                }

                // Сохраняем изменения в тот же файл
                package.Save();
            }
        }

        // Функция форматирования даты
        private string FormatDate(string dateString)
        {
            DateTime date = DateTime.Parse(dateString);
            string formattedDate = date.ToString("dd.MM.yyyy");
            return formattedDate;
        }

        // Таймер оставшегося времени
        private void TimerFunction()
        {
            int totalTime = timeout / 1000;
            int remainingTime = totalTime;

            while (remainingTime > 0)
            {
                // Вычисляем минуты и секунды
                int minutes = remainingTime / 60;
                int seconds = remainingTime % 60;

                if (remainingTimeBox.InvokeRequired)
                    remainingTimeBox.Invoke(new Action(() => remainingTimeBox.Text = $"{minutes:D2}:{seconds:D2}"));
                Thread.Sleep(1000); // задержка на 1 секунду
                remainingTime -= 1; // уменьшаем оставшееся время на 1 секунду
            }
            remainingTimeBox.Invoke(new Action(() => remainingTimeBox.Text = "00:00"));
        }

        // Кнопка очистки введеных полей
        private void ClearAllButton(object sender, EventArgs e)
        {
            // Очистить пользовательские поля ввода
            driverLicense.Text = "";
            issuedDate.Text = "";
            ClearAllOutput();
        }

        // Функция очистки введеных полей
        private void ClearAllOutput()
        {
            // Очистить поля ошибок
            ErrorDriverLicense.Text = "";
            ErrorReceivingDate.Text = "";

            textBoxOutputInfo.Text = "";

            // Очистить поля вывода данных из ГИБДД
            textBoxSeriesAndNumber.Text = "";
            textBoxBirthday.Text = "";
            textBoxReceivingDate.Text = "";
            textBoxEndDate.Text = "";
            textBoxCategories.Text = "";
            textBoxGibddDataFoundComment.Text = "";

            isActiveCategories.Text = "";

            stateDescription1.Text = "";
            comment1.Text = "";
            limitation1.Text = "";
            date1.Text = "";

            stateDescription2.Text = "";
            comment2.Text = "";
            limitation2.Text = "";
            date2.Text = "";
        }

        // Кнопка выбрать файл
        private void buttonGetPath(object sender, EventArgs e)
        {
            // открыть диалоговое окно для получения расположения файла
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx"; //Все файлы (*.*)|*.*|
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                // если в диалоговом выбран файл
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Получить расположение файла
                    string fileName = openFileDialog.FileName;
                    textBoxPath.Text = fileName;
                }
            }
        }

        // Функция генерация токена
        public static string GenerateToken(int age = 60 * 60 * 24)
        {
            string user = "";
            string password = "";
            var timestamp = (int)DateTimeOffset.UtcNow.ToUnixTimeSeconds();

            string passwordHash = Convert.ToBase64String(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(password)));

            string hashWithSalt = $"{timestamp}:{age}:{passwordHash}";
            string saltedHashB64 = Convert.ToBase64String(MD5.Create().ComputeHash(Encoding.UTF8.GetBytes(hashWithSalt)));

            string token = $"{user}:{timestamp}:{age}:{saltedHashB64}";
            string tokenB64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(token));

            return tokenB64;
        }

        private void driverLicenseKeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Regex.IsMatch((sender as TextBox).Text + e.KeyChar, @"^[0-9]+$") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}