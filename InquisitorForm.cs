using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DriverLicenseCheck
{
    public partial class InquisitorForm : Form
    {
        // таймаут получения данных
        private const int timeout = 90000;
        public InquisitorForm()
        {
            InitializeComponent();
            this.Load += async (s, e) => await GetQuotesUsed();
        }

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
                (int code, string message) res = await CreateReportTestDEPRECATED(driverLicense.Text, issuedDate.Text, checkBoxUpdateInfo.Checked);

                if (res.code == 200)
                {
                    // Создаем новый поток на получение времени
                    Thread timerThread = new Thread(TimerFunction);
                    timerThread.Start();

                    // Получить данные из ГИБДД
                    await AfterSendRequestDataFromGIBDDButton();
                }
                else
                {
                    MessageBox.Show("Код " + res.code + ". " + res.message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            await GetQuotesUsed();
        }

        // Функция проверить файл по введенным данным
        private async Task AfterSendRequestDataFromGIBDDButton()
        {
            await Task.Delay(timeout); // ждем указанный таймаут
            // Создать отчет
            (int code, string message) response = await CreateReportTestDEPRECATED(driverLicense.Text, issuedDate.Text, false);

            // Если создание отчета прошло успешно
            if (response.code == 200)
            {
                // Получить информацию из отчета
                Dictionary<string, string> resp_report = await GetInformationTEST(response.message);

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

        //Кнопка "Проверить файл"
        private async void ButtonCheckFile(object sender, EventArgs e)
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

                    textBoxOutputInfo.AppendText("Ожидание получения данных составляет " + (Convert.ToDouble(timeout) / 60000) + " минуты. Пожалуйста, не препятствуйте работе программы. По окончанию получения данных вы получите уведомление." + Environment.NewLine + Environment.NewLine);

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
                            object licence_plate_value = worksheet.Cells[i + 1, 11].Value;
                            object issue_date_licence_plate_value = worksheet.Cells[i + 1, 12].Value;

                            if (licence_plate_value != null && issue_date_licence_plate_value != null)
                            {
                                string issue_date_licence_plate = worksheet.Cells[i + 1, 12].ToText().Replace("/", ".");
                                textBoxOutputInfo.AppendText("Запрос данных по ВУ " + licence_plate_value.ToString() + ", от " + issue_date_licence_plate + Environment.NewLine);

                                ////////////////
                                // СОЗДАТЬ ОТЧЕТ

                                await CreateReportTestDEPRECATED(licence_plate_value.ToString(), issue_date_licence_plate, checkBoxUpdateInfo.Checked);

                                /* // проверка прошло ли 14 дней
                                var response = await CreateReportTest(licence_plate_value.ToString(), issue_date_licence_plate, checkBoxUpdateInfo.Checked);

                                if (!checkBoxUpdateInfo.Checked && forcedVerificationCheckBox.Checked)
                                {
                                    try
                                    {
                                        // получить дату последнего обновления
                                        var json = await response.Content.ReadAsStringAsync();
                                        var date = JsonDocument.Parse(json).RootElement.GetProperty("data")[0].GetProperty("suggest_get").ToString();

                                        // если с последней даты обновления прошло N дней
                                        if ((DateTime.UtcNow - DateTime.Parse(date).ToUniversalTime()).Days > int.Parse(forcedVerificationDate.Text))
                                            // отправить принудительно на проверку
                                            await CreateReportTest(licence_plate_value.ToString(), issue_date_licence_plate, true);
                                    }
                                    catch (Exception) { }
                                }*/

                            }
                        }
                        // Создаем новый поток на получение времени                       
                        Thread timerThread = new Thread(TimerFunction);
                        timerThread.Start();

                        await StartAfterDelay();
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
            // Узнать сколько квот израсходовалось
            await GetQuotesUsed();
        }

        private async Task StartAfterDelay()
        {
            // Добавляем отступ в блоке вывода
            textBoxOutputInfo.Invoke(new Action(() => textBoxOutputInfo.AppendText(Environment.NewLine)));

            // ждем указанный таймаут
            await Task.Delay(timeout);

            // открываем файл Excel
            string fileName = textBoxPath.Text;
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

                // Добавить заголовки в файл
                for (int i = 0; i < headers.Length; i++)
                    worksheet.Cells[1, colCount + i + 1].Value = headers[i];

                // Список для пакетной записи
                List<object[]> allData = new List<object[]>();

                for (int i = 1; i < rowCount; i++)
                {
                    // Получить номер ВУ и дату выдачи
                    object licence_plate_value = worksheet.Cells[i + 1, 11].Value;
                    object issue_date_licence_plate_value = worksheet.Cells[i + 1, 12].Value;

                    if (licence_plate_value != null && issue_date_licence_plate_value != null)
                    {
                        string issue_date_licence_plate = worksheet.Cells[i + 1, 12].ToText().Replace("/", ".");
                        textBoxOutputInfo.Invoke(new Action(() => textBoxOutputInfo.AppendText("Данные получены по ВУ " + licence_plate_value + ", от " + issue_date_licence_plate + Environment.NewLine)));

                        // Создать отчет
                        //var response = await CreateReportTest(licence_plate_value.ToString(), issue_date_licence_plate, false); //new
                        (int code, string message) response = await CreateReportTestDEPRECATED(licence_plate_value.ToString(), issue_date_licence_plate, false); //old
                        if (response.code == 200)
                        {
                            // Получить данные из отчета
                            Dictionary<string, string> responceFromGIBDD = await GetInformationTEST(response.message);

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
                }
                /*if (response.StatusCode == 200) // new
                {
                    var jsonResponse = await response.Content.ReadAsStringAsync();
                    string message = JsonDocument.Parse(jsonResponse).RootElement.GetProperty("data")[0].GetProperty("uid").GetString();
                    // Получить данные из отчета
                    Dictionary<string, string> responceFromGIBDD = await GetInformationTEST(message);
                    /////
                    // Формируем массив значений в нужном порядке
                    object[] rowValues = new object[]
                    {
                        responceFromGIBDD["SeriesAndNumber"],
                        responceFromGIBDD["Birthday"],
                        responceFromGIBDD["IssuedDate"],
                        responceFromGIBDD["EndDate"],
                        responceFromGIBDD["Category"],
                        responceFromGIBDD["Comment"],
                        responceFromGIBDD["CategoryCE"],
                        responceFromGIBDD["stateDescription1"],
                        responceFromGIBDD["comment1"],
                        responceFromGIBDD["limitation1"],
                        responceFromGIBDD["date1"],
                        responceFromGIBDD["stateDescription2"],
                        responceFromGIBDD["comment2"],
                        responceFromGIBDD["limitation2"],
                        responceFromGIBDD["date2"]
                    };
                    allData.Add(rowValues);
                }
                else
                {
                    string errorMessage = "";
                    try
                    {
                        var responseError = await response.Content.ReadAsStringAsync();
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

                    // Если ошибка – заполняем строку сообщением об ошибке
                    object[] errorRow = new object[headers.Length];
                    for (int k = 0; k < headers.Length; k++) errorRow[k] = errorMessage;
                    allData.Add(errorRow);
                }
            }
        }

        // Пакетная запись
        if (allData.Count > 0)
            worksheet.Cells[2, colCount + 1].LoadFromArrays(allData);
        */
                // Сохраняем изменения в файл
                package.Save();

                // Красим строки в красный цвет если есть нарушение
                ChangeRowColorBySearchData(fileName);

                MessageBox.Show("Файл обновлен", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // Покрасить в цвет
        private static void ChangeRowColorBySearchData(string excelFilePath)
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
                        string cellValue = worksheet.Cells[row, col].Text;

                        // Проверяем, содержится ли значение из массива в ячейке
                        if (searchData.Any(data => cellValue.IndexOf(data, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            // Если да, окрашиваем строку в красный цвет
                            using (var range = worksheet.Cells[row, 1, row, colCount])
                            {
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                worksheet.Cells[row, col].Style.Font.Color.SetColor(System.Drawing.Color.White);
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
        private void ButtonGetPath(object sender, EventArgs e)
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

        private async Task GetQuotesUsed()
        {
            string url = $"https://b2b-api.spectrumdata.ru/b2b/api/v1/user/balance/{enviroment}";

            // Установка заголовков
            _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 99999));

            // отправка запроса
            var response = await _httpClient.GetAsync(url);

            string json = await response.Content.ReadAsStringAsync();
            var resdata = JsonDocument.Parse(json).RootElement.GetProperty("data");
            string day = resdata[0].GetProperty("quote_use").ToString();
            string mounth = resdata[1].GetProperty("quote_use").ToString();
            string alltime = resdata[2].GetProperty("quote_use").ToString();

            label15.Text = $"Использовано {mounth} квот за месяц, {day} за сегодня, {alltime} за все время.";
        }

        private void DriverLicenseKeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Regex.IsMatch((sender as TextBox).Text + e.KeyChar, @"^[0-9]+$") && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void InquisitorForm_Load(object sender, EventArgs e)
        {

        }
    }
}