using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DriverLicenseCheck
{
    public partial class InquisitorForm : Form
    {
        // Функция получения отчета из Spectrum
        public async Task<Dictionary<string, string>> GetInformationTEST(string uid)
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

            // Установка заголовков
            _httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("AR-REST", GenerateToken(age: 999999999));

            // Отправляем GET запрос
            HttpResponseMessage response = await _httpClient.GetAsync(url);

            if (response.StatusCode.ToString() == "OK")
            {
                string responseBody = await response.Content.ReadAsStringAsync();

                // Десериализация JSON-ответа
                var jsonResponse = JsonDocument.Parse(responseBody);

                // Если отчет по ВУ нового образца cp_driver_license_v2
                try
                {
                    JsonElement data_driver = jsonResponse.RootElement
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
                    foreach (string item in categories)
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
                    data_from_GIBDD["Comment"] = (data_driver.GetProperty("gibddDataFoundComment").ToString() == "")
                        ? data_driver.GetProperty("deprivationOfManagementRightsComment").ToString()
                        : data_driver.GetProperty("gibddDataFoundComment").ToString();

                    // Получить данные по лишениям
                    try
                    {
                        JsonElement decisions = data_driver.GetProperty("decisions");

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
                        string data_driver = jsonResponse.RootElement
                            .GetProperty("data")[0]
                            .GetProperty("content")
                            .GetProperty("check_person")
                            .GetProperty("piter_driver_license")
                            .GetProperty("category_info")
                            .GetProperty("category_info_list")
                            .ToString().Replace('\'', '\"');

                        List<Dictionary<string, object>> categoriesArray = JsonSerializer.Deserialize<List<Dictionary<string, object>>>(data_driver);

                        // Извлечение значений "category"
                        List<string> categoryValues = new List<string>();
                        foreach (Dictionary<string, object> item in categoriesArray)
                        {
                            // Проверяем, есть ли ключ "category" и добавляем его значение
                            if (item.TryGetValue("category", out object categoryValue))
                            {
                                categoryValues.Add(categoryValue.ToString());
                            }
                        }

                        // Преобразование в массив string[]
                        string[] categoryNames = categoryValues.ToArray();

                        string res_cat = "";
                        // Вывод значений категорий
                        foreach (string category in categoryNames)
                        {
                            res_cat += category + " ";
                        }
                        // если 
                        if (res_cat == "")
                        {
                            string status_spec = jsonResponse.RootElement
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
                        data_from_GIBDD["CategoryCE"] = (categoryNames.Contains("CE") || categoryNames.Contains("СЕ")) ? "Категория CE открыта" : "Категория CE не открыта";

                        JsonElement data_driver_lp = jsonResponse.RootElement
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
                        JsonElement decisions = jsonResponse.RootElement
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
                        try
                        {
                            JsonElement data_driver_lp = jsonResponse.RootElement
                               .GetProperty("data")[0]
                               .GetProperty("query")
                               .GetProperty("data");

                            // Серия и номер ВУ
                            data_from_GIBDD["SeriesAndNumber"] = data_driver_lp.GetProperty("driver_license").ToString();

                            // Дата выдачи ВУ
                            data_from_GIBDD["IssuedDate"] = data_driver_lp.GetProperty("driver_license_date").ToString();
                        }
                        catch
                        {
                            data_from_GIBDD["SeriesAndNumber"] = "Данные не получены";
                            data_from_GIBDD["IssuedDate"] = "Данные не получены";
                        }
                        data_from_GIBDD["Category"] = "Данные не получены";
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

        // Функция форматирования даты
        private string FormatDate(string dateString)
        {
            DateTime date = DateTime.Parse(dateString);
            string formattedDate = date.ToString("dd.MM.yyyy");
            return formattedDate;
        }

    }
}