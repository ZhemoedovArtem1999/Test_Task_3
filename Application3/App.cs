using Application3.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Application3.Controller;
using Application3.Repository.ExcelRepository;
using Application3.ModelView;
using Application3.Models;
using Application3.Service;
using System.Runtime.CompilerServices;

namespace Application3
{
    public class App
    {
        public static void Run()
        {
            InputOutputConsole inputOutputConsole = new InputOutputConsole();
            OutputException exceptionOutputConsole = new OutputException();
            string path = string.Empty;
            Controller.Controller controller = App.CreateControllerInFile(inputOutputConsole, exceptionOutputConsole);
           
            while (true) {
                string task = inputOutputConsole.Input("Выберите действие из доступных:\n" +
                    "1 - По наименованию товара выводить информацию о клиентах, заказавших этот товар, с указанием информации по количеству товара, цене и дате заказа.\n" +
                    "2 - Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица. В результате информация должна быть занесена в этот же документ, в качестве ответа пользователю необходимо выдавать информацию о результате изменений.\n" +
                    "3 - Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.\n" +
                    "4 - Изменить файл для работы\n" +
                    "5 - Выйти");

                switch (task)
                {
                    case "1":
                        string productTitle = inputOutputConsole.Input("Введите наименование товара");
                        ClientProductView clientProductView = controller.ListClients(productTitle);
                        if (clientProductView.Status.State == 200)
                        {
                            foreach (ClientProduct clientProduct in clientProductView.ClientProducts)
                            {
                                inputOutputConsole.Output(clientProduct.ToString());
                            }
                        }
                        else
                        {
                            exceptionOutputConsole.Output(clientProductView.Status.Message);
                        }

                        break;
                    case "2":
                        string organizationTitle = inputOutputConsole.Input("Введите название организации, для которой изменяем контактное лицо");

                        string newContactPerson = inputOutputConsole.Input("Введите ФИО нового контактного лица");

                        UpdateClientView updateClientView = controller.EditClient(organizationTitle, newContactPerson);

                        if (updateClientView.Status.State == 200)
                        {
                            inputOutputConsole.Output(updateClientView.Status.Message);
                        }
                        else
                        {
                            exceptionOutputConsole.Output(updateClientView.Status.Message);
                        }

                        break;
                    case "3":
                        int year;
                        int month;
                        while (true)
                        {
                            bool isYear = int.TryParse(inputOutputConsole.Input("Введите год"), out year);
                            bool isMonth = int.TryParse(inputOutputConsole.Input("Введите месяц"), out month);

                            if (isYear && isMonth && year > 0 && month > 0 && month <= 12)
                            {
                                break;

                            }
                            exceptionOutputConsole.Output("Данные введены неверно. Попробуйте снова.");
                        }

                        GoldenClientView goldenClientView = controller.GetGoldenClientDate(year, month);
                        if (goldenClientView.Status.State == 200)
                        {
                            inputOutputConsole.Output(goldenClientView.Status.Message);
                            inputOutputConsole.Output(goldenClientView.Client.ToString());
                        }
                        else if (goldenClientView.Status.State == 201)
                        {
                            inputOutputConsole.Output(goldenClientView.Status.Message);
                        }
                        else {
                            exceptionOutputConsole.Output(goldenClientView.Status.Message);
                        }


                        break;

                    case "4":
                        controller = App.CreateControllerInFile(inputOutputConsole, exceptionOutputConsole);
                        break;

                    case "5":
                        return;
                        


                    default:
                        exceptionOutputConsole.Output("Действия не существует");
                        break;
                }
            }
        }


        private static Controller.Controller CreateControllerInFile(InputOutputConsole inputOutputConsole, OutputException outputException)
        {
            Controller.Controller controller = null;
            while (true)
            {
                string path = inputOutputConsole.Input("Введите абсолютный путь к файлу");
                if (path == null || path == "")
                {
                    outputException.Output("Путь к файлу пустой");
                    continue;
                }

                controller = new Controller.Controller(new ExcelProductRepository(path), new ExcelClientRepository(path), new ExcelOrderRepository(path), new ValidateExcel(path, new Dictionary<string, List<string>>
                {
                    ["Товары"] = new List<string>() { "Код товара", "Наименование", "Ед. измерения", "Цена товара за единицу" },
                    ["Клиенты"] = new List<string>() { "Код клиента", "Наименование организации", "Адрес", "Контактное лицо (ФИО)" },
                    ["Заявки"] = new List<string>() { "Код заявки", "Код товара", "Код клиента", "Номер заявки", "Требуемое количество", "Дата размещения" },
                }));

                var valid = controller.ValidateFile();
                if (valid == null)
                    break;
                else
                {
                    foreach (var item in valid.Errors)
                    {
                        outputException.Output(item);
                    }
                }
            }
            return controller;
        }

    }
}
