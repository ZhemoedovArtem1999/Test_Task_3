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

namespace Application3
{
    public class App
    {
        public static void Run()
        {
            InputOutputConsole inputOutputConsole = new InputOutputConsole();
            OutputException exceptionOutputConsole = new OutputException();
            string path = inputOutputConsole.Input("Введите абсолютный путь к файлу");
            path = "D:\\Практическое задание для кандидата.xlsx";
            if (path == null || path == "")
            {
                // exception кинуть            
            }
            Controller.Controller controller = new Controller.Controller(new ExcelProductRepository(path), new ExcelClientRepository(path), new ExcelOrderRepository(path));
            while (true) {
                string task = inputOutputConsole.Input("Выберите действие из доступных:\n" +
                    "1 - По наименованию товара выводить информацию о клиентах, заказавших этот товар, с указанием информации по количеству товара, цене и дате заказа.\n" +
                    "2 - Запрос на изменение контактного лица клиента с указанием параметров: Название организации, ФИО нового контактного лица. В результате информация должна быть занесена в этот же документ, в качестве ответа пользователю необходимо выдавать информацию о результате изменений.\n" +
                    "3 - Запрос на определение золотого клиента, клиента с наибольшим количеством заказов, за указанный год, месяц.\n" +
                    "4 - \n" +
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

                    case "5":
                        return;
                        


                    default:
                        exceptionOutputConsole.Output("Действия не существует");
                        break;
                }
            }
        }
    }
}
