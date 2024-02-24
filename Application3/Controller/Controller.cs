using Application3.Exceptions;
using Application3.Models;
using Application3.ModelView;
using Application3.Repository;
using Application3.Repository.Filter;
using Application3.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Application3.Controller
{
    public class Controller
    {
        IProductRepository _productRepository;
        IClientRepository _clientRepository;
        IOrderRepository _orderRepository;


        public Controller(IProductRepository productRepository,
                            IClientRepository clientRepository,
                            IOrderRepository orderRepository


            )
        {
            _productRepository = productRepository;
            _clientRepository = clientRepository;
            _orderRepository = orderRepository;

        }

        public ClientProductView ListClients(string productTtle)
        {
            try
            {
                List<ClientProduct> clientProducts = new();
                Product product = _productRepository.Select(new ProductFilter() { Title = productTtle }).FirstOrDefault();
                if (product == null)
                {
                    return new ClientProductView() { Status = new Status { State = 404, Message = $"Товар с названием {productTtle} не найден!" } };
                }

                List<Order> orders = _orderRepository.Select(new OrderFilter() { ProductId = product.Id }).ToList();

                foreach (Order order in orders)
                {
                    Client client = _clientRepository.GetById(order.ClientId);

                    clientProducts.Add(new ClientProduct() { Client = client, CountProduct = order.RequiredQuantity, PriceProduct = product.Price, Date = order.DatePlacement });
                }

                return new ClientProductView() { Status = new Status { State = 200, Message = "Успешно" }, ClientProducts = clientProducts };
            }
            catch (IOException ex)
            {
                return new ClientProductView() { Status = new Status { State = 404, Message = "Файл не найден или занят" } };
            }
            catch (Exception ex)
            {
                return new ClientProductView() { Status = new Status { State = 400, Message = "Неизвестное исключение. Сообщите в тех. поддержку" } };

            }




        }

        public UpdateClientView EditClient(string organizationTitle, string newContactPerson)
        {
            try
            {
                _clientRepository.UpdateClient(organizationTitle, newContactPerson);
                return new UpdateClientView { Status = new Status { State = 200, Message = "Успешно" } };
            }
            catch (NotFoundDataException ex)
            {
                return new UpdateClientView { Status = new Status { State = 404, Message = "Данных для обновления нет" } };
            }
            catch (UpdateFailedException ex)
            {
                return new UpdateClientView { Status = new Status { State = 404, Message = "Запись для обновления не найдена!" } };

            }
            catch (IOException ex)
            {
                return new UpdateClientView { Status = new Status { State = 404, Message = "Файл не доступен" } };
            }

        }

        public GoldenClientView GetGoldenClientDate(int year, int month)
        {
            try
            {
                List<Order> orders = _orderRepository.Select(new OrderFilter { Year = year, Month = month }).ToList();
                var ordersCountClient = (from order in orders
                                         group order by order.ClientId into clientGroup
                                         let clientId = clientGroup.Key
                                         let orderCount = clientGroup.Count()
                                         orderby orderCount descending
                                         select new { ClientId = clientId, OrderCount = orderCount }).FirstOrDefault();
                if (ordersCountClient is null)
                    return new GoldenClientView { Status = new Status { State = 201, Message = "В выбранную дату не совершались покупки" } };
                Client goldenClient = _clientRepository.GetById(ordersCountClient.ClientId);
                

                return new GoldenClientView { Status = new Status { State = 200, Message = "Успешно" }, Client = goldenClient };
            }
            catch (IOException ex) {
                return new GoldenClientView
                {
                    Status = new Status
                    {
                        State = 404,
                        Message = "Файл не доступен"
                    }
                };
            }
        }
    }
}
