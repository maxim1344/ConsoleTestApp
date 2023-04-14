using ClosedXML.Excel;

namespace ConsoleApp
{
    public class Products
    {
        public int ProductCode { get; set; }
        public string ProductName { get; set; }
        public string Units { get; set; }
        public decimal Price { get; set; }
    }

    public class Customers
    {
        public int CustomerCode { get; set; }
        public string NameOfOrganization { get; set; }
        public string Address { get; set; }
        public string ContactPerson { get; set; }
    }

    public class Orders
    {
        public int OrderCode { get; set; }
        public int ProductCode { get; set; }
        public int CustomerCode { get; set; }
        public int OrderNumber { get; set; }
        public int Quantity { get; set; }
        public DateTime OrderDate { get; set; }
    }

    public class CustomerService
    {
        private Dictionary<int, Products> products = new Dictionary<int, Products>();  //хранение товаров
        private Dictionary<int, Customers> customers = new Dictionary<int, Customers>(); //хранение клиентов
        private List<Orders> orders = new List<Orders>(); //хранение заявок
        private List<IXLWorksheet> worksheetsList = new List<IXLWorksheet>(); //список листов excel файла
        private List<IXLRangeRows> allRowsList = new List<IXLRangeRows>(); //список всех строк с данными excel файла
        public bool exist; //флаг проверки существования файла для загрузки
        public string dataFilePath; //путь файла с данными

        public Dictionary<int, Products> GetProducts //свойство для получения доступа к продуктам
        {
            get { return products; }
        }

        public Dictionary<int, Customers> GetCustomers //свойство для получения доступа к клиентам
        {
            get { return customers; }
        }

        public List<Orders> GetOrders
        {
            get { return orders; }
        } //свойство для получения доступа к заявкам

        public void Clear()
        {
            products.Clear();
            customers.Clear();
            orders.Clear();
            worksheetsList.Clear();
            allRowsList.Clear();
            LoadFile(dataFilePath);
        }

        public void LoadFile(string dataFilePath) //метод загрузки файла с данными
        {          

            if (File.Exists(dataFilePath)) //проверка, существует ли файл по указанному адресу
            {
                exist = true;
                try
                {
                    using (XLWorkbook workBook = new XLWorkbook(dataFilePath)) //загрузка файла с данными
                    {
                        foreach (var worksheet in workBook.Worksheets)
                        {
                            allRowsList.Add(worksheet.RangeUsed().RowsUsed()); //получение списка строк с данными excel файла
                            worksheetsList.Add(worksheet); //получение списка листов excel файла
                        }
                        foreach (var rangeRows in allRowsList) //получение данных
                        {
                            int count = 0;
                            foreach (var row in rangeRows)
                            {
                                if (count == 0)
                                {
                                    count++;
                                    continue;
                                }
                                if (row.Worksheet.Name.Contains("Товары"))
                                {
                                    AddProduct(row, count);
                                }
                                if (row.Worksheet.Name.Contains("Клиенты"))
                                {
                                    AddCustomer(row, count);
                                }
                                if (row.Worksheet.Name.Contains("Заявки"))
                                {
                                    AddOrder(row);
                                }
                                count++;
                            }
                        }
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine($"\nПроизошла ошибка загрузки файла:\n\n\n {e}");
                    exist = false;
                    Console.ReadKey();
                }
            }
            else
            {
                exist = false;
                Console.WriteLine("Файл по указанному пути не найден");
                Console.ReadKey();
            }
        }

        private void AddOrder(IXLRangeRow row) //метод добавления заявки в память программы
        {
            Orders order = new Orders
            {
                OrderCode = (int)row.Cell(1).Value,
                ProductCode = (int)row.Cell(2).Value,
                CustomerCode = (int)row.Cell(3).Value,
                OrderNumber = (int)row.Cell(4).Value,
                Quantity = (int)row.Cell(5).Value,
                OrderDate = (DateTime)row.Cell(6).Value
            };
            orders.Add(order);
        }

        private void AddCustomer(IXLRangeRow row, int count) //метод добавления клиента в память программы
        {
            Customers customer = new Customers
            {
                CustomerCode = (int)row.Cell(1).Value,
                NameOfOrganization = (string)row.Cell(2).Value,
                Address = (string)row.Cell(3).Value,
                ContactPerson = (string)row.Cell(4).Value
            };
            customers.Add(count, customer);
        }

        private void AddProduct(IXLRangeRow row, int count) //метод добавления товаров в память программы
        {
            Products product = new Products
            {
                ProductCode = (int)row.Cell(1).Value,
                ProductName = row.Cell(2).Value.ToString(),
                Units = row.Cell(3).Value.ToString(),
                Price = decimal.Parse(row.Cell(4).Value.ToString())
            };
            products.Add(count, product);
        }

        public void UpdateContactPerson(int organizationIndex, string newContactCustomer) //метод изменения контактного лица клиента
        {
            if (customers.ContainsKey(organizationIndex))
            {
                using (XLWorkbook workBook = new XLWorkbook(dataFilePath))
                {

                    var ws = workBook.Worksheet(2);

                    foreach (var cell in ws.RangeUsed().RowsUsed().Cells())
                    {

                        if (cell.Value.ToString() == customers[organizationIndex].ContactPerson)
                        {
                            cell.SetValue(newContactCustomer);
                        }

                    }
                    ws.Columns().AdjustToContents();

                    workBook.SaveAs(dataFilePath);

                }
                Console.WriteLine($"\nИзменения контактного лица для организации {customers[organizationIndex].NameOfOrganization} успешно \n");
                Console.WriteLine("Для продолжения нажмите любую клавишу\n");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine("Некорректный выбор организации\n");
                Console.WriteLine("Для продолжения нажмите любую клавишу\n");
                Console.ReadKey();
            }
        }

        public void PrintCustomersByProduct(int product) //метод поиска информации о клиентах по наименованию товара
        {
            bool customerIsExist = false;
            if (products.ContainsKey(product))
            {
                foreach (var order in orders)
                {
                    if (order.ProductCode == products[product].ProductCode)
                    {
                        foreach (var customer in customers)
                        {
                            if (order.CustomerCode == customer.Value.CustomerCode)
                            {
                                Console.WriteLine($"Список клиентов сделавших заказ товара - {products[product].ProductName}\n");
                                Console.WriteLine($"Название организации: {customer.Value.NameOfOrganization}");
                                Console.WriteLine($"Контактное лицо: {customer.Value.ContactPerson}");
                                Console.WriteLine($"Количество: {order.Quantity}");
                                Console.WriteLine($"Цена: {products[product].Price}");
                                Console.WriteLine($"Дата заказа: {order.OrderDate.ToString("dd.MM.yyyy")}");
                                customerIsExist = true;
                            }
                        }
                    }
                }
                if (!customerIsExist)
                {
                    Console.WriteLine("\nДля данного товара отсутствуют заказы\n");
                }
                Console.WriteLine("\nДля продолжения нажмите любую клавишу\n");
                Console.ReadKey();
            }
            else
            {
                Console.WriteLine($"\nНекорректный выбор товара\n");
                Console.WriteLine("\nДля продолжения нажмите любую клавишу\n");
                Console.ReadKey();
            }
        }

        public void MainMenu() //главное меню программы
        {
            Console.Clear();
            Console.WriteLine("Выберите опцию:");
            Console.WriteLine("1. Составить список клиентов сделавших заказ по товару");
            Console.WriteLine("2. Изменение контактного лица клиента");
            Console.WriteLine("3. Показать \"Золотого\" клиента с наибольшим количеством заказов, за указанный год, месяц.");
            Console.WriteLine("4. Exit\n");
        }

        public void PrintGoldenCustomer(int year, int month) //метод поиска "Золотого клиента"
        {
            List<int> tempList = new List<int>();
            foreach (var order in orders)
            {
                if (order.OrderDate.Year == year && order.OrderDate.Month == month)
                {
                    tempList.Add(order.CustomerCode);
                }
            }
            if (tempList.Count != 0)
            {
                var result = tempList.GroupBy(n => n).OrderByDescending(g => g.Count()).First();
                foreach (var customer in customers.Values)
                {
                    if (customer.CustomerCode == result.Key)
                    {
                        Console.WriteLine($"\nЗолотой клиент за {year}-{month:00}: {customer.NameOfOrganization}");
                        Console.ReadKey();
                        return;
                    }
                }
            }
            else
            {
                Console.WriteLine($"\nЗа {year}-{month:00} золотых клиентов не было");
                Console.ReadKey();
            }
        }

    }

    public class Program
    {
        static void Main(string[] args)
        {
            CustomerService customerService = new CustomerService();
            while (!customerService.exist)
            {
                Console.Clear();
                Console.WriteLine("Для начала работы программы введите путь до файла с данными:");
                
                customerService.LoadFile(customerService.dataFilePath = Console.ReadLine());
            }


            while (true)
            {
                customerService.MainMenu();

                var option = Console.ReadLine();
                switch (option)
                {
                    case "1":
                        Console.Clear();
                        int count = 0;
                        foreach (var item in customerService.GetProducts)
                        {
                            count++;
                            Console.WriteLine("{0}. {1}", count, item.Value.ProductName);
                        }
                        Console.WriteLine("\n0. Главное меню");
                        Console.WriteLine("\nВыберите товар:");
                        int product = int.Parse(Console.ReadLine());
                        if (product == 0)
                        {
                            customerService.MainMenu();
                            break;
                        }
                        Console.Clear();
                        customerService.PrintCustomersByProduct(product);
                        break;
                    case "2":
                        Console.Clear();
                        int count1 = 0;
                        foreach (var item in customerService.GetCustomers)
                        {
                            count1++;
                            Console.WriteLine("{0}. {1}", count1, item.Value.NameOfOrganization);
                        }
                        Console.WriteLine("\n0. Главное меню");
                        Console.WriteLine("\nВыберите организацию для изменения контактного лица (ФИО):\n");
                        var organizationIndex = int.Parse(Console.ReadLine());
                        if (organizationIndex == 0)
                        {
                            customerService.MainMenu();
                            break;
                        }
                        Console.WriteLine($"\nВведите новое ФИО контактного лица для организации\n");
                        var newContactCustomer = Console.ReadLine();
                        customerService.UpdateContactPerson(organizationIndex, newContactCustomer);
                        customerService.Clear();
                        break;
                    case "3":
                        Console.Clear();
                        Console.WriteLine("Введите год:");
                        var year = int.Parse(Console.ReadLine());
                        Console.WriteLine("\nВведите месяц:");
                        var month = int.Parse(Console.ReadLine());
                        customerService.PrintGoldenCustomer(year, month);
                        break;
                    case "4":
                        Environment.Exit(0);
                        break;
                    default:
                        Console.Clear();
                        Console.WriteLine("Неверный выбор");
                        Console.WriteLine("\nДля продолжения нажмите любую клавишу");
                        Console.ReadKey();
                        customerService.MainMenu();
                        break;
                }
            }
        }


    }
}