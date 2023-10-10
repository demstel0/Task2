using System.Diagnostics;
using System.Formats.Tar;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

class Program
{
    static int productCode;
    protected static int clientCode;
    static float price;
    static int amount;
    static string unit;
    static DateTime dateOrder;
    protected static Excel.Application excelApp;
    protected static Excel.Workbook excelWB;
    protected static Excel._Worksheet excelWS;
    protected static Excel.Range excelRange;
    static string filePath;
    static void Main(string[] args)
    {      
        GetFilePath();
    }
    static void GetFilePath() 
    {
        try
        {
            Console.WriteLine("Пожалуйста, введите путь к файлу:");
            filePath = Console.ReadLine();
            if(filePath.Contains("\""))
            {
                string newPath = filePath.Replace("\"", "");
                filePath = newPath;
            }
            excelApp = new Excel.Application();
            excelWB = excelApp.Workbooks.Open(filePath);
            
        }
        catch (COMException ex)
        {
            Console.WriteLine($"{ex.Message} \nПожалуйста, введите корректный путь к файлу");
            GetFilePath();
        }
        MainMenu();
    }

    protected static void MainMenu()
    {
        Console.WriteLine("Пожалуйста, выберите режим работы: \n1 - Информация о продукте и заказах \n2 - Изменить контактное лицо организации " +
            "\n3 - Найти золотого клиента");
        string n = Console.ReadLine();
        if (n == "1") GetProduct(excelApp, excelWB);
        if (n == "2") Rename.MainRename();
        if (n == "3") GoldClient.MainGoldClient();
        else MainMenu();
    }
   static void GetProduct(Excel.Application excelApp2, Excel.Workbook excelWB2)
    {
        //Ищем товар по введенному пользователем названию
        excelWS = excelWB.Sheets[1];
        excelRange = excelWS.UsedRange;
        int rowCount = excelRange.Rows.Count;
        int columnCount = excelRange.Columns.Count;
        Console.WriteLine("Введите наименование товара: ");
        string name = Console.ReadLine();
        for (int i = 1; i <= rowCount; i++)
        {
            for (int j = 1; j <= columnCount; j++)
            {
                if (excelRange.Cells[i, j].Value.ToString() == name)
                {
                    price = (int)excelRange.Cells[i, j + 2].Value;
                    productCode = (int)excelRange.Cells[i, j - 1].Value();
                    unit = excelRange.Cells[i, j + 1].Value();
                    Console.WriteLine($"Цена за единицу:{price} ");
                }
            }
        }
        if(price==0)
        {
            Console.WriteLine("Товара не существует");
            NextStep();
        }
        GetClientCode();
        Console.ReadKey();
    }

    static void GetClientCode()
    {
        //Получаем код клиента
        excelWS = excelWB.Sheets[3];
        excelRange = excelWS.UsedRange;
        
        int rowCount = excelRange.Rows.Count;
        for (int i = 2; i <= rowCount; i++)
        {
            if(excelRange.Cells[i, 2].Value!=null && excelRange.Cells[i, 2].Value==productCode)
            {
                clientCode = (int)excelRange.Cells[i, 3].Value;
            }
        }
         if(clientCode==0)
        {
            Console.WriteLine("Заказа на данный товар нет");
        }

        GetClient();
    }

    protected static string GetClient(bool isFindOnlyName=false)
    {
        //Получаем название организации
        excelWS = excelWB.Sheets[2];
        excelRange = excelWS.UsedRange;
        string nameOrg=string.Empty;
        int rowCount = excelRange.Rows.Count;
        for (int i = 2; i <= rowCount; i++)
        {
            if (excelRange.Cells[i, 1].Value != null && excelRange.Cells[i, 1].Value == clientCode)
            {
                if(!isFindOnlyName) Console.WriteLine("Организация: " + excelRange.Cells[i, 2].Value);
                nameOrg = excelRange.Cells[i, 2].Value;
            }
        }
        
        if(!isFindOnlyName) GetInfoProduct();
        return nameOrg;
    }

    static void GetInfoProduct()
    {
        //Инфо о заказе и товаре
        excelWS = excelWB.Sheets[3];
        excelRange = excelWS.UsedRange;
        int rowCount = excelRange.Rows.Count;
        for (int i = 2; i <= rowCount; i++)
        {
            if(excelRange.Cells[i, 3].Value==clientCode && excelRange.Cells[i, 2].Value == productCode)
            {
                amount =(int) excelRange.Cells[i, 5].Value;
                dateOrder = excelRange.Cells[i, 6].Value;
                Console.WriteLine($"Требуемое количество: {amount} {unit}");
                Console.WriteLine($"Итоговая цена: {amount*price}");
                Console.WriteLine($"Дата заявки: {dateOrder}");
            }
        }
        NextStep();
    }
    static void NextStep()
    {
        //Выбор продолжения или выхода из программы
        productCode=0;
        clientCode=0;
        price=0;
        amount=0;
        unit=string.Empty;
        dateOrder=DateTime.Now;
        Console.WriteLine("------------------------------------------------");
        Console.WriteLine("Введите: \n1 - продолжить поиск по другим продуктам\n2-Выход в главное меню \nЛюбой другой символ - выйти из программы");
        string n = Console.ReadLine();
        if (n == "1") GetProduct(excelApp, excelWB);
        if (n == "2") MainMenu();
        else QuitApp();
    }
   protected static void QuitApp()
    {

        Console.WriteLine("Завершение работы...");
        excelWB.Close();
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);
        Marshal.ReleaseComObject(excelWB);
        Process.GetCurrentProcess().Kill();
    }
}

class Rename:Program
{
    static int numOrg;
    public static void MainRename()
    {
        excelWS = excelWB.Sheets[2];
        excelRange = excelWS.UsedRange;
        FindOrg();
    }

    static void FindOrg()
    {
        numOrg = 0;
        int rowCount = excelRange.Rows.Count;
        Console.WriteLine("Введите название организации:");
        string nameOrg = Console.ReadLine();
        if(nameOrg==string.Empty)
        {
            Console.WriteLine("Вы ввели пустую строку.");
            FindOrg();
        }
        for (int i = 2; i <= rowCount; i++)
           {
                if (excelRange.Cells[i, 2].Value == nameOrg)
                {
                    numOrg = i;
                    Console.WriteLine($"Текущее контактное лицо: {excelRange.Cells[i, 4].Value}");
                    break;
                }
           }
        RenameContact();
        
    }

    static void RenameContact()
    {
        if (numOrg != 0)
        {
            Console.WriteLine("\nВведите новое контактное лицо:");
            string newName = Console.ReadLine();
            if(newName==string.Empty)
            {
                Console.WriteLine("Вы ввели пустую строку");
                RenameContact();
            }
            excelWS.Cells[numOrg, 4] = newName;
            excelWB.Save();
            Console.WriteLine($"Изменения применены. \nНовое контактное лицо: {newName}");
            Console.WriteLine("Введите: \n1 - Изменить контактное лицо другой организации \n2 - Выход в главное меню \nДругой символ - выход из программы");
            string n = Console.ReadLine();
            if (n=="1")
                FindOrg();
            if (n == "2")
                MainMenu();
            else QuitApp();

        }
        else
        {
            Console.WriteLine("Указанной организации не существует. Пожалуйста, введите название снова:");
            FindOrg();
        }
    }
}

class GoldClient:Program
{    
    public static void MainGoldClient()
    {
        excelWS = excelWB.Sheets[2];
        excelRange = excelWS.UsedRange;
        SelectDateRange();
    }
    static void SelectDateRange()
    {
        Console.WriteLine("Введите год");
        string year = Console.ReadLine();
        Console.WriteLine("Введите месяц цифрой (например 1 - январь)");
        string month = Console.ReadLine();
        if (year == string.Empty || month==string.Empty)
        {
            Console.WriteLine("Вы ввели пустую строку.");
            SelectDateRange();
        }
        
        FindGoldClient(year, month);

    }
    static void FindGoldClient(string year, string month)
    {
        excelWS = excelWB.Sheets[3];
        excelRange = excelWS.UsedRange;
        
        List<int> clientsCode = new List<int>();
        List<DateTime> monthOrder = new List<DateTime>();
        int rowCount = excelRange.Rows.Count;
            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 3].Value!=null && !clientsCode.Contains((int)excelRange.Cells[i, 3].Value))
                {
                    clientsCode.Add((int)excelRange.Cells[i, 3].Value);
                    monthOrder.Add(excelRange.Cells[i, 6].Value);

                }
            }
            int[] numOrders = new int[clientsCode.Count];
            DateTime date = new DateTime(
                                Int32.Parse(year), int.Parse(month), 16);
        bool isNoOrders = true;
            for (int i = 2; i <= rowCount; i++)
            {
                if (excelRange.Cells[i, 4].Value != null && (int)excelRange.Cells[i, 6].Value.Month==date.Month && (int)excelRange.Cells[i, 6].Value.Year == date.Year)
                {
                    for(int j=0; j<clientsCode.Count; j++)
                    {
                        int numInList = 0;
                        if (excelRange.Cells[i, 3].Value == clientsCode[j])
                        {
                            numInList = j;
                            numOrders[numInList]++;
                            isNoOrders = false;
                        } 
                    }
                }
            }
        if (isNoOrders)
        {
            Console.WriteLine("В указанный промежуток заказы отсутствуют");
            MainMenu();
        }

        int max=0;
        int n=0;
        for(int i=0; i<numOrders.Length; i++)
        {
            if (numOrders[i] > max)
            {
                max = numOrders[i];
                n = i;
            }
        }
        clientCode = clientsCode[n];
        Console.WriteLine($"Золотой клиент в указанном месяце {GetClient(true)} с {numOrders[n]} заказов");
        MainMenu();  
    }
}
