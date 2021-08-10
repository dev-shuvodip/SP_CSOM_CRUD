using System;
using System.Configuration;
using System.Security;
using Microsoft.SharePoint.Client;
using log4net;
using log4net.Config;

namespace SP_CSOM_DEMO2
{
    class Program
    {
        private static readonly ILog Log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static void Main(string[] args)
        {
            ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);
            List list = InitiateAuthentication(context);
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            Console.WriteLine($"Website Title : {web.Title}\nWebsite Url : {web.Url}\n\n");
            Console.WriteLine($"List name - {list.Title}\n\n");

            BasicConfigurator.Configure();

            int choice;
            do
            {
                Console.WriteLine("Enter the operation to perform: \n\n");
                Console.Write("1. Retrieve records\n\n2. Create new record\n\n3. Update a record\n\n4. Delete a record\n\n");
                choice = Convert.ToInt32(Console.ReadLine());
                Console.WriteLine();
                switch (choice)
                {
                    case 1:
                        Log.Info("Entering case 1");
                        RetrieveRecords();
                        break;
                    case 2:
                        Log.Info("Entering case 2");
                        CreateRecord();
                        break;
                    case 3:
                        Log.Info("Entering case 3");
                        Console.WriteLine("Update a record - \n\n");
                        Console.WriteLine("List view\n\n");
                        CamlQuery query3 = CamlQuery.CreateAllItemsQuery(100);
                        ListItemCollection items3 = list.GetItems(query3);
                        context.Load(items3);
                        context.ExecuteQuery();
                        Console.WriteLine("Id     Title");
                        foreach (ListItem item in items3)
                        {
                            Console.WriteLine($"{item.Id}   {item["Title"]}\n");
                        }
                        Console.WriteLine("Enter Id of the record to be updated - \n");
                        int itemId1 = Convert.ToInt32(Console.ReadLine());
                        Console.WriteLine();
                        UpdateRecord(itemId1);

                        break;
                    case 4:
                        Log.Info("Entering case 4");
                        Console.WriteLine("Delete a record - \n\n");
                        Console.WriteLine("List view\n\n");
                        CamlQuery query4 = CamlQuery.CreateAllItemsQuery(100);
                        ListItemCollection items4 = list.GetItems(query4);
                        context.Load(items4);
                        context.ExecuteQuery();
                        Console.WriteLine("Id     Title");
                        foreach (ListItem item in items4)
                        {
                            Console.WriteLine($"{item.Id}   {item["Title"]}\n");
                        }
                        Console.WriteLine("Enter Id of the record to be updated - \n");
                        int itemId2 = Convert.ToInt32(Console.ReadLine());
                        Console.WriteLine();
                        DeleteRecord(itemId2);
                        break;
                }
            } while (choice <= 4);
            Console.ReadLine();
        }
        /// <summary>
        /// 
        /// </summary>
        private static void RetrieveRecords()
        {
            ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);
            List list = InitiateAuthentication(context);
            BasicConfigurator.Configure();
            try
            {
                Console.WriteLine("Retrieve records - ");
                CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                ListItemCollection items = list.GetItems(query);
                context.Load(items);
                context.ExecuteQuery();
                foreach (ListItem listItem in items)
                {
                    Console.WriteLine($"Name - {listItem["Title"]}");
                    Console.WriteLine($"Email - {listItem["Email"]}");
                    Console.WriteLine($"Contact - {listItem["Contact"]}");
                    Console.WriteLine($"Subject - {(listItem["Subject"] as FieldLookupValue).LookupValue}");
                    Console.WriteLine($"Branch - {(listItem["Branch"] as FieldLookupValue).LookupValue}");
                    Console.WriteLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Log.Error("Error Message: " + e.Message.ToString(), e);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private static void CreateRecord()
        {
            ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);
            List list = InitiateAuthentication(context);
            BasicConfigurator.Configure();
            try
            {
                Console.WriteLine("Create new record - \n\n");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(itemCreateInfo);
                Console.WriteLine("Enter Title: ");
                newItem["Title"] = Console.ReadLine();
                Console.WriteLine("Enter Email: ");
                newItem["Email"] = Console.ReadLine();
                Console.WriteLine("Enter Contact: ");
                newItem["Contact"] = Console.ReadLine();
                Console.WriteLine("Enter Subject:");
                Console.WriteLine("               Choice 1: Electrical");
                Console.WriteLine("               Choice 2: Mechanical");
                Console.WriteLine("               Choice 3: Civil");
                Console.WriteLine("               Choice 4: Electronics and Communication");
                Console.WriteLine("               Choice 5: Computer Science");
                Console.WriteLine("               Choice 6: Bio-Technology");
                FieldLookupValue subjectId = new FieldLookupValue() { LookupId = Convert.ToInt32(Console.ReadLine()) };
                newItem["Subject"] = subjectId; Console.WriteLine("Enter Branch:");
                Console.WriteLine("              Choice 1: Kolkata");
                Console.WriteLine("              Choice 2: Delhi");
                Console.WriteLine("              Choice 3: Mumbai");
                Console.WriteLine("              Choice 4: Chennai");
                Console.WriteLine("              Choice 5: Bangalore");
                Console.WriteLine("              Choice 6: Ahmedabad");
                Console.WriteLine("              Choice 7: Kerala");
                FieldLookupValue branchId = new FieldLookupValue() { LookupId = Convert.ToInt32(Console.ReadLine()) };
                newItem["Branch"] = branchId;
                newItem.Update();
                context.ExecuteQuery();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Log.Error("Error Message: " + e.Message.ToString(), e);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        private static void UpdateRecord(int id)
        {
            BasicConfigurator.Configure();
            try
            {
                ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);
                List list = InitiateAuthentication(context);
                Console.WriteLine("Enter values: \n");
                ListItem listItem = list.GetItemById(id);
                Console.WriteLine("Enter Title: ");
                listItem["Title"] = Console.ReadLine();
                Console.WriteLine("Enter Email: ");
                listItem["Email"] = Console.ReadLine();
                Console.WriteLine("Enter Contact: ");
                listItem["Contact"] = Console.ReadLine();
                Console.WriteLine("Enter Subject:");
                Console.WriteLine("               Choice 1: Electrical");
                Console.WriteLine("               Choice 2: Mechanical");
                Console.WriteLine("               Choice 3: Civil");
                Console.WriteLine("               Choice 4: Electronics and Communication");
                Console.WriteLine("               Choice 5: Computer Science");
                Console.WriteLine("               Choice 6: Bio-Technology");
                FieldLookupValue subjectId = new FieldLookupValue() { LookupId = Convert.ToInt32(Console.ReadLine()) };
                listItem["Subject"] = subjectId;
                Console.WriteLine("Enter Branch:");
                Console.WriteLine("              Choice 1: Kolkata");
                Console.WriteLine("              Choice 2: Delhi");
                Console.WriteLine("              Choice 3: Mumbai");
                Console.WriteLine("              Choice 4: Chennai");
                Console.WriteLine("              Choice 5: Bangalore");
                Console.WriteLine("              Choice 6: Ahmedabad");
                Console.WriteLine("              Choice 7: Kerala");
                FieldLookupValue branchId = new FieldLookupValue() { LookupId = Convert.ToInt32(Console.ReadLine()) };
                listItem["Branch"] = branchId;
                listItem.Update(); context.ExecuteQuery();
                Console.WriteLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Log.Error("Error Message: " + e.Message.ToString(), e);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        private static void DeleteRecord(int id)
        {
            BasicConfigurator.Configure();
            try
            {
                ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);
                List list = InitiateAuthentication(context);
                ListItem listItem = list.GetItemById(id);
                listItem.DeleteObject();
                context.ExecuteQuery();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Log.Error("Error Message: " + e.Message.ToString(), e);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static string GetSPOUserName()
        {
            try
            {
                return ConfigurationManager.AppSettings["SPOAccount"];
            }
            catch
            {
                throw;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private static SecureString GetSPOSecureStringPassword()
        {
            try
            {
                SecureString secureString = new SecureString();
                foreach (char c in ConfigurationManager.AppSettings["SPOPassword"])
                {
                    secureString.AppendChar(c);
                }
                return secureString;
            }
            catch
            {
                throw;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="ctx"></param>
        /// <returns></returns>
        private static List InitiateAuthentication(ClientContext ctx)
        {
            ClientContext context = ctx;
            context.AuthenticationMode = ClientAuthenticationMode.Default;
            context.Credentials = new SharePointOnlineCredentials(GetSPOUserName(), GetSPOSecureStringPassword());
            List list = context.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["SPOList"]);
            context.Load(list);
            context.ExecuteQuery();
            return (list);
        }
    }
}
