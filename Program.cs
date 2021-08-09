using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace SP_CSOM_DEMO2
{
    class Program
    {
        static void Main(string[] args)
        {
            ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"])
            {
                AuthenticationMode = ClientAuthenticationMode.Default,
                Credentials = new SharePointOnlineCredentials(_getSPOUserName(), _getSPOSecureStringPassword())
            };
            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();
            Console.WriteLine($"Website Title : {web.Title}\nWebsite Url : {web.Url}\n\n");
            List list = context.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["SPOList"]);
            context.Load(list);
            context.ExecuteQuery();
            Console.WriteLine($"List name - {list.Title}\n\n");

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
                                Console.WriteLine($"Subject - {(listItem["Subject"] as FieldLookupValue).LookupValue.ToString()}");
                                Console.WriteLine($"Branch - {(listItem["Branch"] as FieldLookupValue).LookupValue.ToString()}");
                                Console.WriteLine();
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        break;
                    case 2:
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
                            newItem["Subject"] = subjectId;
                            Console.WriteLine("Enter Branch:");
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
                        }
                        break;
                    case 3:
                        try
                        {
                            Console.WriteLine("Update a record - \n\n");
                            Console.WriteLine("List view\n\n");
                            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                            ListItemCollection items = list.GetItems(query);
                            context.Load(items);
                            context.ExecuteQuery();
                            Console.WriteLine("Id     Title");
                            foreach (ListItem item in items)
                            {
                                Console.WriteLine($"{item.Id}   {item["Title"]}\n");
                            }
                            Console.WriteLine("Enter Id of the record to be updated - \n");
                            int itemId = Convert.ToInt32(Console.ReadLine());
                            Console.WriteLine();
                            Console.WriteLine("Enter values: \n");
                            ListItem listItem = list.GetItemById(itemId);
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
                            listItem.Update();
                            context.ExecuteQuery();
                            Console.WriteLine();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        break;
                    case 4:
                        try
                        {
                            Console.WriteLine("Delete a record - \n\n");
                            Console.WriteLine("List view\n\n");
                            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                            ListItemCollection items = list.GetItems(query);
                            context.Load(items);
                            context.ExecuteQuery();
                            Console.WriteLine("Id     Title");
                            foreach (ListItem item in items)
                            {
                                Console.WriteLine($"{item.Id}   {item["Title"]}\n");
                            }
                            Console.WriteLine("Enter Id of the record to be updated - \n");
                            int itemId = Convert.ToInt32(Console.ReadLine());
                            Console.WriteLine();
                            ListItem listItem = list.GetItemById(itemId);
                            listItem.DeleteObject();
                            context.ExecuteQuery();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                        }
                        break;
                }
            } while (choice <= 4);
            Console.ReadLine();
        }

        private static SecureString _getSPOSecureStringPassword()
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

        private static string _getSPOUserName()
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
    }
}
