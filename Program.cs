using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.IdentityModel;
using System.IO;
using System.Diagnostics;




namespace SHarePoiintConnect3
{
    class Program
    {

        static string siteCollectionUrl = "https://account9999.sharepoint.com/";
        static string siteUrl = "https://account9999.sharepoint.com/sites/CesarSite/";
        static string userName = "castanc@account9999.onmicrosoft.com";
        static string password = "Minidisc01";
        static Site site;
        static Web web;
        static ClientContext ctx; 

        public static void ConnectToSharePointOnline(string url)
        {

            // Namespace: Microsoft.SharePoint.Client    
            ctx = new ClientContext(url);

            // Namespace: System.Security  
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);

            // Namespace: Microsoft.SharePoint.Client    
            ctx.Credentials = new SharePointOnlineCredentials(userName, secureString);


            // Namespace: Microsoft.SharePoint.Client    
            site = ctx.Site;
            web = ctx.Web;

            ctx.Load(site);
            ctx.Load(web);
            ctx.ExecuteQuery();

            Console.WriteLine(site.Url.ToString());
        }


        public static void Connect2()
        {
            using (ClientContext spcontext = new ClientContext("https://account9999.sharepoint.com/sites/CesarSite?market=en-US"))
            {
                spcontext.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;

                //SecureString secureString = new SecureString();
                //password.ToList().ForEach(secureString.AppendChar);

                spcontext.FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo(
                    userName, password);


                spcontext.Load(spcontext.Web, w => w.Title, w => w.ServerRelativeUrl, w => w.Lists);
                spcontext.ExecuteQuery();
                Console.WriteLine(spcontext.Web.ServerRelativeUrl);
            }

        }


        static void ImportList()
        {

            // update site url here
            string siteURL = site.Url.ToString();

            SPSecurity.RunWithElevatedPrivileges(delegate () {

                using (SPSite site = new SPSite(siteURL))
                {
                    using (SPWeb web = site.OpenWeb())
                    {

                        SPList listTo = web.Lists["Lookups"]; // new list


                        for(int i=0;i<10;i++)
                        {
                            //SPItem item = new SPItem();
                            // match your field items 
                            SPListItem lookUp = listTo.AddItem();
                            lookUp["Title"] = $"Title_{i}";
                            lookUp["Id"] = i;
                            lookUp.Update();
                        }


                    }
                }

            }); // end run with elevated privilages
        }

        static void ListAlLists()
        {
            ctx.Load(web.Lists,
                         lists => lists.Include(list => list.Title,
                                                list => list.Id));

            // Execute query.
            ctx.ExecuteQuery();

            // Enumerate the web.Lists.
            StringBuilder sb = new StringBuilder();
            int count = 0;
            foreach (List list in web.Lists)
            {
                sb.AppendLine($"{count}\t{list.Id}\t{list.Title}");
                //ctx.Load(list.Fields);
                //ctx.ExecuteQuery();
                //foreach (var f in list.Fields)
                //{
                //    try
                //    {
                //        sb.AppendLine($"\t{f.InternalName}\t{f.Id}\t{f.Title}\t{f.InternalName}");
                //    }
                //    catch (Exception ex)
                //    {
                //        sb.AppendLine($"\tEXCEPTION\t{ex.Message}");
                //    }
                //}
                count++;
            }
            System.IO.File.WriteAllText(@"c:\temp\sharepoint.txt", sb.ToString());
            Console.WriteLine(sb.ToString());
        }

        static void AddItemToList(string listName)
        {
            var list = ctx.Web.Lists.GetByTitle("Lookups");
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View><Query><Where><Contains><FieldRef Name='Title'/><Value Type='Text'>from</Value></Contains></Where></Query></View>";


            ctx.Load(list.Fields);
            var items = list.GetItems(query);
            ctx.ExecuteQuery();

            int itemCOunt = 0;
            foreach (var it in items)
                itemCOunt++;

            Console.WriteLine($"items count {itemCOunt}");
            System.IO.File.WriteAllText(@"c:\temp\itemscount.txt", $"items count {itemCOunt}");


            var fields = list.Fields;

            var sb = new StringBuilder();
            foreach(var f in list.Fields)
            {
                sb.AppendLine($"{f.Id}\t{f.Title}\t{f.InternalName}");
            }
            System.IO.File.WriteAllText(@"c:\temp\Lookups_Fields.txt", sb.ToString());


            int batchCount = 0;
            int maxrecs = 6000;
            int count = 0;
            //Parallel.For(200, maxrecs, i =>
            for(int i=10000; i<16000; i++ )
            {
                try
                {
                    count++;
                    ListItemCreationInformation itemCreationInfo = new ListItemCreationInformation();
                    ListItem newItem = list.AddItem(itemCreationInfo);

                    newItem["Title"] = $"New Item from c# {i}";
                    newItem["Id0"] = i;
                    newItem.Update();
                    ctx.ExecuteQuery();
                    //batchCount++;
                    //if (batchCount > 5000)
                    //{
                    //    ctx.ExecuteQuery();
                    //    batchCount = 0;
                    //}
                    Console.WriteLine($"{count}\t{i}");
                }
                catch(Exception ex)
                {
                    Console.WriteLine("EXCEPTION " + ex.Message);
                    //Console.ReadLine();
                }
            }
        }


        static void Main(string[] args)
        {
            try
            {
                //ConnectToSharePointOnline(siteCollectionUrl);
                //ListAlLists();
                ConnectToSharePointOnline(siteUrl);
                ListAlLists();
                AddItemToList("Lookups");
            }catch(Exception ex)
            {
                string s = ex.Message;
            }
            Console.WriteLine("press enter");
            Console.ReadLine();

        }
    }
}
