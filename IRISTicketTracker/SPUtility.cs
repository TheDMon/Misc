using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace IRISTicketTracker
{
    public class SPUtility
    {
        string siteUrl = Program.GetAppSettingValue("SharePointCRSiteUrl");
        string listTitle = Program.GetAppSettingValue("SharePointCRListName");

        public List<ChangeEntity> ReadData()
        {
            // creates a new SharePoint context 	
            // passing the parameters	
            ClientContext clientContext = new ClientContext(siteUrl);
            clientContext.Credentials = SetCredentials();

            // selects the list by title	
            SP.List list = clientContext.Web.Lists.GetByTitle(listTitle);

            // requests the item collection, without filters	
            SP.ListItemCollection itemCollection = list.GetItems(new CamlQuery());

            // loads the item collection object in the context object	
            clientContext.Load(itemCollection);

            // executes everything that was loaded before (the itemCollection)	
            clientContext.ExecuteQuery();

            // iterates the list printing the title of each list item 	
            List<ChangeEntity> lstChanges = new List<ChangeEntity>();
            foreach (SP.ListItem item in itemCollection)
            {
                var change = new ChangeEntity();
                change.ID = (int)item["ID"];
                change.Application = item["Application_x0020_Name"].ToString().Trim();
                change.Title = item["Title"].ToString().Trim();
                change.SOW = item["SOW_x0020_Number"] != null ? item["SOW_x0020_Number"].ToString().Trim() : "";
                change.Assignee = item["Assignee"] != null ? item["Assignee"].ToString().Trim() : "";
                change.Status = item["Status"] != null ? item["Status"].ToString().Trim(): "";
                change.ResolutionMethod = item["ECC_x002f_CCP_x002f_ECCDC_x0023_"] != null ? item["ECC_x002f_CCP_x002f_ECCDC_x0023_"].ToString().Trim() : "";
                change.ActualResolutionDate = (System.DateTime?)item["Actual_x0020_Resolution_x0020_Da"];
                change.Created = (System.DateTime)item["Created"];
                change.Modified = (System.DateTime)item["Modified"];
                change.Resolved = item["Resolved_x0020_Date"] != null && item["Resolved_x0020_Date"].ToString() != "" ? (DateTime)item["Resolved_x0020_Date"] : (DateTime?)null; 

                lstChanges.Add(change);

            }

            return lstChanges;
        }

        private SharePointOnlineCredentials SetCredentials()
        {
            SecureString passWord = new SecureString();
            foreach (char c in Program.GetAppSettingValue("SharePointLoginPWD").ToCharArray()) passWord.AppendChar(c);
            return new SharePointOnlineCredentials(Program.GetAppSettingValue("SharePointLoginID"), passWord);
        }


    }

    public class ChangeEntity
    {
        public int ID { get; set; }
        public string Application { get; set; }
        public string Title { get; set; }
        public string SOW { get; set; }
        public string Assignee { get; set; }
        public string Status { get; set; }
        public string ResolutionMethod { get; set; }        
        public DateTime? ActualResolutionDate { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public DateTime? Resolved { get; set; }
    }
}
