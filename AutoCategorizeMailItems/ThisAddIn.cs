using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Outlook;

namespace AutoCategorizeMailItems
{
    public partial class ThisAddIn
    {
        private const string CategoryName = "CRM";
        
        private List<Items> ItemsPersister = new List<Items>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            foreach (Folder folder in Application.Session.Folders)
            {
                var inbox = folder.Folders["Inbox"];
                var items = inbox.Items;
                ItemsPersister.Add(items);
                items.ItemAdd += new ItemsEvents_ItemAddEventHandler(CategorizeItem);
            }
        }

        private void CategorizeItem(object item)
        {
            if(item is MailItem)
            {
                var mailItem = (MailItem)item;
                var categories = mailItem.Categories == null
                    ? new List<string>()
                    : mailItem.Categories.Split(',').Select(leadingSpace => leadingSpace.TrimStart(' ')).ToList();
                if (!categories.Contains(CategoryName))
                {
                    categories.Add(CategoryName);
                    mailItem.Categories = string.Join(", ", categories);
                    mailItem.Save();
                }
            }
            else
            {
                throw new NotSupportedException(
                    "Invalid call to CategorizeItem. "
                    + $"Parameter 'item' ({item.GetType()}) must be a MailItem.");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
