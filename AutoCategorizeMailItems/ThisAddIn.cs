/*
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. 
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that. 
You agree: 
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; 
and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code 
*/

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
