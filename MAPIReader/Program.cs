using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Redemption;

namespace MAPIReader
{
    class Program
    {
        static void Main(string[] args)
        {
            //tell the app where the 32 and 64 bit dlls are located
            //by default, they are assumed to be in the same folder as the current assembly and be named
            //Redemption.dll and Redemption64.dll.
            //In that case, you do not need to set the two properties below
            RedemptionLoader.DllLocation64Bit = @"Redemption/redemption64.dll";
            RedemptionLoader.DllLocation32Bit = @"Redemption/redemption.dll";
            //Create a Redemption object and use it
            RDOSession session = RedemptionLoader.new_RDOSession();

            session.Logon(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            
            var stores = session.Stores;
            foreach (RDOStore rdoStore in stores)
            {
                if (rdoStore.Name.Contains("michael"))
                {
                    var folderId = rdoStore.RootFolder.EntryID;
                    PerformMailFix(folderId, session);
                }

            }
            Console.WriteLine(count);
            Console.ReadKey();
            session.Logoff();
           
        }

        private static int count = 0;
        private static void PerformMailFix(string folderId, RDOSession session)
        {
            var folder = session.GetFolderFromID(folderId);

            if (folder.FolderKind == rdoFolderKind.fkSearch)
                return;

            var before=new DateTime(2014,06,30);
            foreach (RDOMail item in folder.Items)
            {
                if (item.ReceivedTime >= before) continue;
                var diff = item.ReceivedTime - item.SentOn;
                if (!(diff.TotalHours > 10)) continue;
                Console.WriteLine("" + item.Subject + " - " + item.ReceivedTime + "    " + item.SentOn);
                count++;
                item.ReceivedTime = item.SentOn;
                item.Save();
            }

            Console.WriteLine(folder.DefaultMessageClass);
            //do the same fix for all subfolders
            foreach (RDOFolder subFolder in folder.Folders)
            {
                PerformMailFix(subFolder.EntryID, session);
            }
            
        }
    }
}
