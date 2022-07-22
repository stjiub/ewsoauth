
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;

namespace EwsOAuth
{
    public class ManagedFolders
    {
        public static void AddManagedFolder(ExchangeService ewsClient, string identity, string rootFolderName, string newFolderName, string tag)
        {
            Folder rootFolder = FindFolder(ewsClient, rootFolderName);
            if (rootFolder == null)
            {
                rootFolder = AddRootFolder(ewsClient, identity, rootFolderName);
            }

            Guid id = Guid.Parse(tag);
            if (TagExists(ewsClient, id))
            {
                Folder subFolder = FindFolder(ewsClient, newFolderName);
                if (subFolder == null)
                {
                    AddSubFolder(ewsClient, newFolderName, id, rootFolder);
                }
                else
                {
                    Console.Error.WriteLine("A folder named '{0}' already exists under '{1}'", newFolderName, rootFolderName);
                    throw new Exception("Cannot create folder of same as existing folder.");
                }
            }
            else
            {
                Console.Error.WriteLine("Retention policy named '{0}' was not found on the mailbox of '{1}'", newFolderName, identity);
                throw new Exception("Retention policy does not exist");
            }
        }

        public static Folder GetManagedFolder(ExchangeService ewsClient, string folderName)
        {
            Folder getFolder = FindFolder(ewsClient, folderName);
            if (getFolder != null)
            {
                Folder folder = Folder.Bind(ewsClient, getFolder.Id);
                return folder;
            }
            else
            {
                return null;
            }
        }

        public static void RemoveManagedFolder(ExchangeService ewsClient, string folderName)
        {
           
            Folder getFolder = FindFolder(ewsClient, folderName);
            if (getFolder != null)
            {
                try
                {
                    Folder managedFolder = Folder.Bind(ewsClient, getFolder.Id);
                    managedFolder.Delete(DeleteMode.SoftDelete);
                }
                catch
                {
                    Console.Error.WriteLine("Failed to remove managed folder {0}", folderName);
                    throw;
                }
            }
        }

        public static void SetManagedFolderTag(ExchangeService ewsClient, string identity, string setFolderName, string tag)
        {
            Folder managedFolder = GetManagedFolder(ewsClient, setFolderName);
            if (managedFolder != null)
            {
                Guid id = Guid.Parse(tag);

                try
                {
                    managedFolder.PolicyTag = new PolicyTag(true, id);
                    managedFolder.Update();
                }
                catch
                {
                    Console.Error.WriteLine("Failed to set the policy tag for folder '{0}' to GUID value {1}", setFolderName, tag);
                    throw;
                }
            }
            
        }

        public static Folder FindFolder(ExchangeService ewsClient, string name)
        {
            FolderView view = new FolderView(1);
            view.PropertySet = new PropertySet(FolderSchema.DisplayName);
            view.PropertySet.Add(FolderSchema.Id);
            view.PropertySet.Add(FolderSchema.ArchiveTag);
            view.Traversal = FolderTraversal.Deep;

            SearchFilter searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, name);
            FindFoldersResults foldersResults = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, searchFilter, view);

            try
            {
                return foldersResults.Folders[0];
            }
            catch
            {
                return null;
            }
        }

        private static Folder AddRootFolder(ExchangeService ewsClient, string identity, string rootFolderName)
        {
            var managedFolder = new Folder(ewsClient)
            {
                DisplayName = rootFolderName,
                FolderClass = "IPF.Note",
            };

            var managedFolderId = new FolderId(WellKnownFolderName.MsgFolderRoot, identity);
            var ewsParentFolder = Folder.Bind(ewsClient, managedFolderId);
            managedFolder.Save(ewsParentFolder.Id);

            return managedFolder;
        }

        private static Folder AddSubFolder(ExchangeService ewsClient, string name, Guid id, Folder rootFolder)
        {
            var managedFolder = new Folder(ewsClient)
            {
                DisplayName = name,
                FolderClass = "IPF.Note",
                PolicyTag = new PolicyTag(true, id)
            };
            managedFolder.Save(rootFolder.Id);

            return managedFolder;
        }

        private static bool TagExists(ExchangeService ewsClient, Guid id)
        {
            var tags = ewsClient.GetUserRetentionPolicyTags().RetentionPolicyTags;
            for (int i = 0; i < tags.Length; i++)
            {
                if (String.Equals(tags[i].RetentionId, id))
                {
                    return true;
                }
            }
            return false;
        }
    }
}