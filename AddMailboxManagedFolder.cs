using System.Management.Automation;

namespace EwsOAuth
{
    [Cmdlet(VerbsCommon.Add, "MailboxManagedFolder")]
    public class AddMailboxManagedFolder : Cmdlet
    {
        [Parameter(Position = 0, Mandatory = true, ValueFromPipelineByPropertyName = true)]
        public string Identity { get; set; }

        [Parameter(Position = 1, Mandatory = true, ValueFromPipelineByPropertyName = true)]
        public string FolderName { get; set; }

        [Parameter(Position = 2, Mandatory = true, ValueFromPipelineByPropertyName = true)]
        public string RetentionTagGuid { get; set; }

        [Parameter(Position = 3, Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public string AppId { get; set; }

        [Parameter(Position = 4, Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public string ClientSecret { get; set; }

        [Parameter(Position = 5, Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public string TenantId { get; set; }

        [Parameter(Position = 5, Mandatory = false, ValueFromPipelineByPropertyName = true)]
        public string RootFolderName { get; set; } = "Managed Folders";

        private string[] scopes = { "https://outlook.office365.com/.default" };
        private string authToken;

        protected override void BeginProcessing()
        {
            var cca = new AADConfidentialClient(AppId, ClientSecret, TenantId);

            WriteVerbose("Getting Azure AD authorization token...");
            authToken = ((cca.GetAuthToken(scopes)).Result).ToString();

            WriteDebug("Authorization Token recieved.");
        }

        protected override void ProcessRecord()
        {
            WriteVerbose("Connecting to Exchange Web Service...");
            var ewsClient = EwsService.EwsClient(Identity, authToken);

            WriteVerbose("Adding managed folder...");
            ManagedFolders.AddManagedFolder(ewsClient, Identity, RootFolderName, FolderName, RetentionTagGuid);
        }

    }
}
