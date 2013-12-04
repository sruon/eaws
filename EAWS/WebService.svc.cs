using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using EAWS.Properties;

namespace EAWS
{
    public class WebService : IWebService
    {
        #region WS Interface

        #region Generic

        public string HealthCheck()
        {
            return "Looks like I'm OKAY, thanks for worrying though.";
        }

        #endregion

        #region Mailbox

        public List<WsKeyValue> CreateMailbox(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Enable-Mailbox", parameters);
        }

        public List<WsKeyValue> DisableMailbox(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Disable-Mailbox", parameters);
        }

        public List<WsKeyValue> DetailMailbox(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Get-MailboxStatistics", parameters, false);
        }

        public List<WsKeyValue> ListMailbox()
        {
            return RunPowershell("Get-Mailbox");
        }

        public List<WsKeyValue> AddEmailToMailbox(string login, string email)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login},
                {"EmailAddresses", "@{add=\"" + email + "\"}"}
            };
            return RunPowershellWithParameters("Set-Mailbox", parameters);
        }

        #endregion

        #region Distribution Group

        public List<WsKeyValue> CreateDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Name", login},
                {"IgnoreNamingPolicy", null},
                {"OrganizationalUnit", Constants.DG_OU}
            };
            List<WsKeyValue> result = RunPowershellWithParameters("New-DistributionGroup", parameters);
            parameters.Clear();
            parameters.Add("Identity", login);
            parameters.Add("RequireSenderAuthenticationEnabled", "FALSE");
            result.Concat(RunPowershellWithParameters("Set-DistributionGroup", parameters));

            return result;
        }

        public List<WsKeyValue> SetDistributionGroupOwner(string dg, string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", dg},
                {"ManagedBy", login},
                {"BypassSecurityGroupManagerCheck", null}
            };
            return RunPowershellWithParameters("Set-DistributionGroup", parameters);
        }

        public List<WsKeyValue> CreateContact(string email)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Name", email},
                {"ExternalEmailAddress", email},
                {"OrganizationalUnit", Constants.ALIAS_OU}
            };
            List<WsKeyValue> result = RunPowershellWithParameters("New-MailContact", parameters);
            return result;
        }

        public List<WsKeyValue> AddUserToDistributionGroup(string dg, string login, bool external = false)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", dg},
                {"Member", login},
                {"BypassSecurityGroupManagerCheck", null}
            };
            if (external)
                CreateContact(login);
            List<WsKeyValue> result = RunPowershellWithParameters("Add-DistributionGroupMember", parameters);
            if (!external)
                {
                    parameters.Clear();
                    parameters.Add("Identity", dg);
                    parameters.Add("ExtendedRights", "Send-As");
                    parameters.Add("User", login);
                    parameters.Add("AccessRights", "ExtendedRight");
                    result.Concat(RunPowershellWithParameters("Add-ADPermission", parameters));
                }
            return result;
        }

        public List<WsKeyValue> RemoveUserFromDistributionGroup(string dg, string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", dg},
                {"Member", login},
                {"BypassSecurityGroupManagerCheck", null}
            };
            List<WsKeyValue> result = RunPowershellWithParameters("Remove-DistributionGroupMember", parameters);
            parameters.Clear();
            parameters.Add("Identity", dg);
            parameters.Add("ExtendedRights", "Send-As");
            parameters.Add("User", login);
            parameters.Add("AccessRights", "ExtendedRight");
            result.Concat(RunPowershellWithParameters("Remove-ADPermission", parameters));

            return result;
        }

        public List<WsKeyValue> DisableDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Disable-DistributionGroup", parameters);
        }

        public List<WsKeyValue> DeleteDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Remove-DistributionGroup", parameters);
        }

        public List<WsKeyValue> EnableDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Enable-DistributionGroup", parameters, false);
        }

        public List<WsKeyValue> DetailDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Get-DistributionGroup", parameters, false);
        }

        public List<WsKeyValue> GetMembersOfDistributionGroup(string login)
        {
            var parameters = new Dictionary<string, string>
            {
                {"Identity", login}
            };
            return RunPowershellWithParameters("Get-DistributionGroupMember", parameters, false);
        }

        public List<WsKeyValue> ListDistributionGroup()
        {
            return RunPowershell("Get-DistributionGroup");
        }

        #endregion

        #region Dynamic Distribution Group

        #endregion

        #endregion

        #region Powershell API

        public List<WsKeyValue> RunPowershell(string cmdlet)
        {
            var secured = new SecureString();
            foreach (char x in Constants.WS_PASSWORD)
                secured.AppendChar(x);
            var exchangeCredential = new PSCredential(Constants.WS_USER, secured);
            var connectionInfo = new WSManConnectionInfo(new Uri(Constants.PS_URI),
                "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
                exchangeCredential);

            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            PowerShell powershell = PowerShell.Create();

            var command = new PSCommand();

            command.Commands.Add(cmdlet);
            powershell.Commands = command;
            runspace.Open();
            powershell.Runspace = runspace;
            Collection<PSObject> results = powershell.Invoke();

            var returnvalue = new List<WsKeyValue>();

            int i = 0;
            foreach (PSObject psObject in results)
                {
                    foreach (PSPropertyInfo psPropertyInfo in psObject.Properties)
                        {
                            if (psPropertyInfo.Name != null && psPropertyInfo.Value != null &&
                                psPropertyInfo.Value.ToString() != null)
                                {
                                    string propertyName;
                                    if (i == 0)
                                        propertyName = psPropertyInfo.Name;
                                    else
                                        propertyName = psPropertyInfo.Name + "_" + i.ToString();
                                    returnvalue.Add(new WsKeyValue(propertyName, psPropertyInfo.Value.ToString()));
                                }
                        }
                    ++i;
                }
            runspace.Close();
            return returnvalue;
        }

        public List<WsKeyValue> RunPowershellWithParameters(string cmdlet, Dictionary<string, string> parameters,
            Boolean force = true)
        {
            var secured = new SecureString();
            foreach (char x in Constants.WS_PASSWORD)
                secured.AppendChar(x);
            var exchangeCredential = new PSCredential(Constants.WS_USER, secured);
            var connectionInfo = new WSManConnectionInfo(new Uri(Constants.PS_URI),
                "http://schemas.microsoft.com/powershell/Microsoft.Exchange",
                exchangeCredential);

            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Basic;
            PowerShell powershell = PowerShell.Create();

            var command = new PSCommand();

            command.AddCommand(cmdlet);
            foreach (var entry in parameters)
                {
                    if (entry.Value != null && entry.Value != "FALSE")
                        command.AddParameter(entry.Key, entry.Value);
                    else if (entry.Value == "FALSE")
                        command.AddParameter(entry.Key, false);
                    else
                        command.AddParameter(entry.Key);
                }
            if (force)
                command.AddParameter("Confirm", false);
            powershell.Commands = command;
            runspace.Open();
            powershell.Runspace = runspace;
            Collection<PSObject> results = powershell.Invoke();

            var returnvalue = new List<WsKeyValue>();
            int i = 0;
            foreach (PSObject psObject in results)
                {
                    foreach (PSPropertyInfo psPropertyInfo in psObject.Properties)
                        {
                            if (psPropertyInfo.Name != null && psPropertyInfo.Value != null &&
                                psPropertyInfo.Value.ToString() != null)
                                {
                                    string propertyName;
                                    if (i == 0)
                                        propertyName = psPropertyInfo.Name;
                                    else
                                        propertyName = psPropertyInfo.Name + "_" + i.ToString();
                                    returnvalue.Add(new WsKeyValue(propertyName, psPropertyInfo.Value.ToString()));
                                }
                        }
                    ++i;
                }
            runspace.Close();
            return returnvalue;
        }

        #endregion
    }
}