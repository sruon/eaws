using System.Collections.Generic;
using System.Runtime.Serialization;
using System.ServiceModel;

namespace EAWS
{
    [ServiceContract]
    public interface IWebService
    {
        [OperationContract]
        string HealthCheck();

        #region Mailbox

        [OperationContract]
        List<WsKeyValue> CreateMailbox(string login);

        [OperationContract]
        List<WsKeyValue> DisableMailbox(string login);

        [OperationContract]
        List<WsKeyValue> DetailMailbox(string login);

        [OperationContract]
        List<WsKeyValue> ListMailbox();

        [OperationContract]
        List<WsKeyValue> AddEmailToMailbox(string login, string email);

        #endregion

        #region Distribution Group

        [OperationContract]
        List<WsKeyValue> ListDistributionGroup();

        [OperationContract]
        List<WsKeyValue> CreateDistributionGroup(string login);

        [OperationContract]
        List<WsKeyValue> SetDistributionGroupOwner(string dg, string login);

        [OperationContract]
        List<WsKeyValue> AddUserToDistributionGroup(string dg, string login, bool external);

        [OperationContract]
        List<WsKeyValue> RemoveUserFromDistributionGroup(string dg, string login);

        [OperationContract]
        List<WsKeyValue> DisableDistributionGroup(string login);

        [OperationContract]
        List<WsKeyValue> DeleteDistributionGroup(string login);

        [OperationContract]
        List<WsKeyValue> EnableDistributionGroup(string login);

        [OperationContract]
        List<WsKeyValue> DetailDistributionGroup(string login);

        [OperationContract]
        List<WsKeyValue> GetMembersOfDistributionGroup(string login);

        #endregion
    }

    [DataContract]
    public struct WsKeyValue
    {
        public WsKeyValue(string key, string value) : this()
        {
            Key = key;
            Value = value;
        }

        [DataMember]
        public string Key { get; set; }

        [DataMember]
        public string Value { get; set; }
    }
}