from suds.client import Client

import copy

'''
Splits large dictionary of KV pairs into a list of subdictionnaries
Used to get multiple mailboxes and other informations
Subdictionnaries are matched on the suffix appended to the key
Exemple : EmailAddresses_1 indicates this is part of the second entry
First dictionnary has regular key naming, therefore we append it at the end
FIXME : Match key on something better/more reliable than CustomAttribute1
'''
def objlist(mydict):
    i = 1
    l = [ ]
    dictcopy = copy.deepcopy(mydict)
    while mydict.has_key("CustomAttribute1_" + str(i)):
        subdict = { }
        suffix = "_" + str(i)
        for key, value in mydict.iteritems():
            if key.endswith(suffix):
                newkey = key.replace(suffix, '')
                subdict[newkey] = value
                del dictcopy[key]
        l.insert(0, subdict)
        i += 1
    l.append(dictcopy)
    return l

_clients = {}

def get_client(name):
    if name not in _clients:
        _clients[name] = Client("http://mail.contoso.com/EAWS/WebService.svc?wsdl")
    return _clients[name]

def connect():
    wsdl = "http://mail.contoso.com/EAWS/WebService.svc?wsdl"
    return (get_client("API"))
        
def DetailMailbox(login):
    return to_dict(connect().service.DetailMailbox(login))

def CreateMailbox(login):
    return to_dict(connect().service.CreateMailbox(login))

def GetMembersOfDistributionGroup(dg):
    return objlist(to_dict(connect().service.GetMembersOfDistributionGroup(dg)))

def AddMemberToDistributionGroup(dg, login, external):
    return to_dict(connect().service.AddUserToDistributionGroup(dg, login, external))

def RemoveMemberFromDistributionGroup(dg, login):
    return to_dict(connect().service.RemoveUserFromDistributionGroup(dg, login))

def CreateDistributionGroup(dg):
    return to_dict(connect().service.CreateDistributionGroup(dg))

def SetDistributionGroupOwner(dg, login):
    return to_dict(connect().service.SetDistributionGroupOwner(dg, login))

def EnableDistributionGroup(dg):
    return to_dict(connect().service.EnableDistributionGroup(dg))

def DisableDistributionGroup(dg):
    try:
        to_dict(connect().service.DisableDistributionGroup(dg))
    except:
        pass
    return "OK"

def DeleteDistributionGroup(dg):
    return to_dict(connect().service.DeleteDistributionGroup(dg))

def ListMailbox():
    return objlist(to_dict(connect().service.ListMailbox()))

def ListDistributionGroup():
    return objlist(to_dict(connect().service.ListDistributionGroup()))

def DetailDistributionGroup(dg):
    return to_dict(connect().service.DetailDistributionGroup(dg))

def to_dict(answer):
    d = { }
    if hasattr(answer, 'WsKeyValue'):
        for p in answer.WsKeyValue:
            d[p.Key] = p.Value
        return d
    else:
        return None
