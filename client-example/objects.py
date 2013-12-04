import eaws

class DistributionGroup(object):
    enabled = None
    restriction = None
    email = None
    restrict_to_internal = None
    name = None
    manager = None
    login = None
    hidden = None
    def __init__(self, dg, mydict = None):
        if dg is None:
            if mydict is None:
                raise Exception('Invalid __init__() parameter for distribution group')
        if mydict is None:
            d = eaws.DetailDistributionGroup(dg)
        else:
            d = mydict
        self.login = d["Name"]
        self.enabled = d["IsValid"]
        self.restriction = d["MemberJoinRestriction"]
        self.email = d["PrimarySmtpAddress"]
        self.restrict_to_internal = d["RequireSenderAuthenticationEnabled"]
        self.name = d["Name"]
        self.manager = d["ManagedBy"]
        self.hidden = d["HiddenFromAddressListsEnabled"]
    @classmethod
    def from_dict(cls, mydict):
        return cls(None, mydict)
    def __unicode__(self):
        return self.name
    def delete(self):
        eaws.DisableDistributionGroup(self.name)
    def fdelete(self): 
        eaws.DeleteDistributionGroup(self.name)
    def set_manager(self, login):
        eaws.SetDistributionGroupOwner(self.name, login)
        self.manager = login
    def get_manager(self):
        return (self.manager.split('/'))[-1]
    def get_members(self):
        members = [ ]
        d = eaws.GetMembersOfDistributionGroup(self.name)
        for user in d:
            proper_email = user["PrimarySmtpAddress"]
            members.append(proper_email)
        return members
    def add_member(self, member):
        eaws.AddMemberToDistributionGroup(self.name, member, False)
    def add_external(self, member):
        eaws.AddMemberToDistributionGroup(self.name, member, True)
    def remove_member(self, member):
        eaws.RemoveMemberFromDistributionGroup(self.name, member)

class Mailbox(object):
    enabled = None
    name = None
    server = None
    login = None
    email = None
    size = None
    limit = None

    def __init__(self, login, mydict = None):
        if login is None:
            if mydict is None:
                raise Exception('Invalid __init__() parameter for mailbox')
        if mydict is None:
            d = eaws.DetailMailbox(login)
            self.size = d["TotalItemSize"]
            self.limit = d["DatabaseProhibitSendQuota"]
            self.login = login
            self.email = login
        else:
            d = mydict
            self.email = d["WindowsEmailAddress"]
            self.login = d["SamAccountName"]
        self.enabled = d["IsValid"]
        self.name = d["DisplayName"]
        self.server = d["ServerName"]
    def __unicode__(self):
        return self.name
    @classmethod
    def from_dict(cls, mydict):
        return cls(None, mydict)
    def get_fields(self):
        return self.__dict__
    def get_dg(self):
        result = [ ]
        list_of_dg = eaws.ListDistributionGroup()
        for dg in list_of_dg:
            d = eaws.GetMembersOfDistributionGroup(dg["SamAccountName"])
            for user in d:
                if self.login in user["EmailAddresses"]:
                    result.append(dg["SamAccountName"])
        return result
