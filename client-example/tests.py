from objects import DistributionGroup
import eaws

print "###Creating distribution group"
eaws.CreateDistributionGroup("test-create-dg")
print "###Retrieving distribution group"
dg = DistributionGroup("test-create-dg")
print dg
print dg.name
print "###Printing manager"
print "Expecting: Exchange Admin Web Service"
print "Result: " + dg.get_manager()
print "###Setting manager to myuser"
dg.set_manager("myuser")
print "###Refreshing DG data"
dg = DistributionGroup("test-create-dg")
print dg
print "###Printing manager"
print "Expecting: myuser"
print "Result: " + dg.get_manager()
print "###Adding myotheruser to group"
dg.add_member("myotheruser")
print "###Adding myuser to group"
dg.add_member("myuser")
print "###Adding myextuser@outlook.com to group"
dg.add_external("myextuser@outlook.com")
print "###Printing members list"
print "Expecting: myextuser@outlook.com"
print "Expecting: myotheruser@mail.contoso.com"
print "Expecting: myuser@mail.contoso.com"
for member in dg.get_members():
    print "Result: " + member
print "###Removing distribution group"
dg.fdelete()
