Exchange Admin Web Service
====

Any Exchange Server administrator in a mixed environments knows how terrible Powershell is at being bridged with other OS.
Most solutions require custom SSH server for Windows that are either expensive or do not include the features we need.
This project allows anyone to integrate common Exchange Server administration tasks into their usual user lifecycle tools.

Requirements
====

- Exchange server 2013 (this might work with 2010)
- IIS server with powershell support and .Net 4
- A webservice user with enough rights to perform actions on Exchange, Active Directory
 
Usage
====

Create a new website on IIS in a new application pool, both of which should be owned by your webservice user. Then you should be able to access it at http(s)://iis.contoso.com/EAWS/WebService.svc?wsdl.
An example of client implementation in python using python-suds is included.

Notes
====

This isn't limited to Exchange, you could perform anything Powershell related.


Disclaimer
====

This project might not be suitable for your needs or environment. I recommand you read the code thoroughly (it's short) and to understand the code makes assumption at times based on what my needs were at the time such as the way the information is serialized/deserialized. It is also very likely it isn't optimized.
