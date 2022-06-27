# ProxyLogon
ProxyLogon is the formally generic name for CVE-2021-26855, a vulnerability on Microsoft Exchange Server that allows an attacker bypassing the authentication and impersonating as the admin. We have also chained this bug with another post-auth arbitrary-file-write vulnerability, CVE-2021-27065, to get code execution.


## ProxyLogon For Python3
```python
usage:
    sudo python3 proxylogon.py --host=target.com --mail=admin@target.com
    sudo python3 proxylogon.py --host=target.com --mails=./mails.txt
```

#### Tips:
* 1) recon target to find valid email address
* 2) if you do not find any email, use bruteforce target with your email file.


#### references
* [proxylogon](https://proxylogon.com/)
