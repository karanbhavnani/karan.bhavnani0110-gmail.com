
import requests, threading, queue

def is_tally(ip,port,timeout=0.5):
    try:
        xml="<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC><REPORTNAME>List of Accounts</REPORTNAME></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"
        r = requests.post(f"http://{ip}:{port}",data=xml,timeout=timeout)
        return r.status_code==200 and "<ENVELOPE>" in r.text
    except:
        return False

def scan_lan_for_tally(port=9000,base="192.168.1."):
    q=queue.Queue()
    def worker(ip):
        if is_tally(ip,port):
            q.put(ip)
    threads=[]
    for i in range(1,255):
        ip=f"{base}{i}"
        t=threading.Thread(target=worker,args=(ip,))
        t.start()
        threads.append(t)
    for t in threads:
        t.join()
    results=[]
    while not q.empty():
        results.append(q.get())
    return results
