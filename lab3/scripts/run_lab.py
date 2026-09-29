import pexpect, io, time, json, re, sys
PROMPT='$ '
class Term:
    def __init__(s, title):
        s.title=title; s.log=io.StringIO()
        s.p=pexpect.spawn('bash',['--norc','--noprofile','--noediting'],encoding='utf-8',env={'PS1':PROMPT,'PATH':'/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin','TERM':'dumb','HOME':'/root'},dimensions=(50,200))
        s.p.expect_exact(PROMPT); s.p.logfile_read=s.log; s.log.write(PROMPT)
    def run(s, cmd, wait=True, t=15):
        s.p.sendline(cmd)
        if wait: s.p.expect_exact('\r\n'+PROMPT, timeout=t)
        else: time.sleep(0.8)
    def interact(s, cmd, pairs):
        s.p.sendline(cmd)
        for exp, ans in pairs:
            s.p.expect(exp); s.p.sendline(ans)
        s.p.expect_exact('\r\n'+PROMPT)
    def ctrl_c(s):
        time.sleep(1); s.p.sendcontrol('c'); s.p.expect_exact(PROMPT); 
    def text(s):
        t=s.log.getvalue().replace('\r\n','\n').replace('\r','')
        return t.rstrip()
out={}
def save(key, *terms): out[key]=[(t.title,t.text()) for t in terms]

part=sys.argv[1]
if part=='q12':
    sub=Term('Terminal 1 (subscribe)'); pub=Term('Terminal 2 (publish)')
    sub.run('mosquitto_sub -t /a/test -v', wait=False)
    pub.run('mosquitto_pub -t a/test -m test_q1'); time.sleep(1.5)
    pub.run('mosquitto_pub -t /a/test -m test_q1_compare'); time.sleep(1.5)
    sub.ctrl_c(); save('q1',sub,pub)
    subs=[Term(f'Subscriber {i+1}') for i in range(4)]
    topics=['+/roof/brightness/day','house1/firstfloor/#','house2/+/+/night','+/roof/#']
    for s,tp in zip(subs,topics): s.run(f"mosquitto_sub -v -t '{tp}'", wait=False)
    pub=Term('Publisher')
    pub.run('for p in house1 house2; do for f in firstfloor secondfloor roof; do for s in temp humid brightness; do for t in day night; do mosquitto_pub -t "$p/$f/$s/$t" -m "$p-$f-$s-$t"; done; done; done; done; echo "published 36 messages"', t=60)
    time.sleep(2)
    for s in subs: s.ctrl_c()
    save('q2',*subs,pub)
elif part=='q3':
    a=Term('Terminal 1 (設定 broker)')
    a.run('ls /etc/mosquitto')
    a.interact('sudo mosquitto_passwd -c /etc/mosquitto/passwd 112652006', [('Password:','123456'),('Reenter password:','123456')])
    a.run('ls /etc/mosquitto')
    a.run('sudo chown mosquitto:mosquitto /etc/mosquitto/passwd')
    a.run('sudo cat /etc/mosquitto/passwd')
    a.run("printf '\\nlistener 1883\\npassword_file /etc/mosquitto/passwd\\nallow_anonymous false\\n' | sudo tee -a /etc/mosquitto/mosquitto.conf > /dev/null")
    a.run('cat /etc/mosquitto/mosquitto.conf')
    import subprocess; subprocess.run('pkill -x mosquitto; sleep 1',shell=True)
    a.run('sudo service mosquitto restart'); time.sleep(1)
    s=Term('Terminal 2 (subscribe)'); p=Term('Terminal 3 (publish)')
    s.run('mosquitto_sub -v -t test')
    p.run('mosquitto_pub -t test -m hello')
    p.run('mosquitto_pub -t test -m hello -u 112652006 -P wrongpw')
    s.run('mosquitto_sub -v -t test -u 112652006 -P 123456', wait=False)
    p.run('mosquitto_pub -t test -m hello_from_112652006 -u 112652006 -P 123456'); time.sleep(1.5)
    s.ctrl_c()
    save('q3',a,s,p)
json.dump(out, open(f'{part}.json','w'), ensure_ascii=False, indent=1)
for k,v in out.items():
    for title,txt in v: print('=====',k,title); print(txt)
if part=='q4':
    import subprocess
    c=Term('Broker A (VM) 設定')
    c.run("printf '\\nlistener 1883\\nallow_anonymous true\\n\\nconnection bridge-01\\naddress 127.0.0.1:1884\\ntopic # out 1 out/ fromA/\\ntopic # in 1 fromB/\\n' | sudo tee -a /etc/mosquitto/mosquitto.conf > /dev/null")
    c.run('tail -8 /etc/mosquitto/mosquitto.conf')
    subprocess.run('pkill -x mosquitto; sleep 1; mosquitto -c b_pi.conf -d; sleep 0.5',shell=True)
    c.run('sudo service mosquitto restart'); time.sleep(2)
    As=Term('Broker A (VM) subscribe'); Ap=Term('Broker A (VM) publish')
    Bs=Term('Broker B (Pi) subscribe'); Bp=Term('Broker B (Pi) publish')
    As.run("mosquitto_sub -v -t '#'", wait=False); Bs.run("mosquitto_sub -h 127.0.0.1 -p 1884 -v -t '#'", wait=False); time.sleep(1)
    for tp,m in [('in','msg1'),('out','msg2'),('out/msg','msg3'),('/out','msg4')]:
        Ap.run(f'mosquitto_pub -t {tp} -m {m}'); time.sleep(1)
    for tp,m in [('out','msg5'),('fromB','msg6')]:
        Bp.run(f'mosquitto_pub -h 127.0.0.1 -p 1884 -t {tp} -m {m}'); time.sleep(1)
    time.sleep(1); As.ctrl_c(); Bs.ctrl_c()
    save('q4',c,As,Ap,Bs,Bp)
    json.dump(out, open('q4.json','w'), ensure_ascii=False, indent=1)
    for k,v in out.items():
        for title,txt in v: print('=====',k,title); print(txt)
