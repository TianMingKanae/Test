import copy, docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
SRC='/root/.claude/uploads/de2f076b-18ee-50d2-8724-46a3b7c13fa0/bda20359-Lab_3_report.docx'
d=docx.Document(SRC)
BLUE=RGBColor(0x1F,0x4E,0x79)
def font(run, code=False, bold=False, color=BLUE, size=None):
    rpr=run._r.get_or_add_rPr(); f=rpr.find(qn('w:rFonts'))
    if f is None: f=OxmlElement('w:rFonts'); rpr.insert(0,f)
    a='Courier New' if code else 'Times New Roman'
    for k,v in [('w:ascii',a),('w:hAnsi',a),('w:cs',a),('w:eastAsia','標楷體')]: f.set(qn(k),v)
    run.bold=bold
    if color: run.font.color.rgb=color
    if size: run.font.size=Pt(size)
def new_par_after(p):
    np=OxmlElement('w:p'); p._p.addnext(np); return Paragraph(np,p._parent)
class Writer:
    def __init__(s,anchor): s.cur=anchor
    def par(s, parts, indent=0, align=None, space_after=4):
        p=new_par_after(s.cur); s.cur=p
        if indent: p.paragraph_format.left_indent=Inches(indent)
        p.paragraph_format.space_after=Pt(space_after)
        if align: p.alignment=align
        if isinstance(parts,str): parts=[parts]
        for part in parts:
            if isinstance(part,str): part=(part,{})
            t,o=part; r=p.add_run(t); font(r,**o)
        return p
    def img(s,path,w=5.75,caption=None):
        p=new_par_after(s.cur); s.cur=p; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path,width=Inches(w))
        if caption: s.par([(caption,{'size':10,'color':RGBColor(0x59,0x59,0x59)})],align=WD_ALIGN_PARAGRAPH.CENTER,space_after=8)
C={'code':True}; B={'bold':True}
body=d.element.body
def paras(): return d.paragraphs
def find(text):
    for p in paras():
        if p.text.strip().startswith(text): return p
    raise KeyError(text)
def clear_blanks_after(p):
    # remove consecutive empty paragraphs following p (keep none)
    nxt=p._p.getnext()
    while nxt is not None and nxt.tag==qn('w:p') and not ''.join(t.text or '' for t in nxt.iter(qn('w:t'))).strip() and not list(nxt.iter(qn('w:drawing'))):
        n2=nxt.getnext(); body.remove(nxt); nxt=n2
def blank(w): w.par('',space_after=0)

# ---------- header
_ts=list(d.paragraphs[0]._p.iter(qn('w:t')))
for a_,b_ in zip(_ts,_ts[1:]):
    if a_.text=='學號' and b_.text==':_____': b_.text=': 112652006'
    if a_.text=='姓名' and b_.text==':_____': b_.text=': 陳昱翰'
# ---------- Q1
a=find('當訂閱的topic為 /a/test'); clear_blanks_after(a); w=Writer(a)
w.par([('Ans：不能。',B)])
w.par(['MQTT 的 topic 以「/」作為階層分隔符號，開頭的「/」本身也會切出一個「空字串」的層級。因此 ',('/a/test',C),' 共有 3 層（""、"a"、"test"），而 ',('a/test',C),' 只有 2 層（"a"、"test"）。訂閱的 topic 沒有使用萬用字元（+、#）時，必須每一層都完全相同才算匹配；兩者層數不同、第一層也不同（"" ≠ "a"），所以訂閱 ',('/a/test',C),' 的 subscriber 收不到發佈到 ',('a/test',C),' 的訊息。'])
w.par(['實驗驗證：Terminal 1 訂閱 ',('/a/test',C),'，Terminal 2 先 publish 到 ',('a/test',C),'（test_q1），subscriber 沒有任何輸出；再 publish 到 ',('/a/test',C),'（test_q1_compare）作對照，subscriber 才收到訊息。'])
w.img('img_q1.png',caption='圖 1　Q1 測試：只有 /a/test 的訊息被收到，a/test 的 test_q1 沒有出現')
blank(w)

# ---------- Q2
a=find('所有roof的所有內容'); clear_blanks_after(a); w=Writer(a)
w.par([('Ans：',B)])
for i,(q,t,alt,n) in enumerate([
    ('所有 roof 在 day 的 brightness','+/roof/brightness/day',None,2),
    ('house1 firstfloor 的所有內容','house1/firstfloor/#','house1/firstfloor/+/+',6),
    ('house2 在 night 的所有內容','house2/+/+/night',None,9),
    ('所有 roof 的所有內容','+/roof/#','+/roof/+/+',12)],1):
    parts=[f'({i}) {q}：',(t,{'code':True,'bold':True})]
    if alt: parts+=['（亦可寫成 ',(alt,C),'）']
    parts+=[f'　→ 實測收到 {n} 則']
    w.par(parts,indent=0.2)
w.par(['說明：「+」只能代表「單一」階層，可放在任何位置；「#」代表其後的任意多個階層，只能放在最後一層。第 (1)(3) 題要固定某幾層、其他單層任意，所以用「+」；第 (2)(4) 題要某層以下的所有內容，所以在結尾用「#」。將 2×3×3×2 = 36 種 topic 全部 publish 一次，四個 subscriber 分別只收到 2、6、9、12 則，與預期相符。'])
w.par(['補充：若題目格式中最後的「/」代表實際 topic 結尾真的有一個「/」（例如 ',('house1/roof/brightness/day/',C),'），則 (1)(3) 需寫成 ',('+/roof/brightness/day/',C),'、',('house2/+/+/night/',C),'（多一個空的最後層），(2)(4) 用「#」的寫法不受影響。'])
w.img('img_q2.png',caption='圖 2　Q2 測試：四個 subscriber 同時訂閱，publisher 送出全部 36 種 topic')
blank(w)

# ---------- Q3
a=find('排除匿名使用者截圖'); clear_blanks_after(a); w=Writer(a)
w.par([('步驟：',B)])
for s in [['用 ',('sudo mosquitto_passwd -c /etc/mosquitto/passwd <帳號>',C),' 建立帳號（本次帳號設為學號 112652006，密碼 123456；-c 會新建 passwd 檔），建立後 /etc/mosquitto 下多了 passwd，內容為「帳號:雜湊後的密碼」。'],
          ['在 ',('/etc/mosquitto/mosquitto.conf',C),' 加入 ',('password_file /etc/mosquitto/passwd',C),' 與 ',('allow_anonymous false',C),'，並 ',('sudo service mosquitto restart',C),'。'],
          ['注意：mosquitto 2.x 若設定檔中沒有任何 listener，會進入 local only mode，此時 allow_anonymous false 不會生效，因此要另外加上 ',('listener 1883',C),'。'],
          ['結果：未帶帳密（匿名）或密碼錯誤的 sub / pub 都被拒絕（Connection Refused: not authorised）；帶正確的 ',('-u',C),' / ',('-P',C),' 後即可正常訂閱與發佈。']]:
    w.par(s,indent=0.2)
w.img('img_q3a.png',caption='圖 3　建立帳號、檢視 passwd、修改 mosquitto.conf 並重啟 broker')
w.img('img_q3b.png',caption='圖 4　匿名 / 錯誤密碼被拒；使用帳號密碼後 subscriber 成功收到訊息')
blank(w)

# ---------- Q4 table
tbl=d.tables[0]
res=[('in　msg1','（收不到）'),('out　msg2','fromA/out　msg2'),('out/msg　msg3','fromA/msg　msg3'),
     ('/out　msg4','（收不到）'),('fromB/out　msg5','out　msg5'),('fromB/fromB　msg6','fromB　msg6')]
for row,(ra,rb) in zip(tbl.rows[1:],res):
    for cell,val in zip(row.cells[1:],(ra,rb)):
        p=cell.paragraphs[0]; r=p.add_run(val); font(r,code=not val.startswith('（'),size=11)
w=Writer(Paragraph(tbl._tbl,d._body))
# insert after table: need paragraph-like anchor
anchor=OxmlElement('w:p'); tbl._tbl.addnext(anchor); w=Writer(Paragraph(anchor,d._body))
w.par([('Bridge 設定（設在 Broker A / VM）：',B)])
for l in ['connection bridge-01','address <Pi板 IP>:1883','topic # out 1 out/ fromA/','topic # in 1 fromB/']:
    w.par([(l,C)],indent=0.3,space_after=0)
blank(w)
w.par([('結果說明：',B)])
for s in [['out 規則：local-prefix 為 out/、remote-prefix 為 fromA/，A 會在本地訂閱 ',('out/#',C),'，符合的訊息轉送給 B，並把開頭的 out/ 換成 fromA/。所以 msg3（',('out/msg',C),'）在 B 顯示為 ',('fromA/msg',C),'。'],
          ['msg2 的 topic 是 ',('out',C),'：MQTT 的「#」也會匹配它的上一層本身，所以 ',('out',C),' 符合 ',('out/#',C),' 而被轉送；但它並不是以 "out/" 開頭，沒有前綴可以拿掉，於是只在前面加上 fromA/，B 看到 ',('fromA/out',C),'。'],
          ['msg1（',('in',C),'）與 msg4（',('/out',C),'，第一層是空字串而不是 out）都不符合 out/#，只留在 A，B 收不到。'],
          ['in 規則：只有 local-prefix fromB/，沒有 remote-prefix，所以 A 在 B 上訂閱「#」，B 上所有訊息都會被帶回 A，並在 topic 前面加上 fromB/：msg5 → ',('fromB/out',C),'、msg6 → ',('fromB/fromB',C),'。B 自己的 subscriber 則照原 topic 顯示。'],
          ['B 上的 fromA/… 訊息不會再被 in 規則帶回 A 形成迴圈，因為 mosquitto 的 bridge 連線會告知遠端 broker 不要把 bridge 自己送出的訊息再回送（no-local / try_private）。'],
          ['A 的 subscriber 訂閱「#」，本地 publish 的訊息都會原樣顯示（out 規則只影響送往 B 的訊息，不影響 A 本地）。']]:
    w.par(s,indent=0.2)
w.img('img_q4a.png',caption='圖 5　在 Broker A 的 mosquitto.conf 加入 bridge 設定並重啟')
w.img('img_q4b.png',caption='圖 6　兩邊 client 皆訂閱 #，依表格順序 publish 的結果')
w.par([('註：此次截圖中的 Broker B 以同一台機器上 port 1884 的 mosquitto 代替 Pi 板（位址 127.0.0.1:1884），bridge 行為與連到 Pi 板相同。',{'size':10,'color':RGBColor(0x59,0x59,0x59)})])
# remove leftover blanks after (anchor chain) up to Q5
clear_blanks_after(w.cur); blank(w)
body.remove(anchor)

# ---------- Q5
a=find('在MQTT bridge configuration中，QoS level'); clear_blanks_after(a); w=Writer(a)
w.par([('Ans：',B),'QoS 有 0、1、2 三種等級，數字越高傳遞保證越強，但交握次數與負擔也越大。bridge 設定 ',('topic <pattern> <direction> <qos>',C),' 中的 qos，是兩個 broker 之間轉送該 topic 時使用的 QoS。'])
for s in [[('QoS 0 – At most once（最多一次）：',B),'送出 PUBLISH 後就不管了，接收端不回覆確認，也不重送。速度最快、負擔最小，但網路不穩時訊息可能遺失。適合頻繁更新、掉一筆也無所謂的感測資料。'],
          [('QoS 1 – At least once（至少一次）：',B),'接收端收到後回 PUBACK；傳送端在收到 PUBACK 前會保存訊息，逾時就重送（DUP 旗標）。保證送達，但可能因重送而收到重複訊息，接收端需能容忍重複。'],
          [('QoS 2 – Exactly once（剛好一次）：',B),'以 PUBLISH → PUBREC → PUBREL → PUBCOMP 四次交握，雙方用 Packet ID 記錄狀態，保證訊息只被處理一次、不遺失也不重複。最可靠但延遲與負擔最大，適合計費、控制指令等不能重複的訊息。']]:
    w.par(s,indent=0.2)
w.par(['實際送達的 QoS 會取「發佈端的 QoS」與「訂閱（或 bridge 設定）的 QoS」兩者中較低者。本實驗 bridge 設 QoS 1，代表兩個 broker 之間的轉送保證至少送達一次。'])
blank(w)

# ---------- Q6
a=find('什麼是"保留消息'); clear_blanks_after(a); w=Writer(a)
w.par([('Ans：',B)])
w.par([('(1) true / false 的作用',B)])
w.par([('cleansession true：',B),'bridge 以「乾淨的 session」連線到遠端 broker。連線中斷時，遠端 broker 會清除這個 bridge 的所有訂閱以及尚未送出的訊息；重新連線時一切重新開始，bridge 重新訂閱，斷線期間的訊息會遺失。'],indent=0.2)
w.par([('cleansession false（預設）：',B),'使用「持久 session」。遠端 broker 以 client id 記住 bridge 的訂閱，斷線時仍保留，並替它暫存斷線期間 QoS 1/2 的訊息；重連後繼續沿用原訂閱並補送這些訊息，不容易掉資料。'],indent=0.2)
w.par([('(2) cleansession 為 false 時更改訂閱主題的意外行為',B)])
w.par(['現象：修改 bridge 的 topic 設定（例如把 ',('topic sensor/# in',C),' 改成 ',('topic light/# in',C),'）並重連後，除了新的主題，仍會繼續收到「舊主題」的訊息，而且在設定檔裡已經找不到這條規則。'],indent=0.2)
w.par(['原因：session 是保存在「遠端 broker」上的。false 時重連會沿用舊 session，遠端 broker 仍記得舊的訂閱；bridge 重連時只會送出新設定中的 SUBSCRIBE，並不知道要去 UNSUBSCRIBE 已經從設定檔刪掉的舊主題，所以新舊訂閱同時存在。'],indent=0.2)
w.par(['解決：先把 cleansession 設為 true 並重啟讓 bridge 重新連線，遠端就會清掉舊 session（含舊訂閱）；確認之後再改回 false、重啟一次，恢復正常的持久 session。'],indent=0.2)
w.par([('(3) 保留訊息（retained messages）與大量重送',B)])
w.par(['保留訊息：publish 時設定 retain 旗標（如 ',('mosquitto_pub -r',C),'），broker 會為該 topic 保存「最後一筆」保留訊息；之後任何 client 一訂閱到符合的 topic，broker 就立刻把這筆訊息送給它，讓新訂閱者不用等下一次更新就能拿到最新狀態（例如裝置目前的溫度、開關狀態）。'],indent=0.2)
w.par(['為何 true 時會大量發送：cleansession 為 true 時，每次斷線後遠端都會把訂閱清掉，所以 bridge 每次重連都必須「重新訂閱」。而每一次新的訂閱都會觸發 broker 把所有符合主題的保留訊息整批送出；bridge 常用「#」這類範圍很大的萬用字元，涵蓋的保留訊息可能非常多，網路不穩頻繁重連時，就會一次又一次收到大量重複的保留訊息。false 時訂閱一直保留、不需重新訂閱，就不會發生這種情況。'],indent=0.2)
blank(w)

# ---------- Q7
a=find('Q7.心得'); w=Writer(a)
w.par(['這次實驗實際操作了 MQTT 的 publish / subscribe、帳號驗證與 bridge。Q1 讓我注意到 topic 開頭多一個「/」就等於多了一層空字串，看起來很像的 /a/test 和 a/test 其實完全不同；Q2 則熟悉了「+」只代表單一層、「#」代表之後所有層的差別，用實際發送 36 種 topic 驗證每個 subscriber 收到的筆數，比只在紙上推比較有把握。'])
w.par(['設定帳號密碼時遇到一個坑：新版 mosquitto（2.x）如果設定檔沒有寫 listener，會進入 local only mode，allow_anonymous false 不會生效，匿名使用者照樣可以連線，加上 listener 1883 之後才正確擋下。另外 passwd 檔不存在或 broker 沒有權限讀取時，mosquitto 會直接啟動失敗，client 只看到 Connection refused，要去看 log 才知道原因。'])
w.par(['Bridge 的部分最有收穫。out/in 方向搭配 local-prefix、remote-prefix 可以讓兩個 broker 之間只分享部分主題並重新命名，例如 out/msg 到 B 變成 fromA/msg、B 的所有訊息到 A 都加上 fromB/。其中 topic 為 out 的訊息也會被轉送成 fromA/out，是因為「#」也會匹配上一層本身，這是實際測試後才發現的細節。整體而言，MQTT 架構簡單、傳輸量小，透過 broker 讓發佈者與訂閱者互不需要知道對方位址，很適合 IoT 裝置之間的溝通。'])

d.save('Lab_3_report_filled.docx'); print('saved')
