import email.message,smtplib,msal,requests,random,os,time
from dotenv import load_dotenv
load_dotenv()

msg=email.message.EmailMessage()
subject_list=["特急件！","特別篇預算","絕佳預算","重要會議通知","調查表","帳戶權益通知","包裹發送通知","產品即將推出","重要更新","帳單到期通知","滿意度調查"]
object_list=["您的帳單已經到期！請及時付款以避免延遲費用！","我們期待著您的回饋和評論！","我們已經準備好與您簽署合同了！","關於新產品的最新更新！","請填寫滿意度調查表，幫助我們了解您對我們的服務的感受。","我們期待著您的回饋和評論！","特別通知：我們的辦公時間已更改！","您的訂單已確認，現在處理中！","您的包裹已發貨！請查收！","您的帳戶安全有問題嗎？立即解決！","重要通知：我們即將更新我們的隱私政策。","我們將在下星期一舉行一個重要會議。"]
msg["From"]=os.getenv("FROM_EMAIL")
msg["To"]=os.getenv("TO_EMAIL")
msg["Subject"]=random.choice(subject_list)
msg.add_alternative("<h3>HTML內容</h3>"+random.choice(object_list),subtype="html") #HTML信件內容

server=smtplib.SMTP_SSL("smtp.gmail.com",465) #建立gmail連驗
server.login(os.getenv("FROM_EMAIL"),os.getenv("GOOGLE_PASSWORD"))
server.send_message(msg)
server.close() #發送完成後關閉連線
print("發送成功")

for i in range(0,random.randint(1,10)):
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    tenant_id = os.getenv("TENANT_ID")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority)

    scopes = ["https://graph.microsoft.com/.default"]
    result = None
    result = app.acquire_token_silent(scopes, account=None)

    if not result:
        print(
            "No suitable token exists in cache. Let's get a new one from Azure Active Directory.")
        result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        userId = os.getenv("USER_ID")
        endpoint = f'https://graph.microsoft.com/v1.0/users/{userId}/messages?$select=sender,subject'
        r = requests.get(endpoint,
                        headers={'Authorization': 'Bearer ' + result['access_token']})
        if r.ok:
            print('Retrieved emails successfully')
            data = r.json()
            for email in data['value']:
                print(email['subject'] + ' (' + email['sender']
                    ['emailAddress']['name'] + ')')
        else:
            print(r.json())
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))
    time.sleep(1)
