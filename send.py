import msal
import requests
import random

subject_list=["特急件！","特別篇預算","絕佳預算","重要會議通知","調查表","帳戶權益通知","包裹發送通知","產品即將推出","重要更新","帳單到期通知","滿意度調查"]
object_list=["您的帳單已經到期！請及時付款以避免延遲費用！","我們期待著您的回饋和評論！","我們已經準備好與您簽署合同了！","關於新產品的最新更新！","請填寫滿意度調查表，幫助我們了解您對我們的服務的感受。","我們期待著您的回饋和評論！","特別通知：我們的辦公時間已更改！","您的訂單已確認，現在處理中！","您的包裹已發貨！請查收！","您的帳戶安全有問題嗎？立即解決！","重要通知：我們即將更新我們的隱私政策。","我們將在下星期一舉行一個重要會議。"]

client_id = 'bbd4abc6-a81c-4539-8dfe-c1ca7ce92fe5'
client_secret = 'FlW8Q~qJN3cASbI5tdtWe9HSH32o4w91ytdpvdzC'
tenant_id = '14c677a8-d0d7-4701-97a9-176e66472585'
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

# if "access_token" in result:
#     print("Access token is " + result["access_token"])


if "access_token" in result:
    userId = "6f45e16f-da48-486e-b7fa-cea641555e01"
    endpoint = f'https://graph.microsoft.com/v1.0/users/{userId}/sendMail'
    toUserEmail = "tim20060112@gmail.com"
    email_msg = {'Message': {'Subject': random.choice(subject_list),
                             'Body': {'ContentType': 'Text', 'Content': random.choice(object_list)},
                             'ToRecipients': [{'EmailAddress': {'Address': toUserEmail}}]
                             },
                 'SaveToSentItems': 'true'}
    r = requests.post(endpoint,
                      headers={'Authorization': 'Bearer ' + result['access_token']}, json=email_msg)
    if r.ok:
        print('Sent email successfully')
    else:
        print(r.json())
else:
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))