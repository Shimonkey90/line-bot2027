from flask import Flask, request, abort
import os
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

app = Flask(__name__)

CHANNEL_SECRET = os.environ["CHANNEL_SECRET"]
CHANNEL_ACCESS_TOKEN = os.environ["CHANNEL_ACCESS_TOKEN"]

# config.pyで設定したチャネルアクセストークン
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
# config.pyで設定したチャネルシークレット
handler = WebhookHandler(CHANNEL_SECRET)

@app.route('/')
def hello_world():
    return 'Hello, World!'

@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
 
    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)
 
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)
 
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    if event.reply_token == "00000000000000000000000000000000":
        return
    
    df = pd.read_excel("./test.xls")
    df = df.drop(columns=df.columns[[0, 2,3,6,7,8,9,11,13,14,15,16,17,18,19,20,21]])
    df['商品ｺｰﾄﾞ'] = df['商品ｺｰﾄﾞ'].str.strip()
    def search(code):
        for i in df['商品ｺｰﾄﾞ']:
          if i == code:
            hit = df[df['商品ｺｰﾄﾞ'] == i]
            item = hit['商品名'].values
            zaikosu = hit['現在庫'].values
            genka = hit['仕入原価'].values
            return item[0],zaikosu[0], genka[0]

    try:
        a,b,c = search(event.message.text)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text='{0}の在庫数は{1}。原価は{2}です。'.format(a,b,c))
    )
    except TypeError:
        TextSendMessage(text='入力に誤りがあります。')

if __name__ == "__main__":
    app.run()