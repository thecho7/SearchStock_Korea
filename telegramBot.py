# Done! Congratulations on your new bot.
# You will find it at t.me/ChorrenBuffettBot. You can now add a description, about section and profile picture for your bot, see /help for a list of commands. 
# By the way, when you've finished creating your cool bot, ping our Bot Support if you want a better username for it. Just make sure the bot is fully operational before you do this.

# Use this token to access the HTTP API:
# 1468902047:AAF30AYxDuZM1YeiRTqAFZEhptIJ0Bv9CJ0
# Keep your token secure and store it safely, it can be used by anyone to control your bot.
# thecho7 Chat id: 487955608
# 추월차선 chat id: -1001372378459

import telegram

class telegramMachine():
    def __init__(self, telgm_token, chat_id):
        print('Start Telegram Chat Bot')
        self.bot = telegram.Bot(token = telgm_token)
        self.chat_id = chat_id

    def updateBot(self):
        updates = self.bot.getUpdates()

    def sendMsg(self, msg):
        self.bot.sendMessage(chat_id = 'YOUR_CHAT_ID', text=msg)
        # self.updateBot()

    def sendImage(self, img_path):
        self.bot.send_photo(chat_id = 'YOUR_CHAT_ID', photo=open(img_path, 'rb'))