import asyncio
import telegram

async def sendTelegramMsg(msg,chat_id):
    bot = telegram.Bot("6557815543:AAFYQ1itku7fMSbBFirk3YQIQRtNE54WbEM")
    async with bot:
        await bot.send_message(text=msg, chat_id=chat_id)
        #print((await bot.get_updates())[4])
        #print(await bot.get_me())
        
async def sendTelegramMsgWithDocuments(chat_id):
    bot = telegram.Bot("6557815543:AAFYQ1itku7fMSbBFirk3YQIQRtNE54WbEM")
    async with bot:        
        await bot.send_document(chat_id=chat_id, document='C:\\BOT INSUMO BASE\\Input BOT 001 V1.xlsx')
        #await bot.send_document(chat_id=chat_id, document='logs/super_log.txt')
        #await bot.send_media_group(chat_id=chat_id, media=['C:\\BOT INSUMO BASE\\Input BOT 001 V1.xlsx','logs/super_log.txt'])
        #print((await bot.get_updates())[4])
        #print(await bot.get_me())        
        
# Generic Functions        
# async def main():
#     bot = telegram.Bot("6557815543:AAFYQ1itku7fMSbBFirk3YQIQRtNE54WbEM")
#     async with bot:
#         #print((await bot.get_updates())[4])
#         #5970685607 personal chat ID
#         await bot.send_message(text='Hola bienvenidos a las notificaciones RPA para el BPO!', chat_id=-1002019721248)
#         #print(await bot.get_me())

# if __name__ == '__main__':
#     asyncio.run(main())