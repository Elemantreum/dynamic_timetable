import openpyxl as oxl
import datetime as dt
import excel2img
import telebot
from telebot import types
import config

def last_col(table):
    count = 0
    for x in table['A']:
        if x.value == "-":
            break
        else:
            count += 1
    return count

def last_row(table):
    count = 0
    for x in table['1']:
        if x.value == "-":
            break
        else:
            count += 1
    return chr(ord("A") + count - 1)

def send_to_tg():
    if config.mode == 'test':
        channel = '-1001835166654' # ИД Канала, где бот
        main_mes = 2 # Основное сообщение с caption
        sec_mes = 3 # Сообщение с расписанием на завтра (no caption)
    elif config.mode == 'main':
        channel = '-1001831144699' # ИД Канала, где бот
        main_mes = 15 # Основное сообщение с caption
        sec_mes = 16 # Сообщение с расписанием на завтра (no caption)


    bot = telebot.TeleBot(config.token, parse_mode=None)

    table = oxl.load_workbook('Изменения в расписании - 3.xlsx')
    today = dt.date.today()
    week_number = dt.date.today().weekday()
    dotw = {0: 'пн 1',
            1: 'вт 1',
            2: 'ср 1',
            3: 'чт 1',
            4: 'пт 1',
            5: 'сб1' ,
            6: 'пн 1',
            7: 'пн 1',
            8: 'вт 1'}


    today_table, tomorrow_table = table[dotw[week_number+config.after_17]], table[dotw[week_number+1+config.after_17]]

    # Расписание на сегодня
    excel2img.export_img("Изменения в расписании - 3.xlsx",
                         f"Расписание на {dt.date(today.year,today.month,today.day + config.after_17)}.png",
                         dotw[week_number+config.after_17],
                         f'A1:{last_row(today_table)}{last_col(today_table)}')

    # Расписание на завтра
    excel2img.export_img("Изменения в расписании - 3.xlsx",
                         f"Расписание на {dt.date(today.year,today.month,today.day + 1 + config.after_17)}.png",
                         dotw[week_number + 1 + config.after_17],
                         f'A1:{last_row(tomorrow_table)}{last_col(tomorrow_table)}')

    today_table_img = open(f"Расписание на {dt.date(today.year,today.month,today.day + config.after_17)}.png", 'rb')
    tomorrow_table_img = open(f"Расписание на {dt.date(today.year,today.month,today.day + 1 + config.after_17)}.png", 'rb')
    try:
        bot.edit_message_media(media=types.InputMedia(type='photo', media=today_table_img),
                               chat_id=channel,
                               message_id=main_mes)
    except Exception:
        pass
    bot.edit_message_caption(caption=f'*📅САМОЕ СВЕЖЕЕ РАСПИСАНИЕ* \n\nЗдесь расписание появляется быстрее, чем на первом этаже 😌 \n\nПост обновляется автоматически при поддержке Холявко А.Н. \n\n_Последнее обновление {dt.datetime.now().strftime("%d.%m.%Y, %H:%M")}_',
                          chat_id=channel,
                          message_id=main_mes,
                          parse_mode='markdown')
    try:
        bot.edit_message_media(media=types.InputMedia(type='photo', media=tomorrow_table_img),
                               chat_id=channel,
                               message_id=sec_mes)
    except Exception:
        pass


















