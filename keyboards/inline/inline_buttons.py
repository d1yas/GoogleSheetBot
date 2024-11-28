from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton


student_list = InlineKeyboardMarkup(
    inline_keyboard=[
        [
            InlineKeyboardButton(text="Shamsiddin",callback_data="shamsiddin_call"),
        ],
        [
            InlineKeyboardButton(text="Alisher", callback_data="alisher_call"),
        ],
        [
            InlineKeyboardButton(text="Baxtiyor", callback_data="baxtiyor_call"),
        ],
        [
            InlineKeyboardButton(text="Diyas", callback_data="diyas_call"),
        ],
        [
            InlineKeyboardButton(text="Behruz", callback_data="behruz_call"),
        ],
        [
            InlineKeyboardButton(text="Sardor", callback_data="sardor_call"),
        ],
        [
            InlineKeyboardButton(text="Sarvar", callback_data="sarvar_call"),
        ],
    ]
)