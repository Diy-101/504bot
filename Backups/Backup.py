import os
import pandas as pd
from aiogram import Bot, types
from aiogram.dispatcher.dispatcher import Dispatcher as AiogramDispatcher  # Импорт Dispatcher из aiogram.dispatcher.dispatcher

# Замените 'your_bot_token' на токен вашего бота
API_TOKEN = '6930270526:AAFvENZg4Md_jQRp48NhcJEfZVSbt3jSb0w'

bot = Bot(token=API_TOKEN)
dp = AiogramDispatcher(bot)  # Используйте AiogramDispatcher вместо обычного Dispatcher

# Замените 'your_excel_file.xlsx' на путь к вашему файлу Excel
excel_file_path = r'C:\Users\abuti\Bot\Tables\ТableOne.xlsx'
df = pd.read_excel(excel_file_path, header=None)

def get_data(class_name, day_of_week):
    # Найти соответствующую ячейку в таблице
    class_col_index = df.iloc[:, 0][df.iloc[:, 0] == class_name].index.tolist()
    day_row_index = df.iloc[0][df.iloc[0] == day_of_week].index.tolist()

    if not class_col_index or not day_row_index:
        return "Класс или день недели не найдены в таблице."

    class_col_index = class_col_index[0] + 1
    day_row_index = day_row_index[0] + 1

    # Получить значение из ячейки
    cell_value = df.iat[class_col_index - 1, day_row_index - 1]

    # Проверить, является ли значение числом
    if pd.api.types.is_number(cell_value):
        # Если значение является числом, преобразуем его в строку
        cell_value = str(cell_value)

    # Форматировать данные для вывода
    formatted_data = f"Расписание:\n{cell_value}"

    return formatted_data

@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message):
    await message.answer("Привет! Я бот для работы с данными из Excel таблицы. Попробуйте команду /get_data.")

@dp.message_handler(commands=['get_data'])
async def cmd_get_data(message: types.Message):
    await message.answer("Введите класс и день недели через пробел (например, 9В Понедельник):")
    # Устанавливаем следующий обработчик, который будет ожидать ввод класса и дня недели
    dp.current_state(user=message.from_user.id, chat=message.chat.id, state="*")
    await message.answer("Введите класс и день недели через пробел (например, 9В Понедельник):")

@dp.message_handler(state="*")
async def process_data_input(message: types.Message):
    user_input = message.text

    if not user_input:
        await message.answer("Введите класс и день недели через пробел (например, 9В Понедельник):")
        return

    # Разбиваем введенные данные
    class_name, day_of_week = user_input.split(maxsplit=1)

    # Получаем расписание для выбранного класса и дня недели
    result = get_data(class_name, day_of_week)

    # Отправляем результат пользователю
    await message.answer(result)

    # Сбрасываем состояние
    await dp.get_current().reset_state()

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
