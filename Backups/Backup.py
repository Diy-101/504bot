import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram import executor
import logging

# Устанавливаем уровень логирования
logging.basicConfig(level=logging.INFO)

# Замените 'your_bot_token' на токен вашего бота
API_TOKEN = ''

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

# Замените 'your_excel_file.xlsx' на путь к вашему файлу Excel
excel_file_path = r'C:\Users\abuti\Bot\Tables\ТableOne.xlsx'
df = pd.read_excel(excel_file_path, header=None)

def get_data(user_input):
    # Разбиваем строку пользователя на класс и день недели
    class_name, day_of_week = user_input.split(maxsplit=1)
    
    # Ищем класс в столбце A (начиная со второй ячейки и до первой пустой)
    class_col_index = df.iloc[:, 0][df.iloc[:, 0] == class_name].index.tolist()
    
    # Если класс не найден, возвращаем сообщение с информацией о ячейке
    if not class_col_index:
        return f"Класс не найден. Поиск производился в ячейке ({df.shape[0] + 1}, 1)."
    
    class_col_index = class_col_index[0] + 1  # Добавляем 1, так как индексы в DataFrame начинаются с 0
    
    # Ищем день недели в строке 1 (начиная со второй ячейки и до первой пустой)
    day_row_index = df.iloc[0][df.iloc[0] == day_of_week].index.tolist()
    
    # Если день недели не найден, возвращаем сообщение с информацией о ячейке
    if not day_row_index:
        return f"День недели не найден. Поиск производился в ячейке (1, {df.shape[1] + 1})."
    
    day_row_index = day_row_index[0] + 1  # Добавляем 1, так как индексы в DataFrame начинаются с 0
    
    # Находим значение в пересекающейся ячейке
    cell_value = df.iat[class_col_index - 1, day_row_index - 1]
    
    # Проверяем, является ли значение числом
    if pd.api.types.is_number(cell_value):
        # Если значение является числом, преобразуем его в строку
        cell_value = str(cell_value)
    
    # Форматируем данные для вывода
    formatted_data = f"Расписание:\n{cell_value}"
    
    return formatted_data

# Обработчик команды /start
@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message):
    await message.answer("Привет! Я бот для работы с данными из Excel таблицы. Попробуйте команду /get_data.")

# Заменяем cmd_get_data на get_data
@dp.message_handler(commands=['get_data'])
async def get_data_command(message: types.Message):
    user_input = message.get_args()
    
    if not user_input:
        await message.answer("Введите класс и день недели после команды /get_data, например: /get_data 9В Понедельник")
        return
    
    result = get_data(user_input)
    await message.answer(result)

# Обработчик неизвестных команд
@dp.message_handler(lambda message: message.text.startswith('/'))
async def cmd_unknown(message: types.Message):
    await message.answer("Извините, я не понимаю эту команду.")

# Обработчик текстовых сообщений
@dp.message_handler(lambda message: not message.text.startswith('/'))
async def handle_text(message: types.Message):
    await message.answer("Извините, я понимаю только команды. Попробуйте /start или /get_data.")

if __name__ == '__main__':
    from aiogram import executor

    executor.start_polling(dp, skip_updates=True)