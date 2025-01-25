# Xcells

A simple, yet powerful library for working with Excel files in Python.

Простая, но мощная библиотека для работы с файлами Excel на Python.

## Key Features / Основные возможности

- **Read Excel files**: Extract data from `.xlsx` files easily.
  
  **Чтение файлов Excel**: Легко извлекайте данные из файлов `.xlsx`.

- **Access worksheets**: Get structured data for every sheet in the workbook.
  
  **Доступ к листам**: Получите структурированные данные для каждого листа книги.

- **Read individual cells**: Retrieve cell data using row and column references.
  
  **Чтение отдельных ячеек**: Извлекайте данные ячеек с помощью ссылок на строки и столбцы.

- **Designed for Developers**: Simple and intuitive API to streamline your workflow.
  
  **Разработано для разработчиков**: Простой и интуитивно понятный API для упрощения работы.

## Installation / Установка

The library is not yet published to PyPI. To use it, clone the repository:

Библиотека еще не опубликована в PyPI. Чтобы использовать её, клонируйте репозиторий:

```bash
git clone https://github.com/horhe-dvlp/xcells.git
cd xcells
pip install .
```

## Getting Started / Как начать

### Reading a Workbook / Чтение книги Excel

```python
from xcells.core.reader import Reader

# Initialize the Reader / Инициализация Reader
reader = Reader()

# Read an Excel file / Чтение Excel файла
workbook = reader.read("Book1.xlsx")

# Access sheets / Доступ к листам
sheets = workbook.get_worksheets()

for sheet in sheets:
    for row in sheet.cells:
        for cell in row: 
            print(cell.value)
```

## Why Use This Library? / Почему стоит использовать эту библиотеку?

- **Efficiency**: Focused on essential features without unnecessary overhead.
  
  **Эффективность**: Сосредоточена на основных функциях без лишнего.

- **Open-source**: MIT-licensed and free for all use cases.
  
  **Открытый исходный код**: Лицензия MIT, свободно для любых случаев использования.

- **Built for Developers**: Clean and understandable code structure designed for extensibility.
  
  **Создана для разработчиков**: Чистая и понятная структура кода, созданная для расширяемости.

## Contributing / Вклад

We welcome contributions! Please follow these steps to contribute:

Мы приветствуем вклад сообщества! Следуйте этим шагам, чтобы внести изменения:

1. Fork the repository.
   
   Склонируйте репозиторий.
2. Create a new branch for your feature or bug fix.
   
   Создайте новую ветку для функции или исправления ошибки.
3. Write tests to ensure your changes are robust.
   
   Напишите тесты, чтобы убедиться, что изменения устойчивы.
4. Submit a pull request with a detailed explanation of your changes.
   
   Отправьте pull request с подробным объяснением ваших изменений.

## License / Лицензия

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

Этот проект распространяется под лицензией MIT. Подробности смотрите в файле [LICENSE](LICENSE).

## Contact / Контакты

For questions, suggestions, or feedback, feel free to contact me:

По вопросам, предложениям или отзывам, свяжитесь со мной:

- **Email**: your-email@example.com
  
  **Электронная почта**: your-email@example.com
- **GitHub**: [horhe-dvlp](https://github.com/horhe-dvlp)
  

