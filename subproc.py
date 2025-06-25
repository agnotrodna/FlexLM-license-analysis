import re
from collections import defaultdict
from datetime import datetime
import os



def parse_flexlm_log(log_content):
    # Паттерны для извлечения данных
    date_patterns = [
        r'TIMESTAMP (\d{1,2}/\d{1,2}/\d{4})',  # Формат 5/21/2025
        r'Time: (\w{3} \w{3} \d{1,2} \d{4})',  # Формат Mon Jun 9 2025
    ]

    # Словари для хранения данных
    date_entries = defaultdict(list)
    current_date = None

    for line in log_content.split('\n'):
        line = line.strip()
        if not line:
            continue

        # Проверяем все форматы дат
        date_found = False
        for pattern in date_patterns:
            date_match = re.search(pattern, line)
            if date_match:
                date_str = date_match.group(1)
                try:
                    # Пробуем разные форматы дат
                    for fmt in ('%m/%d/%Y', '%d/%m/%Y', '%a %b %d %Y'):
                        try:
                            current_date = datetime.strptime(date_str, fmt).date()
                            date_found = True
                            break
                        except ValueError:
                            continue
                except Exception:
                    continue
                break

        # Если это строка с датой, добавляем ее в записи
        if date_found:
            date_entries[current_date].append(f"----{current_date.strftime('%Y-%m-%d')}----")
            continue

        # Пропускаем строки без даты (если дата еще не установлена)
        if current_date is None:
            continue

        # Добавляем строки IN и OUT
        if 'IN:' in line or 'OUT:' in line:
            date_entries[current_date].append(line)

    return date_entries

appdatafile = '%s\\FlexLM License Analyst\\' %  os.environ['APPDATA']
if not os.path.exists(appdatafile):
    os.makedirs(appdatafile)

file_path = '%slicense_report.txt' % appdatafile


def write_log_report(date_entries, output_file=file_path):
    with open(output_file, 'w', encoding='utf-8') as f:
        for date in sorted(date_entries.keys()):
            for entry in date_entries[date]:
                f.write(entry + '\n')
            f.write('\n')  # Добавляем пустую строку между датами


def main(user_input):
    try:
        input_file = user_input
        output_file = file_path

        with open(input_file, 'r', encoding='utf-8') as file:
            log_content = file.read()

        date_entries = parse_flexlm_log(log_content)
        write_log_report(date_entries, output_file)

        print(f"Отчет успешно создан в файле: {output_file}")

    except FileNotFoundError:
        print(f"Ошибка: файл {input_file} не найден")
    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")


if __name__ == "__main__":
    main()