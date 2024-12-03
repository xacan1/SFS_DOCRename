import pathlib
import json
import re
from zipfile import BadZipFile
import win32com.client
from win32com.client import CDispatch
from pythoncom import com_error
from multiprocessing import Queue
import docx2txt
import config
from timeout import timeout


VALIDATE_FILENAME = re.compile(r'[^\w\-_\. ]')


def get_default_keywords() -> dict:
    keywords = {
        'Аннотирован$ANN': ['Профиль/специализация', 'Факультет', 'Направление/специальность подготовки', 'Направленность', 'Направление'],
        'Кейс-задача$KZ': ['Профиль/специализация', 'Факультет', 'Направление/специальность подготовки', 'Направленность', 'Направление'],
        'Выпускная квалификационная работа$VKR': ['Тема ВКР', 'Тема выпускной квалификационной работы', 'тему', 'по теме', 'ТЕМА', 'ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА'],
        'Научно-квалификационная работа$VKR': ['Тема ВКР', 'Тема выпускной квалификационной работы', 'выполненная на тему:', 'на тему:', 'тему:', 'по теме', 'ТЕМА', 'ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА'],
        'ДИПЛОМНЫЙ$VKR': ['Тема ВКР', 'Тема выпускной квалификационной работы', 'на тему', 'по теме', 'ТЕМА', 'ВЫПУСКНАЯ КВАЛИФИКАЦИОННАЯ РАБОТА'],
        'Курсовая работа$KR': ['На тему', 'По теме', 'Тема', 'по дисциплине'],
        'Контрольно-курсовое$KKZ': ['Дисциплина'],
        'Экзаменационная дисциплина$EZ': ['Экзаменационная дисциплина'],
        'ПМ.$PM': ['ПМ.'],
        'Лабораторный практикум$LP': ['Профиль/специализация', 'Факультет', 'Направление/специальность подготовки', 'Направленность', 'Направление'],
        'Лабораторная работа$LR': ['Профиль/специализация', 'Факультет', 'Направление/специальность подготовки', 'Направленность', 'Направление'],
        'Ситуационный практикум$SP': ['Ситуационный практикум'],
        'Общепсихологический практикум$OPP': ['Общепсихологический практикум'],
        'Содержание индивидуального задания на практику в$KZUGP': ['Содержание индивидуального задания на практику в'],
        'Проективная методика$OPP': ['Проективная методика'],
        'Практикум по психодиагностике$OPP': ['Практикум по психодиагностике'],
        'Профиль/специализация$PRF': ['Профиль/специализация'],
        '(профиль:$PRF': ['(профиль:'],
        'Специальность$PRF': ['Специальность'],
        'Направление подготовки$NP': ['Направление подготовки'],
        'Практикум$PRUM': ['Практикум'],
        'Факультет$NP': ['Факультет'],
    }

    return keywords


def get_default_stop_words() -> list[str]:
    stop_words = [
        'обучения заочная',
        '(наименование',
        'Утверждена приказом',
        'Структура ВКРВведение',
        'Структура ВКР_ВведениеГлава 1_ Теоретические ас',
        'Тема прописывается',
        'Структура ВКР_Введение1_ Аналитическая',
        '2_ Структура',
        'ВКР_ВведениеВведениеГлава 1_',
        'Исходные данные по',
        'Обучающий',
    ]

    return stop_words


def get_default_min_symbols_in_doc() -> str:
    min_symbols_in_doc = 6000
    return min_symbols_in_doc


def load_config() -> tuple[str, dict, list, str]:
    path_directory = ''
    keywords = {}
    stop_words = []
    min_symbols_in_doc = ''

    try:
        with open('config.cfg', mode='r', encoding='utf-8') as f:
            data_config_json = f.read()

            if data_config_json:
                data_config = json.loads(data_config_json)
                path_directory = data_config.get('path_directory', '')
                keywords = data_config.get('keywords', {})
                stop_words = data_config.get('stop_words', [])
                min_symbols_in_doc = data_config.get('min_symbols_in_doc', '0')
    except FileNotFoundError:
        f = open('config.cfg', mode='w', encoding='utf-8')
        f.close()

    if not keywords:
        keywords = get_default_keywords()

    if not stop_words:
        stop_words = get_default_stop_words()

    if not min_symbols_in_doc:
        min_symbols_in_doc = get_default_min_symbols_in_doc()

    return (path_directory, keywords, stop_words, min_symbols_in_doc)


def save_config(data_config: dict) -> None:
    with open('config.cfg', mode='w', encoding='utf-8') as f:
        keywords_json = json.dumps(data_config, ensure_ascii=False)
        f.write(keywords_json)


def get_ms_word_com() -> CDispatch | None:
    msword = None

    try:
        msword = win32com.client.Dispatch('Word.Application')
    except com_error as exp:
        if config.DEBUG:
            print(f'Не установлен Word. {exp}')

    return msword


def msword_installed() -> bool:
    result = False
    msword = get_ms_word_com()

    if msword:
        result = True
        msword.Quit()

    return result


def convert_doc_to_docx(mpqueue: Queue) -> None:
    max_count = get_count_files('*.doc')
    path_directory, _, _, _ = load_config()
    path = pathlib.Path(path_directory)
    msword = get_ms_word_com()

    if not msword:
        return

    msword.visible = config.DEBUG
    msword.DisplayAlerts = False
    counter = 0

    for file_path in path.glob('*.doc'):
        doc_file = str(file_path)
        counter += 1
        mpqueue.put((counter, max_count, f'Конвертация {file_path.name}'))

        try:
            wb = msword.Documents.Open(doc_file)
        except com_error as exp:
            if config.DEBUG:
                print(exp)

            continue

        try:
            wb.SaveAs2(doc_file + 'x', FileFormat=16)
        except com_error as exp:
            if config.DEBUG:
                print(exp)

            wb.Close()
            continue

        try:
            wb.Close()
        except com_error as exp:
            if config.DEBUG:
                print(exp)

            continue

        file_path.unlink()

    msword.Quit()


# Выведу в отдельную функцию работу с сохранением файла для реализации таймаута
# ПОКА ЧТО НЕ РАБОТАЕТ
@timeout(10)
def save_doc(wb, doc_file: str) -> bool:
    result = True

    try:
        wb.SaveAs2(doc_file + 'x', FileFormat=16)
    except com_error as exp:
        if config.DEBUG:
            print(exp)

        result = False
        wb.Close()

    return result


def get_count_files(file_mask: str) -> int:
    path_directory, _, _, _ = load_config()
    path = pathlib.Path(path_directory)
    count_files = len(list(path.glob(file_mask)))
    return count_files


def rename_docx(mpqueue: Queue) -> None:
    path_directory, keywords, stop_words, min_symbols_in_doc = load_config()
    min_symbols_in_doc = int(min_symbols_in_doc)
    path = pathlib.Path(path_directory)
    max_count = get_count_files('*.docx')
    counter = 0

    for file_path in path.glob('*.docx'):
        counter += 1
        mpqueue.put(
            (counter, max_count, 'Анализ и переименовывание файлов DOCX...'))

        if 'find_' in file_path.name or 'notfound_' in file_path.name or 'notlim_' in file_path.name:
            continue

        if config.DEBUG:
            print(file_path)

        try:
            text = docx2txt.process(file_path)
        except BadZipFile:
            if config.DEBUG:
                print('файл не является DOCX')

            continue
        except KeyError:
            if config.DEBUG:
                print('Проблемы с содержимым DOCX файла')

            continue

        if len(text) < min_symbols_in_doc:
            file_path.replace(
                f'{str(file_path.parent)}\\notlim_{file_path.name}')

            continue

        text = text.casefold()
        new_name_file = get_new_name_file(text, keywords, stop_words)

        if new_name_file:
            if config.DEBUG:
                print(new_name_file)

            file_path.replace(f'{str(file_path.parent)}\\{new_name_file}')
        else:
            if config.DEBUG:
                print(f'notfound_{file_path.name}')

            file_path.replace(
                f'{str(file_path.parent)}\\notfound_{file_path.name}')


# Выполняет всю работу подряд
def start_work(mpqueue: Queue) -> None:
    convert_doc_to_docx(mpqueue)
    rename_docx(mpqueue)


def get_new_name_file(text: str, keywords: dict, stop_words: list[str]) -> str:
    name_file = ''

    # сначала найдем тип работы по ключевому слову
    for key, subword in keywords.items():
        key_info = key.split('$')
        type_work = key_info[0].casefold()
        symbol_work = key_info[1]

        if type_work in text:
            if config.DEBUG:
                print(f'type_work:[ {type_work} ]')

            name_work = ''

            # теперь найдем начало предполагаемого названия темы
            for word in subword:
                word = word.casefold()

                if word in text:
                    if config.DEBUG:
                        print(f'word:[ {word} ]')

                    word = safe_symbols_re(word)
                    result = re.search(f'{word}[\W\s]*((.+))', text)

                    if result:
                        name_work = result.group(1).strip()
                        break

            if name_work:
                name_file = get_validate_filename(f'find_{symbol_work}_{name_work}',
                                                  stop_words)
                break

    return name_file


def get_validate_filename(filename: str, stop_words: list[str]) -> str:
    validate_filename = VALIDATE_FILENAME.sub('', filename)
    validate_filename = del_words_from_filename(validate_filename, stop_words)
    validate_filename = truncate_filname(validate_filename)
    validate_filename += '.docx'
    return validate_filename


def del_words_from_filename(filename: str, stop_words: list[str]) -> str:
    for stop_word in stop_words:
        filename = filename.replace(stop_word, '')

    return filename


def truncate_filname(filename: str) -> str:
    return filename[0:120].strip()


# экранирует спец символы регулярок если они есть
def safe_symbols_re(word: str) -> str:
    symbols = '^$*+?{}[]\|()'

    for symbol in symbols:
        word = word.replace(symbol, f'\{symbol}')

    return word
