from pathlib import Path
import chardet


class UnknownEncoding(Exception):
    pass


def detect_encoding(file: Path):
    with file.open('rb') as stream:
        content = stream.read()
        detected = chardet.detect(content)
        if detected['confidence'] < 0.95:
            raise UnknownEncoding(detected)
        return detected['encoding']


def check_utf_or_win_encoding(file: Path):
    try:
        with open(file, encoding='utf-16') as stream:
            data = stream.read()
            line_list = data.split('\n')
            if any('charset = unicode' in s for s in line_list[:10]):
                return 'utf-16'
    except (UnicodeError, UnicodeDecodeError):
        with open(file, encoding='windows-1251') as stream:
            data = stream.read()
            line_list = data.split('\n')
            if any('charset=windows-1251' in s for s in line_list[:10]):
                return 'windows-1251'
        raise None


def add_mail_thread_to_final_table(file: Path, encoding, content, aim_file):
    with file.open(encoding=encoding) as stream:
        data = stream.read()
        data = data.replace("<div class=WordSection1>", f"<div class=WordSection1>{content}")
        with aim_file.open('w', encoding=encoding) as stream2:
            stream2.write(data)
