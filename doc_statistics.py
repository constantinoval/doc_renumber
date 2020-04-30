import re
from doctools_lib import paragraph_iterator
from docx import Document
from colorama import Fore, init, Back
from pprint import pprint


def analize_figures(lines):
    stop_tag = re.compile(r'<f+stop>')
    continue_tag = re.compile(r'<f+continue>')
    pat = re.compile(
        r'^\s*(Рис\.|Рисунок)\s+(?P<num>[\d.]+?)\s*([–-].*|[. ]*$)')
    rez = []
    do_analysis = True
    for p in lines:
        if stop_tag.search(p):
            do_analysis = False
        if continue_tag.search(p):
            do_analysis = True
        if not do_analysis:
            continue
        m = pat.match(p)
        if m:
            rez.append(m.group('num'))
    return rez


def analize_formulas(lines):
    stop_tag = re.compile(r'<e+stop>')
    continue_tag = re.compile(r'<e+continue>')
    pat = re.compile(r'[^а-яА-Я]*\((?P<num>[\d.-]+)\)\s*$')
    rez = []
    do_analysis = True
    for p in lines:
        if stop_tag.search(p):
            do_analysis = False
        if continue_tag.search(p):
            do_analysis = True
        if not do_analysis:
            continue
        m = pat.match(p)
        if m:
            rez.append(m.group('num'))
    return rez


def analize_tables(lines):
    stop_tag = re.compile(r'<t+stop>')
    continue_tag = re.compile(r'<t+continue>')
    pat = re.compile(
        r'^\s*(Таб\.|Таблица|Табл.)\s+(?P<num>[\d.]+?)\s*([–-].*|[. ]*$)')
    rez = []
    do_analysis = True
    for p in lines:
        if stop_tag.search(p):
            do_analysis = False
        if continue_tag.search(p):
            do_analysis = True
        if not do_analysis:
            continue
        m = pat.match(p)
        if m:
            rez.append(m.group('num'))
    return rez


def get_lines(doc):
    rez = []
    for p in paragraph_iterator(doc):
        rez.append(p.text)
    return rez


def doc_analysis(doc):
    init(autoreset=True)
    print(Fore.BLACK+Back.WHITE+'Reading text...')
    lines = get_lines(Document(doc))
    print(Fore.BLACK+Back.WHITE+'Done...')
    f = analize_figures(lines)
    print(Fore.GREEN + 'Figures:', len(f))
    pprint(f, indent=5, width=40, compact=True)
    e = analize_formulas(lines)
    print(Fore.GREEN + 'Equations:', len(e))
    pprint(e, indent=5, width=40, compact=True)
    t = analize_tables(lines)
    print(Fore.GREEN + 'Tables:', len(t))
    pprint(t, indent=5, width=40, compact=True)

    print(Back.YELLOW+Fore.BLACK + 'Figures:' +
          Fore.WHITE+Back.BLACK+'\t'+str(len(f)))
    print(Back.YELLOW+Fore.BLACK + 'Equations:' +
          Fore.WHITE+Back.BLACK+'\t'+str(len(e)))
    print(Back.YELLOW+Fore.BLACK + 'Tables:' +
          Fore.WHITE+Back.BLACK+'\t'+str(len(t)))


if __name__ == '__main__':
    doc_analysis(r"D:\oneDrive\work\НИИМ\НИАГАРА\Отчет по 218\all\2.docx")
