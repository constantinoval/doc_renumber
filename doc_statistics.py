from renumber_figures import analize_figures
from renumber_formulas import analize_formulas
from renumber_tables import analize_tables
from docx import Document
from colorama import Fore, init
from pprint import pprint


def doc_analysis(doc):
    init(autoreset=True)
    d = Document(doc)
    f = analize_figures(d)
    print(Fore.GREEN + 'Figures:', len(f))
    pprint(list(f.keys()), indent=5, width=40, compact=True)
    e = analize_formulas(d)
    print(Fore.GREEN + 'Equations:', len(e))
    pprint(list(e.keys()), indent=5, width=40, compact=True)
    t = analize_tables(d)
    print(Fore.GREEN + 'Tables:', len(t))
    pprint(list(t.keys()), indent=5, width=40, compact=True)


if __name__ == '__main__':
    doc_analysis(r"D:\oneDrive\work\НИИМ\НИАГАРА\Отчет по 218\all\2.docx")
