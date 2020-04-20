# %%
import docx
import re
from doctools_lib import replace_text_in_runs, paragraph_iterator

formula = re.compile(r'\(([\d.]+)\)')


def get_eq_params(input_str):
    kw = re.compile(r'\]->(.*eq.*)<-\[')
    s = kw.search(input_str)
    if s:
        s = s[1]
        kws = {}
        pref = re.compile(r'eq_prefix=([\d.-]+)')
        start = re.compile(r'eq_start=([\d]+)')
        kws['prefix'] = pref.search(s)[1] if pref.search(s) else ''
        kws['start'] = int(start.search(s)[1]) if start.search(s) else 1
        return kws
    else:
        return None


def renumber_formulas(inp, output, prefix, start):
    formulas = {}
    doc = docx.Document(inp)
    for p in paragraph_iterator(doc):  # doc.paragraphs:
        kw = get_eq_params(p.text)
        if kw:
            start = kw['start']
            prefix = kw['prefix']
            continue
        ss = formula.search(p.text)
        while ss:
            n = ss[0]
            if not n in formulas:
                formulas[n] = [prefix, start]
                start += 1
            new_str = f'{formulas[n][0]}{formulas[n][1]}'
            print(ss[1], '->', new_str)
            replace_text_in_runs(p.runs, ss.span(1)[0],
                                 ss.span(1)[1], new_str)
            ss = formula.search(p.text, pos=ss.span(1)[1]+1)
    doc.save(output)


if __name__ == '__main__':
    from argparse import ArgumentParser
    parser = ArgumentParser()
    parser.add_argument('input', help='Input docx file')
    parser.add_argument('--start', '-n', default=1,
                        help='Start number for figures numbering', type=int)
    parser.add_argument('--prefix', '-p', default='',
                        help='Default prefix for numbers')
    parser.add_argument(
        '--output', '-o', default='rez.docx', help='Output file')
    args = parser.parse_args()
    renumber_formulas(args.input, args.output, args.prefix, args.start)
