# %%
import docx
import re
from doctools_lib import unpack_ref, pack_ref, replace_text_in_runs, paragraph_iterator


def get_tab_params(input_str):
    kw = re.compile(r'\]->(.*tab.*)<-\[')
    s = kw.search(input_str)
    if s:
        s = s[1]
        kws = {}
        pref = re.compile(r'tab_prefix=([\d.-]+)')
        start = re.compile(r'tab_start=([\d]+)')
        kws['prefix'] = pref.search(s)[1] if pref.search(s) else ''
        kws['start'] = int(start.search(s)[1]) if start.search(s) else 1
        return kws
    else:
        return None


tab_in_text = re.compile(
    r'(табл((.)|(иц((е)|(ы)|(ах)|(а)|())))\s+)(([\s,и]*[\d.\-–]+)*[^\D])', re.IGNORECASE)


def renumber_tables(inp, output, prefix, start):
    tables = {}
    doc = docx.Document(inp)
    for p in paragraph_iterator(doc):  # doc.paragraphs:
        kw = get_tab_params(p.text)
        if kw:
            start = kw['start']
            prefix = kw['prefix']
            continue
        ss = tab_in_text.search(p.text)
        while ss:
            #body = ss[1]
            _, _, nn = unpack_ref(ss[11])
            for n in nn:
                if not n in tables:
                    tables[n] = [prefix, start]
                    start += 1
            new_str = pack_ref([tables[n][1] for n in nn], tables[n][0])
            print(ss[0], '->', new_str)
            replace_text_in_runs(p.runs, ss.span(11)[0],
                                 ss.span(11)[1], new_str)
            ss = tab_in_text.search(p.text, pos=ss.span(11)[1]+1)
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
    renumber_tables(args.input, args.output, args.prefix, args.start)
