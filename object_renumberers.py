from colorama import Fore, init
import docx
import re
from os.path import exists
from doctools_lib import replace_text_in_runs, \
    unpack_ref, pack_ref, paragraph_iterator


class doc_element_renumberer(object):
    def __init__(self, doc_file=None, title_pattern=None,
                 inline_pattern=None, params_tag=None,
                 start=1, prefix='',
                 element_name='Element',
                 verbose=True):
        self.doc_file = doc_file
        self.title_pattern = title_pattern
        self.inline_pattern = inline_pattern
        self.params_pattern = re.compile(
            r'<%s>(.*)<\\%s>' % (params_tag, params_tag))
        self.stop_tag = re.compile(r'<%s?stop>' % (params_tag, ))
        self.continue_tag = re.compile(r'<%s?continue>' % (params_tag, ))
        self.start_pattern = re.compile(r'start\s*=\s*([\d]+)')
        self.prefix_pattern = re.compile(r'prefix\s*=\s*([\d.]+)')
        self.start = start
        self.prefix = prefix
        self.output = None
        self.element_name = element_name
        self.verbose = verbose
        init(autoreset=True)

    def get_renumber_params(self, input_str):
        _ = self.params_pattern.search(input_str)
        if not _:
            return None
        self.prefix = self.prefix_pattern.search(_[1])[1] \
            if self.prefix_pattern.search(_[1]) else ''
        self.start = int(self.start_pattern.search(_[1])[1]
                         ) if self.start_pattern.search(_[1]) else 1

    def analize(self):
        if not exists(self.doc_file):
            print(self.doc_file, 'does not exist')
            return
        if self.verbose:
            print(Fore.GREEN+f'Analising {self.element_name.lower()}s...')
        document = docx.Document(self.doc_file)
        rez = {}
        do_analysis = True
        for p in paragraph_iterator(document):
            if self.stop_tag.search(p.text):
                do_analysis = False
            if self.continue_tag.search(p.text):
                do_analysis = True
            if not do_analysis:
                continue
            self.get_renumber_params(p.text)
            m = self.title_pattern.match(p.text)
            if m:
                rez[m.group('num')] = [self.prefix, self.start]
                self.start += 1
        if self.verbose:
            print(f'\t{len(rez)} {self.element_name.lower()}s in document')
        self.title_dict = rez

    def renumber(self):
        if not exists(self.doc_file):
            print(self.doc_file, 'does not exist')
            return
        self.analize()
        if self.verbose:
            print(Fore.GREEN+f'Renumbering {self.element_name.lower()}s...')
        doc = docx.Document(self.doc_file)
        do_replacement = True
        for p in paragraph_iterator(doc):
            if self.stop_tag.search(p.text):
                do_replacement = False
            if self.continue_tag.search(p.text):
                do_replacement = True
            if not do_replacement:
                continue
            ss = self.inline_pattern.search(p.text)
            while ss:
                _, _, nn = unpack_ref(ss[10])
                for n in nn:
                    if n not in self.title_dict:
                        print(
                            Fore.RED+f'{self.element_name} {n} is not found in document')
                        nn.remove(n)
                if nn:
                    new_str = pack_ref([self.title_dict[n][1]
                                        for n in nn], self.title_dict[nn[0]][0])
                    if self.verbose:
                        print('\t', ss[0], '->', new_str)
                    replace_text_in_runs(p.runs, ss.span(10)[0],
                                         ss.span(10)[1], new_str)
                ss = self.inline_pattern.search(p.text, ss.span(10)[1]+1)
        self.output = doc

    def save(self, path):
        if self.verbose:
            print(Fore.GREEN+f'Saveing to '+Fore.WHITE+f'{path}')
        self.output.save(path)


class doc_figures_renumberer(doc_element_renumberer):
    def __init__(self, doc_file=None, start=1, prefix=''):
        title_pattern = re.compile(
            r'^\s*(Рис\.|Рисунок)\s+(?P<num>[\d.]+?)\s*([–-].*|[. ]*$)')
        inline_pattern = re.compile(
            r'(рис((.)|(ун((ке)|(ок)|(ка)|(ках))))\s+)(([\s,и]*[\d.\-–]+)*[^\D])', re.IGNORECASE)
        super().__init__(doc_file=doc_file,
                         element_name='Figure',
                         start=start,
                         prefix=prefix,
                         params_tag='f',
                         inline_pattern=inline_pattern,
                         title_pattern=title_pattern)


class doc_table_renumberer(doc_element_renumberer):
    def __init__(self, doc_file=None, start=1, prefix=''):
        title_pattern = re.compile(
            r'^\s*(Таб\.|Таблица|Табл.)\s+(?P<num>[\d.]+?)\s*([–-].*|[. ]*$)')
        inline_pattern = re.compile(
            r'(табл((.)|(иц((е)|(ы)|(ах)|(а)|())))\s+)(([\s,и]*[\d.\-–]+)*[^\D])', re.IGNORECASE)
        super().__init__(doc_file=doc_file,
                         element_name='Table',
                         start=start,
                         prefix=prefix,
                         params_tag='t',
                         inline_pattern=inline_pattern,
                         title_pattern=title_pattern)


class doc_equation_renumberer(doc_element_renumberer):
    def __init__(self, doc_file=None, start=1, prefix=''):
        title_pattern = re.compile(r'[^а-яА-Я]*\((?P<num>[\d.-]+)\)\s*$')
        inline_pattern = re.compile(r'\(([\d.]+)\)')
        super().__init__(doc_file=doc_file,
                         element_name='Equation',
                         start=start,
                         prefix=prefix,
                         params_tag='e',
                         inline_pattern=inline_pattern,
                         title_pattern=title_pattern)


if __name__ == '__main__':
    ren = doc_figures_renumberer(doc_file='test.docx', start=10, prefix='2.')
    ren.renumber()
    ren.save('test2.docx')
