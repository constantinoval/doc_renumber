from colorama import Fore, Back, init
if __name__ == '__main__':
    init(autoreset=True)
    from argparse import ArgumentParser
    parser = ArgumentParser(description="""Script to renumber formulas, figures, tables and references in docx files.
You can use special symbols in text to add prefixes and start numbers for formalas, figures and tables
at specific place of word files. It is <e><\e> for equations, <f><\f> - for figures, <><\t> - for tables.
Tags <estop>, <tstop> and <fstop> - stops analysis and renumbering of equations, tables and figures.
Tags <econtinue>, <tcontinue>, <fcontinue> - continues analysis and renumbering of equations, tables and figures.
Tags <stop> and <continue> - stops and continues e,f, and t.
For example: <e>prefix=2.1. start=5<\e> - for equations.

Usage example: python doc_renumber test.docx -m fe --output test.docx
               python doc_renumber test.docx
               python doc_renumber test.docx --refs literature.xlsx --new_refs new_lit.xlsx
""")
    parser.add_argument('input', help='Input docx file')
    parser.add_argument('--mode', '-m', default='eft',
                        help='Items to renumber. e - equations, f - figures, t - tables, r - references, s - show statistics.\n For example: fr - figures and references')
    parser.add_argument('--start', '-n', default=1,
                        help='Start number of new numbering', type=int)
    parser.add_argument('--prefix', '-p', default='',
                        help='Prefix to add to numbering')
    parser.add_argument(
        '--refs', '-r', default=None, help='xlsx file with references. Must have cols: n and title. Or it can have same name as input docx')
    parser.add_argument('--new_refs', default='new_lit.xlsx',
                        help='Output xls file with new references numbering')
    parser.add_argument('--output', '-o', default='rez.docx',
                        help='Output docx file')
    args = parser.parse_args()
    inp = args.input
    if 'R' in args.mode.upper():
        from renumber_literature import renumber_refs
        print(Fore.GREEN+'*** Renumbering references ***')
        renumber_refs(inp, args.output, args.refs, args.new_refs, args.start)
        inp = args.output
    if 'F' in args.mode.upper():
        from renumber_figures import renumber_figures
        print(Fore.GREEN+'*** Renumbering figures ***')
        renumber_figures(inp, args.output, args.prefix, args.start)
        inp = args.output
    if 'T' in args.mode.upper():
        from renumber_tables import renumber_tables
        print(Fore.GREEN+'*** Renumbering tables ***')
        renumber_tables(inp, args.output, args.prefix, args.start)
        inp = args.output
    if 'E' in args.mode.upper():
        from renumber_formulas import renumber_formulas
        print(Fore.GREEN+'*** Renumbering formulas ***')
        renumber_formulas(inp, args.output, args.prefix, args.start)
        inp = args.output
    if 'S' in args.mode.upper():
        from doc_statistics import doc_analysis
        doc_analysis(args.input)
