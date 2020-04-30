from zipfile import ZipFile
from pathlib import Path
from shutil import copyfile


file_path = Path() / 'doc_renumber.pyz'
z = ZipFile(file_path, 'w')
for f in ['doctools_lib.py',
          'iter_blocks.py',
          'renumber_figures.py',
          'renumber_formulas.py',
          'renumber_literature.py',
          'renumber_tables.py',
          'doc_statistics.py',
          '__main__.py'
          ]:
    z.write(f, f)
z.close()
copyfile(file_path, 'd:/programs/doc_renumber.pyz')
