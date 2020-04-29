import docx
from iter_blocks import iter_block_items
import re


def replace_text_in_runs(runs, start, end, new_text):
    do_replacement = False
    ii = 0
    for j, r in enumerate(runs):
        if ii > end:
            return
        i = len(r.text)
        if ii < start and (ii+i) > end:
            runs[j].text = runs[j].text[:start-ii] + \
                new_text+runs[j].text[end-ii:]
            do_replacement = True
            break
        if ii >= start and (ii+i) <= end:
            if not do_replacement:
                runs[j].text = new_text
                do_replacement = True
            else:
                runs[j].text = ''
            ii += i
            continue
        if start <= ii <= end and (ii+i) >= end:
            runs[j].text = r.text[end-ii:]
            if not do_replacement:
                runs[j].text = new_text+runs[j].text
                do_replacement = True
            ii += i
            continue
        if ii <= start and start <= (ii+i) <= end:
            runs[j].text = r.text[:start-ii]
            if not do_replacement:
                runs[j].text += new_text
                do_replacement = True
            ii += i
            continue
        ii += i


def unpack_ref(s):
    p = re.compile(',|;|и')
    tir = re.compile('-|–')
    rez = []
    for ss in p.split(s):
        sss = tir.split(ss)
        if len(sss) > 1:
            rez.extend(range(int(sss[0].split('.')[-1]),
                             int(sss[1].split('.')[-1])+1))
        else:
            rez.append(int(ss.split('.')[-1]))
    pr = sss[0]
    prefix = ('.'.join(pr.split('.')[:-1])
              if len(pr.split('.')) >= 2 else '').strip()
    if prefix:
        prefix += '.'
    nums = rez
    full_nums = [f'{prefix}{i}' for i in nums]

    return prefix, nums, full_nums


def pack_ref(s, prefix=''):
    s = list(s)
    s.sort()
    while(-1) in s:
        s.pop(0)
    if len(s) == 0:
        return ''
    ss = [False]+[abs((s[i+1]-s[i])) > 1 for i in range(len(s)-1)]
    rez = [[]]
    for i, sss in enumerate(ss):
        if sss:
            rez.append([])
            rez[-1].append(s[i])
        else:
            rez[-1].append(s[i])
    rez2 = []
    for r in rez:
        if len(r) > 2:
            rez2.append(f'{prefix}{r[0]}-{prefix}{r[-1]}')
        else:
            rez2.extend(map(lambda x: f'{prefix}'+str(x), r))
    return ', '.join(rez2)


def paragraph_iterator(doc):
    for block in iter_block_items(doc):
        if isinstance(block, docx.text.paragraph.Paragraph):
            yield block
        if isinstance(block, docx.table.Table):
            ncols = len(block.columns)
            nrows = len(block.rows)
            for r in range(nrows):
                for c in range(ncols):
                    try:
                        cell = block.cell(r, c)
                        for p in cell.paragraphs:
                            yield p
                    except:
                        pass


if __name__ == '__main__':
    s = f"3.1.1- 3.1.5"
    print(s, unpack_ref(s))
