from .misc import _doc2html,insert_tag
import os
from bs4 import BeautifulSoup

desired_type = ['docx','doc','txt']
def is_desired_type(fname):
    for i in desired_type:
        if fname.endswith(i):
            return True
    return False

def process_title(html_path):
    fname = '.'.join(html_path.split('\\')[-1].split('.')[:-1])
    with open(html_path,'r',encoding='cp936') as fin:
        buff = ''.join(fin.readlines())
        soup = BeautifulSoup(buff,'lxml')
        head_tag = soup.find('head')
        title_tags = head_tag.find_all('title')
        for t in title_tags:
            t.extract()
        insert_tag(head_tag,'title', fname)
    with open(html_path,'w',encoding='cp936') as fout:
        fout.write(soup.prettify(encoding='gb2312').decode('gb2312'))    

def doc2html(src_path,des_path,process_html = None):
    print((src_path,des_path))
    src_path_length = len(src_path.split(os.path.sep))
    scnt,ecnt = 0,0
    for root_dir,sub_dir,files in os.walk(src_path):
        infiles = [ifile for ifile in files if ( not ifile.startswith('~') and is_desired_type(ifile ))]
        cur_dir = os.path.sep.join((root_dir.split(os.sep)[src_path_length:]))
        des_dir = os.path.join(des_path,cur_dir)
        if not os.path.exists(des_dir):
            os.makedirs(des_dir)
        print(cur_dir)
        for ifile in infiles:
            doc_path = os.path.join(root_dir,ifile)
            html_path = os.path.join(des_dir,'.'.join(ifile.split('.')[:-1]) + '.html')
            if os.path.exists(html_path):
                continue
            else:
                try:
                    print((doc_path, html_path))
                    _doc2html(doc_path, html_path)
                    scnt += 1
                    print((scnt,ifile))
                    if process_html is not None:
                        process_html(html_path)
                except Exception as e:
                    ecnt += 1
                    print(e)
                    print(('error! infile = ' , ifile))
    print(('successed conveted ' + str(scnt) + ' documents!')) 
    print(('failed conveted ' + str(ecnt) + ' documents!')) 
        

    