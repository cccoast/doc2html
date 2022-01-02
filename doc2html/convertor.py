#coding:utf-8

import ConfigParser as cp
import misc
import os
from bs4 import BeautifulSoup

desired_type = ['docx','doc','txt']
def is_desired_type(fname):
    for i in desired_type:
        if fname.endswith(i):
            return True
    return False

def process_title(html_path):
    fname = '.'.join(html_path.split('\\')[-1].split('.')[:-1]).decode('cp936')
    print html_path.decode('cp936')
    with open(html_path,'r') as fin:
        buffer = ''.join(fin.readlines())
        soup = BeautifulSoup(buffer,'lxml')
        head_tag = soup.find('head')
        title_tags = head_tag.find_all('title')
        for t in title_tags:
            t.extract()
        misc.insert_tag(head_tag,'title', fname)
    with open(html_path,'w') as fout:
        fout.write(soup.encode('cp936'))    

def doc2html(src_path,des_path):
    src_path_length = len(src_path.split(os.path.sep))
    scnt,ecnt = 0,0
    for root_dir,sub_dir,files in os.walk(src_path):
        infiles = [ifile for ifile in files if ( not ifile.startswith('~') and is_desired_type(ifile ))]
        cur_dir = os.path.sep.join((root_dir.split(os.sep)[src_path_length:]))
        des_dir = os.path.join(des_path,cur_dir)
        if not os.path.exists(des_dir):
            os.makedirs(des_dir)
        print cur_dir.decode('cp936')
        for ifile in infiles:
            doc_path = os.path.join(root_dir,ifile)
            html_path = os.path.join(des_dir,'.'.join(ifile.split('.')[:-1]) + '.html')
            if os.path.exists(html_path):
                continue
            else:
                try:
                    misc.doc2html(doc_path, html_path)
                    scnt += 1
                    print scnt,'\t',os.path.join(cur_dir,ifile).decode('cp936')
                    process_title(html_path)
                except:
                    ecnt += 1
                    print 'error! infile = ' , os.path.join(cur_dir,ifile).decode('cp936')
    print 'successed conveted ' + str(scnt) + ' documents!' 
    print 'failed conveted ' + str(ecnt) + ' documents!' 
    misc.word.Quit()
        
if __name__ == '__main__':
    config = cp.ConfigParser()
    config.read('config.txt')
    src_path = config.get('config','src_path')
    des_path = config.get('config','des_path')
    doc2html(src_path,des_path)
    