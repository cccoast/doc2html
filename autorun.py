import configparser as cp
from doc2html.convertor import doc2html,process_title

if __name__ == '__main__':
    config = cp.ConfigParser()
    config.read('config.txt')
    src_path = config.get('config','src_path')
    des_path = config.get('config','des_path')
    doc2html(src_path,des_path,process_title)