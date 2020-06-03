import os, sys, shutil
from configparser import ConfigParser


class MyPath():

    # 获取路径，获取当前应用程序所在的路径
    @staticmethod
    def get_cur_path(fileName=None):
        path_dir = os.path.dirname(os.path.realpath(sys.argv[0]))
        if fileName == None:
            return path_dir
        path = os.path.join(path_dir, fileName)
        return path

    # 获取路径, 打包后资源文件所在路径。
    @staticmethod
    def get_res_path(fileName=None):
        # 方法一（如果要将资源文件打包到app中，使用此法可以获取资源文件路径）
        # 注意资源文件可能不在当前应用程序目录下。
        bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(
                os.path.dirname(__file__)))
        if fileName == None:
            return bundle_dir
        path = os.path.join(bundle_dir, fileName)
        return path

    # 复制文件到指定路径
    @staticmethod
    def copy_file(filename='default.pptx'):
        oldname = MyPath.get_res_path(filename)
        newname = MyPath.get_cur_path(filename)
        # 如果资源文件目录中存在该文件，则复制到应用程序目录中来。
        if os.path.exists(oldname) and os.path.isfile(oldname) and oldname != newname:
            # 保复制文件数据
            shutil.copyfile(oldname, newname)

    # 复制文件夹或文件到指定路径
    @staticmethod
    def copy_files(oldname, newname):
        # 复制文件夹及里面的文件
        if os.path.isdir(oldname):
            # 如果新目录已存在，则先删除。（即覆盖模式）
            if os.path.exists(newname):
                shutil.rmtree(newname)
            # 复制整个目录树，目标路径不能已存在，否则报错
            try:
                shutil.copytree(oldname, newname)
                print('复制文件夹及里面的文件')
                return True
            except Exception as e:
                print('复制文件夹出错，原因：', e)
            
        # 复制文件
        if os.path.isfile(oldname) and oldname != newname:
            try:
                # 复制文件数据和文件的权限模式
                shutil.copy(oldname, newname)
                return True
            except Exception as e:
                print('复制文件出错，原因：', e)

        return False

# 获取配置文件config.ini的数据库配置信息并返回
# def getConfigs():
#     cfg = ConfigParser()
#     cfg.read(MyPath.get_cur_path('config.ini'), encoding="utf-8")
#
#     d_provider = cfg.get('db_provider', 'provider')
#     print(d_provider)
#
#     mydb = {}
#     if d_provider == 'sqlite':
#         provider = cfg.get(d_provider, 'provider')
#         filename = cfg.get(d_provider, 'filename')
#         create_db = cfg.getboolean(d_provider, 'create_db')
#         mydb = dict(provider=provider, filename=filename, create_db=create_db)
#
#     elif d_provider == 'mysql':
#         provider = cfg.get(d_provider, 'provider')
#         host = cfg.get(d_provider, 'host')
#         user = cfg.get(d_provider, 'user')
#         passwd = cfg.get(d_provider, 'passwd')
#         db = cfg.get(d_provider, 'db')
#         charset = cfg.get(d_provider, 'charset')
#         mydb = dict(
#             provider=provider,
#             host=host,
#             user=user,
#             passwd=passwd,
#             db=db,
#             charset=charset)
#
#     return mydb

# 获取配置文件某个值
def getConfig(section, option):
    cfg = ConfigParser()
    cfg.read(MyPath.get_cur_path('config.ini'), encoding="utf-8")

    if section in cfg.sections() and option in cfg.options(section):
        val = cfg.get(section, option)
        return val
    return None


# 修改配置文件某个值
def setConfig(section, option, val):
    cfg = ConfigParser()
    filename = MyPath.get_cur_path('config.ini')
    cfg.read(filename, encoding='utf-8')  # 如果有中文，则需要encoding='utf-8'
    if section in cfg.sections() and option in cfg.options(section):
        cfg.set(section, option, val)
        # 只有当执行conf.write()方法的时候，才会修改ini文件内容
        cfg.write(open(filename, 'r+', encoding="utf-8"))  # set修改ini文件,用r+模式写入
        return True
    return False

# 打印配置文件内容
def printConfig():
    cfg = ConfigParser()
    cfg.read(MyPath.get_cur_path('config.ini'), encoding='utf-8')
    # 获取所有的section
    sections = cfg.sections()
    print(sections)
    for section in sections:
        print(cfg.options(section))
        items = cfg.items(section)
        print(items)


# 初次运行程序，写入config.ini文件
def writeConfig():
    conf = ConfigParser()
    filename = MyPath.get_cur_path('config.ini')
    conf.read(filename, encoding='utf-8')

    # 添加一个section
    conf.add_section('baiduwenku')
    # 往section添加option和value，下载方式
    conf.set('baiduwenku', 'download_mode', 'mode1')

    # 保存路径
    conf.add_section('path')
    conf.set('path', 'savepath', '')

    conf.write(open(filename, 'w'))  # 删除原文件重新写入


# 初始化config.ini文件
def initConfig():
    # 先判断文件是否存在
    if not os.path.exists(MyPath.get_cur_path('config.ini')):
        # 初次运行程序，写入config.ini文件
        writeConfig()


# 恢复默认设置
def defaultConfig():
    # 先判断文件是否存在
    filename = MyPath.get_cur_path('config.ini')
    if not os.path.exists(filename):
        # 初次运行程序，写入config.ini文件
        writeConfig()
    else:
        os.remove(filename)
        writeConfig()
    return True


if __name__ == '__main__':

    # 初始化config.ini文件
    initConfig()

    # 恢复默认设置
    defaultConfig()

    # 打印配置文件内容
    printConfig()

    # print(os.path.dirname('/Users/wenlong/GitHub/exam_demo不上传/copy/内务.xlsx'))
    # print(getConfigSection('mysql'))
    # print(getConfigs())

    # setConfig('path', 'import_path', 'ui')
    # print(getConfig('path', 'import_path'))

    # 路径相关测试
    # oldname = MyPath.get_cur_path('config.py')
    # newname = MyPath.get_cur_path('ui/')
    # print(oldname, newname)
    # print(os.path.isdir(oldname), os.path.isdir(newname), os.path.exists(oldname))
    # MyPath.copy_files(oldname, newname)

