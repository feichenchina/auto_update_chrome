import subprocess
import sys
import urllib.request
import urllib.parse
import winreg  # 和注册表交互
import re  # 正则模块
import zipfile
from configparser import RawConfigParser
import os

class WebDriver():
    def __init__(self,path):
        self.version_re = re.compile(r'^[1-9]\d*\.\d*.\d*')  # 匹配前3位版本号的正则表达式
        self.path = path
        self.ini = self.get_ini()
        self.updateAbsPath = self.ini.get('driver', 'updateAbsPath')
        self.addAbsPath = self.ini.get('driver', 'addAbsPath')

    def get_ini(self):
        _ini = RawConfigParser()
        _ini.read(self.path, encoding='utf-8')
        self.ini = _ini
        return _ini

    # 获取 chrome 版本
    def getChromeVersion(self):
        try:
            # 从注册表中获得版本号
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')

            _v, type = winreg.QueryValueEx(key, 'version')

            print('Current Chrome Version: {}'.format(_v))  # 这步打印会在命令行窗口显示
            self.version = self.version_re.findall(_v)[0]
            return self.version_re.findall(_v)[0]  # 返回前3位版本号

        except WindowsError as e:
            print('check Chrome failed:{}'.format(e))

    # 获取 chromedrive 版本
    def getDriverVersion(self,absPath):
        """
        @param absPath: chromedriver.exe的绝对路径
        """
        cmd = r'{} --version'.format(absPath)  # 拼接成cmd命令

        try:
            # 执行cmd命令并接收命令回显
            out, err = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()
            out = out.decode('utf-8')

            _v = out.split(' ')[1]  # 拆分回显字符串，获取版本号
            print('Current chromedriver Version:{}'.format(_v))

            return self.version_re.findall(_v)[0]
        except IndexError as e:
            print('check chromedriver failed:{}'.format(e))
            return 0

    def downLoadDriver(self,__v, url, save_d):
        # 访问淘宝镜像首页
        rep = urllib.request.urlopen(url).read().decode('utf-8')
        # '<a href="/mirrors/chromedriver/84.0.4147.30/">84.0.4147.30/</a>'
        directory = re.compile(r'>(\d.*?/)</a>').findall(rep)  # 匹配文件夹（版本号）

        # 获取期望的文件夹（版本号）
        match_list = []
        for i in directory:
            v = self.version_re.findall(i)[0]
            if __v == v:
                match_list.append(i)

        # http://npm.taobao.org/mirrors/chromedriver/83.0.4103.39/chromedriver_win32.zip
        dirUrl = urllib.parse.urljoin(url, match_list[-1])

        downUrl = urllib.parse.urljoin(dirUrl, 'chromedriver_win32.zip')  # 拼接出下载路径
        print('will download {}'.format(downUrl))
        print(save_d)
        # 指定下载的文件名和保存位置
        file = os.path.join(save_d, os.path.basename(downUrl))
        print('will saved in {}'.format(file))

        # 开始下载，并显示下载进度(progressFunc)
        urllib.request.urlretrieve(downUrl, file, self.progressFunc)

        # 下载完成后解压
        zFile = zipfile.ZipFile(file, 'r')
        for fileM in zFile.namelist():
            zFile.extract(fileM, os.path.dirname(file))
        zFile.close()

        input('Complete, press Enter to exit')

    def progressFunc(self,blocknum, blocksize, totalsize):
        '''作回调函数用
        @blocknum: 已经下载的数据块
        @blocksize: 数据块的大小
        @totalsize: 远程文件的大小
        '''
        percent = 100.0 * blocknum * blocksize / totalsize

        if percent > 100:
            percent = 100
        downsize = blocknum * blocksize

        if downsize >= totalsize:
            downsize = totalsize

        s = "%.2f%%" % (percent) + "====>" + "%.2f" % (downsize / 1024 / 1024) + "M/" + "%.2f" % (
                    totalsize / 1024 / 1024) + "M \r"
        sys.stdout.write(s)
        sys.stdout.flush()

        if percent == 100:
            print('')

    def checkVersionMatch(self,absPath,temp=True):
        # 读取conf.ini中的配置项
        # absPath = self.ini.get('driver', 'absPath')
        print('Chrome version compares with chromedriver version')
        c_v = self.getChromeVersion()
        d_v = ""
        if temp:
            d_v = self.getDriverVersion(absPath)

        if c_v == d_v:
            # 若匹配，在命令行窗口提示下面的信息
            input(f'path {absPath} Chrome and chromedriver are matched. Press Enter to exit.')
        else:
            # 若不匹配，走下面的流程去下载chromedriver
            _v = c_v

            url = self.ini.get('driver', 'url')  # 从conf.ini中获取淘宝镜像url
            if os.path.isdir(absPath):
                save_d = absPath
            else:
                save_d = os.path.dirname(absPath)  # 下载文件的保存路径，与chromedriver同级
            print("保存位置：",save_d,absPath)
            self.downLoadDriver(_v, url, save_d)  # call下载文件的方法

    def run(self):
        updateAbsPaths = self.updateAbsPath.split(",")
        addAbsPaths = self.addAbsPath.split(",")

        for absPath in updateAbsPaths:
            if absPath:
                if os.path.exists(absPath):
                    self.checkVersionMatch(absPath)
                else:
                    input(f"{absPath} not exists")

        for absPath in addAbsPaths:
            if absPath:
                if os.path.isdir(absPath):
                    self.checkVersionMatch(absPath,False)
                else:
                    input(f"{absPath} is not a directory or {absPath} not exists")

if __name__ == '__main__':
    path = os.path.join(os.getcwd(),"chrome.ini")
    driver = WebDriver(path)
    driver.run()