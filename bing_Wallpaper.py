import requests, json, os
import win32api, win32gui, win32con
import sys

path = "D:/Bing_Wallpaper/"


def Httpget():
    url = 'http://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=zh-CN'
    res = requests.get(url).text
    res = json.loads(res).get('images')[0].get('url')
    return res


def getimage(url):
    try:
        # print(url)
        name = url.replace('/az/hprichbg/rb/', '')
        # name=re.match(r'(.*)1920x1080.jpg',url).group()
        # print(name)
        image_url = 'http://www.bing.com' + url
        res = requests.get(image_url)
        if res.status_code == 200:
            open(path + name, 'wb').write(res.content)
            if os.access(path + name, os.F_OK):
                print('图片下载成功!')
            else:
                print('图片下载失败！')
        return path + name
    except:
        print('图片下载失败！')


def setting_wallpaper(image_path):
    try:
        k = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0, win32con.KEY_SET_VALUE)
        win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, image_path, win32con.SPIF_SENDWININICHANGE)
        return True
    except:
        return False


def setting():
    try:
        if os.path.exists("D:\Bing_Wallpaper") == True:
            pass
            return True
        else:
            print('目录不存在！')
            print('创建目录...')
            os.mkdir("D:\Bing_Wallpaper")
            if os.path.exists("D:\Bing_Wallpaper") == True:
                print('目录创建成功！')
                return True
            else:
                print('目录创建失败！')
                return False
    except:
        print('程序初始化异常！')
        return False


if __name__ == "__main__":
    print('程序初始化...')
    k = setting()
    if k == True:
        print('程序初始化成功！')
    else:
        print('程序初始化失败，请检查权限设置！')
    print('正在获取请求...')
    image_url = Httpget()
    if image_url == "":
        print('请求失败！')
    else:
        print('请求成功！')
    print('正在下载图片...')
    image_path = getimage(image_url)
    print('正在设置桌面壁纸...')
    result = setting_wallpaper(image_path)
    if result == True:
        print('设置壁纸成功！')
    else:
        print('设置壁纸失败！')
    print('程序退出...')
    sys.exit(2)
