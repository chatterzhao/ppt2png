# ppt2png
ppt转图片和pdf，并且可以拼接成长图

# 在 python2 或 python3 版本上运行都可以

# 操作
克隆本项目到本地，把待转图片的 PPT 文件放到本项目的根目录下，也就是与 .py 文件并排放置，打开Visual Studio Code 编辑器，打开本项目目录（文件-打开文件夹），项目在 VSCode 加载后，在VSCode 的终端（顶部菜单-终端）上执行命令`python ppt2png.py` 或 `python3 ppt2png.py`,如 python 或代码中用到的库之前在你其他项目都安装正常，则代码会正常执行，输出图片，如果报错，则根据报错进行处理，下面就不将 python 本身的报错，那一般是你还没安装 python 到官网下载安装即可，建议安装 python3，下面主要将库的报错的解决办法。 

## 一、安装 win32com
win32com 是 python 的一个第三方库，它是操作 Windows 的 com 库。前一条命令是在python2下安装，后一条是在python3下安装。先在命令行执行：`python -m pip install pypiwin32` 或 `python3 -m pip install pypiwin32` 安装 win32com

## 二、安装 图像库 Pillow
我们要安装一个第三方库—— Python Imaging Library，这是 Python 下非常强大的处理图像的工具库。不过，PIL 目前只支持到 Python 2.7，并且有年头没有更新了，因此我们要装Pillow，它是基于 PIL 的，Pillow 项目开发非常活跃，并且同时支持最新的 Python 3。
给 Python 安装 Pillow 非常简单。

在命令行使用 PIP 安装：
` pip install Pillow `
 或 ` pip3 install Pillow ` 或 `python3 -m pip install Pillow`

安装完成后即可。安装完后，项目中的代码不用修改，但是代码中的疑惑这里要解释一下,注意还是 from PIL 而不是 from Pillow，也就是下面这句代码，不用修改，继续使用它。
``` 
from PIL import Image
```
```
#python2 
import Image 
#python3（因为是派生的PIL库，所以要导入PIL中的Image） 
from PIL import Image
```

## 如果提示 pip 版本过低，执行下面的命令
升级 pip `python -m pip install --upgrade pip`
或升级 pip3，但是这里给出升级 pip3 的命令 `python3 -m pip install --upgrade pip`，注意pip3 是对应 python3 的。

## 三、将PPT放置在本项目的根目录，也就是与.py文件并列，然后在终端上 cd 到这个目录，然后执行命令：
`python ppt2png.py` 或 `python3 ppt2png.py`,执行完毕，会生成一个目录（里面是每页PPT的图片），还有一个pdf文件，有一个合并的长图

# 待解决的问题
## 一、选择待生成长图的PPT文件，而不是将文件放在项目文件夹下
## 二、最后删除每页图片的目录
## 三、如果不是指定转PDF文件，不转
## 四、bug 拼接的长图图片顺序不对
