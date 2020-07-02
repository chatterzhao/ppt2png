# ppt2png
ppt转图片和pdf，并且可以拼接成长图

# 在 python2 版本上运行即可

# 操作


## 一、安装 win32com
先在命令行执行：`python -m pip install pypiwin32` 安装 win32com，win32com 是 python 的一个第三方库，它是操作 Windows com 的库

## 二、安装 图像库 Pillow
我们要安装一个第三方库—— Python Imaging Library，这是 Python 下非常强大的处理图像的工具库。不过，PIL 目前只支持到 Python 2.7，并且有年头没有更新了，因此，基于 PIL 的 Pillow 项目开发非常活跃，并且支持最新的 Python 3。
给 Python 安装 Pillow 非常简单，使用 pip 或 easy_install 只要一行代码即可。

在命令行使用 PIP 安装：
` pip install Pillow `

或在命令行使用 easy_install 安装：
` easy_install Pillow `

安装完成后即可。安装完后，项目中的代码不用修改，但是代码中的疑惑这里要解释一下,注意还是 from PIL 而不是 from Pillow，也就是下面这句代码，不用修改，继续使用它。
``` 
from PIL import Image
```

## 如果提示 pip 版本过低，执行下面的命令
升级 pip `python -m pip install --upgrade pip`
本次不用升级 pip3，但是这里给出升级 pip3 的命令 `python3 -m pip install --upgrade pip`，注意pip3 是对应 python3 的。

## 三、将PPT放置在本项目的根目录，也就是与.py文件并列，然后在终端上 cd 到这个目录，然后执行命令：
`python ppt2png.py`,执行完毕，会生成一个目录（里面是每页PPT的图片），还有一个pdf文件，有一个合并的长图

# 待解决的问题
## 一、选择待生成长图的PPT文件，而不是将文件放在项目文件夹下
## 二、最后删除每页图片的目录
## 三、如果不是指定转PDF文件，不转
