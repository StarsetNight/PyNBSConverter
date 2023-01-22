# PyNBSConverter
## 开发说明 DEVELOPING INSTRUCTIONS
我马上就要建一个`AdvancedDev`分支，以后我的最新改动都写在这条分支里。  
本项目基于`pynbs`和`xlwings`前置Python库。  
## 介绍 INTRODUCE
众所周知，将传统的Minecraft Noteblock Studio（后文用“NBS”代替）工程转为 **“第四代编码格式”** 是一件吃力不讨好的事情，于是，本人便开发了这款全自动转译脚本。  
目前我就写了这么些功能，以后还会继续更新的，同时，开源永远都是最好的！  
## 注意 NOTIFY
输入的文件需要经过预处理，处理步骤如下：  
1. 开头的音符转换成编码肯定放不下，在NBS软件内使用`Ctrl`+`A`全选音符，并后移。  
2. 本程序暂时无法进行更精细的编码压缩，请先尽量删除和弦轨道。  
3. 输入的NBS文件最好为10t/s（这个没有测试如果不这样设置会不会出问题，应该没事）。  

输出的文件为`.xlsx`文档，使用`WPS表格`或`Microsoft Excel`即可打开，编码代表内容如下：
- 蓝色：普通音符，也就是要在编码器播放的音符。  
- 黄色：无延迟的执行编码。  
- 绿色：有延迟的执行编码，有几个就代表延迟几个gt。  

## 如何构建 HOW TO BUILD
将本项目使用Git的`git clone`命令克隆至本地，然后使用`pip install -r requirements.txt`安装所有依赖库，即可运行本程序。  
## 下载 DOWNLOADS
（暂无下载，源码开放下载）  
## 贡献者名单 CONTRIBUTORS
- Advanced_Killer: [GitHub](https://github.com/ThirdBlood) 
[BiliBili](https://space.bilibili.com/477677552)
## 鸣谢名单 THANKS
（暂无鸣谢名单）  
**Powered by Advanced_Killer**  