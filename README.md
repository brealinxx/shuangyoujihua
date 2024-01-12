Hey 这是一个练习项目，通过特定 excel 生成特定图表。

## 实例

软件和示例表格已经打包好，请点击 [Release](https://github.com/brealinxx/shuangyoujihua/releases/tag/v1.0) 下载

## 手动编译

以 Visual Studio Code 演示

### 拉取项目

`git clone https://github.com/brealinxx/shuangyoujihua.git`

### 进入项目目录

`cd shuangyoujihua`

### 创建 python 虚拟环境

1. Mac 按下 `command` + `shift` + `p` 打开面板再输入 `Python: Create Environment`，选择 `.venv`，建议选择 python 3.9+ 版本的解释器创建环境

2. 输入命令 `source path/to/your/.venv/bin/activate`

>记得修改为自己的路径再激活 activate

3. 安装相关依赖，输入命令 `pip3 install -r requirements.txt` 

### 根据特定要求修改 matplotlib 包的源码

> 本项目需要特定实现文字出现在饼图内且要控制字体大小

修改 `path/to/your/.venv/lib/python3.12/site-packages/matplotlib/axes/_axes.py` 脚本

#### Part1

位于 3260 行，是 `textprops = {}`

在花括号可以加入 `'fontsize': 20`,`'color' = 'white'`，以控制字体大小和颜色

#### Part2

位于 3313 行，**注释以下内容**

```python
if isinstance(autopct, str):
	s = autopct % (100. * frac)
elif callable(autopct):
	s = autopct(100. * frac)
else:
  raise TypeError(
    'autopct must be callable or a format string')
```

紧接着在下面一行**再加入一行代码** `s=label`

### 加入中文字体包

找到位于本项目的 Assets 文件夹下的 `SimHei.ttf` 字体文件，将它移动到 `/path/to/your/.venv/lib/pythonX.X/site-packages/matplotlib/mpl-data/fonts/ttf/` 文件夹下

### 编译

输入命令 `pyinstaller launch.spec`

编译后的应用程序位于当前文件夹下 `dist/双优计划生成器`

## License

MIT
