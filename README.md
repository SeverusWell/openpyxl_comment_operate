[openpyxl Excel表格批注操作](https://gitlab.com/newonezero/openpyxl_comment_operate)
==

> 用于优化Excel表格批注修改操作

### 环境初始化

#### 安装环境

  略

#### 安装依赖

  略



### 拆分

![](https://ws2.sinaimg.cn/large/006tNc79ly1g03hfzg88sj306p047gln.jpg)

  将原始Excel表格更名为`source.xlsx`放到项目目录下,执行拆分命令:

```
python openpyxl_split.py
```

  执行成功后将会生成文件`split.xlsx`,之后在此文件中进行修改操作.

![image-20190212113319583](https://ws4.sinaimg.cn/large/006tNc79ly1g03hjhnl0gj309u08hq36.jpg)

### 合并

  确保文件`split.xlsx`依然存在,执行命令

```
python openpyxl_merge.py
```

  执行成功后生成合并后的目标文件`source.xlsx`.

![image-20190212113346005](https://ws4.sinaimg.cn/large/006tNc79ly1g03hjms57lj30bf083aao.jpg)
