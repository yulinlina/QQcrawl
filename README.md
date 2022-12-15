# QQcrawl
## 使用方法
运行`main.py` 即可
## 适用条件
**首先得用QQ打开需要爬取的群** 【重要】
## 爬取不同的群
``` python
if __name__ =="__main__":

    grouptext = ["2022Fall_数据库设计_小组"]        #  需要爬取的群组 确保提前打开
    qq =QQhandle()
    qq.crawl(grouptext)
```
修改grouptext即可
## 
一个是`txt` 文档  
另一个是excel表
![image](https://user-images.githubusercontent.com/77262518/207834349-c6fbc735-b31a-4908-b2a1-b2c70ad389ec.png)
