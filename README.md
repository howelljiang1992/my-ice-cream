# my-ice-cream

受同学所托，要爬一个网站

 第一步很自然的，查看网页元素，看渲染的能不能直接抓下来，

但是用上request，cherrio，一查元素，为空？(页面上明明有)，把request的body打印出来，才发现是ajax，应对方法：打开chrome，用Network，终于查看到了请求，复制链接地址，竟然直接浏览器打开就可以获得到json数据...

接下来就是漫长的数据整理工作了以及一直要弄明白的回调地狱的问题 

