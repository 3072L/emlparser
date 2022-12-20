# emlparser

rust 实现的一个eml邮件格式解析的小工具，用来统计各个发件人出现的次数

主要逻辑是从文件夹(包括子文件夹)中的eml文件的headers中提取From，Sender，Reply-To等键值对的value值，并存在hashmap中，最后再写入到表格文件中
