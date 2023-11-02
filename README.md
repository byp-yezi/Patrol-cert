# python 巡检域名证书过期时间

### 1、创建一个名为domains.xlsx的工作表，Sheet中第一列加上需要巡检的域名

| baidu.com       |      |      |      |      |      |
| --------------- | ---- | ---- | ---- | ---- | ---- |
| taobao.com      |      |      |      |      |      |
| jd.com          |      |      |      |      |      |
| aliyun.com      |      |      |      |      |      |
| huaweicloud.com |      |      |      |      |      |

**注意事项：**domains.xlsx中的工作表sheet命名对应main方法中的sheet变量

### 2、运行main.py，生成out.xls文件，如果按照上述的例子执行，则返回的结果大致如下：

| baidu.com       | 2023/12/8 23:59:59  |      |      |      |      |
| --------------- | ------------------- | ---- | ---- | ---- | ---- |
| taobao.com      | 2024/6/8 23:59:59   |      |      |      |      |
| jd.com          | 2024/5/8 23:59:59   |      |      |      |      |
| aliyun.com      | 2023/11/11 23:59:59 |      |      |      |      |
| huaweicloud.com | 2023/12/12 23:59:59 |      |      |      |      |