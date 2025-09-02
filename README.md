## 使用流程

### 1. 创建虚拟环境
终端运行
`python -m venv venv_name`
### 2. 激活虚拟环境
等待虚拟环境创建完成后激活虚拟环境
`.\venv_name\Scripts\activate`
### 3. 安装依赖
`pip install -r requirements.txt`
建议使用镜像安装
`pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple`
### 4. 设置本地打印机名称
* 打开打印机和扫描仪设置界面，找到要使用的打印机，
![alt text](image-1.png)
* 修改打印机名称为“辅助打证”
![alt text](image.png)
### 5. 运行代码
终端运行
`python app.py`

***

## 注意事项

1. 本地需要安装有Edge浏览器，且浏览器版本为：139.0.3405.125（在浏览器设置中查看）
2. 目前只有法人登录的第一个证件类型（食品安全许可证）可以正常使用。
   ![alt text](image-2.png)
3. 单个账号单日连续5次密码错误会被锁定
