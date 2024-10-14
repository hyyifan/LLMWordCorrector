# WordLLMChecker: 专业的文档纠错助手

## 项目简介

WordLLMChecker 是一款基于 Python 开发的 Windows 端应用，主要是利用阿里云百炼大模型[qwen-plus API](https://bailian.console.aliyun.com/?spm=5176.29311086.J_RY_4Q8--sru4dMV7o3lqS.1.24873123CqDVFV#/model-market/detail/qwen-plus-latest)提供中文错词错句纠错功能。这款软件旨在帮助用户自动检测和纠正文档中的语法错误、错别字和用词不当，能够直接在原文上进行批注，从而提升文档的专业性和准确性。

## 功能特点

- **自动读取与分析**：用户只需上传 docx 或 doc 格式的文件，WordLLMChecker 即可自动读取内容并分析潜在的语文错误。
- **智能批注**：软件会在原文上直接进行批注，指出错误并提供修改建议，使错误一目了然。
- **数据隐私保护**：本软件不做任何数据库存储，直接通过互联网，将您的文本交由 qwen-plus进行分析。

## 开源贡献

我们鼓励开源社区的贡献者参与到 WordLLMChecker 的开发中来。无论是代码改进、bug 修复还是新功能的建议，我们都欢迎您的参与。请访问我们的 GitHub 仓库了解更多详情。

## 如何使用

1. **安装**：下载并安装 WordLLMChecker 应用，链接：
2. **上传文件**：打开应用，上传您的 docx 或 doc 格式文件。
![image](https://github.com/user-attachments/assets/19ee362d-6472-4d83-898e-94cc2460c073)
3.**输入Qwen大模型key**: 前往[阿里云百炼大模型平台](https://bailian.console.aliyun.com/?spm=5176.29311086.J_RY_4Q8--sru4dMV7o3lqS.1.24873123nvuVmw#/home)注册账号，获得api key，并进行输入。
   ![image](https://github.com/user-attachments/assets/9f8f4d74-ec4c-4029-9602-a88df6b9ea7d)
4.**开始处理**：点击后软件即开始运行，运行过程之中请关注进度内容。
5. **自动纠错**：软件将自动分析文件内容，并在原文上进行批注。
6. **查看结果**：打开你的word文件，并关注进度内容。 目前存在技术难题，部分批注无法显示在word文件之中，需要在进度栏查看。

 
## 联系我们

如果您有任何问题或想要了解更多关于 WordLLMChecker 的信息，请通过以下方式联系我们：

- **邮箱**：hyyifan7@163.com
- **GitHub 仓库**：https://github.com/hyyifan/LLMWordCorrector
