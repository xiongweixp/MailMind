# 智邮 MailMind

MailMind 是一个基于 WinUI 3 的 Windows 桌面邮件客户端。它通过 IMAP 连接邮箱，提供邮件浏览、阅读和本地缓存；AI 助手与邮件发送界面目前仍在开发中。

## 当前功能

- 配置一个邮箱账户，保存前测试 IMAP 连接，支持 SSL/TLS 和 STARTTLS。
- 浏览邮箱文件夹、按日期分组查看邮件，滚动加载邮件列表。
- 阅读纯文本和 HTML 邮件，显示未读状态并尝试将已读状态同步到服务器。
- 在已加载邮件中搜索和按日期筛选。
- 使用 SQLite 缓存邮件目录、列表和已读取的正文，连接失败时可查看已有缓存。

> 写邮件界面和 SMTP 配置已提供，但发送尚未实现；回复、转发、删除、归档以及 AI 搜索和助手功能尚未启用。

## 环境要求

- Windows 10（最低版本 1809，内部版本 17763）或更高版本。
- Visual Studio，安装 .NET 桌面开发、Windows 应用开发相关组件和 Windows SDK。
- .NET 8 SDK；项目使用 Windows App SDK 和 WinUI 3，支持 x86、x64 和 ARM64。
- 可访问 IMAP 的邮箱，必要时需在邮箱服务商处启用 IMAP 并获取授权码。

## 构建和运行

1. 使用 Visual Studio 打开根目录下的 `MailMind.slnx`。
2. 等待 NuGet 包还原，选择与设备匹配的平台（例如 `x64`）。
3. 选择 `MailMind` 应用项目或打包项目作为启动项目，生成并运行。

解决方案同时包含 Windows 应用打包项目，可使用 Visual Studio 的打包工具创建应用包。项目依赖的主要 NuGet 包包括 Windows App SDK、MailKit 和 Microsoft.Data.Sqlite。

## 配置邮箱

1. 启动应用，点击左下角的“配置”。
2. 填写邮箱地址、用户名、IMAP 服务器和端口，以及密码或邮箱授权码；按服务商要求选择连接方式。
3. 可填写 SMTP 服务器信息供后续使用。保存时应用会测试 IMAP 连接，但不会测试或发送 SMTP 邮件。
4. 连接成功后，在左侧选择文件夹并点击邮件阅读。首次离线使用前，需要先联网加载相应邮件和正文。

服务器地址、端口、加密方式和授权码获取方式请以邮箱服务商的说明为准。取消勾选“使用 SSL/TLS”时，应用使用 STARTTLS。

## 本地数据

账户配置保存在 `%LOCALAPPDATA%\MailMind\account.json`，邮件缓存保存在同一目录下的 `mailmind.db`。密码或授权码通过 Windows 凭据管理器保存。缓存可能包含邮件正文，请按个人数据妥善保护设备和备份。

## 项目结构

```text
MailMind.slnx                       Visual Studio 解决方案
MailMind/MailMind/                  WinUI 应用
  MainWindow.xaml(.cs)              界面和交互
  MailService.cs                   IMAP 账户与邮件服务
  MailCache.cs                     SQLite 缓存
  InlineImageResolver.cs           邮件内嵌图片处理
MailMind/MailMind (Package)/        Windows 应用打包项目
docs/设计文档/                     需求、概要设计和界面设计资料
```
