# GitHub 使用最佳实践

这份文档用于后续新项目快速复用，目标是减少 GitHub 连接失败、认证混乱、私有仓库拉取失败、推送失败等问题。

## 推荐结论

优先使用两种稳定方式：

1. **个人长期开发：SSH key**
2. **脚本、自动化、Codex/命令行临时访问：GitHub Personal Access Token**

不建议使用：

- GitHub 账号密码。GitHub 已不支持用账号密码直接访问 Git 仓库。
- 把 token 写死在仓库文件、README、脚本或命令历史里。
- 多个项目混用不清楚来源的 token、代理、credential helper。

## 为什么有的项目 GitHub 不好用

常见原因：

- 远端地址不对，例如本地项目指向了旧仓库或同名但不同内容的仓库。
- 私有仓库没有访问权限。
- 使用 HTTPS 但本机没有保存有效 token。
- token 过期、权限不足，或没有 `repo` 权限。
- 使用 SSH 但没有配置 key，或 GitHub 没有登记对应公钥。
- 本机代理环境变量错误，例如 `HTTP_PROXY`、`HTTPS_PROXY`、`ALL_PROXY` 指向不可用端口。
- Git credential helper 里缓存了旧账号或旧 token。
- 本地分支和远端分支历史无关，需要明确是合并、拉取还是覆盖。
- GitHub 仓库是 private，但同事没有 collaborator 权限。

## 每个项目先检查这些

进入项目目录后：

```bash
git remote -v
git branch --show-current
git status --short --branch
git log --oneline -5
```

确认：

- `origin` 是否指向正确仓库。
- 当前分支是否是预期分支，通常是 `main`。
- 工作区是否有未提交改动。
- 本地提交历史是否是这个项目自己的历史。

查看远端分支关系：

```bash
git fetch origin
git status --short --branch
git log --oneline --decorate --graph --left-right origin/main...main
```

## HTTPS + Token 方式

适合：

- Codex/脚本访问 GitHub。
- 私有仓库拉取或推送。
- 不想配置 SSH key 的环境。

### Token 权限

私有仓库常用权限：

- 经典 token：需要 `repo` 权限。
- Fine-grained token：至少给目标仓库 `Contents: Read and write`。

如果只是只读安装 skill 或 clone 私有仓库，只需要读权限。

### 不推荐的写法

不要把 token 长期写进 remote URL：

```bash
git remote set-url origin https://TOKEN@github.com/user/repo.git
```

原因：

- 容易进入 `.git/config`。
- 容易泄露到日志、截图、命令历史。
- 换 token 时排查困难。

### 推荐：使用 Git Credential Manager

第一次访问时，让 Git 弹出登录或保存 token。之后本机安全存储凭据。

检查 credential helper：

```bash
git config --global credential.helper
```

Windows 上常见：

```bash
manager
manager-core
```

Linux/WSL 可按环境配置。若使用 Windows Git Credential Manager，优先让 Windows Git 统一保存凭据。

### 推荐：脚本中使用环境变量

临时访问 GitHub API：

```bash
curl -H "Authorization: Bearer $GITHUB_TOKEN" \
  -H "Accept: application/vnd.github+json" \
  https://api.github.com/repos/OWNER/REPO
```

临时 push 时可以使用一次性 URL，但不要保存到 remote：

```bash
git push https://x-access-token:$GITHUB_TOKEN@github.com/OWNER/REPO.git main
```

注意：

- 不要把这种命令写进公开文档。
- 不要提交包含 token 的脚本。
- shell 历史可能记录命令，敏感环境下要谨慎。

## SSH Key 方式

适合：

- 长期开发。
- 经常 push/pull。
- 不想频繁处理 HTTPS token。

生成 key：

```bash
ssh-keygen -t ed25519 -C "your_email@example.com"
```

查看公钥：

```bash
cat ~/.ssh/id_ed25519.pub
```

把公钥添加到 GitHub：

```text
GitHub -> Settings -> SSH and GPG keys -> New SSH key
```

测试：

```bash
ssh -T git@github.com
```

项目 remote 使用：

```bash
git remote set-url origin git@github.com:OWNER/REPO.git
```

常见问题：

- WSL 和 Windows 是两套 SSH 环境，key 不一定共享。
- 多个 GitHub 账号需要配置 `~/.ssh/config`。
- 公司网络可能拦截 SSH，需要改用 HTTPS。

## 私有仓库给同事使用

推荐优先级：

1. **邀请同事成为 collaborator**
2. **发 release/zip 包**
3. **只读 deploy key**
4. **临时只读 token**

不建议把自己的长期 token 发给同事。

### Collaborator 方式

GitHub 仓库：

```text
Settings -> Collaborators -> Add people
```

同事接受邀请后，用自己的账号拉取：

```bash
git clone https://github.com/OWNER/REPO.git
```

或：

```bash
git clone git@github.com:OWNER/REPO.git
```

### Zip 包方式

如果同事只需要使用 skill，不需要参与开发，最简单是发 zip：

```text
dist/claude-word-style-standardization.zip
```

优点：

- 不需要 GitHub 权限。
- 不暴露 token。
- 不需要处理 git 冲突。

缺点：

- 更新时需要重新发新版 zip。

## 新项目标准流程

### 本地已有项目，推到 GitHub 新仓库

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/OWNER/REPO.git
git push -u origin main
```

如果用 SSH：

```bash
git remote add origin git@github.com:OWNER/REPO.git
git push -u origin main
```

### 本地项目覆盖 GitHub 同名仓库

先确认远端：

```bash
git remote -v
git fetch origin
git log --oneline --decorate --graph --left-right origin/main...main
```

如果远端历史和本地是同一条线，只需要：

```bash
git push origin main
```

如果远端是无关项目，且你明确要用本地覆盖远端：

```bash
git push --force-with-lease origin main
```

注意：

- `--force-with-lease` 比 `--force` 更安全。
- 覆盖前最好确认远端没有需要保留的内容。
- 不要在没确认的情况下强推。

## 常用排查命令

检查远端：

```bash
git remote -v
```

检查当前登录或 token 是否能访问仓库：

```bash
curl -H "Authorization: Bearer $GITHUB_TOKEN" \
  -H "Accept: application/vnd.github+json" \
  https://api.github.com/repos/OWNER/REPO
```

返回中重点看：

```json
"private": true,
"visibility": "private",
"permissions": {
  "pull": true,
  "push": true
}
```

检查代理环境变量：

```bash
env | grep -i proxy
```

临时绕过代理：

```bash
env -u ALL_PROXY -u HTTPS_PROXY -u HTTP_PROXY curl https://api.github.com
```

检查本地和远端提交：

```bash
git rev-parse HEAD origin/main
```

查看是否领先/落后：

```bash
git status --short --branch
```

## Token 保存建议

安全优先级：

1. GitHub 官方 credential manager / 系统钥匙串
2. SSH key
3. 环境变量，例如 `$GITHUB_TOKEN`
4. 临时命令行传入 token
5. 写入文件或 remote URL，尽量避免

如果一定要本机保存 token：

- 保存到系统凭据管理器，而不是项目文件。
- 不要放进 `.env` 后提交。
- 如果放 `.env`，必须确保 `.gitignore` 包含 `.env`。
- token 权限尽量小，能只读就不要给写权限。
- 定期轮换 token。

## 本机建议配置

设置默认分支名：

```bash
git config --global init.defaultBranch main
```

设置用户名邮箱：

```bash
git config --global user.name "Your Name"
git config --global user.email "your_email@example.com"
```

开启凭据管理器，具体值按环境决定：

```bash
git config --global credential.helper manager
```

或使用 SSH，避免 HTTPS 凭据问题。

## 对 Codex/自动化的建议

- 优先使用环境变量 `$GITHUB_TOKEN`。
- 不把 token 写入仓库。
- 每次操作前先看：

```bash
git status --short --branch
git remote -v
```

- 推送前确认：

```bash
git diff --stat
git log --oneline -3
```

- 私有仓库查询优先用 GitHub API 验证权限。
- 如果网络失败，先检查代理，再决定是否绕过代理。

## 快速模板

新项目推送：

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/OWNER/REPO.git
git push -u origin main
```

已有项目更新：

```bash
git status --short --branch
git add .
git commit -m "Update project"
git push origin main
```

确认仓库私有/公开：

```bash
curl -H "Authorization: Bearer $GITHUB_TOKEN" \
  -H "Accept: application/vnd.github+json" \
  https://api.github.com/repos/OWNER/REPO
```

看返回：

```json
"private": true,
"visibility": "private"
```
