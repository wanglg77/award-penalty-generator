# GitHub操作注意事项

## 用户环境特点
- 网络封锁了SSH(22)和直接HTTPS(443)
- 但GitHub CLI (`gh`) 可以正常工作

## 正确流程
1. 使用 `gh` CLI 操作GitHub，不要用MCP工具
2. 创建仓库：`gh repo create <name> --public --source=. --push`
3. 如果仓库已存在：
   ```bash
   gh repo clone wanglg77/<repo-name>
   # 复制文件进去
   cd <repo-name>
   git add . && git commit -m "message"
   git push
   ```

## 失败的尝试
- MCP create_repository: 认证失效
- HTTPS git push: schannel handshake失败
- SSH git push: port 22被封锁
