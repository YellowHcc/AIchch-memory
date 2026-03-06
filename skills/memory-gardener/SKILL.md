---
name: memory-gardener
description: 智能记忆园丁 - 自动汇总每日对话，提炼精华到长期记忆，去除闲聊噪音。
---

# Memory Gardener 🌱

智能记忆管理系统，自动从每日记录中提炼有价值的内容，维护清晰的长期记忆。

## 功能

- 📖 **读取** - 读取当日/指定日期的 memory 文件
- 🧠 **提炼** - 提取关键信息（学到的东西、决定、待办、情感洞察）
- 📝 **汇总** - 生成简洁的 MEMORY.md 格式条目
- 🔄 **备份** - 可选：备份原始记录后清理

## 使用方法

### 手动整理今天
```bash
cd ~/.openclaw/workspace
./skills/memory-gardener/garden.sh today
```

### 整理指定日期
```bash
./skills/memory-gardener/garden.sh 2026-03-06
```

### 整理最近7天（批量）
```bash
./skills/memory-gardener/garden.sh week
```

### 查看 MEMORY.md 预览（不写入）
```bash
./skills/memory-gardener/garden.sh preview
```

## 输出格式

MEMORY.md 更新示例：
```markdown
## 2026-03-06

**学到的**:
- 焦虑型依恋的核心信念："如果我不主动，他就会离开我"
- 延迟回复实验证明焦虑会自然下降

**决定的**:
- 重新设置GitHub自动备份（每天23:00）

**待办**:
- [ ] 小红书文案发布

**标签**: #情感自救 #心理成长 #工具配置
```

## 集成到备份

在 `AIchch-memory/backup.sh` 中加入：
```bash
# 先整理，再备份
cd ~/.openclaw/workspace
./skills/memory-gardener/garden.sh today
```

## 配置

编辑 `~/.openclaw/workspace/skills/memory-gardener/config.json`：
```json
{
  "workspace": "/root/.openclaw/workspace",
  "keep_raw_days": 30,
  "categories": ["学到的", "决定的", "待办", "情感", "项目"],
  "auto_backup_after_garden": true
}
```
