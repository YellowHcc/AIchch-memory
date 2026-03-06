#!/bin/bash

# 🤖 AIchch 记忆备份脚本
# 用法: ./backup.sh

set -e

REPO_URL="https://github.com/YellowHcc/AIchch-memory.git"
WORKSPACE="/root/.openclaw/workspace"
REPO_DIR="$WORKSPACE/AIchch-memory"

echo "🔄 开始备份记忆文件..."

# 检查是否在正确的目录
if [ ! -d "$REPO_DIR" ]; then
    echo "❌ 错误: 找不到仓库目录 $REPO_DIR"
    echo "请先克隆仓库: git clone $REPO_URL"
    exit 1
fi

# 🌱 第一步：整理今日记忆
echo "🌱 整理今日记忆..."
cd "$WORKSPACE"
./skills/memory-gardener/garden.sh today 2>/dev/null || echo "⚠️ 整理记忆跳过或失败"

cd "$REPO_DIR"

# 拉取最新变更（避免冲突）
echo "📥 拉取远程更新..."
git pull origin main 2>/dev/null || echo "⚠️ 可能是空仓库，跳过拉取"

# 同步记忆文件
echo "📋 同步记忆文件..."

# 复制 workspace 的记忆文件（如果存在）
if [ -f "$WORKSPACE/MEMORY.md" ]; then
    cp "$WORKSPACE/MEMORY.md" ./MEMORY.md
    echo "  ✓ 已同步 MEMORY.md"
fi

# 复制每日记忆
if [ -d "$WORKSPACE/memory" ]; then
    cp -r "$WORKSPACE/memory/"* ./memory/ 2>/dev/null || true
    echo "  ✓ 已同步 memory/ 目录"
fi

# 复制 skills/memory-gardener
if [ -d "$WORKSPACE/skills/memory-gardener" ]; then
    mkdir -p ./skills/memory-gardener
    cp -r "$WORKSPACE/skills/memory-gardener/"* ./skills/memory-gardener/ 2>/dev/null || true
    echo "  ✓ 已同步 memory-gardener skill"
fi

# 复制项目文件
if [ -f "$WORKSPACE/notes.html" ]; then
    cp "$WORKSPACE/notes.html" ./projects/notes/index.html
    echo "  ✓ 已同步 notes.html"
fi

# 复制调研报告（支持多种命名）
for file in "$WORKSPACE"/*调研报告.md "$WORKSPACE"/*趋势*.md; do
    if [ -f "$file" ]; then
        cp "$file" ./
        echo "  ✓ 已同步 $(basename "$file")"
    fi
done

# 提交变更
git add -A

if git diff --cached --quiet; then
    echo "✅ 没有变更需要提交"
else
    git commit -m "🤖 auto: memory backup $(date +'%Y-%m-%d %H:%M:%S')"
    git push origin main
    echo "✅ 备份完成！已推送到 GitHub"
fi
