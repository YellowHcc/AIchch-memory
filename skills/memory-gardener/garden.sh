#!/bin/bash
#
# 🌱 Memory Gardener - 智能记忆整理脚本
# 用法: ./garden.sh [today|YYYY-MM-DD|week|preview|all]

set -e

# 配置
WORKSPACE="/root/.openclaw/workspace"
MEMORY_DIR="$WORKSPACE/memory"
MEMORY_FILE="$WORKSPACE/MEMORY.md"

log_info() { echo "ℹ $1"; }
log_success() { echo "✓ $1"; }
log_warn() { echo "⚠ $1"; }
log_error() { echo "✗ $1"; }

# 显示帮助
show_help() {
    cat << 'EOF'
🌱 Memory Gardener - 智能记忆整理

用法:
  ./garden.sh today          整理今天的记忆
  ./garden.sh YYYY-MM-DD     整理指定日期
  ./garden.sh week           整理最近7天
  ./garden.sh all            整理所有历史记忆
  ./garden.sh help           显示此帮助
EOF
}

# 解析记忆文件
parse_memory() {
    local file="$1"
    local date=$(basename "$file" .md)
    
    if [ ! -f "$file" ]; then
        return 1
    fi
    
    # 读取文件内容
    local content=$(cat "$file")
    
    # 提取学到的（包含"学到"、"了解"、"知道"、"发现"的行）
    local learned=$(echo "$content" | grep -E "(学到|了解|知道|发现|原来|明白了)" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' | head -5 || true)
    
    # 提取决定/做的（包含"决定"、"设置"、"创建"、"完成"的行）
    local decided=$(echo "$content" | grep -E "(决定|设置|创建|完成|搞定|重新设置)" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' | head -5 || true)
    
    # 提取情感/洞察（包含"感觉"、"焦虑"、"依恋"、"害怕"）
    local emotions=$(echo "$content" | grep -E "(感觉|焦虑|依恋|害怕|恐惧|喜欢|爱|情绪|上头)" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' | head -5 || true)
    
    # 提取项目/工具（包含"项目"、"工具"、"脚本"、"备份"）
    local projects=$(echo "$content" | grep -E "(项目|工具|脚本|备份|skill|Skill|仓库|GitHub)" | sed 's/^[[:space:]]*//;s/[[:space:]]*$//' | head -5 || true)
    
    # 提取标签
    local tags=$(echo "$content" | grep -o '#[^ ]*' | grep -E '^#[A-Za-z]' | sort -u | tr '\n' ' ' || echo "")
    
    # 生成格式化输出
    local output="## $date"
    local has_content=false
    
    if [ -n "$learned" ]; then
        output="$output

**学到的**:
$(echo "$learned" | sed 's/^/- /')"
        has_content=true
    fi
    
    if [ -n "$decided" ]; then
        output="$output

**决定的**:
$(echo "$decided" | sed 's/^/- /')"
        has_content=true
    fi
    
    if [ -n "$emotions" ]; then
        output="$output

**情感洞察**:
$(echo "$emotions" | sed 's/^/- /')"
        has_content=true
    fi
    
    if [ -n "$projects" ]; then
        output="$output

**项目/工具**:
$(echo "$projects" | sed 's/^/- /')"
        has_content=true
    fi
    
    if [ -n "$tags" ]; then
        output="$output

**标签**: $tags"
    fi
    
    if [ "$has_content" = true ]; then
        echo "$output"
        return 0
    else
        return 1
    fi
}

# 更新 MEMORY.md
update_memory() {
    local entry="$1"
    
    # 确保 MEMORY.md 存在
    if [ ! -f "$MEMORY_FILE" ]; then
        log_info "创建 MEMORY.md..."
        cat > "$MEMORY_FILE" << 'HEADER'
# 🧠 长期记忆

_自动汇总的精华记忆，由 Memory Gardener 维护_

HEADER
    fi
    
    # 检查是否已存在该日期
    local date=$(echo "$entry" | head -1 | sed 's/## //')
    if grep -q "^## $date$" "$MEMORY_FILE" 2>/dev/null; then
        log_warn "$date 已存在于 MEMORY.md，跳过"
        return 1
    fi
    
    # 找到标题后的空行，插入新内容
    local temp_file=$(mktemp)
    
    # 写入头部（前3行）
    head -3 "$MEMORY_FILE" > "$temp_file"
    echo "" >> "$temp_file"
    
    # 写入新条目
    echo "$entry" >> "$temp_file"
    echo "" >> "$temp_file"
    
    # 追加原有内容（从第4行开始）
    tail -n +4 "$MEMORY_FILE" >> "$temp_file" 2>/dev/null || true
    
    mv "$temp_file" "$MEMORY_FILE"
    log_success "已更新 MEMORY.md"
}

# 处理单个文件
process_file() {
    local file="$1"
    local date=$(basename "$file" .md)
    
    log_info "正在整理: $date"
    
    local entry=$(parse_memory "$file")
    if [ $? -eq 0 ]; then
        update_memory "$entry"
        return 0
    else
        log_warn "$date 没有可提取的精华内容"
        return 1
    fi
}

# 处理今天的记忆
garden_today() {
    local today=$(date +%Y-%m-%d)
    local file="$MEMORY_DIR/$today.md"
    
    if [ ! -f "$file" ]; then
        log_warn "今天还没有记忆记录 ($today)"
        return 1
    fi
    
    process_file "$file"
}

# 处理指定日期
garden_date() {
    local date="$1"
    local file="$MEMORY_DIR/$date.md"
    
    process_file "$file"
}

# 处理最近7天
garden_week() {
    log_info "整理最近7天的记忆..."
    local count=0
    
    for i in {0..6}; do
        local date=$(date -d "$i days ago" +%Y-%m-%d 2>/dev/null || date -v-${i}d +%Y-%m-%d)
        local file="$MEMORY_DIR/$date.md"
        
        if [ -f "$file" ]; then
            if process_file "$file"; then
                count=$((count + 1))
            fi
        fi
    done
    
    log_success "整理了 $count 天的记忆"
}

# 处理所有
garden_all() {
    log_info "整理所有历史记忆..."
    local count=0
    
    for file in "$MEMORY_DIR"/[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9].md; do
        if [ -f "$file" ]; then
            if process_file "$file"; then
                count=$((count + 1))
            fi
        fi
    done
    
    log_success "整理了 $count 天的记忆"
}

# 主程序
main() {
    # 确保目录存在
    mkdir -p "$MEMORY_DIR"
    
    case "${1:-today}" in
        today)
            garden_today
            ;;
        week)
            garden_week
            ;;
        all)
            garden_all
            ;;
        help|--help|-h)
            show_help
            ;;
        [0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9])
            garden_date "$1"
            ;;
        *)
            log_error "未知命令: $1"
            show_help
            exit 1
            ;;
    esac
}

main "$@"
