#!/bin/bash
# 🚀 一键推送知识库到 GitHub（简化版）
#
# 用法（需要两个环境变量）：
#   export GH_USERNAME="你的GitHub用户名"
#   export GH_PAT="ghp_你的token"
#   export REPO_NAME="mind-growth-content"   # 可选，默认此名
#   bash push_to_github.sh
#
# 说明：日常使用其实不用自己跑，直接对 AI 说
#       "把知识库同步到 GitHub" 即可，AI 会代为执行。

set -e
cd /workspace/知识库

GH_USERNAME="${GH_USERNAME:?请先 export GH_USERNAME}"
GH_PAT="${GH_PAT:?请先 export GH_PAT}"
REPO_NAME="${REPO_NAME:-mind-growth-content}"
BRANCH="main"

echo "=== 准备推送 ==="
echo "账号: $GH_USERNAME"
echo "仓库: $REPO_NAME"
echo "文件数: $(git ls-files | wc -l)"
echo "内容大小: $(du -sh --exclude=.git . | cut -f1)"
echo ""

# 已存在 origin 则更新，否则添加
if git remote get-url origin >/dev/null 2>&1; then
    git remote set-url origin "https://${GH_USERNAME}:${GH_PAT}@github.com/${GH_USERNAME}/${REPO_NAME}.git"
else
    git remote add origin "https://${GH_USERNAME}:${GH_PAT}@github.com/${GH_USERNAME}/${REPO_NAME}.git"
fi

echo "=== 推送中 ==="
git push -u origin "$BRANCH"

echo ""
echo "🎉 完成！访问 https://github.com/${GH_USERNAME}/${REPO_NAME} 查看"
