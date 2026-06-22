#!/bin/bash
# ============================================
# Cloudflare Pages 部署脚本
# Excel 多关键词筛选工具
# ============================================

set -e

PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_NAME="excel-multi-keyword-filter"
OUTPUT_DIR="$PROJECT_DIR/out"

# 颜色输出
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${GREEN}=== Excel 多关键词筛选工具 - Cloudflare Pages 部署 ===${NC}"

# 检查是否已认证
if [ -z "$CLOUDFLARE_API_TOKEN" ]; then
    echo -e "${YELLOW}⚠ 未检测到 CLOUDFLARE_API_TOKEN 环境变量${NC}"
    echo ""
    echo "请选择认证方式："
    echo "  1) 设置 CLOUDFLARE_API_TOKEN 环境变量（推荐用于 CI/CD）"
    echo "  2) 运行 wrangler login 进行交互式登录"
    echo ""
    echo "方式1 - 创建 API Token："
    echo "  访问 https://developers.cloudflare.com/fundamentals/api/get-started/create-token/"
    echo "  创建一个具有 Cloudflare Pages 编辑权限的 Token"
    echo "  然后运行: export CLOUDFLARE_API_TOKEN=你的token"
    echo ""
    echo "方式2 - 交互式登录："
    echo "  运行: wrangler login"
    echo ""
    
    # 尝试交互式登录
    if [ -t 0 ]; then
        echo -e "${YELLOW}检测到交互式终端，尝试运行 wrangler login...${NC}"
        wrangler login
    else
        echo -e "${RED}✘ 非交互式环境，请先设置 CLOUDFLARE_API_TOKEN${NC}"
        exit 1
    fi
fi

# 检查输出目录
if [ ! -d "$OUTPUT_DIR" ]; then
    echo -e "${YELLOW}⚠ 未找到静态输出目录，开始构建...${NC}"
    cd "$PROJECT_DIR"
    npm install
    npx next build
    echo -e "${GREEN}✓ 构建完成${NC}"
fi

# 部署
echo -e "${GREEN}开始部署到 Cloudflare Pages...${NC}"
cd "$PROJECT_DIR"
wrangler pages deploy out --project-name "$PROJECT_NAME" --branch main

echo ""
echo -e "${GREEN}=== 部署完成！ ===${NC}"
echo ""
echo "后续操作："
echo "  - 在 Cloudflare Dashboard 中绑定自定义域名"
echo "  - 访问 https://dash.cloudflare.com 查看项目"
