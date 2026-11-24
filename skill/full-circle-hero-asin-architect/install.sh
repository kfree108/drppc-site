#!/usr/bin/env bash
set -euo pipefail

# ==============================================================================
# Full Circle Hero ASIN Campaign Architect — Claude Code Skill Installer
# By Full Circle ($500M+ Managed Amazon Revenue) × Full Circle Orbit (Dr. PPC + Ava)
# Usage: curl -fsSL https://drppc.ai/skill/full-circle-hero-asin-architect/install.sh | bash
# ==============================================================================

BOLD='\033[1m'
GREEN='\033[0;32m'
GOLD='\033[0;33m'
CYAN='\033[0;36m'
RED='\033[0;31m'
DIM='\033[2m'
NC='\033[0m'

SKILL_NAME="full-circle-hero-asin-architect"
SKILL_DIR="${HOME}/.claude/skills/${SKILL_NAME}"
REPO_RAW_BASE="${FC_SKILL_BASE_URL:-https://drppc.ai/skill/${SKILL_NAME}}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]:-$0}")" 2>/dev/null && pwd || true)"

echo ""
echo -e "${GREEN}${BOLD}┌────────────────────────────────────────────────────────────────────┐${NC}"
echo -e "${GREEN}${BOLD}│  FULL CIRCLE × ORBIT — Hero ASIN Campaign Architect Installer      │${NC}"
echo -e "${GREEN}${BOLD}│  Full Circle 8-Segment Architecture + Perfect CPC + 3x Shield      │${NC}"
echo -e "${GREEN}${BOLD}└────────────────────────────────────────────────────────────────────┘${NC}"
echo ""

# 1. Detect the operating system (macOS, Linux, or Windows via WSL / Git Bash)
OS_NAME="$(uname -s 2>/dev/null || echo unknown)"
echo -e "${GREEN}✓${NC} System detected (${BOLD}${OS_NAME}${NC})"

# 2. Check Claude Code CLI
if command -v claude >/dev/null 2>&1; then
  CLAUDE_VER="$(claude --version 2>/dev/null | head -n 1 || echo 'installed')"
  echo -e "${GREEN}✓${NC} Claude Code CLI found (${DIM}${CLAUDE_VER}${NC})"
else
  echo -e "${GOLD}!${NC} Claude Code CLI not found in PATH."
  echo -e "  Install Claude Code first with:"
  echo -e "    ${CYAN}npm install -g @anthropic-ai/claude-code${NC}  ${DIM}(or curl -fsSL https://claude.ai/install.sh | bash)${NC}"
  echo -e "  Continuing skill installation to ${BOLD}${SKILL_DIR}${NC} so it's ready once Claude Code is installed..."
fi

# 3. Check Python 3 + openpyxl (for .xlsx generation)
if command -v python3 >/dev/null 2>&1; then
  PY_VER="$(python3 --version 2>&1)"
  echo -e "${GREEN}✓${NC} ${PY_VER} found"
  if python3 -c "import openpyxl" >/dev/null 2>&1; then
    echo -e "${GREEN}✓${NC} openpyxl installed (ready for 6-tab .xlsx generation)"
  else
    echo -e "${GOLD}→${NC} Installing ${BOLD}openpyxl${NC} for Excel workbook generation..."
    python3 -m pip install --user --quiet openpyxl 2>/dev/null || \
      python3 -m pip install --break-system-packages --user --quiet openpyxl 2>/dev/null || \
      echo -e "${GOLD}!${NC} Could not auto-install openpyxl. Run: ${CYAN}pip3 install openpyxl${NC}"
  fi
else
  echo -e "${GOLD}!${NC} python3 not found. Install Python 3 to compile .xlsx bulk upload spreadsheets."
fi

# 4. Install skill files to ~/.claude/skills/full-circle-hero-asin-architect/
echo -e "${CYAN}→${NC} Installing skill to ${BOLD}${SKILL_DIR}${NC}..."
mkdir -p "${SKILL_DIR}"

install_asset() {
  local filename="$1"
  if [[ -n "${SCRIPT_DIR}" && -f "${SCRIPT_DIR}/${filename}" ]]; then
    cp "${SCRIPT_DIR}/${filename}" "${SKILL_DIR}/${filename}"
  else
    curl -fsSL "${REPO_RAW_BASE}/${filename}" -o "${SKILL_DIR}/${filename}"
  fi
}

install_asset "SKILL.md"
install_asset "build_campaign_sheet.py"
chmod +x "${SKILL_DIR}/build_campaign_sheet.py"

echo -e "${GREEN}✓${NC} Installed ${BOLD}SKILL.md${NC}"
echo -e "${GREEN}✓${NC} Installed ${BOLD}build_campaign_sheet.py${NC}"

# 5. Print quick-start instructions
echo ""
echo -e "${GREEN}${BOLD}Installation complete!${NC}"
echo ""
echo -e "${BOLD}Next 2 steps to build your 1 Free Hero ASIN Architecture:${NC}"
echo -e "  ${GOLD}1.${NC} Connect your Amazon account (read-only) to ${BOLD}Dr. PPC + Ava${NC} in Orbit:"
echo -e "     ${CYAN}https://fullcircleorbit.com/signup?from=drppc&utm_source=creator&utm_campaign=hormozi_skill_funnel&code=FOUNDER10${NC}"
echo -e "  ${GOLD}2.${NC} Launch Claude Code in any folder and run:"
echo -e "     ${CYAN}claude \"/full-circle-hero-asin-architect\"${NC}"
echo ""
echo -e "${DIM}Tip: Want to test the 6-tab Excel generator locally right now? Run:${NC}"
echo -e "     ${CYAN}python3 \"${SKILL_DIR}/build_campaign_sheet.py\"${NC}"
echo ""
echo -e "${DIM}By Full Circle (\$500M+ Managed Amazon Revenue) · Dr. PPC + Ava + Orbit${NC}"
echo -e "${DIM}Daily autopilot: \$1 first 30 days, then \$299/mo + 2% of ad spend · Code FOUNDER10 = 10% off for life${NC}"
echo ""
