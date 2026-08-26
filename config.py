"""
Word Agent 配置文件
在这里填写你的 API Key 与模型推理参数
"""

import os
from pathlib import Path

# ============================================================
# 🎯 API 密钥（从环境变量 / .env 文件读取，不硬编码入库）
# ============================================================
# 安全说明：密钥不写入本文件（避免提交到远程仓库泄露）。
# 请在项目根目录创建 .env 文件（已被 .gitignore 忽略），内容：
#   ARK_API_KEY=你的密钥
# 或设置系统环境变量 ARK_API_KEY。
# 变量名沿用 ARK_API_KEY（兼容环境变量覆盖），值为 DeepSeek Key。

# 加载 config.py 同目录的 .env 文件（若存在），不依赖 python-dotenv
_env_path = Path(__file__).resolve().parent / ".env"
if _env_path.exists():
    try:
        for _line in _env_path.read_text(encoding="utf-8").splitlines():
            _line = _line.strip()
            if not _line or _line.startswith("#") or "=" not in _line:
                continue
            _key, _val = _line.split("=", 1)
            _key = _key.strip()
            _val = _val.strip().strip('"').strip("'")
            if _key and _key not in os.environ:
                os.environ[_key] = _val
    except Exception:
        pass

ARK_API_KEY = os.getenv("ARK_API_KEY", "")

# ============================================================
# 🤖 模型配置（2026-08：切换为 DeepSeek 官方 + flash 模型）
# ============================================================
# ✅ 当前：DeepSeek 官方 API，flash 模型（快、省）
DEFAULT_MODEL = "deepseek-v4-flash"
DEFAULT_BASE_URL = "https://api.deepseek.com"


## 硅基流动 API
# DEFAULT_MODEL = "deepseek-ai/DeepSeek-V4-Flash"
# DEFAULT_BASE_URL = "https://api.siliconflow.cn/v1"

# ⚠️ 备选（火山方舟 ARK coding 端点，取消注释并换回 ark Key 即可）
# DEFAULT_MODEL = "deepseek-v4-flash-ga-260731"   # 质量高，reasoning_effort 真实生效
# DEFAULT_BASE_URL = "https://ark.cn-beijing.volces.com/api/coding/v3"
#   doubao-seed-2-1-turbo-260628 / glm-5-2-260617 / deepseek-v4-pro-ga-260813 亦可
# ⚠️ 基准测试提示：deepseek flash 系列打分离谱（曾出现 10/10），已靠 prompt 严格度校准兜底

# ============================================================
# 🎚️ 推理参数（基准测试结论，详见 benchmark_configs.py 输出）
# ============================================================
#   - reasoning_effort=high 使 v4pro 诊断最贴近老师（抓住"论证多叙述少评论/事实堆砌"）
#   - temperature 0.0~0.7 对批改质量几乎无影响 → 取 0.3 稳定中间值
#   - 所有模型都系统性高估分数 → 靠 prompt 里的"批改标准要比考研阅卷严格"来校准
TEMPERATURE = 0.3        # 0.0~0.7；实测无显著差异，取稳定中间值 0.3
REASONING_EFFORT = None  # None/"low"/"medium"/"high"；仅 deepseek-v4-pro 真实生效，flash 不支持 → 置 None 避免报错
MAX_TOKENS = 393000       # 单次回复最大输出 token（deepseek-v4-flash 上限 393K，每句都有评语输出量大）

# ============================================================
# 📝 批改行为
# ============================================================
GRADING_MODE = "agentic"  # "agentic"=多轮逐句（精细） / "single"=单次结构化（省 token）
# 评论作业评分维度（与老师返修一致，50 分制）
COMMENT_DIMENSIONS = ["立意", "标题", "论证逻辑", "行文结构", "语言表达"]
SCORE_TOTAL = 50

# 采访作业评分维度（与老师返修一致，20 分制）
# 采访部分 12 分（核心：采访对象及问题）、其他部分 5 分（背景/主题/方式/时间/地点等）、表述及逻辑 3 分
INTERVIEW_DIMENSIONS = ["采访部分", "其他部分", "表述及逻辑"]
INTERVIEW_SCORE_TOTAL = 20

# （旧备注已上移：ARK 端点信息见上方「备选」注释）