"""Stage 6 (optional, looped): render the actual .pptx slide and have the
model visually review it against what was intended. This closes the spec's
architecture loop ("render -> QA -> fix") against the real deliverable
rather than the HTML intermediate."""

from __future__ import annotations

from pathlib import Path

from pptx_agent.llm_client import complete_json_with_image
from pptx_agent.models import DesignTokens, SlideContent, SlideQaVerdict

SYSTEM_PROMPT = """\
你是簡報視覺品質檢查員。你會看到一張「實際在 PowerPoint 裡渲染出來的投影片截圖」,以及這張投影片
「原本應該呈現的結構化內容」與「整份簡報的設計令牌」。你的工作是老實指出這張截圖有沒有視覺缺陷,\
而不是重新評論設計美感偏好。

只回報以下這幾類客觀、可驗證的問題:
1. 文字互相重疊,或文字被其他元素(裝飾圖形、其他文字、卡片邊框)蓋住、部分看不見。
2. 文字被裁切/溢出容器邊界,或緊貼容器邊緣到不舒服的程度。
3. 文字與背景對比度太低,難以閱讀(例如淺色文字疊在淺色背景、深色文字疊在深色背景)。
4. 內容缺漏或跟原始結構化內容不符(例如 title/bullets/data_points/comparison/timeline/quote 裡的文字,\
截圖上完全沒出現,或明顯被截斷)。
5. 元素明顯跑版:例如卡片大小不一致到不自然、元素超出投影片畫布邊界、大片不合理的空白或元素互相穿插。

不要回報以下這些(這些是合理的設計選擇,不算問題):
- 你個人不喜歡的配色、字型、排版風格,只要對比度足夠、沒有重疊遮蔽。
- 裝飾性圖形位置只是「不是你會選的位置」但沒有實際遮住文字。
- 圖表/圖片本身的視覺設計。

has_issues=false 時,issues 留空陣列,fix_instructions 留 null。
has_issues=true 時,issues 列出具體看到的問題(每項一句話,講清楚是哪個元素、什麼問題),\
fix_instructions 則是要交給「重新設計這張投影片版面的人」的具體修正指示(例如「數據卡文字太靠右邊界,\
卡片內距要加大」「標題被右上角的裝飾線條蓋住,裝飾元素要移到不會碰到標題文字的位置」),不要只是重複 issues,\
要講清楚具體怎麼改。
"""


def run_qa_agent(slide_png: Path, slide: SlideContent, tokens: DesignTokens) -> SlideQaVerdict:
    user_prompt = (
        f"這是投影片 #{slide.index} 實際在 PowerPoint 中渲染出來的畫面。\n\n"
        f"這張投影片應呈現的結構化內容(JSON):{slide.model_dump_json()}\n\n"
        f"整份簡報的設計令牌(JSON):{tokens.model_dump_json()}\n\n"
        "請檢查這張截圖有沒有前述定義的視覺缺陷。"
    )
    verdict = complete_json_with_image(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
        image_path=slide_png,
        response_model=SlideQaVerdict,
    )
    verdict.index = slide.index
    return verdict
