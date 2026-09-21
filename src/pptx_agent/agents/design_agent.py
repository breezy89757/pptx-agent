"""Stage 2: generate one design-token set for the whole deck, so every slide
draws from the same palette/type/spacing instead of choosing its own."""

from __future__ import annotations

from pptx_agent.color_contrast import enforce_token_contrast
from pptx_agent.fonts import DEFAULT_BODY_FONT, DEFAULT_HEADING_FONT, SAFE_FONTS, clamp_to_safe_font
from pptx_agent.llm_client import complete_json
from pptx_agent.models import Deck, DesignTokens

_FONT_LIST = "\n".join(f"  - {name}:{desc}" for name, desc in SAFE_FONTS.items())

SYSTEM_PROMPT = f"""\
你是視覺設計系統負責人。根據簡報的主題與內容,設計「一套」貫穿全場的設計令牌(design tokens),\
後續每一張投影片都必須只使用這套令牌,不能自行選色或選字。

規則:
- primary_color / secondary_color / accent_color / background_color / surface_color / text_color / muted_text_color \
皆為 6 碼 hex(不帶 #)。text_color 與 muted_text_color 是唯二會被拿來排「大段可讀文字」的顏色,\
兩者對 background_color 與對 surface_color 的對比度都必須達到 WCAG AA 標準(對比比值至少 4.5:1),\
不可為了風格好看犧牲可讀性——這是硬性規定,系統會在你輸出後再做一次程式檢查並強制修正不合格的顏色。
- surface_color 用於卡片/區塊背景,通常介於 background_color 與 text_color 之間的中性色,或 background_color 的近似變體。
- font_heading / font_body 只能從下面這個安全清單裡選(不可自己發明或使用其他 Google Fonts 名稱)。這些都是
  Windows/Office 內建字型,確保瀏覽器預覽與最終在使用者電腦開啟的 PowerPoint 顯示同一套字,不會因為對方電腦
  沒裝某個網路字型而被置換、導致文字量測跑掉、換行/溢出跟預覽不一致:
{_FONT_LIST}
  font_body 建議選「內文適用」的字型(標楷體不適合內文,只能給 font_heading 當標題強調用);font_heading 可以跟
  font_body 相同,也可以用不同字型製造標題/內文的層次感。
- spacing_unit_px 是一個 8 的倍數(例如 8/12/16),作為全場間距的基礎單位。
- radius_px 是卡片/按鈕的圓角基礎值。
- motif 用一句話描述一個會在多張投影片重複出現的視覺元素(例如「左上角一道對角強調色塊」「數字強調用等寬大字」),\
且此描述要具體到後續 HTML 設計者能直接照做。
- 設計要呼應簡報主題的調性,但不要預設任何一種主題就該用深色系——實務上企業簡報、提案、報告多半是淺色/白底,
  深色系只適合特定情境(例如產品發表、夜間活動、遊戲/創意產業)。除非主題內容明確指向「科技感、未來感、夜間、
  極客」這類調性,否則預設應該考慮淺色或中性背景(例如近白、米色、淺灰)搭配深色文字,而不是自動選深色底。
  拿到題目後,先問自己「這份簡報的受眾與場合,實際上比較常見淺色還是深色簡報?」再決定。
"""


def run_design_agent(deck: Deck) -> DesignTokens:
    slide_summary = "\n".join(f"- ({s.content_type.value}) {s.title}" for s in deck.slides)
    user_prompt = (
        f"簡報標題:{deck.deck_title}\n\n投影片內容概覽:\n{slide_summary}\n\n"
        "請為這份簡報設計一套設計令牌。"
    )
    tokens = complete_json(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
        response_model=DesignTokens,
    )
    tokens = tokens.model_copy(
        update={
            "font_heading": clamp_to_safe_font(tokens.font_heading, fallback=DEFAULT_HEADING_FONT),
            "font_body": clamp_to_safe_font(tokens.font_body, fallback=DEFAULT_BODY_FONT),
        }
    )
    return enforce_token_contrast(tokens)
