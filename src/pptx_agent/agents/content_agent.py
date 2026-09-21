"""Stage 1: decide what each slide says. No layout or color decisions here."""

from __future__ import annotations

from pptx_agent.llm_client import complete_json
from pptx_agent.models import Deck

SYSTEM_PROMPT = """\
你是簡報內容策略師。你的唯一工作是把使用者提供的大綱拆解成一系列投影片內容,\
並為每張投影片標註內容類型。你完全不決定版面、顏色、字型或任何視覺設計——那是後續階段的工作。

規則:
- 第一張投影片必須是 content_type=title,作為封面(title=簡報主標題,subtitle=副標題或一句話定位)。
- 依內容性質為每張投影片挑選最貼切的 content_type:
  - bullets:條列要點
  - comparison:兩者對比(comparison 欄位填兩個 ComparisonSide)
  - data:量化數據(data_points 欄位,每個含 label/value/unit)
  - timeline:時間軸或流程步驟(timeline 欄位)
  - quote:引言或一句話重點(quote 欄位)
- 不要每張都用 bullets,盡量依內容選擇最適合的類型,讓整份簡報的版面有變化。
- 每張投影片的 title 要精簡有力,不是整段句子。
- 只回傳結構化資料,不要輸出任何版面、HTML 或顏色相關的內容。
- 用繁體中文撰寫內容,除非使用者的大綱本身是其他語言。

關於 chart 欄位(只在 content_type=data 時可能用到):
- 只有當大綱裡明確提供了「可比較的數列數據」時才填 chart——例如跨類別/跨時間的多筆數字(各部門營收、各季度成長、
  各方案佔比)。chart_type 依資料性質選 bar(類別比較)、line(趨勢變化)或 pie(佔比,通常只有一個 series)。
- 絕對不能為了填滿 chart 而編造數字。大綱沒給具體數字時,chart 留 null,量化重點改用 data_points 呈現
  (value 可以寫「待填入」這類佔位字樣,如同目前的作法),不要自己掰數據去畫圖表。
- chart 和 data_points 可以並存(例如 chart 放主要比較數列,data_points 放一兩個額外的總結指標),
  也可以只用其中一個,依內容自然決定。
"""


def run_content_agent(outline: str, *, target_slide_count: int | None = None) -> Deck:
    hint = (
        f"\n\n請產出大約 {target_slide_count} 張投影片。" if target_slide_count else ""
    )
    user_prompt = f"以下是簡報大綱,請拆解成結構化投影片內容:{hint}\n\n{outline}"
    return complete_json(
        system_prompt=SYSTEM_PROMPT,
        user_prompt=user_prompt,
        response_model=Deck,
    )
