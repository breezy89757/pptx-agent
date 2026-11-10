"""Stage 3: pick a layout per content type and write it as HTML/CSS using the
design tokens. This only borrows the browser's layout engine — the HTML/CSS
output here is not the final deliverable (v1 extracts it into a native PPTX)."""

from __future__ import annotations

import json

from pptx_agent.charts import GRAMMAR_PROMPT, validate_chart_directives
from pptx_agent.llm_client import complete_json
from pptx_agent.models import Dataset, DesignTokens, SlideContent, SlideHtml

MAX_CHART_RETRIES = 2

CANVAS_W = 1280
CANVAS_H = 720

SYSTEM_PROMPT = f"""\
你是 HTML/CSS 排版設計師。你會拿到「一張投影片的結構化內容」與「整份簡報共用的設計令牌」,\
任務是把這張投影片畫成一個自我完整的 HTML 片段。

硬性規則:
1. 畫布固定 {CANVAS_W}x{CANVAS_H}px(16:9)。你輸出的最外層 div 必須是
   `<div style="position:absolute;inset:0;width:{CANVAS_W}px;height:{CANVAS_H}px;overflow:hidden;...">`,
   內部用 flexbox 或 grid 排版,不要用 <style> 標籤或 class 選擇器——所有樣式一律寫在
   `style="..."` 行內屬性,因為後續程式會逐一取每個元素的 computed style。
2. 顏色、字型、間距只能用以下 CSS 變數,不能自己發明新顏色或字型:
   var(--primary) var(--secondary) var(--accent) var(--bg) var(--surface) var(--text) var(--muted)
   var(--font-heading) var(--font-body) var(--space)(基礎間距 px)var(--radius)(基礎圓角 px)
   需要更大間距時用 var(--space) 的倍數,例如 calc(var(--space) * 3)。
3. 每一張投影片都必須包含至少一個「非純文字」的視覺元素,例如:
   - 用 div 畫的資料卡(帶底色/邊框的區塊,顯示數字或關鍵字)
   - 用 div 堆疊/漸層/圓形畫的簡單圖示或進度指示
   - 用横向 flex 排出的時間軸節點與連接線
   - 用左右分欄 + 對比色塊呈現 comparison
   純粹一個標題加一串 <ul><li> 條列且沒有任何色塊/卡片包裝,不符合這條規則。
4. 標題文字用 var(--font-heading),內文/條列用 var(--font-body)。
5. 依 motif 描述(設計令牌裡的 motif 欄位)在這張投影片上重現該視覺元素,讓整份簡報有一致的識別度。
6. 不同投影片即使 content_type 相同,也要讓版面配置(欄位切分、元素排列方式)有變化,不要每張都套用一模一樣的骨架。
7. 只輸出 HTML,不要輸出 Markdown code fence、不要輸出說明文字。
8. 文字內容必須完整呈現使用者提供的 title/subtitle/bullets/data_points/comparison/timeline/quote,不可省略或改寫。
9. 避免元素互相重疊:文字長度不可預先得知確切換行後的高度,所以任何文字區塊(標題、描述、label)都不能用
   絕對座標緊貼相鄰的裝飾元素(圓點、連接線、圖示、其他文字區塊)。規則:
   - 文字與相鄰裝飾元素(圓點/線條/圖示)之間至少保留 `calc(var(--space) * 2)` 的間距,寧可留白也不要貼近。
   - 需要容納不定長文字的容器,高度盡量用 auto 或足夠寬裕的固定值(而非剛好貼合單行文字的高度),讓文字有
     換行 2-3 行的空間而不會溢出容器或蓋到下一個元素。
   - 時間軸/多節點版面每個節點之間的水平或垂直間距,要以「文字最多換行 3 行」的高度來抓,不要以單行高度抓。
   - 絕對定位(position:absolute)的裝飾元素(圓點、連接線)只能放在明確不會被文字佔用的區域,或放在文字容器
     的外側,不要讓文字容器的估計邊界與裝飾元素的座標重疊。
10. 文字對比度是硬性規定,不是美感取捨:
    - 任何承載文字的元素,文字顏色只能用 var(--text) 或 var(--muted)(一般內文/次要文字用這兩色,已保證
      對 var(--bg) 與 var(--surface) 都有足夠對比)。var(--primary)/var(--secondary)/var(--accent) 只能用在
      「字級較大的標題、數字、標籤」等強調用途上,且底色必須是 var(--bg) 或 var(--surface)(不要疊在同樣鮮豔
      的強調色塊上)。絕對不要把文字顏色設成跟它所在背景明暗相近的顏色。
    - 絕對不要對「帶有文字內容的元素」套用 `opacity` 小於 1(例如 opacity:0.6 的 label、caption、數字都不行)——
      opacity 會直接拉低文字與背景的對比,可能導致文字幾乎看不見。opacity < 1 只能用在純裝飾元素上
      (圓點、線條、色塊、圖示),絕對不能用在任何 <div>/<span>/<p>/<h1> 等承載可讀文字的元素本身。
      如果想要「文字看起來次要」的效果,改用 var(--muted) 顏色本身,而不是調低不透明度。
11. 如果這張投影片附有資料集(下方「本頁可用資料集」不是空的),代表有真實數據。要用圖表呈現時,必須使用
    「原生圖表版位」,不要自己用 div 畫長條/折線去模擬;你可以依訊息自行決定圖表類型、數量與位置。
    圖表版位大小視重要性而定:只有一張圖時抓整張投影片 40%-60% 面積,多張圖時各自縮小,其餘空間留給標題、
    重點說明與卡片。是否用圖表由你判斷——只有 1-2 個數字時,用資料卡(大數字)可能比圖表更好。
{GRAMMAR_PROMPT}"""


def _content_type_hint(slide: SlideContent) -> str:
    hints = {
        "title": "封面頁:大標題 + 副標題,善用 motif 與 accent 色塊營造視覺焦點,不要只是置中文字。",
        "bullets": "條列頁:避免單純 <ul><li>,改用卡片網格或帶編號色塊的清單,每個要點給一個視覺容器。",
        "comparison": "對比頁:左右兩欄,用不同底色或邊框色區分兩側,標題置頂。",
        "data": (
            "數據頁:若本頁有資料集,依規則 11 選最能傳達訊息的原生圖表(可多張),標題/data_points(如果有)排在旁邊或上方,"
            "並用一段重點說明點出圖表想表達的結論;不要自己畫圖表模擬數列。沒有資料集時,每個 data_point 做成一張"
            "數據卡(大數字 + 單位 + 標籤),用 grid 排列。"
        ),
        "timeline": (
            "時間軸頁:橫向或縱向時間軸,節點用圓點+連接線,每個節點旁放 label 與 description。"
            "每個節點的文字卡與圓點/連接線用各自獨立的區塊分開放置(例如圓點在上、卡片在下,中間留白間隔),"
            "不要讓卡片的絕對定位座標去貼近圓點座標;卡片本身高度要抓夠(容納 description 換行 2-3 行),"
            "節點與節點之間的水平/垂直間距也要以最長 description 換行後的高度預留,避免相鄰節點的卡片互相重疊。"
        ),
        "quote": "引言頁:大字級引言置中或偏一側,搭配 accent 色的引號圖形或色塊裝飾,attribution 放小字。",
    }
    return hints.get(slide.content_type.value, "")


def _slide_datasets(slide: SlideContent, datasets: list[Dataset] | None) -> list[Dataset]:
    wanted = set(slide.dataset_ids or [])
    return [d for d in (datasets or []) if d.id in wanted]


def run_layout_agent(
    slide: SlideContent,
    tokens: DesignTokens,
    *,
    datasets: list[Dataset] | None = None,
    extra_instructions: str | None = None,
    design_instructions: str | None = None,
) -> SlideHtml:
    """Write one slide as HTML. Any `data-chart` directives are validated
    against the slide's datasets; on failure the concrete errors are fed back
    and the model retries (up to MAX_CHART_RETRIES), so a bad chart spec never
    reaches the Node stages."""
    slide_datasets = _slide_datasets(slide, datasets)
    datasets_json = json.dumps([d.model_dump() for d in slide_datasets], ensure_ascii=False)

    def build_prompt(note: str | None) -> str:
        revision_note = (
            f"\n\n這是重新設計(修正版):上一版實際渲染成 PowerPoint 後,視覺 QA 發現以下問題,這次務必避開,"
            f"必要時可以整個換一種版面配置,不必修修補補:\n{note}\n"
            if note
            else ""
        )
        design_note = (
            f"\n\n必須遵守的設計系統規範與可用素材:\n{design_instructions}"
            if design_instructions
            else ""
        )
        return (
            f"設計令牌(JSON):{tokens.model_dump_json()}\n\n"
            f"這張投影片的內容類型:{slide.content_type.value}\n"
            f"版面提示:{_content_type_hint(slide)}\n\n"
            f"投影片結構化內容(JSON):{slide.model_dump_json()}\n\n"
            f"本頁可用資料集(JSON,圖表只能引用這些;沒有就代表本頁不能有 data-chart):{datasets_json}"
            f"\n\n設計參考圖片(JSON):{json.dumps([asset.model_dump() for asset in tokens.reference_assets], ensure_ascii=False)}"
            f"{design_note}"
            f"{revision_note}\n\n"
            "請畫出這張投影片的 HTML。"
        )

    note = extra_instructions
    result: SlideHtml | None = None
    for attempt in range(MAX_CHART_RETRIES + 1):
        result = complete_json(
            system_prompt=SYSTEM_PROMPT,
            user_prompt=build_prompt(note),
            response_model=SlideHtml,
        )
        errors = validate_chart_directives(result.html, slide_datasets)
        if not errors:
            break
        note = (
            (extra_instructions + "\n" if extra_instructions else "")
            + "上一版的 data-chart 圖表規格有以下錯誤,請修正後重新輸出完整 HTML:\n"
            + "\n".join(f"- {e}" for e in errors)
        )
    else:  # exhausted retries: fail loudly rather than ship a chart that will not build
        raise RuntimeError(f"slide #{slide.index}: chart spec still invalid after retries:\n" + "\n".join(errors))
    result.index = slide.index
    return result
