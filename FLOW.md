# System & Business Flow

## システムフロー（技術）

```mermaid
flowchart TD
    A([Client / User]) -->|JSON input| B[CLI: memo_generator generate]
    B --> C{Input Validation\nPropertyInput model}
    C -->|Invalid| ERR([Error: validation failed])
    C -->|Valid| D[Financial Calculator\ncalculate_all_metrics]
    D --> E[Build Property Context\n+ metrics dict]
    E --> F[Claude API Client\nget_client]

    F --> G1[Section: Executive Summary]
    F --> G2[Section: Investment Thesis]
    F --> G3[Section: Financial Analysis]
    F --> G4[Section: Risk Factors]
    F --> G5[Section: Market Overview]

    G1 & G2 & G3 & G4 & G5 --> H[Aggregate memo_data dict]

    H --> I{Output Format?}
    I -->|markdown| J[Markdown Renderer\n→ .md file]
    I -->|pdf| K[HTML Renderer\n→ WeasyPrint\n→ .pdf file]

    J --> OUT([Delivered to Client])
    K --> OUT
```

---

## ビジネスフロー（収益化）

```mermaid
flowchart LR
    subgraph Day1["Day 1 (3/20) — インフラ"]
        A1[APIキーローテーション] --> A2[GitHub commit\n+ PDF デモ]
        A2 --> A3[Upwork Gig公開\n3プラン設定]
    end

    subgraph Day2["Day 2 (3/21) — 獲得"]
        B1[Upwork Proposals\n25件送信] --> B2[LinkedIn / Reddit\n投稿・拡散]
        B2 --> B3[返信・スコープ確認]
    end

    subgraph Day3["Day 3 (3/22) — クロージング"]
        C1[デモセッション\n実施] --> C2[JSON受取\n→ 即時生成\n→ PDF納品]
        C2 --> C3[初回収益\n$99〜$249]
    end

    Day1 --> Day2 --> Day3

    subgraph Scale["Week 2以降 — スケール"]
        D1[リピーター獲得] --> D2[Enterprise契約\n$499/mo]
        D2 --> D3[Claude並列エージェント\n大量処理対応]
    end

    Day3 --> Scale
```

---

## Claude並列エージェント活用フロー

```mermaid
flowchart TD
    REQ([大量物件リクエスト]) --> SPLIT[スプリッター\n物件ごとに分割]

    SPLIT --> AG1[Agent 1\n物件A処理]
    SPLIT --> AG2[Agent 2\n物件B処理]
    SPLIT --> AG3[Agent 3\n物件C処理]
    SPLIT --> AGN[Agent N\n物件N処理]

    AG1 --> MERGE[マージャー\n全PDFを結合]
    AG2 --> MERGE
    AG3 --> MERGE
    AGN --> MERGE

    MERGE --> DELIVER([一括納品\n処理時間: 1/N に短縮])
```
