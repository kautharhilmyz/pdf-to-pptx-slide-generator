# JSON Schemas for `pdf-to-pptx` Pipeline

## 1. `<deck_dir>/workspace/chart_plan_{ts}.json`

```json
{
  "theme": "classic_navy_teal",
  "charts": [
    {
      "slide_number": 4,
      "output_file": "assets/chart_slide_4_20261124_160000.png",
      "type": "bar",
      "title": "Revenue by Region ($M)",
      "xlabel": "Region",
      "ylabel": "USD Millions",
      "categories": ["APAC", "EMEA", "Americas"],
      "series": [
        {"name": "FY2025", "values": [42.5, 38.0, 61.2]},
        {"name": "FY2026", "values": [54.1, 41.5, 73.8]}
      ],
      "value_format": "{:.1f}"
    }
  ]
}
```

Supported `type` values:
- `bar`: Vertical single or grouped bar chart
- `barh`: Horizontal single or grouped bar chart
- `line`: Multi-series line chart with markers and data labels
- `donut`: Clean donut chart (uses first series values + categories)

---

## 2. `<deck_dir>/workspace/slide_plan_{ts}.json`

```json
{
  "theme": "classic_navy_teal",
  "brand_file": "brand.json",
  "template_pptx": null,
  "slides": [
    {
      "slide_number": 1,
      "layout": "title_slide",
      "title": "Annual Operations Review",
      "subtitle": "FY2026 Performance & Strategic Priorities",
      "meta": "Strategy Team | November 2026",
      "speaker_notes": "Opening overview. [Source: annual_report.pdf, p. 1]"
    },
    {
      "slide_number": 2,
      "layout": "section_header",
      "section_number": "01",
      "title": "Financial Performance",
      "subtitle": "Revenue growth and margin expansion across regions",
      "speaker_notes": "Transition to financial results. [Source: annual_report.pdf, p. 4]"
    },
    {
      "slide_number": 3,
      "layout": "key_metrics_cards",
      "title": "Key Performance Indicators",
      "cards": [
        {"label": "Total Revenue", "value": "$169.4M", "delta": "+19.5% YoY"},
        {"label": "Operating Margin", "value": "22.4%", "delta": "+210 bps"},
        {"label": "Net Retention", "value": "118%", "delta": "Enterprise tier"}
      ],
      "bullets": [
        "**Enterprise Expansion:** New multi-year contracts drove 64% of net-new ARR.",
        "**Cost Discipline:** Cloud infrastructure spend fell 8% per active seat."
      ],
      "speaker_notes": "Detailed breakdown of KPIs. [Source: annual_report.pdf, p. 5-7]"
    },
    {
      "slide_number": 4,
      "layout": "two_column_chart",
      "title": "Regional Revenue Growth",
      "bullets": [
        "**Americas Lead:** Remains the largest market at $73.8M (+20.6% YoY).",
        "**APAC Acceleration:** Fastest-growing region at +27.3% YoY."
      ],
      "visual_path": "assets/chart_slide_4_20261124_160000.png",
      "speaker_notes": "Regional numbers verified against audited tables. [Source: annual_report.pdf, p. 9]"
    },
    {
      "slide_number": 5,
      "layout": "two_column_image",
      "title": "Platform Architecture",
      "bullets": [
        "**Modular Core:** Decoupled ingestion and analytics services.",
        "**Zero-Downtime:** Active-active failover across three availability zones."
      ],
      "visual_path": "assets/pdf_img_f0_p12_0.png",
      "speaker_notes": "Extracted diagram from technical appendix. [Source: annual_report.pdf, p. 12]"
    },
    {
      "slide_number": 6,
      "layout": "comparison_table",
      "title": "Tier Comparison",
      "headers": ["Capability", "Standard", "Enterprise"],
      "rows": [
        ["SLA Uptime", "99.5%", "99.99%"],
        ["SSO & Audit Logs", "Optional", "Included"],
        ["Dedicated Support", "Business Hours", "24/7 Named TAM"]
      ],
      "speaker_notes": "Pricing and SLA matrix. [Source: pricing_guide.docx]"
    },
    {
      "slide_number": 7,
      "layout": "executive_summary",
      "title": "Executive Summary",
      "cards": [
        {"title": "Growth", "body": "Revenue reached $169.4M with APAC leading at +27.3%."},
        {"title": "Efficiency", "body": "Operating margin improved 210 bps to 22.4%."},
        {"title": "Outlook", "body": "FY2027 pipeline stands at $210M across 45 enterprise accounts."}
      ],
      "takeaway": "Prioritize APAC enterprise onboarding and self-serve automation in H1 FY2027.",
      "speaker_notes": "Synthesis for board review. [Source: annual_report.pdf, p. 2-3]"
    },
    {
      "slide_number": 8,
      "layout": "timeline_roadmap",
      "title": "FY2027 Execution Roadmap",
      "milestones": [
        {"period": "Q1 2027", "title": "APAC Hub Launch", "detail": "Establish Singapore delivery team"},
        {"period": "Q2 2027", "title": "v4.0 Release", "detail": "Ship real-time analytics engine"},
        {"period": "Q3 2027", "title": "SOC2 Type II", "detail": "Complete EU & APAC compliance"},
        {"period": "Q4 2027", "title": "Partner Channel", "detail": "Onboard 15 regional SI partners"}
      ],
      "speaker_notes": "Milestones aligned with board targets. [Source: strategy_memo.md]"
    },
    {
      "slide_number": 9,
      "layout": "process_flow",
      "title": "Client Onboarding Workflow",
      "steps": [
        {"step": "1", "title": "Discovery", "detail": "Audit data sources & security scope"},
        {"step": "2", "title": "Provision", "detail": "Deploy isolated tenant & SSO"},
        {"step": "3", "title": "Migrate", "detail": "Validate historical datasets"},
        {"step": "4", "title": "Go-Live", "detail": "Train admins & hand over SLA"}
      ],
      "speaker_notes": "Standard 4-week onboarding playbook. [Source: ops_manual.pdf, p. 18]"
    },
    {
      "slide_number": 10,
      "layout": "matrix_2x2",
      "title": "Initiative Prioritization Matrix",
      "x_axis": "Implementation Effort (Low -> High)",
      "y_axis": "Business Impact (Low -> High)",
      "quadrants": {
        "top_left": {"title": "Quick Wins", "items": ["Self-serve SSO", "Usage alerts"]},
        "top_right": {"title": "Strategic Bets", "items": ["APAC Hub", "v4.0 Engine"]},
        "bottom_left": {"title": "Fill-Ins", "items": ["UI theme tweaks", "CSV export"]},
        "bottom_right": {"title": "Deprioritize", "items": ["Legacy SOAP bridge"]}
      },
      "speaker_notes": "2x2 prioritization from steering committee. [Source: strategy_memo.md]"
    }
  ]
}
```

---

## 3. Optional `<deck_dir>/brand.json`

```json
{
  "company_name": "Acme Advisory",
  "logo_path": "source/logo.png",
  "primary": "#1E3A8A",
  "accent": "#0D9488",
  "background": "#F8FAFC",
  "dark": "#0F172A",
  "body": "#334155",
  "font_title": "Segoe UI",
  "font_body": "Calibri",
  "chart_palette": ["#1E3A8A", "#0D9488", "#E05A47", "#D97706", "#3B82F6", "#475569"]
}
```
