"""
frame_chat.py
Chat-with-your-results panel for the Frame Services Calculator.

- Whole result set (4 tables + premises) is placed in the system prompt
  (data is small, so no vector DB / retrieval is needed). Prompt caching is on.
- What-if questions are NOT calculated by the LLM. Claude calls the
  `run_what_if` tool, which re-runs your real calculator code with the
  changed assumptions, and the before/after diff is returned + shown in UI.

.env:
    ANTHROPIC_API_KEY=sk-ant-...
    CLAUDE_MODEL=claude-sonnet-5-5      # optional override
"""

import copy
import os
from unittest import mock

import anthropic
import pandas as pd
import streamlit as st
from dotenv import load_dotenv

import roi_excel_calculator_v2_dynamic_premises_fixed_v8 as calc

load_dotenv()

MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-5-5")
MAX_TOKENS = 2000
MAX_TOOL_ROUNDS = 4
MAX_DIFF_ROWS = 80
STATE_VERSION = "2"  # bump when cached input structure changes

# Tool "field" name -> column name used by calc.read_excel_tables()
FIELD_MAP = {
    "talent_services": "Talent Number of same services",
    "talent_fte": "Talent % FTE/Service",
    "talent_xyz": "Talent value (X, Y or Z)",
    "ai_services": "AI Number of same services",
    "ai_fte": "AI % FTE/Service",
    "ai_xyz": "Talent value AI (X, Y or Z)",
}
SCHEDULE_ROWS = {"row1": 0, "row2": 1, "row4": 2}


# =========================================================
# SCENARIO ENGINE (re-uses the real calculator functions)
# =========================================================

def load_inputs(input_path):
    """Read raw inputs once so scenarios can be re-run without the file."""
    current_input, future_input = calc.read_excel_tables(input_path)
    return {
        "premises": calc.read_premises_values(input_path),
        "current_input": current_input,
        "future_input": future_input,
        "schedule_fte": calc.read_future_service_schedule(input_path),
    }


def _apply_row_change(df, ch):
    field = ch["field"]
    if field not in FIELD_MAP:
        raise ValueError(f"Unknown field '{field}'. Use one of {list(FIELD_MAP)}.")
    col = FIELD_MAP[field]

    names = df["Service Improvement"].astype(str)
    mask = names.str.lower().str.contains(
        str(ch["service_name"]).lower(), regex=False
    )
    hits = names[mask].tolist()
    if not hits:
        raise ValueError(
            f"No service matches '{ch['service_name']}'. Available: {names.tolist()}"
        )
    if len(hits) > 1:
        raise ValueError(
            f"'{ch['service_name']}' matches several rows: {hits}. "
            "Use a more specific service_name."
        )

    if field.endswith("xyz"):
        value = str(ch["value"]).strip().upper()
        if value not in {"X", "Y", "Z"}:
            raise ValueError("X/Y/Z fields only accept X, Y or Z.")
    else:
        value = float(ch["value"])

    df[col] = df[col].astype(object)
    df.loc[mask, col] = value
    return f"{ch['service_name']} -> {field} = {value}"


def run_scenario(inputs, premises_changes=None, row_changes=None,
                 schedule_fte_changes=None):
    """Re-run the full pipeline with overrides. Always starts from the original data."""
    premises = copy.deepcopy(inputs["premises"])
    current_input = inputs["current_input"].copy()
    future_input = inputs["future_input"].copy()
    schedule_fte = copy.deepcopy(inputs["schedule_fte"])  # {"talent": [3], "ai": [3]}
    if isinstance(schedule_fte, (list, tuple)):  # legacy shape: one shared list
        schedule_fte = {"talent": list(schedule_fte), "ai": list(schedule_fte)}
    applied = []

    for ch in premises_changes or []:
        try:
            premises[ch["period"]][ch["category"]][ch["metric"]][ch["level"]] = float(
                ch["value"]
            )
        except KeyError as exc:
            raise ValueError(f"Invalid premises change {ch}: missing key {exc}")
        applied.append(
            f"Premises {ch['period']} {ch['category']} {ch['metric']} "
            f"{ch['level']} = {ch['value']}"
        )

    for ch in row_changes or []:
        df = current_input if ch["section"] == "current" else future_input
        applied.append(f"{ch['section'].title()} row: " + _apply_row_change(df, ch))

    for key, value in (schedule_fte_changes or {}).items():
        # keys: talent_row1 / ai_row4 ...; a bare "row1" changes both sides
        side, _, row = key.rpartition("_")
        if row not in SCHEDULE_ROWS or side not in ("", "talent", "ai"):
            raise ValueError(
                "schedule_fte_changes keys must look like talent_row1, "
                "ai_row2, ai_row4 (rows 1, 2, 4)."
            )
        for target in (["talent", "ai"] if side == "" else [side]):
            schedule_fte[target][SCHEDULE_ROWS[row]] = float(value)
        applied.append(f"Schedule % FTE {key} = {value}")

    # calculate_future_service_schedule() reads its 3 inputs from the file;
    # feed it the (possibly overridden) values instead.
    with mock.patch.object(calc, "read_future_service_schedule", lambda _f: schedule_fte):
        schedule_df = calc.calculate_future_service_schedule(None, premises)

    current_df = calc.process_table(current_input, "Current", premises)
    future_df = calc.process_table(future_input, "Future", premises)
    margin_df = calc.calculate_future_current_service_margin(
        current_df, future_df, schedule_df
    )
    return {
        "current": current_df,
        "future": future_df,
        "schedule": schedule_df,
        "margin": margin_df,
        "applied": applied,
    }


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return float("nan")


def _diff(base, scen, table, last_row_only=False):
    if base.empty or scen.empty or len(base) != len(scen):
        return []
    label_col = "Service Improvement"
    idx = [len(base) - 1] if last_row_only else range(len(base))
    rows = []
    for i in idx:
        label = str(base.iloc[i][label_col])
        for col in base.columns:
            if col == label_col:
                continue
            b, s = _num(base.iloc[i][col]), _num(scen.iloc[i][col])
            b_nan, s_nan = pd.isna(b), pd.isna(s)
            if b_nan and s_nan:
                continue
            if not b_nan and not s_nan and abs(b - s) < 0.005:
                continue
            both = not b_nan and not s_nan
            rows.append({
                "Table": table,
                "Row": label,
                "Metric": col,
                "Before": None if b_nan else round(b, 2),
                "After": None if s_nan else round(s, 2),
                "Change": round(s - b, 2) if both else None,
                "Change %": round((s - b) / abs(b) * 100, 1) if both and b != 0 else None,
            })
    return rows


def build_diff(base, scen):
    rows = (
        _diff(base["margin"], scen["margin"], "Future-Current Margin")
        + _diff(base["schedule"], scen["schedule"], "Cost & Schedule")
        + _diff(base["current"], scen["current"], "Current Services (Total)", True)
        + _diff(base["future"], scen["future"], "Future Services (Total)", True)
    )
    return pd.DataFrame(rows)


# =========================================================
# CLAUDE: PROMPT + TOOL
# =========================================================

TOOLS = [{
    "name": "run_what_if",
    "description": (
        "Re-run the real calculator with changed assumptions and return the "
        "before/after difference. ALWAYS use this for any 'what if / if we "
        "change / suppose / increase / decrease' question - never compute "
        "scenario numbers yourself. Changes are always applied to the ORIGINAL "
        "uploaded data (not to previous what-ifs), so include every change you "
        "want combined in a single call."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "premises_changes": {
                "type": "array",
                "description": "Change X/Y/Z rate assumptions from the Premises tab.",
                "items": {
                    "type": "object",
                    "properties": {
                        "period": {"type": "string", "enum": ["current", "future"]},
                        "category": {"type": "string", "enum": ["talent", "ai"]},
                        "metric": {"type": "string", "enum": ["cost", "value"]},
                        "level": {"type": "string", "enum": ["X", "Y", "Z"]},
                        "value": {"type": "number"},
                    },
                    "required": ["period", "category", "metric", "level", "value"],
                },
            },
            "row_changes": {
                "type": "array",
                "description": "Change an input of one service row.",
                "items": {
                    "type": "object",
                    "properties": {
                        "section": {"type": "string", "enum": ["current", "future"]},
                        "service_name": {
                            "type": "string",
                            "description": "Unique part of the Service Improvement name.",
                        },
                        "field": {"type": "string", "enum": list(FIELD_MAP)},
                        "value": {
                            "type": ["number", "string"],
                            "description": "Number, or X/Y/Z for *_xyz fields.",
                        },
                    },
                    "required": ["section", "service_name", "field", "value"],
                },
            },
            "schedule_fte_changes": {
                "type": "object",
                "description": (
                    "Change the input % FTE/Service values of the Cost & Schedule "
                    "table. Talent and AI have separate inputs (row 3 is derived "
                    "as row1+row2 per side)."
                ),
                "properties": {
                    "talent_row1": {"type": "number"},
                    "talent_row2": {"type": "number"},
                    "talent_row4": {"type": "number"},
                    "ai_row1": {"type": "number"},
                    "ai_row2": {"type": "number"},
                    "ai_row4": {"type": "number"},
                },
            },
        },
    },
}]

FORMULAS = """
## How every number is calculated (exactly as coded)

### Current / Future Services tables (per row, per side; side = Talent or AI)
- Total % FTE = Number of services x % FTE/Service
- Rates come from the Premises tab: pick the row's X/Y/Z level for that
  period (Current/Future) and side -> FTE Cost rate and FTE Value rate.
- Service Cost (K$) = Total % FTE x Cost rate / 100
- Service Value (K$): Talent = Service Cost x Value rate / 100
                      AI     = Total % FTE x Value rate / 100
- Yearly Margin = Service Value - Service Cost ; Monthly Margin = Yearly / 12
- Talent+AI columns = Talent + AI. Last row "Total" = column sums.

### Cost and Schedule to reach Future services (4 rows)
- The workbook has two separate % FTE/Service input columns: one for Talent, one
  for AI (rows 1, 2, 4). Row 3 = row1 + row2, computed separately per side.
- Talent cost = Talent % FTE x Current Talent FTE Cost (X);
  AI cost = AI % FTE x Current AI FTE Cost (X).
- Monthly Margin only on row 4 (recurring) = row 4 cost / 12 (Talent, AI, and sum).

### Future-Current Service Margin Contribution (5 rows)
Row 1 Net Gain/Loss: Talent & AI = Future total - Current total (yearly and monthly).
   Talent+AI yearly = (Future T+AI yearly) - (Current T+AI yearly).
   Talent+AI monthly = Current Talent monthly + Current AI monthly (NOT a difference).
Row 2 NRE: yearly = Total Investment cost (schedule row 3) / 100;
   monthly = Recurring cost (schedule row 4) / 100. Talent+AI = Talent + AI.
Row 3 Net Margin Year 1 = Row1 yearly - Row2 monthly - Row2 yearly/12 (Talent, AI);
   Talent+AI = Talent + AI. No monthly value.
Row 4 ROI %: Talent = Row3 / Row2 yearly x 100;
   AI = Row2 yearly / Row3 x 100 (inverted vs. the row label);
   Talent+AI = Row1 yearly / Row2 yearly x 100.
Row 5 Months to Breakeven: Talent & AI = Row2 yearly / (Row2 monthly - Row1 monthly);
   Talent+AI = Row2 yearly / (Row1 Talent+AI monthly - Row2 yearly).
A divide by zero leaves the cell blank.
"""

INSTRUCTIONS = """You are the analysis assistant inside the Frame Services Calculator.
The user has just uploaded an Excel file and sees four result tables. Help them
understand and explore those results.

Rules:
- Answer only from the data and formulas below. If something is not in them, say so.
- Explaining a value: name the formula, plug in the actual numbers from the tables,
  and show the arithmetic step by step. Name the Premises rate (X/Y/Z) that was used.
- Any hypothetical ("what if", "if we change", "suppose", "increase by 10%") MUST go
  through the run_what_if tool. Never calculate scenario results by hand.
  Convert relative changes (e.g. +10%) into absolute values using the baseline numbers.
  If a request is ambiguous (which row / which rate?), ask one short question first.
- After a what-if: state what you changed, give the headline before -> after numbers
  (with change and %), then briefly explain WHY they moved. The app already shows the
  full before/after table, so do not repeat every cell.
- Formulas are described exactly as coded. If a formula looks inconsistent with its
  row label (e.g. AI ROI, Months to Breakeven), say what the code does and flag the
  inconsistency once, briefly, without lecturing.
- Units are K$ unless stated. Keep answers concise and plain-language; use short
  paragraphs or a small table. Your reply is limited to about 1500 words, aim for far less.
"""


def _csv(df):
    return df.round(2).to_csv(index=False) if not df.empty else "(empty)"


def build_system(inputs, base):
    p = inputs["premises"]
    data = (
        "\n# DATA (baseline results)\n\n"
        f"## Premises (rates per X/Y/Z level)\n{p}\n\n"
        f"## Schedule input % FTE values (rows 1, 2, 4) per side\n{inputs['schedule_fte']}\n\n"
        f"## Current Services\n{_csv(base['current'])}\n"
        f"## Future Services\n{_csv(base['future'])}\n"
        f"## Cost and Schedule to reach Future services\n{_csv(base['schedule'])}\n"
        f"## Future-Current Service Margin Contribution\n{_csv(base['margin'])}\n"
    )
    return [{
        "type": "text",
        "text": INSTRUCTIONS + FORMULAS + data,
        "cache_control": {"type": "ephemeral"},  # data is re-sent every turn -> cache it
    }]


def _run_tool(inputs, base, args):
    scen = run_scenario(
        inputs,
        args.get("premises_changes"),
        args.get("row_changes"),
        args.get("schedule_fte_changes"),
    )
    diff = build_diff(base, scen)
    text = "Applied changes:\n- " + "\n- ".join(scen["applied"] or ["(none)"]) + "\n\n"
    if diff.empty:
        text += "No numbers changed."
    else:
        text += "Changed cells (unchanged cells omitted):\n"
        text += diff.head(MAX_DIFF_ROWS).to_csv(index=False)
        if len(diff) > MAX_DIFF_ROWS:
            text += f"... {len(diff) - MAX_DIFF_ROWS} more rows truncated"
    return text, diff


def ask_claude(system, history, user_text, inputs, base):
    """Returns (answer_text, comparison_dfs, new_history)."""
    client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY
    msgs = list(history) + [{"role": "user", "content": user_text}]
    comparisons = []
    resp = None

    for _ in range(MAX_TOOL_ROUNDS):
        resp = client.messages.create(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            system=system,
            tools=TOOLS,
            messages=msgs,
        )
        msgs.append({"role": "assistant", "content": resp.content})
        if resp.stop_reason != "tool_use":
            break

        results = []
        for block in resp.content:
            if block.type != "tool_use":
                continue
            try:
                out, diff = _run_tool(inputs, base, block.input)
                if not diff.empty:
                    comparisons.append(diff)
                results.append(
                    {"type": "tool_result", "tool_use_id": block.id, "content": out}
                )
            except Exception as exc:  # let Claude see the error and retry
                results.append({
                    "type": "tool_result",
                    "tool_use_id": block.id,
                    "content": f"Error: {exc}",
                    "is_error": True,
                })
        msgs.append({"role": "user", "content": results})

    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    if resp.stop_reason == "max_tokens":
        text += "\n\n_(Answer was cut at the token limit - ask me to continue or be more specific.)_"
    if not text:
        text = "I couldn't finish that scenario. Try rephrasing or splitting it into smaller changes."
    return text, comparisons, msgs


# =========================================================
# STREAMLIT UI
# =========================================================

def _show_comparisons(comparisons):
    for diff in comparisons:
        with st.expander("Before vs after (changed cells only)", expanded=True):
            st.dataframe(diff, use_container_width=True, hide_index=True)


def _reset_chat():
    st.session_state.chat_display = []
    st.session_state.chat_api = []


def render_chat_panel(input_path, file_key, current_df, future_df,
                      schedule_df, margin_df):
    st.divider()
    st.subheader("💬 Chat with your results")
    st.caption(
        "Ask how any value was calculated, or try a what-if, e.g. "
        "“What if Future AI cost Y changes to 80?” or "
        "“What if the first Future service has 3 more AI services?”"
    )

    if not os.getenv("ANTHROPIC_API_KEY"):
        st.warning("ANTHROPIC_API_KEY not found. Add it to your .env file.")
        return

    ss = st.session_state
    file_key = f"{file_key}|v{STATE_VERSION}"  # stale session data from older code is dropped
    if ss.get("chat_file_key") != file_key:  # new upload -> fresh chat + context
        ss.chat_file_key = file_key
        ss.chat_inputs = load_inputs(input_path)
        ss.chat_base = run_scenario(ss.chat_inputs)
        ss.chat_system = build_system(ss.chat_inputs, ss.chat_base)
        _reset_chat()

    st.button("Clear chat", on_click=_reset_chat)

    for m in ss.chat_display:
        with st.chat_message(m["role"]):
            st.markdown(m["content"])
            _show_comparisons(m.get("comparisons", []))

    prompt = st.chat_input("Ask about the calculations or try a what-if…")
    if not prompt:
        return

    with st.chat_message("user"):
        st.markdown(prompt)
    ss.chat_display.append({"role": "user", "content": prompt})

    with st.chat_message("assistant"):
        with st.spinner("Thinking…"):
            try:
                text, comps, new_hist = ask_claude(
                    ss.chat_system, ss.chat_api, prompt, ss.chat_inputs, ss.chat_base
                )
                ss.chat_api = new_hist
            except Exception as exc:
                text, comps = f"⚠️ Could not get a response: {exc}", []
        st.markdown(text)
        _show_comparisons(comps)
    ss.chat_display.append(
        {"role": "assistant", "content": text, "comparisons": comps}
    )