from __future__ import annotations

import json
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from fundigraph_review_workflow import (
    TRIPLE_PATTERN,
    apply_review,
    build_review_payload,
    review_html_path,
    reviewed_json_path,
)


BG = "#f5f7fb"
PANEL = "#ffffff"
PANEL_2 = "#f8fafc"
CARD = "#ffffff"
ACCENT = "#2563eb"
TEXT = "#0f172a"
MUTED = "#64748b"
BORDER = "#dbe3f0"
ERROR = "#e11d48"
LIST_SELECT = "#e8f0ff"
LIST_SELECT_TEXT = "#0f172a"
FONT_UI = "Segoe UI"
FONT_MONO = "Consolas"


class ReviewApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Triplet Confirmer")
        self.root.geometry("1280x820")
        self.root.configure(bg=BG)

        self.payload = build_review_payload()
        self.reviewed_path = reviewed_json_path()
        self.state = self._build_state()
        self.current_index = 0
        self.filter_mode = tk.StringVar(value="pending")
        self.visible_indices: list[int] = []
        self._wheel_remainder = 0.0

        self.right_shell: tk.Frame | None = None
        self.detail_canvas: tk.Canvas | None = None
        self._detail_window = None

        self._configure_style()
        self._build_ui()
        self._load_existing_review()
        self._refresh_list()
        if self.visible_indices:
            self._show_item(self.visible_indices[0])
        elif self.state:
            self._show_item(0)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background=BG)
        style.configure("Panel.TFrame", background=PANEL)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=(FONT_UI, 20, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=(FONT_UI, 10))
        style.configure("Section.TLabel", background=PANEL, foreground=TEXT, font=(FONT_UI, 10, "bold"))
        style.configure(
            "Primary.TButton",
            background=ACCENT,
            foreground="#ffffff",
            borderwidth=0,
            focuscolor=BG,
            padding=(12, 8),
        )
        style.map("Primary.TButton", background=[("active", "#1d4ed8")])
        style.configure(
            "Secondary.TButton",
            background=PANEL_2,
            foreground=TEXT,
            bordercolor=BORDER,
            lightcolor=PANEL_2,
            darkcolor=PANEL_2,
            padding=(12, 8),
        )
        style.map("Secondary.TButton", background=[("active", "#eef2f7")])
        style.configure(
            "Accent.TButton",
            background="#eff6ff",
            foreground=ACCENT,
            borderwidth=0,
            padding=(12, 8),
        )
        style.map("Accent.TButton", background=[("active", "#dbeafe")])
        style.configure(
            "Danger.TButton",
            background="#fff1f2",
            foreground=ERROR,
            borderwidth=0,
            padding=(12, 8),
        )
        style.map("Danger.TButton", background=[("active", "#ffe4e6")])
        style.configure(
            "App.TCombobox",
            fieldbackground=PANEL_2,
            background=PANEL_2,
            foreground=TEXT,
            arrowcolor=ACCENT,
            bordercolor=BORDER,
            lightcolor=PANEL_2,
            darkcolor=PANEL_2,
        )

    def _build_state(self) -> list[dict]:
        return [
            {
                "id": item["id"],
                "decision": "pending",
                "corrected_triple": item["triple"],
                "original": item,
            }
            for item in self.payload["items"]
        ]

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, style="App.TFrame", padding=18)
        top.pack(fill="x")

        ttk.Label(top, text="Triplet Confirmer", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            top,
            text="A compact review desk for confirming, correcting, and applying extracted knowledge.",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(4, 14))

        toolbar = ttk.Frame(top, style="App.TFrame")
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="Save Review", style="Secondary.TButton", command=self.save_review).pack(side="left")
        ttk.Button(toolbar, text="Save And Apply", style="Primary.TButton", command=self.save_and_apply).pack(
            side="left", padx=(10, 0)
        )
        ttk.Button(toolbar, text="Copy HTML Path", style="Accent.TButton", command=self.copy_html_path).pack(
            side="left", padx=(10, 0)
        )
        ttk.Label(toolbar, text="Filter", style="Sub.TLabel").pack(side="left", padx=(20, 8))
        filter_box = ttk.Combobox(
            toolbar,
            textvariable=self.filter_mode,
            values=["pending", "new_only", "error", "confirmed", "all"],
            state="readonly",
            width=12,
            style="App.TCombobox",
        )
        filter_box.pack(side="left")
        filter_box.bind("<<ComboboxSelected>>", lambda _e: self.on_filter_change())
        self.summary_var = tk.StringVar()
        ttk.Label(toolbar, textvariable=self.summary_var, style="Sub.TLabel").pack(side="right")

        body = ttk.Frame(self.root, style="App.TFrame", padding=(18, 0, 18, 18))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 14))
        left.configure(width=340)
        left.grid_propagate(False)

        self.right_shell = tk.Frame(body, bg=PANEL, highlightbackground=BORDER, highlightthickness=1)
        self.right_shell.grid(row=0, column=1, sticky="nsew")
        self.right_shell.grid_columnconfigure(0, weight=1)
        self.right_shell.grid_rowconfigure(0, weight=1)

        self.detail_canvas = tk.Canvas(self.right_shell, bg=PANEL, highlightthickness=0, bd=0, relief="flat")
        detail_scrollbar = ttk.Scrollbar(self.right_shell, orient="vertical", command=self.detail_canvas.yview)
        self.detail_canvas.configure(yscrollcommand=detail_scrollbar.set)
        self.detail_canvas.grid(row=0, column=0, sticky="nsew")
        detail_scrollbar.grid(row=0, column=1, sticky="ns")

        right = tk.Frame(self.detail_canvas, bg=PANEL)
        right.grid_columnconfigure(0, weight=1)
        self._detail_window = self.detail_canvas.create_window((0, 0), window=right, anchor="nw")
        right.bind("<Configure>", self._on_detail_frame_configure)
        self.detail_canvas.bind("<Configure>", self._on_detail_canvas_configure)

        ttk.Label(left, text="Queue", style="Section.TLabel").pack(anchor="w", padx=14, pady=(14, 8))
        self.listbox = tk.Listbox(
            left,
            bg=PANEL,
            fg=TEXT,
            selectbackground=LIST_SELECT,
            selectforeground=LIST_SELECT_TEXT,
            activestyle="none",
            relief="flat",
            highlightthickness=0,
            bd=0,
            font=(FONT_UI, 10),
        )
        self.listbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        header = tk.Frame(right, bg=PANEL)
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))
        header.grid_columnconfigure(0, weight=1)
        self.status_var = tk.StringVar()
        self.position_var = tk.StringVar()
        tk.Label(header, textvariable=self.status_var, bg=PANEL, fg=TEXT, font=(FONT_UI, 14, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        tk.Label(header, textvariable=self.position_var, bg=PANEL, fg=ACCENT, font=(FONT_UI, 10)).grid(
            row=0, column=1, sticky="e"
        )

        nav = tk.Frame(right, bg=PANEL)
        nav.grid(row=1, column=0, sticky="ew", padx=16)
        ttk.Button(nav, text="Previous", style="Secondary.TButton", command=self.show_previous).pack(side="left")
        ttk.Button(nav, text="Next", style="Secondary.TButton", command=self.show_next).pack(side="left", padx=(10, 0))
        tk.Label(nav, text="Confirm/Error will auto-advance", bg=PANEL, fg=MUTED, font=(FONT_UI, 9)).pack(side="right")

        self.meta = tk.Label(
            right,
            bg=CARD,
            fg=MUTED,
            justify="left",
            anchor="nw",
            padx=14,
            pady=12,
            font=(FONT_UI, 10),
            highlightbackground=BORDER,
            highlightthickness=1,
        )
        self.meta.grid(row=2, column=0, sticky="ew", padx=16, pady=(12, 10))

        self.evidence = self._text_block(right, "Evidence", row=3, height=5, readonly=True)
        self.original_triple = self._text_block(right, "Original Triple", row=4, height=5, readonly=True)

        actions = tk.Frame(right, bg=PANEL)
        actions.grid(row=5, column=0, sticky="ew", padx=16, pady=(4, 10))
        ttk.Button(actions, text="Confirm", style="Accent.TButton", command=lambda: self.set_decision("confirm")).pack(
            side="left"
        )
        ttk.Button(actions, text="Error", style="Danger.TButton", command=lambda: self.set_decision("error")).pack(
            side="left", padx=(10, 0)
        )
        ttk.Button(
            actions, text="Reset", style="Secondary.TButton", command=lambda: self.set_decision("pending", False)
        ).pack(side="left", padx=(10, 0))

        self.corrected = self._text_block(
            right,
            "Corrected Triple",
            row=6,
            height=14,
            readonly=False,
            outer_pady=(6, 18),
            inner_pad=(14, 14),
            helper_text="Edit only when the extracted relation needs manual correction.",
        )

        self._bind_global_scroll_events()

    def _text_block(
        self,
        parent: tk.Widget,
        title: str,
        row: int,
        height: int,
        readonly: bool,
        outer_pady: tuple[int, int] = (0, 10),
        inner_pad: tuple[int, int] = (12, 12),
        helper_text: str = "",
    ) -> tk.Text:
        wrap = tk.Frame(parent, bg=PANEL)
        wrap.grid(row=row, column=0, sticky="ew", padx=16, pady=outer_pady)
        tk.Label(wrap, text=title, bg=PANEL, fg=TEXT, font=(FONT_UI, 10, "bold")).pack(anchor="w", pady=(0, 6))
        if helper_text:
            tk.Label(wrap, text=helper_text, bg=PANEL, fg=MUTED, font=(FONT_UI, 9)).pack(anchor="w", pady=(0, 8))
        text = tk.Text(
            wrap,
            height=height,
            wrap="word",
            bg=CARD,
            fg=TEXT,
            insertbackground=TEXT,
            selectbackground="#dbeafe",
            relief="flat",
            bd=0,
            padx=inner_pad[0],
            pady=inner_pad[1],
            highlightbackground=BORDER,
            highlightthickness=1,
            font=(FONT_MONO, 10),
        )
        text.pack(fill="both", expand=True)
        if readonly:
            text.configure(state="disabled")
        return text

    def _load_existing_review(self) -> None:
        if not self.reviewed_path.exists():
            return
        try:
            data = json.loads(self.reviewed_path.read_text(encoding="utf-8"))
        except Exception:
            return
        by_id = {item["id"]: item for item in data.get("items", [])}
        for row in self.state:
            saved = by_id.get(row["id"])
            if not saved:
                continue
            row["decision"] = saved.get("decision", "pending")
            row["corrected_triple"] = saved.get("corrected_triple", row["corrected_triple"])

    def _row_tint(self, row: dict) -> str:
        if row["decision"] == "confirm":
            return "ok"
        if row["decision"] == "error":
            return "fix"
        if row["original"]["planned_action"] != "skip":
            return "new"
        return "hold"

    def _refresh_list(self) -> None:
        self._persist_current_text()
        self.listbox.delete(0, tk.END)
        self.visible_indices = []
        for index, row in enumerate(self.state):
            if not self._matches_filter(row):
                continue
            self.visible_indices.append(index)
            self.listbox.insert(tk.END, f"{row['id']:02d}  {self._row_tint(row).upper():<4}  {row['original']['relation']}")

        confirm = sum(1 for row in self.state if row["decision"] == "confirm")
        error = sum(1 for row in self.state if row["decision"] == "error")
        pending = sum(1 for row in self.state if row["decision"] == "pending")
        self.summary_var.set(
            f"Confirmed {confirm}   Error {error}   Pending {pending}   Visible {len(self.visible_indices)}/{len(self.state)}"
        )
        if not self.visible_indices:
            self._clear_detail("No items match the current filter.")
            return
        if self.current_index not in self.visible_indices:
            self._show_item(self.visible_indices[0])
        else:
            self._sync_list_selection()
            self._show_item(self.current_index, update_selection=False)

    def _matches_filter(self, row: dict) -> bool:
        mode = self.filter_mode.get()
        if mode == "pending":
            return row["decision"] == "pending"
        if mode == "confirmed":
            return row["decision"] == "confirm"
        if mode == "error":
            return row["decision"] == "error"
        if mode == "new_only":
            return row["original"]["planned_action"] != "skip"
        return True

    def _clear_detail(self, message: str) -> None:
        self.status_var.set(message)
        self.position_var.set("0 / 0")
        self.meta.configure(text="")
        self._set_text(self.evidence, "")
        self._set_text(self.original_triple, "")
        self.corrected.configure(state="normal")
        self.corrected.delete("1.0", tk.END)

    def _sync_list_selection(self) -> None:
        if self.current_index not in self.visible_indices:
            return
        visible_pos = self.visible_indices.index(self.current_index)
        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(visible_pos)
        self.listbox.activate(visible_pos)

    def _set_text(self, widget: tk.Text, value: str) -> None:
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        widget.insert("1.0", value)
        if widget is not self.corrected:
            widget.configure(state="disabled")

    def _show_item(self, index: int, update_selection: bool = True) -> None:
        if not self.state:
            return
        index = max(0, min(index, len(self.state) - 1))
        self.current_index = index
        row = self.state[index]
        item = row["original"]
        self.status_var.set(f"Item {row['id']}   {row['decision'].upper()}   {item['relation']}")
        self.meta.configure(
            text="\n".join(
                [
                    f"Source: {item['source_file']}",
                    f"Confidence: {item['confidence']:.2f}",
                    f"Overlap: {item['overlap']['matched']}",
                    f"Best Match: {item['overlap']['value'] or 'N/A'}",
                    f"Gemini: {item.get('gemini_precheck', {}).get('verdict', 'unavailable')}",
                    f"Planned Action: {item['planned_action']}",
                    f"Gemini Reason: {item.get('gemini_precheck', {}).get('reason', '')}",
                ]
            )
        )
        self._set_text(self.evidence, item["source_excerpt"])
        self._set_text(self.original_triple, item["triple"])
        self.corrected.configure(state="normal")
        self.corrected.delete("1.0", tk.END)
        self.corrected.insert("1.0", row["corrected_triple"])
        if self.visible_indices and index in self.visible_indices:
            self.position_var.set(f"{self.visible_indices.index(index) + 1} / {len(self.visible_indices)}")
        else:
            self.position_var.set("0 / 0")
        if update_selection:
            self._sync_list_selection()
        self.corrected.focus_set()

    def on_select(self, _event=None) -> None:
        selection = self.listbox.curselection()
        if not selection or not self.visible_indices:
            return
        self._persist_current_text()
        self._show_item(self.visible_indices[selection[0]])

    def _persist_current_text(self) -> None:
        if not self.state:
            return
        self.state[self.current_index]["corrected_triple"] = self.corrected.get("1.0", tk.END).strip()

    def set_decision(self, decision: str, advance: bool = True) -> None:
        self._persist_current_text()
        self.state[self.current_index]["decision"] = decision
        previous_index = self.current_index
        next_index = self._next_visible_index(previous_index) if advance and decision in {"confirm", "error"} else previous_index
        self._refresh_list()
        if self.visible_indices:
            target = next_index if next_index in self.visible_indices else self.visible_indices[0]
            self._show_item(target)

    def _next_visible_index(self, current: int) -> int:
        if current not in self.visible_indices:
            return self.visible_indices[0] if self.visible_indices else current
        pos = self.visible_indices.index(current)
        if pos + 1 < len(self.visible_indices):
            return self.visible_indices[pos + 1]
        return current

    def on_filter_change(self) -> None:
        self._refresh_list()

    def show_previous(self) -> None:
        if not self.visible_indices:
            return
        if self.current_index not in self.visible_indices:
            self._show_item(self.visible_indices[0])
            return
        pos = self.visible_indices.index(self.current_index)
        self._persist_current_text()
        self._show_item(self.visible_indices[max(0, pos - 1)])

    def show_next(self) -> None:
        if not self.visible_indices:
            return
        if self.current_index not in self.visible_indices:
            self._show_item(self.visible_indices[0])
            return
        pos = self.visible_indices.index(self.current_index)
        self._persist_current_text()
        self._show_item(self.visible_indices[min(len(self.visible_indices) - 1, pos + 1)])

    def save_review(self) -> None:
        self._persist_current_text()
        data = {
            "generated_on": self.payload["generated_on"],
            "source_review_json": str(review_html_path().name),
            "items": [
                {
                    "id": row["id"],
                    "decision": row["decision"],
                    "corrected_triple": row["corrected_triple"],
                    "source_file": row["original"]["source_file"],
                    "source_excerpt": row["original"]["source_excerpt"],
                    "confidence": row["original"]["confidence"],
                }
                for row in self.state
            ],
        }
        self.reviewed_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo("Saved", f"Saved review to:\n{self.reviewed_path}")

    def save_and_apply(self) -> None:
        self.save_review()
        invalid = []
        for row in self.state:
            if row["decision"] not in {"confirm", "error"}:
                continue
            if not TRIPLE_PATTERN.search(row["corrected_triple"]):
                invalid.append(row["id"])
        if invalid:
            messagebox.showerror("Invalid Triple", f"These items have invalid corrected triples: {invalid}")
            return
        try:
            report = apply_review()
        except Exception as exc:
            messagebox.showerror("Apply Failed", str(exc))
            return
        row_results = report.get("row_results", [])
        neo4j_sync = report.get("neo4j_sync", {})
        summary_lines = [f"Item {idx + 1}: {result['status']}" for idx, result in enumerate(row_results)]
        summary_lines.append(f"Neo4j sync: {neo4j_sync.get('status', 'unknown')}")
        if neo4j_sync.get("status") == "ok":
            summary_lines.append(
                f"Rows synced: {neo4j_sync.get('row_count', 0)} | Relationships merged: {neo4j_sync.get('relationship_count', 0)}"
            )
        elif neo4j_sync.get("reason"):
            summary_lines.append(str(neo4j_sync["reason"]))
        summary = "\n".join(summary_lines)
        messagebox.showinfo("Apply Finished", summary or "No changes were applied.")

    def copy_html_path(self) -> None:
        path = str(Path(review_html_path()).resolve())
        self.root.clipboard_clear()
        self.root.clipboard_append(path)
        messagebox.showinfo("Copied", f"Copied HTML path:\n{path}")

    def _bind_global_scroll_events(self) -> None:
        self.root.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.root.bind_all("<Shift-MouseWheel>", self._on_mousewheel_horizontal, add="+")
        self.root.bind_all("<Button-4>", self._on_button4_scroll, add="+")
        self.root.bind_all("<Button-5>", self._on_button5_scroll, add="+")

    def _is_inside_right_panel(self, widget: tk.Widget | None) -> bool:
        if widget is None or self.right_shell is None:
            return False
        current = widget
        while current is not None:
            if current == self.right_shell:
                return True
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                break
        return False

    def _on_detail_frame_configure(self, _event=None) -> None:
        if self.detail_canvas is not None:
            self.detail_canvas.configure(scrollregion=self.detail_canvas.bbox("all"))

    def _on_detail_canvas_configure(self, event) -> None:
        if self.detail_canvas is not None and self._detail_window is not None:
            self.detail_canvas.itemconfigure(self._detail_window, width=event.width)

    def _scroll_detail_units(self, units: int) -> str | None:
        if self.detail_canvas is None or units == 0:
            return None
        self.detail_canvas.yview_scroll(units, "units")
        return "break"

    def _on_mousewheel(self, event) -> str | None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if not self._is_inside_right_panel(widget):
            return None
        delta = getattr(event, "delta", 0)
        if not delta:
            return None
        steps = delta / 120.0
        if abs(steps) < 1:
            steps = 0.35 if steps > 0 else -0.35
        self._wheel_remainder += steps
        whole_units = int(self._wheel_remainder)
        if whole_units == 0:
            whole_units = 1 if self._wheel_remainder > 0 else -1
            self._wheel_remainder = 0.0
        else:
            self._wheel_remainder -= whole_units
        return self._scroll_detail_units(-whole_units)

    def _on_mousewheel_horizontal(self, event) -> str | None:
        return self._on_mousewheel(event)

    def _on_button4_scroll(self, event) -> str | None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if not self._is_inside_right_panel(widget):
            return None
        return self._scroll_detail_units(-1)

    def _on_button5_scroll(self, event) -> str | None:
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if not self._is_inside_right_panel(widget):
            return None
        return self._scroll_detail_units(1)


def main() -> None:
    root = tk.Tk()
    ReviewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
