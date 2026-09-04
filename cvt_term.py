"""Shared dark terminal UI: black bars + light purple / pink text."""

from __future__ import annotations

import os
import sys
import threading
import time


def ui_width() -> int:
    try:
        cols = os.get_terminal_size().columns
    except OSError:
        cols = 100
    return max(80, min(cols - 1, 140))


class Term:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    BG = "\033[40m"  # black bar behind every colored line
    # Light pink / purple on black (readable on dark bars).
    GREEN = "\033[38;5;183m"  # soft lavender (body / tips)
    NEON = "\033[95m"  # bright magenta / pink (headers / success)
    RED = "\033[38;5;218m"  # light pink (wait / reopen / alerts)


def _fit(text: str, width: int | None = None) -> str:
    width = ui_width() if width is None else width
    if len(text) <= width:
        return text.ljust(width)
    if width <= 1:
        return text[:width]
    return text[: width - 1] + "…"


def c(color: str, text: str, *, fill: bool = True) -> str:
    body = _fit(text) if fill else text
    return f"{Term.BG}{color}{body}{Term.RESET}"


def blank() -> None:
    print(c(Term.GREEN, ""))


def rule(char: str = "─") -> None:
    print(c(Term.GREEN, char * ui_width()))


def banner(step: int | None, title: str, *, total: int = 3) -> None:
    width = ui_width()
    blank()
    print(c(Term.BOLD + Term.NEON, f"┌{'─' * (width - 2)}┐"))
    if step is None:
        label = f"  {title}"
    else:
        label = f"  {step}/{total}  {title}"
    print(c(Term.BOLD + Term.NEON, f"│{_fit(label, width - 2)}│"))
    print(c(Term.BOLD + Term.NEON, f"└{'─' * (width - 2)}┘"))
    blank()


def kv(label: str, value: str, *, color: str = Term.GREEN) -> None:
    """Print label + value without cutting the filename; wrap value if needed."""
    width = ui_width()
    left = f"  {label:<10}"
    first = f"{left}{value}"
    if len(first) <= width:
        print(c(color, first))
        return
    print(c(color, left.rstrip()))
    chunk = width - 4
    text = value
    while text:
        print(c(color, f"    {text[:chunk]}"))
        text = text[chunk:]


def ok(msg: str) -> None:
    kv("✓", msg, color=Term.BOLD + Term.NEON)


def warn(msg: str) -> None:
    kv("!", msg, color=Term.BOLD + Term.RED)


def wait(msg: str) -> None:
    kv("…", msg, color=Term.RED)


def tip(msg: str) -> None:
    print(c(Term.GREEN, f"  · {msg}"))


def metric(label: str, value: object, *, alert: bool = False) -> None:
    color = Term.BOLD + Term.RED if alert else Term.BOLD + Term.NEON
    print(c(color, f"  {label:<32}{value}"))


def table(
    headers: list[str],
    rows: list[list[str]],
    *,
    widths: list[int] | None = None,
    align_right: bool = False,
) -> None:
    """Print a simple aligned table with header + rule under the black bar UI."""
    if not headers:
        return
    cols = len(headers)
    if widths is None:
        widths = [len(h) for h in headers]
        for row in rows:
            for i, cell in enumerate(row[:cols]):
                widths[i] = max(widths[i], len(str(cell)))
    # Cap each column so the row still fits the terminal.
    max_w = ui_width() - 4
    total = sum(widths) + 2 * (cols - 1)
    if total > max_w and total > 0:
        scale = max_w / total
        widths = [max(4, int(w * scale)) for w in widths]

    def fmt(cells: list[str]) -> str:
        parts: list[str] = []
        for i, cell in enumerate(cells[:cols]):
            text = str(cell)
            w = widths[i]
            if align_right and i == cols - 1 and text.replace(",", "").isdigit():
                parts.append(text.rjust(w)[:w])
            else:
                parts.append(text.ljust(w)[:w])
        return "  ".join(parts)

    print(c(Term.BOLD + Term.NEON, f"  {fmt(headers)}"))
    rule_parts = ["─" * w for w in widths]
    print(c(Term.GREEN, f"  {'  '.join(rule_parts)}"))
    for row in rows:
        alert = False
        if cols >= 2:
            try:
                alert = int(str(row[-1]).replace(",", "")) > 0
            except ValueError:
                alert = False
        color = Term.BOLD + Term.RED if alert else Term.BOLD + Term.NEON
        print(c(color, f"  {fmt([str(c) for c in row])}"))


_active_progress: Progress | None = None


def stop_interrupted() -> None:
    """Clear in-place progress line and exit cleanly after Ctrl+C / EOF."""
    global _active_progress
    try:
        if _active_progress is not None:
            _active_progress.stop_work()
            _active_progress = None
        sys.stdout.write("\r" + " " * ui_width() + "\r" + Term.RESET + "\n")
        sys.stdout.flush()
        blank()
        warn("Stopped.")
        blank()
    except Exception:  # noqa: BLE001 — best-effort cleanup on interrupt
        pass
    raise SystemExit(130)


def ask(question: str) -> str:
    """Single-line prompt; no full-width fill (avoids empty black/white strips)."""
    # fill=False + RESET before input(): macOS Terminal paints a white bar if
    # BG/color stay open across input(), and full-width _fit looks like blank space.
    sys.stdout.write(
        f"{Term.BG}{Term.BOLD}{Term.NEON}  {question} › {Term.RESET}"
    )
    sys.stdout.flush()
    try:
        return input().strip()
    except (KeyboardInterrupt, EOFError):
        stop_interrupted()
        raise  # unreachable; keeps type checkers happy


class Progress:
    """In-place percent / work-in-progress bar (no braille spinner)."""

    def __init__(self, *, bar_width: int = 28) -> None:
        self.bar_width = bar_width
        self._lock = threading.Lock()
        self._work_stop: threading.Event | None = None
        self._work_thread: threading.Thread | None = None
        self._alive = False
        self._started = time.monotonic()

    def _activate(self) -> None:
        global _active_progress
        _active_progress = self

    def _paint(self, text: str) -> None:
        sys.stdout.write("\r" + c(Term.BOLD + Term.NEON, text))
        sys.stdout.flush()

    def bar(self, current: int, total: int, label: str) -> None:
        self.stop_work()
        total = max(total, 1)
        current = max(0, min(current, total))
        frac = current / total
        filled = int(self.bar_width * frac)
        block = "█" * filled + "░" * (self.bar_width - filled)
        pct = int(100 * frac)
        elapsed = int(time.monotonic() - self._started)
        self._paint(f"  [{block}] {pct:3d}%  {label}  ({elapsed}s)")
        self._alive = True
        self._activate()

    def start_work(self, label: str) -> None:
        """Animated work-in-progress bar while a long step runs (unknown %)."""
        self.stop_work()
        stop = threading.Event()
        self._work_stop = stop
        self._activate()

        def run() -> None:
            i = 0
            while not stop.wait(0.15):
                pos = i % (self.bar_width + 3)
                block = ["░"] * self.bar_width
                for j in range(3):
                    idx = pos - j
                    if 0 <= idx < self.bar_width:
                        block[idx] = "█"
                elapsed = int(time.monotonic() - self._started)
                with self._lock:
                    self._paint(
                        f"  [{''.join(block)}]  still working · {label}  ({elapsed}s)"
                    )
                i += 1

        thread = threading.Thread(target=run, daemon=True)
        self._work_thread = thread
        self._alive = True
        thread.start()

    def stop_work(self) -> None:
        if self._work_stop is not None:
            self._work_stop.set()
        if self._work_thread is not None:
            self._work_thread.join(timeout=1.0)
        self._work_stop = None
        self._work_thread = None

    def finish(self, label: str = "Done") -> None:
        global _active_progress
        self.stop_work()
        if self._alive:
            self.bar(1, 1, label)
            sys.stdout.write("\n")
            sys.stdout.flush()
            self._alive = False
        if _active_progress is self:
            _active_progress = None
